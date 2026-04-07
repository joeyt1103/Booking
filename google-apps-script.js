// ═══════════════════════════════════════════════════════════════
// Google Apps Script — Tech Help Booking Backend
// Deploy this as a Web App in Google Apps Script (script.google.com)
// See setup-guide.md for detailed instructions.
// ═══════════════════════════════════════════════════════════════

// ── Configuration ──────────────────────────────────────────────
// Use 'primary' for your main calendar, or paste a specific calendar ID
var CALENDAR_ID = "primary";
var TIMEZONE = "America/New_York";

// Outlook calendar ICS feed URL — used to check Joey's existing availability
var OUTLOOK_ICS_URL = "https://outlook.office365.com/owa/calendar/bd1ef5c549bb44d0aa6f665b5c609a70%40actslife.org/b16836e090744536845b98df6f447d849103338040331570038/calendar.ics";

// Booking window: 1:00 PM to 4:00 PM, Monday–Friday
var BOOKING_START_HOUR = 13; // 1 PM
var BOOKING_END_HOUR   = 16; // 4 PM (last slot ends at 4 PM)
var SLOT_INCREMENT     = 30; // minutes

// ── Handle POST requests (create a booking) ────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var name     = data.name;
    var email    = data.email;
    var date     = data.date;      // "2026-04-15"
    var time     = data.time;      // "14:30" (24-hour)
    var duration = data.duration;  // 30 or 60
    var notes    = data.notes || "";

    if (!name || !email || !date || !time || !duration) {
      return jsonResponse({ success: false, error: "Missing required fields" });
    }

    var startTime = parseDateTime(date, time);
    var endTime   = new Date(startTime.getTime() + duration * 60000);

    // Get the Google Calendar
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      return jsonResponse({ success: false, error: "Calendar not found" });
    }

    // Check for conflicts on Google Calendar
    var existing = calendar.getEvents(startTime, endTime);
    if (existing.length > 0) {
      return jsonResponse({ success: false, error: "Time slot is no longer available" });
    }

    // Create the event
    var event = calendar.createEvent(
      "Tech Help: " + name,
      startTime,
      endTime,
      {
        description: "Booked by: " + name +
                     "\nEmail: " + email +
                     "\nDuration: " + duration + " minutes" +
                     "\n\nNotes: " + (notes || "None"),
        guests: email,
        sendInvites: true
      }
    );

    return jsonResponse({ success: true, eventId: event.getId() });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ── Handle GET requests (availability check) ───────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";

  if (action === "availability") {
    return getAvailability(e.parameter.date, parseInt(e.parameter.duration || "30", 10));
  }

  return jsonResponse({ status: "Booking API is running" });
}

// ── Get available slots for a given date ───────────────────────
function getAvailability(dateStr, duration) {
  if (!dateStr) {
    return jsonResponse({ available: [] });
  }

  // Parse the requested date
  var parts = dateStr.split("-");
  var year  = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  var day   = parseInt(parts[2], 10);
  var requestedDate = new Date(year, month, day);

  // Block weekends
  var dayOfWeek = requestedDate.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return jsonResponse({ available: [] });
  }

  // Fetch busy times from Outlook ICS feed
  var busyPeriods = fetchOutlookBusyTimes(dateStr);

  // Also check Google Calendar for existing bookings
  var dayStart = new Date(year, month, day, BOOKING_START_HOUR, 0, 0);
  var dayEnd   = new Date(year, month, day, BOOKING_END_HOUR, 0, 0);

  try {
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (calendar) {
      var gcalEvents = calendar.getEvents(dayStart, dayEnd);
      for (var g = 0; g < gcalEvents.length; g++) {
        busyPeriods.push({
          start: gcalEvents[g].getStartTime().getTime(),
          end:   gcalEvents[g].getEndTime().getTime()
        });
      }
    }
  } catch (err) {
    // If calendar check fails, continue with just Outlook data
  }

  // Generate all possible slots and filter out busy ones
  var available = [];
  var lastStartMinutes = (BOOKING_END_HOUR * 60) - duration;

  for (var mins = BOOKING_START_HOUR * 60; mins <= lastStartMinutes; mins += SLOT_INCREMENT) {
    var hour = Math.floor(mins / 60);
    var min  = mins % 60;

    var slotStart = new Date(year, month, day, hour, min, 0).getTime();
    var slotEnd   = slotStart + (duration * 60000);

    // Check if this slot overlaps with any busy period
    var isBusy = false;
    for (var b = 0; b < busyPeriods.length; b++) {
      if (slotStart < busyPeriods[b].end && slotEnd > busyPeriods[b].start) {
        isBusy = true;
        break;
      }
    }

    if (!isBusy) {
      available.push(pad(hour) + ":" + pad(min));
    }
  }

  return jsonResponse({ available: available });
}

// ── Fetch and parse Outlook ICS feed ───────────────────────────
function fetchOutlookBusyTimes(dateStr) {
  var busyPeriods = [];

  try {
    var response = UrlFetchApp.fetch(OUTLOOK_ICS_URL, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return busyPeriods;
    }

    var icsText = response.getContentText();
    var events = parseICS(icsText);

    // Filter events that overlap with the requested date
    var parts = dateStr.split("-");
    var year  = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1;
    var day   = parseInt(parts[2], 10);
    var dayStart = new Date(year, month, day, 0, 0, 0).getTime();
    var dayEnd   = new Date(year, month, day, 23, 59, 59).getTime();

    for (var i = 0; i < events.length; i++) {
      var evt = events[i];
      // Include if the event overlaps with this day at all
      if (evt.start < dayEnd && evt.end > dayStart) {
        busyPeriods.push({ start: evt.start, end: evt.end });
      }
    }
  } catch (err) {
    // If ICS fetch fails, return empty (all slots shown)
  }

  return busyPeriods;
}

// ── Parse ICS content into event objects ───────────────────────
// Handles single events and expands recurring events (RRULE)
function parseICS(icsText) {
  var events = [];
  // Unfold continuation lines (lines starting with space/tab are continuations)
  var lines = icsText.replace(/\r\n[ \t]/g, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");

  var inEvent = false;
  var dtStart = null;
  var dtEnd = null;
  var rrule = null;
  var exdates = [];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    if (line === "BEGIN:VEVENT") {
      inEvent = true;
      dtStart = null;
      dtEnd = null;
      rrule = null;
      exdates = [];
      continue;
    }

    if (line === "END:VEVENT") {
      if (inEvent && dtStart && dtEnd) {
        if (rrule) {
          // Expand recurring event into individual occurrences
          var expanded = expandRRule(dtStart, dtEnd, rrule, exdates);
          for (var e = 0; e < expanded.length; e++) {
            events.push(expanded[e]);
          }
        } else {
          events.push({ start: dtStart, end: dtEnd });
        }
      }
      inEvent = false;
      continue;
    }

    if (!inEvent) continue;

    if (line.indexOf("DTSTART") === 0) {
      dtStart = parseICSDate(line);
    } else if (line.indexOf("DTEND") === 0) {
      dtEnd = parseICSDate(line);
    } else if (line.indexOf("RRULE:") === 0) {
      rrule = line.substring(6);
    } else if (line.indexOf("EXDATE") === 0) {
      var exVal = line.substring(line.lastIndexOf(":") + 1).trim();
      // EXDATE can have multiple comma-separated dates
      var exParts = exVal.split(",");
      for (var x = 0; x < exParts.length; x++) {
        var exTs = parseICSDateValue(exParts[x].trim());
        if (exTs) exdates.push(exTs);
      }
    }
  }

  return events;
}

// ── Expand a recurring event using RRULE ───────────────────────
// Supports FREQ=DAILY, WEEKLY, MONTHLY, YEARLY with COUNT, UNTIL, INTERVAL, BYDAY
function expandRRule(dtStart, dtEnd, rruleStr, exdates) {
  var results = [];
  var duration = dtEnd - dtStart;

  // Parse RRULE parameters
  var params = {};
  var parts = rruleStr.split(";");
  for (var p = 0; p < parts.length; p++) {
    var kv = parts[p].split("=");
    if (kv.length === 2) params[kv[0]] = kv[1];
  }

  var freq     = params["FREQ"] || "";
  var interval = parseInt(params["INTERVAL"] || "1", 10);
  var count    = params["COUNT"] ? parseInt(params["COUNT"], 10) : null;
  var until    = params["UNTIL"] ? parseICSDateValue(params["UNTIL"]) : null;
  var byday    = params["BYDAY"] ? params["BYDAY"].split(",") : null;

  // Map day abbreviations to JS getDay() values
  var dayMap = { "SU": 0, "MO": 1, "TU": 2, "WE": 3, "TH": 4, "FR": 5, "SA": 6 };

  // Limit expansion to 1 year out to avoid infinite loops
  var maxDate = new Date();
  maxDate.setFullYear(maxDate.getFullYear() + 1);
  var maxTs = maxDate.getTime();
  if (until && until < maxTs) maxTs = until;

  var startDate = new Date(dtStart);
  var occurrences = 0;
  var maxOccurrences = count || 365; // safety limit

  if (freq === "WEEKLY") {
    // For WEEKLY with BYDAY, iterate week by week and check each day
    var targetDays = [];
    if (byday) {
      for (var bd = 0; bd < byday.length; bd++) {
        var dayCode = byday[bd].replace(/[^A-Z]/g, "");
        if (dayMap[dayCode] !== undefined) targetDays.push(dayMap[dayCode]);
      }
    } else {
      targetDays.push(startDate.getDay());
    }

    // Start from the week of the original event
    var cursor = new Date(startDate);
    // Go to the start of the week (Sunday)
    cursor.setDate(cursor.getDate() - cursor.getDay());
    cursor.setHours(startDate.getHours(), startDate.getMinutes(), startDate.getSeconds(), 0);

    while (cursor.getTime() <= maxTs && occurrences < maxOccurrences) {
      for (var td = 0; td < targetDays.length; td++) {
        var eventDate = new Date(cursor);
        eventDate.setDate(cursor.getDate() + targetDays[td]);
        eventDate.setHours(startDate.getHours(), startDate.getMinutes(), startDate.getSeconds(), 0);
        var ts = eventDate.getTime();

        if (ts < dtStart) continue;
        if (ts > maxTs) break;
        if (isExcluded(ts, exdates)) continue;

        results.push({ start: ts, end: ts + duration });
        occurrences++;
        if (count && occurrences >= count) break;
      }
      if (count && occurrences >= count) break;
      // Advance by interval weeks
      cursor.setDate(cursor.getDate() + (7 * interval));
    }

  } else if (freq === "DAILY") {
    var cursor = new Date(startDate);
    while (cursor.getTime() <= maxTs && occurrences < maxOccurrences) {
      var ts = cursor.getTime();
      if (!isExcluded(ts, exdates)) {
        if (byday) {
          var dayCode2 = ["SU","MO","TU","WE","TH","FR","SA"][cursor.getDay()];
          if (byday.indexOf(dayCode2) !== -1) {
            results.push({ start: ts, end: ts + duration });
            occurrences++;
          }
        } else {
          results.push({ start: ts, end: ts + duration });
          occurrences++;
        }
      }
      cursor.setDate(cursor.getDate() + interval);
    }

  } else if (freq === "MONTHLY") {
    var cursor = new Date(startDate);
    while (cursor.getTime() <= maxTs && occurrences < maxOccurrences) {
      var ts = cursor.getTime();
      if (!isExcluded(ts, exdates)) {
        results.push({ start: ts, end: ts + duration });
        occurrences++;
      }
      cursor.setMonth(cursor.getMonth() + interval);
    }

  } else if (freq === "YEARLY") {
    var cursor = new Date(startDate);
    while (cursor.getTime() <= maxTs && occurrences < maxOccurrences) {
      var ts = cursor.getTime();
      if (!isExcluded(ts, exdates)) {
        results.push({ start: ts, end: ts + duration });
        occurrences++;
      }
      cursor.setFullYear(cursor.getFullYear() + interval);
    }

  } else {
    // Unknown freq — just return the single instance
    results.push({ start: dtStart, end: dtEnd });
  }

  return results;
}

// ── Check if a timestamp is in the exclusion list ──────────────
function isExcluded(ts, exdates) {
  // Compare dates only (ignore time component for EXDATE matching)
  var d = new Date(ts);
  var dateOnly = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  for (var i = 0; i < exdates.length; i++) {
    var exd = new Date(exdates[i]);
    var exDateOnly = new Date(exd.getFullYear(), exd.getMonth(), exd.getDate()).getTime();
    if (dateOnly === exDateOnly) return true;
  }
  return false;
}

// ── Parse an ICS date line (e.g. "DTSTART;TZID=...:20260415T130000") ──
function parseICSDate(line) {
  var colonIdx = line.lastIndexOf(":");
  if (colonIdx === -1) return null;
  var value = line.substring(colonIdx + 1).trim();
  return parseICSDateValue(value);
}

// ── Parse a raw ICS date value into a timestamp ────────────────
// Handles: 20260415T130000Z, 20260415T130000, 20260415
function parseICSDateValue(value) {
  if (!value) return null;

  // All-day event: YYYYMMDD
  if (value.length === 8) {
    var y = parseInt(value.substring(0, 4), 10);
    var m = parseInt(value.substring(4, 6), 10) - 1;
    var d = parseInt(value.substring(6, 8), 10);
    return new Date(y, m, d, 0, 0, 0).getTime();
  }

  // DateTime: YYYYMMDDTHHMMSS or YYYYMMDDTHHMMSSZ
  var year   = parseInt(value.substring(0, 4), 10);
  var month  = parseInt(value.substring(4, 6), 10) - 1;
  var day    = parseInt(value.substring(6, 8), 10);
  var hour   = parseInt(value.substring(9, 11), 10);
  var minute = parseInt(value.substring(11, 13), 10);
  var second = parseInt(value.substring(13, 15), 10) || 0;

  // If UTC (ends with Z), convert to local via Date.UTC
  if (value.charAt(value.length - 1) === "Z") {
    return Date.UTC(year, month, day, hour, minute, second);
  }

  // Otherwise treat as local time
  return new Date(year, month, day, hour, minute, second).getTime();
}

// ── Helpers ────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseDateTime(dateStr, timeStr) {
  var dateParts = dateStr.split("-");
  var timeParts = timeStr.split(":");
  var year   = parseInt(dateParts[0], 10);
  var month  = parseInt(dateParts[1], 10) - 1;
  var day    = parseInt(dateParts[2], 10);
  var hour   = parseInt(timeParts[0], 10);
  var minute = parseInt(timeParts[1], 10);
  return new Date(year, month, day, hour, minute, 0);
}

function pad(n) {
  return (n < 10 ? "0" : "") + n;
}
