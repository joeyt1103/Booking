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
function parseICS(icsText) {
  var events = [];
  var lines = icsText.replace(/\r\n /g, "").replace(/\r/g, "\n").split("\n");

  var inEvent = false;
  var dtStart = null;
  var dtEnd = null;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    if (line === "BEGIN:VEVENT") {
      inEvent = true;
      dtStart = null;
      dtEnd = null;
      continue;
    }

    if (line === "END:VEVENT") {
      if (inEvent && dtStart && dtEnd) {
        events.push({ start: dtStart, end: dtEnd });
      }
      inEvent = false;
      continue;
    }

    if (!inEvent) continue;

    if (line.indexOf("DTSTART") === 0) {
      dtStart = parseICSDate(line);
    } else if (line.indexOf("DTEND") === 0) {
      dtEnd = parseICSDate(line);
    }
  }

  return events;
}

// ── Parse an ICS date line into a timestamp ────────────────────
// Handles formats like:
//   DTSTART:20260415T130000Z
//   DTSTART;TZID=America/New_York:20260415T130000
//   DTSTART;VALUE=DATE:20260415
function parseICSDate(line) {
  // Extract the date value (everything after the last colon)
  var colonIdx = line.lastIndexOf(":");
  if (colonIdx === -1) return null;
  var value = line.substring(colonIdx + 1).trim();

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
