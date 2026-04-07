// ═══════════════════════════════════════════════════════════════
// Google Apps Script — Tech Help Booking Backend
// Deploy this as a Web App in Google Apps Script (script.google.com)
// See setup-guide.md for detailed instructions.
// ═══════════════════════════════════════════════════════════════

// ── Configuration ──────────────────────────────────────────────
// Use 'primary' for your main calendar, or paste a specific calendar ID
// (e.g. "abc123@group.calendar.google.com") if you created a dedicated one.
var CALENDAR_ID = "primary";
var TIMEZONE = "America/New_York"; // Change to your timezone

// ── Handle POST requests from the booking page ─────────────────
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

    // Parse start and end times
    var startTime = parseDateTime(date, time);
    var endTime   = new Date(startTime.getTime() + duration * 60000);

    // Get the calendar
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      return jsonResponse({ success: false, error: "Calendar not found" });
    }

    // Check for conflicts
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

// ── Handle GET requests (health check) ─────────────────────────
function doGet(e) {
  return jsonResponse({ status: "Booking API is running" });
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
  var month  = parseInt(dateParts[1], 10) - 1; // JS months are 0-indexed
  var day    = parseInt(dateParts[2], 10);
  var hour   = parseInt(timeParts[0], 10);
  var minute = parseInt(timeParts[1], 10);
  return new Date(year, month, day, hour, minute, 0);
}
