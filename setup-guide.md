# Setup Guide — Tech Help Booking Page

## Step 1: Create a Google Apps Script

1. Go to [script.google.com](https://script.google.com)
2. Click **New project**
3. Delete any existing code in `Code.gs`
4. Copy and paste the entire contents of `google-apps-script.js` into the editor
5. Update the `CALENDAR_ID` and `TIMEZONE` variables at the top if needed:
   - Use `"primary"` to add events to your main Google Calendar
   - Or create a dedicated calendar (see Step 6) and use its ID
   - Change `TIMEZONE` to match your timezone (e.g. `"America/New_York"`, `"America/Chicago"`, etc.)
6. Click **Save** (Ctrl+S)

## Step 2: Deploy as a Web App

1. Click **Deploy** → **New deployment**
2. Click the gear icon next to "Select type" and choose **Web app**
3. Set these options:
   - **Description**: Tech Help Booking
   - **Execute as**: Me
   - **Who has access**: Anyone
4. Click **Deploy**
5. You'll be asked to authorize — click **Authorize access** and follow the prompts
6. **Copy the Web app URL** — it will look like:
   `https://script.google.com/macros/s/XXXXXXX/exec`

## Step 3: Connect the Booking Page

1. Open `index.html` in a text editor
2. Find this line near the top of the `<script>` section:
   ```js
   const GOOGLE_SCRIPT_URL = "YOUR_GOOGLE_APPS_SCRIPT_URL_HERE";
   ```
3. Replace `YOUR_GOOGLE_APPS_SCRIPT_URL_HERE` with the URL you copied in Step 2
4. Save the file

## Step 4: Test It

1. Open `index.html` in your browser (just double-click the file)
2. Select a duration, pick a date and time, enter test info, and click **Book Session**
3. Check your Google Calendar — you should see a new event titled "Tech Help: [Name]"
4. The email you entered should receive a Google Calendar invite

## Step 5: Host the Page (Optional)

To make it publicly accessible, you can host it on GitHub Pages:

1. Push this repo to GitHub
2. Go to **Settings** → **Pages**
3. Set source to your branch and root folder
4. Your booking page will be live at `https://yourusername.github.io/Booking/`

## Step 6: Create a Dedicated Calendar (Recommended)

Instead of using your primary calendar, create a separate "Tech Help" calendar:

1. In Google Calendar, click **+** next to "Other calendars" → **Create new calendar**
2. Name it **Tech Help Bookings**
3. Click **Create calendar**
4. Go to the new calendar's **Settings** (three dots → Settings)
5. Scroll to **Integrate calendar** and copy the **Calendar ID**
   - It looks like: `abc123xyz@group.calendar.google.com`
6. Paste this ID into `google-apps-script.js` as the `CALENDAR_ID` value
7. Re-deploy the Apps Script (Deploy → Manage deployments → Edit → Update)

## Step 7: Subscribe in Outlook

To see bookings in Outlook, subscribe to the Google Calendar's iCal feed:

### Get the iCal URL from Google Calendar:
1. Open [Google Calendar Settings](https://calendar.google.com/calendar/r/settings)
2. Click on the calendar you're using (Primary or Tech Help Bookings)
3. Scroll to **Integrate calendar**
4. Copy the **Secret address in iCal format** URL
   - It looks like: `https://calendar.google.com/calendar/ical/.../basic.ics`

### Add to Outlook (Desktop):
1. Open Outlook
2. Go to **File** → **Account Settings** → **Account Settings**
3. Click the **Internet Calendars** tab
4. Click **New** and paste the iCal URL
5. Click **Add** → **OK**

### Add to Outlook (Web / outlook.com):
1. Go to [Outlook Calendar](https://outlook.live.com/calendar)
2. Click **Add calendar** (left sidebar)
3. Select **Subscribe from web**
4. Paste the iCal URL
5. Name it "Tech Help Bookings" and click **Import**

Bookings will now appear in both your Google Calendar and Outlook.

---

## Troubleshooting

**"Calendar not found" error:**
- Make sure the `CALENDAR_ID` in the Apps Script matches your actual calendar
- If using a dedicated calendar, double-check you copied the correct Calendar ID

**Events not appearing:**
- Check that you authorized the Apps Script correctly (it needs Calendar permissions)
- Make sure you re-deployed after any code changes

**No calendar invite email:**
- Google may not send invites to your own Google account
- Test with a different email address to verify invites are working

**Outlook not syncing:**
- iCal subscriptions in Outlook refresh every few hours by default
- New bookings may take up to 24 hours to appear in Outlook
- You can manually refresh by removing and re-adding the subscription
