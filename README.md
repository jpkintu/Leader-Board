# Staff Referral Leaderboard Tracker

A Google Apps Script solution that generates a weekly leaderboard for staff referrals with beautiful email notifications and spreadsheet tracking.

## Features

- � **Weekly Leaderboard** - Automatically calculates weekly top performers
- ✉️ **Email Notifications** - Sends beautifully formatted HTML emails
- 📊 **Spreadsheet Integration** - Maintains historical records in Google Sheets
- 🏆 **Top 3 Recognition** - Highlights gold, silver, and bronze performers
- ⏰ **Automated Scheduling** - Runs every Monday at 8 AM automatically

## Prerequisites

- Google Workspace account
- Google Sheets with form responses data
- Form responses should include:
  - Timestamp (Column A)
  - Staff Name (Column B)
  - Status (Column J, with "Delivered" status)

## Installation

1. **Create a new Google Sheet** with:
   - A sheet named "Form Responses 1" for your form data
   - (Optional) A sheet named "Leaderboard" for historical records

2. **Open Script Editor**:
   - Go to `Extensions > Apps Script`
   - Paste the entire code from `code.gs`

3. **Set up the trigger**:
   - Run `setupWeeklyTrigger()` once to create the Monday 8 AM schedule
   - Alternatively, run `createWeeklyLeaderboard()` manually for testing

## Configuration

Customize these variables in the code:

```javascript
const THEME_COLOR = "#00a3e0"; // Main brand color
const SECONDARY_COLOR = "#0078a5"; // Darker blue for gradients
const RECIPIENT_EMAIL = Session.getActiveUser().getEmail(); // Change to distribution list if needed
