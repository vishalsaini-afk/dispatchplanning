# Lingel · Dispatch Planning

Separate app for dispatch planning. It sends its data to Team Availability (To Assign → Ready & paid).

**Link (one for everyone):** https://vishalsaini-afk.github.io/dispatchplanning/

Sign in with your name + password. The password decides what you see (checked by the script):

| Login | Sees | Can do |
|---|---|---|
| Planning (`DP_KEY`) | Waiting payment · Plan dispatch · Dispatch plan | Payment received, plan date + installation, change region/date |
| Dispatch (`DP_DISPATCH_KEY`) | Dispatch plan only (planned + loading confirmed) | View |

Backend: the Lingel invoice Apps Script (`readyLogin`, `readyList`, `readyApprove`, `readyPlan`, …), "Ready Orders" tab in the Installation invoice sheet.
