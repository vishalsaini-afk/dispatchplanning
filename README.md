# Lingel · Dispatch Planning

Separate app for the dispatch planning team. It sends its data to Team Availability (To Assign → Ready & paid).

- **Planning:** https://vishalsaini-afk.github.io/dispatchplanning/ (office password)
  Waiting payment → Plan dispatch (date + installation needed / supply only) → Dispatch plan
- **Dispatch:** https://vishalsaini-afk.github.io/dispatchplanning/?role=dispatch
  What to load, by date and region (only orders whose loading the installation team confirmed)

Backend: the Lingel invoice Apps Script (`readyList`, `readyApprove`, `readyPlan`, …), "Ready Orders" tab in the Installation invoice sheet.
