# Team Calendar Feed

Turns a Winnipeg Central Hockey League team schedule page into an `.ics` file you can
subscribe to with `webcal://`. Games update in your calendar when the league changes the page.

## Files

- `Update-TeamCalendar.ps1` : downloads the page, builds the calendar
- `config.json` : settings (the schedule URL lives here)
- `.github/workflows/update-calendar.yml` : runs the script every 4 hours on GitHub

## config.json

| Setting | What it does |
|---|---|
| ScheduleUrl | Team schedule page. Change this for a new season or team. |
| TeamName | Leave blank to take it from `Team=` in the URL |
| CalendarName | Name shown in the calendar app. Blank = team name (Cobra HC) |
| OutputPath | Where the .ics is written |
| GameLengthMinutes | Event length (60) |
| ReminderMinutes | Alerts before each game, e.g. `[120, 60]`. `[]` for none |
| ArenaAddresses | Optional map of arena name to street address, so maps links work |

Example:

```json
"ArenaAddresses": {
  "Sargent Park": "999 Sargent Ave, Winnipeg, MB"
}
```

## Option A: GitHub (free, no PC needed)

1. Create a public repo on GitHub and upload all these files, keeping the `.github` folder.
2. Settings > Pages > Source: "Deploy from a branch", branch `main`, folder `/docs`.
3. Actions tab > "Update calendar" > Run workflow (first run creates `docs/team.ics`).
4. Subscribe with:
   `webcal://YOURNAME.github.io/REPONAME/team.ics`

To change the URL later, edit `config.json` on GitHub. The workflow runs right away.

GitHub pauses scheduled workflows in repos with no activity for 60 days and emails you.
Re-enable it from the Actions tab if that happens.

## Option B: Your own PC or server

Run it once to test:

```powershell
.\Update-TeamCalendar.ps1 -ShowGames
```

Schedule it every 4 hours with Task Scheduler:

```powershell
schtasks /Create /TN "Team Calendar" /SC HOURLY /MO 4 /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\TeamCalendar\Update-TeamCalendar.ps1"
```

Set `OutputPath` to a folder your web server serves, then subscribe with
`webcal://yourdomain/path/team.ics`.

## Switches

- `-ShowGames` prints what was parsed
- `-Force` rewrites the file even if nothing changed
- `-Url <url>` overrides the URL for one run
- `-HtmlFile page.html` parses a saved copy of the page (handy if the layout changes)

## Notes

- If the page can't be read or no games are found, the old .ics is left alone.
- Event IDs are based on date and teams, so a time or rink change updates the existing event.
- Finished games show the result in the title, e.g. `Cobra HC vs Hotwings (W 4-2)`.
- Google Calendar only refreshes subscribed feeds every 12 to 24 hours. Apple and Outlook are faster.
