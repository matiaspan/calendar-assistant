# Calendar Assistant

A Google Apps Script that automatically manages your calendar by:

1. **Drive time blocks** — Adds "Drive to" and "Drive from" events around meetings that have a physical location, calculated from your office address.
2. **Buffer blocks** — Inserts a 15-minute break after any 2+ hour stretch of back-to-back meetings.

Runs every 5 minutes on Google's infrastructure. No server, no cost, no maintenance.

## Setup

### 1. Create the Apps Script project

Go to [script.google.com](https://script.google.com) and click **New project**.

### 2. Add the code

Delete the default `Code.gs` content. Create one script file for each `.gs` file in this repo and paste the contents:

- `Code.gs`
- `Utils.gs`
- `DriveTime.gs`
- `Buffers.gs`

Also replace `appsscript.json` (click the gear icon > **Show "appsscript.json" manifest file** in the editor settings).

### 3. Set your office address

Go to **Project Settings** (gear icon) > **Script Properties** > **Add script property**:

| Property | Value |
|---|---|
| `OFFICE_ADDRESS` | Your office address (e.g. `123 Main St, City, State`) |

### 4. Run setup

Select `setup` from the function dropdown in the editor toolbar and click **Run**. This will:
- Write default configuration values
- Create the recurring 5-minute trigger
- Prompt you to authorize calendar and maps access

### 5. Verify

Select `main` and click **Run**. Check the execution log (View > Execution log) for any errors. Your calendar should now have drive time and buffer blocks.

## Configuration

All settings are in **Project Settings > Script Properties**. Defaults are applied by `setup()`.

| Property | Default | Description |
|---|---|---|
| `OFFICE_ADDRESS` | *(required)* | Starting point for drive time calculations |
| `LOOK_AHEAD_DAYS` | `7` | How many days ahead to scan |
| `DRIVE_TIME_BUFFER_PERCENT` | `25` | Extra % added on top of the traffic-aware drive estimate |
| `BUFFER_DURATION_MINUTES` | `15` | Length of break blocks in minutes |
| `MEETING_BLOCK_THRESHOLD_MINUTES` | `120` | Minimum consecutive meeting duration to trigger a buffer |
| `CONSECUTIVE_GAP_MINUTES` | `15` | Max gap between events to count as back-to-back |

## How it works

- `main()` runs every 5 minutes via a time-driven trigger
- It scans the next 7 days of your default calendar
- **Source events**: non-all-day events you own or accepted
- **Managed events**: events created by this script (tagged in the description)
- For each source event with a location, it calculates drive time using Google Maps and creates purple "Drive to/from" blocks
- For consecutive meeting stretches >= 2 hours, it creates green "Break" blocks
- Orphaned blocks (source event deleted/changed) are automatically cleaned up

## Uninstall

Run `uninstall()` from the editor. This removes the trigger and deletes all managed events from the next 90 days.

## Alternative: deploy with clasp

```bash
npm install -g @google/clasp
clasp login
clasp create --title "Calendar Assistant" --type standalone
clasp push
clasp open
```

Then run `setup` from the browser editor.
