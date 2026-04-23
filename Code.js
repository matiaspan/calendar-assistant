const DEFAULT_CONFIG = {
  OFFICE_ADDRESS: '',
  LOOK_AHEAD_DAYS: '7',
  DRIVE_TIME_BUFFER_PERCENT: '10',
  BUFFER_DURATION_MINUTES: '30',
  MEETING_BLOCK_THRESHOLD_MINUTES: '120',
  CONSECUTIVE_GAP_MINUTES: '30',
  // Comma-separated list of additional calendar IDs to process (e.g. a
  // personal calendar overlaid on the work calendar). Events from these
  // calendars get drive blocks written back to them, and a single generic
  // "Busy" block on the work calendar covering event + drive time.
  SOURCE_CALENDAR_IDS: '',
};

/**
 * Load configuration from Script Properties, falling back to defaults.
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const config = {};

  for (const key in DEFAULT_CONFIG) {
    const value = props.getProperty(key);
    config[key] = value !== null ? value : DEFAULT_CONFIG[key];
  }

  return config;
}

/**
 * Main entry point — called by the time-driven trigger.
 *
 * Guarded by a script-wide lock so a manual editor run and a scheduled
 * trigger invocation can't race each other. Apps Script does NOT serialize
 * concurrent trigger/editor executions by default, and when both runs
 * interleave they each see the same old state, each create a fresh set of
 * managed blocks, and nothing in the normal cleanup path recognizes the
 * extras as orphaned (same tag source on both copies).
 */
function main() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('Another main() is already running. Skipping this invocation.');
    return;
  }

  try {
    runMain();
  } finally {
    lock.releaseLock();
  }
}

function runMain() {
  const config = getConfig();

  if (!config.OFFICE_ADDRESS) {
    Logger.log('ERROR: OFFICE_ADDRESS not set. Go to Project Settings > Script Properties.');
    return;
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const lookAheadDays = parseInt(config.LOOK_AHEAD_DAYS, 10);
  const endDate = new Date(now.getTime() + minutesToMs(lookAheadDays * 24 * 60));

  Logger.log(`Scanning ${calendar.getName()} from ${now.toLocaleString()} to ${endDate.toLocaleString()}`);

  const sourceEvents = getSourceEvents(calendar, now, endDate);
  let managedEvents = getManagedEvents(calendar, now, endDate);

  // Remove any duplicate managed events that an earlier concurrent run
  // left behind (same tag source appearing more than once per type).
  managedEvents = dedupeManagedEvents(managedEvents, [
    'drive-to',
    'drive-from',
    'buffer',
    'busy-mirror',
  ]);

  const { drive, buffer, busyMirror } = categorizeManagedEvents(managedEvents);

  const externalCalendarIds = (config.SOURCE_CALENDAR_IDS || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);

  Logger.log(
    `Found ${sourceEvents.length} source events, ${drive.length} drive blocks, ` +
      `${buffer.length} buffer blocks, ${busyMirror.length} busy mirrors, ` +
      `${externalCalendarIds.length} external calendars`
  );

  processDriveTimeBlocks(calendar, sourceEvents, drive, config);
  processBufferBlocks(calendar, sourceEvents, buffer, config);

  if (externalCalendarIds.length > 0) {
    processExternalCalendars(
      calendar,
      externalCalendarIds,
      now,
      endDate,
      config,
      busyMirror
    );
  }

  Logger.log('Done.');
}

/**
 * One-time setup: write default config and create the recurring trigger.
 */
function setup() {
  const props = PropertiesService.getScriptProperties();

  for (const key in DEFAULT_CONFIG) {
    if (props.getProperty(key) === null && DEFAULT_CONFIG[key] !== '') {
      props.setProperty(key, DEFAULT_CONFIG[key]);
    }
  }

  const triggers = ScriptApp.getProjectTriggers();
  const hasMainTrigger = triggers.some((t) => t.getHandlerFunction() === 'main');

  if (!hasMainTrigger) {
    ScriptApp.newTrigger('main').timeBased().everyMinutes(5).create();
    Logger.log('Created trigger: main() every 5 minutes');
  } else {
    Logger.log('Trigger already exists');
  }

  if (!props.getProperty('OFFICE_ADDRESS')) {
    Logger.log('ACTION REQUIRED: Set OFFICE_ADDRESS in Project Settings > Script Properties');
  }

  Logger.log('Setup complete.');
}

/**
 * Remove all triggers and delete managed events (on both the primary
 * calendar and any configured external calendars).
 */
function uninstall() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Removed main trigger');
    }
  }

  const config = getConfig();
  const now = new Date();
  const endDate = new Date(now.getTime() + minutesToMs(90 * 24 * 60)); // 90 days ahead

  let total = deleteManagedEventsFromCalendar(CalendarApp.getDefaultCalendar(), now, endDate);

  const externalCalendarIds = (config.SOURCE_CALENDAR_IDS || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
  for (const id of externalCalendarIds) {
    try {
      const extCal = CalendarApp.getCalendarById(id);
      if (extCal) total += deleteManagedEventsFromCalendar(extCal, now, endDate);
    } catch (e) {
      Logger.log(`Failed to clean external calendar ${id}: ${e.message}`);
    }
  }

  Logger.log(`Removed ${total} managed events. Uninstall complete.`);
}

function deleteManagedEventsFromCalendar(calendar, startDate, endDate) {
  const managedEvents = getManagedEvents(calendar, startDate, endDate);
  let count = 0;
  for (const event of managedEvents) {
    try {
      event.deleteEvent();
      count++;
    } catch (e) {
      Logger.log(`Failed to delete event: ${e.message}`);
    }
  }
  return count;
}
