const DEFAULT_CONFIG = {
  OFFICE_ADDRESS: '',
  LOOK_AHEAD_DAYS: '7',
  DRIVE_TIME_BUFFER_PERCENT: '25',
  BUFFER_DURATION_MINUTES: '15',
  MEETING_BLOCK_THRESHOLD_MINUTES: '120',
  CONSECUTIVE_GAP_MINUTES: '15',
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
 */
function main() {
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
  const managedEvents = getManagedEvents(calendar, now, endDate);
  const { drive, buffer } = categorizeManagedEvents(managedEvents);

  Logger.log(`Found ${sourceEvents.length} source events, ${drive.length} drive blocks, ${buffer.length} buffer blocks`);

  processDriveTimeBlocks(calendar, sourceEvents, drive, config);
  processBufferBlocks(calendar, sourceEvents, buffer, config);

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
 * Remove all triggers and delete managed events.
 */
function uninstall() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Removed main trigger');
    }
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const endDate = new Date(now.getTime() + minutesToMs(90 * 24 * 60)); // 90 days ahead
  const managedEvents = getManagedEvents(calendar, now, endDate);

  for (const event of managedEvents) {
    try {
      event.deleteEvent();
    } catch (e) {
      Logger.log(`Failed to delete event: ${e.message}`);
    }
  }

  Logger.log(`Removed ${managedEvents.length} managed events. Uninstall complete.`);
}
