/**
 * Process events from additional (e.g. personal) calendars:
 *
 *   1. Write drive-to / drive-from blocks to the external calendar itself,
 *      so the user sees meaningful "Drive to X" entries where the event lives.
 *
 *   2. Write ONE generic "Busy" block on the primary (work) calendar per
 *      external event. The busy window is expanded to
 *      [eventStart − driveTo, eventEnd + driveFrom] so coworkers looking at
 *      free/busy see the entire outing (drive + appointment + drive back)
 *      as one contiguous busy block and can't schedule inside it.
 *
 * Personal details (title, location, attendees, description) are NEVER
 * copied to the work calendar — only the time window and an opaque source
 * ID are stored in the mirror's tag.
 *
 * External calendars must be writable by the account running the script
 * (typically the user's own personal calendar shared into their work
 * account with "Make changes to events" permission). Read-only calendars
 * will log a permission error for the drive-block step but still get
 * mirrored on the work calendar.
 */
function processExternalCalendars(
  workCalendar,
  externalCalendarIds,
  startDate,
  endDate,
  config,
  busyMirrorManagedEvents
) {
  const processedMirrorSourceIds = {};
  const bufferPercent = parseInt(config.DRIVE_TIME_BUFFER_PERCENT, 10);

  for (const id of externalCalendarIds) {
    let extCal;
    try {
      extCal = CalendarApp.getCalendarById(id);
    } catch (e) {
      Logger.log(`Error accessing external calendar ${id}: ${e.message}`);
      continue;
    }
    if (!extCal) {
      Logger.log(`External calendar not accessible: ${id}`);
      continue;
    }

    const externalEvents = extCal
      .getEvents(startDate, endDate)
      .filter((e) => !e.isAllDayEvent() && !isManagedEvent(e));
    Logger.log(`  External ${extCal.getName()}: ${externalEvents.length} events`);

    // Step 1: drive blocks on the external calendar itself.
    try {
      const externalManaged = getManagedEvents(extCal, startDate, endDate);
      const { drive: externalDriveManaged } = categorizeManagedEvents(externalManaged);
      processDriveTimeBlocks(extCal, externalEvents, externalDriveManaged, config);
    } catch (e) {
      Logger.log(`Failed to manage drive blocks on ${id} (is it writable?): ${e.message}`);
    }

    // Step 2: expanded "Busy" mirror on the work calendar.
    for (const event of externalEvents) {
      const sourceId = event.getId();
      processedMirrorSourceIds[sourceId] = true;

      const eventStart = event.getStartTime();
      const eventEnd = event.getEndTime();
      const location = event.getLocation();

      let mirrorStart = eventStart;
      let mirrorEnd = eventEnd;
      let driveMs = 0;

      if (location && location.trim().length > 0) {
        const driveMinutes = getDriveTimeMinutes(
          config.OFFICE_ADDRESS,
          location,
          eventStart,
          bufferPercent
        );
        if (driveMinutes !== null) {
          driveMs = minutesToMs(driveMinutes);
          mirrorStart = new Date(eventStart.getTime() - driveMs);
          mirrorEnd = new Date(eventEnd.getTime() + driveMs);
        }
      }

      const metadata = {
        eventStart: String(eventStart.getTime()),
        eventEnd: String(eventEnd.getTime()),
        driveMs: String(driveMs),
      };
      const description = createTag('busy-mirror', sourceId, metadata);
      const existing = findManagedEvent(busyMirrorManagedEvents, 'busy-mirror', sourceId);

      if (existing) {
        updateManagedEvent(existing, 'Busy', mirrorStart, mirrorEnd, description);
      } else {
        createManagedEvent(
          workCalendar,
          'Busy',
          mirrorStart,
          mirrorEnd,
          description,
          CalendarApp.EventColor.GRAY
        );
      }
    }
  }

  cleanupOrphanedEvents(busyMirrorManagedEvents, processedMirrorSourceIds, ['busy-mirror']);
}
