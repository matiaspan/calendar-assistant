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

  // Track whether every configured external calendar fetched cleanly.
  // If any transient failure (calendar inaccessible, getEvents throws, etc.)
  // skipped a calendar, we skip the busy-mirror cleanup at the end — otherwise
  // every mirror event would be treated as orphaned and deleted.
  let allCalendarsFetchedOk = true;

  for (const id of externalCalendarIds) {
    let extCal;
    try {
      extCal = CalendarApp.getCalendarById(id);
    } catch (e) {
      Logger.log(`Error accessing external calendar ${id}: ${e.message}`);
      allCalendarsFetchedOk = false;
      continue;
    }
    if (!extCal) {
      Logger.log(`External calendar not accessible: ${id}`);
      allCalendarsFetchedOk = false;
      continue;
    }

    let externalEvents;
    try {
      externalEvents = extCal
        .getEvents(startDate, endDate)
        .filter((e) => !e.isAllDayEvent() && !isManagedEvent(e));
    } catch (e) {
      Logger.log(`Failed to fetch events from ${extCal.getName()}: ${e.message}`);
      allCalendarsFetchedOk = false;
      continue;
    }

    const mirrorColor = closestEventColor(extCal.getColor());
    Logger.log(
      `  External ${extCal.getName()}: ${externalEvents.length} events, ` +
        `mirror color ${mirrorColor} (from ${extCal.getColor()})`
    );

    // Step 1: drive blocks on the external calendar itself. Only let
    // processDriveTimeBlocks run its own cleanup if we successfully fetched
    // events for this calendar (we did, if we got here).
    try {
      let externalManaged = getManagedEvents(extCal, startDate, endDate);
      externalManaged = dedupeManagedEvents(externalManaged, ['drive-to', 'drive-from']);
      const { drive: externalDriveManaged } = categorizeManagedEvents(externalManaged);
      processDriveTimeBlocks(extCal, externalEvents, externalDriveManaged, config);
    } catch (e) {
      Logger.log(`Failed to manage drive blocks on ${id} (is it writable?): ${e.message}`);
      allCalendarsFetchedOk = false;
    }

    // Step 2: expanded "Busy" mirror on the work calendar.
    for (const event of externalEvents) {
      try {
        const sourceId = getEventInstanceId(event);
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
          updateManagedEvent(existing, 'Busy', mirrorStart, mirrorEnd, description, mirrorColor);
        } else {
          createManagedEvent(
            workCalendar,
            'Busy',
            mirrorStart,
            mirrorEnd,
            description,
            mirrorColor
          );
        }
      } catch (e) {
        Logger.log(`Error mirroring external event: ${e.message}`);
        allCalendarsFetchedOk = false;
      }
    }
  }

  if (allCalendarsFetchedOk) {
    cleanupOrphanedEvents(busyMirrorManagedEvents, processedMirrorSourceIds, ['busy-mirror']);
  } else {
    Logger.log(
      'Skipping busy-mirror cleanup: at least one external calendar had a fetch error. ' +
        'Stale mirrors will be cleaned up on a subsequent successful run.'
    );
  }
}
