/**
 * Calculate drive time in minutes from origin to destination.
 * Returns the buffered (pessimistic) estimate, or null if the route can't be found.
 */
function getDriveTimeMinutes(origin, destination, departureTime, bufferPercent) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `drive_${destination}_${departureTime.getDay()}_${departureTime.getHours()}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    Logger.log(`Cache hit for ${destination}: ${cached} min`);
    return parseInt(cached, 10);
  }

  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(origin)
      .setDestination(destination)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .setDepart(departureTime)
      .getDirections();

    if (!directions || !directions.routes || directions.routes.length === 0) {
      Logger.log(`No route found to ${destination}`);
      return null;
    }

    const leg = directions.routes[0].legs[0];
    const durationSeconds = leg.duration_in_traffic
      ? leg.duration_in_traffic.value
      : leg.duration.value;

    let minutes = Math.ceil(durationSeconds / 60);
    minutes = Math.ceil(minutes * (1 + bufferPercent / 100));

    // Round up to nearest 5 minutes for cleaner calendar blocks
    minutes = Math.ceil(minutes / 5) * 5;

    cache.put(cacheKey, minutes.toString(), 21600); // 6-hour TTL
    Logger.log(`Drive time to ${destination}: ${minutes} min (with ${bufferPercent}% buffer)`);

    return minutes;
  } catch (e) {
    Logger.log(`Error calculating drive time to ${destination}: ${e.message}`);
    return null;
  }
}

/**
 * Check if existing drive blocks still match the source event's time and location.
 * Returns true if both blocks exist and nothing changed (safe to skip).
 */
function driveBlocksMatchEvent(driveManagedEvents, sourceId, eventStart, eventEnd, location) {
  const driveTo = findManagedEvent(driveManagedEvents, 'drive-to', sourceId);
  const driveFrom = findManagedEvent(driveManagedEvents, 'drive-from', sourceId);

  if (!driveTo || !driveFrom) return false;

  const toTag = parseTag(driveTo.getDescription());
  const fromTag = parseTag(driveFrom.getDescription());

  return (
    toTag &&
    fromTag &&
    toTag.eventStart === eventStart &&
    toTag.eventEnd === eventEnd &&
    toTag.location === location &&
    fromTag.eventStart === eventStart &&
    fromTag.eventEnd === eventEnd &&
    fromTag.location === location
  );
}

/**
 * Process all drive-to and drive-from blocks: create, update, or clean up.
 * Only recalculates drive time when the source event's time or location changed.
 */
function processDriveTimeBlocks(calendar, sourceEvents, driveManagedEvents, config) {
  const eventsWithLocation = getSourceEventsWithLocation(sourceEvents);
  const processedSourceIds = {};
  const bufferPercent = parseInt(config.DRIVE_TIME_BUFFER_PERCENT, 10);

  for (const event of eventsWithLocation) {
    const sourceId = event.getId();
    const location = event.getLocation();
    const eventStart = String(event.getStartTime().getTime());
    const eventEnd = String(event.getEndTime().getTime());

    processedSourceIds[sourceId] = true;

    // Skip if both blocks exist and the source event hasn't changed
    if (driveBlocksMatchEvent(driveManagedEvents, sourceId, eventStart, eventEnd, location)) {
      Logger.log(`Skipping unchanged: ${event.getTitle()}`);
      continue;
    }

    const driveMinutes = getDriveTimeMinutes(
      config.OFFICE_ADDRESS,
      location,
      event.getStartTime(),
      bufferPercent
    );
    if (driveMinutes === null) continue;

    const driveMs = minutesToMs(driveMinutes);
    const metadata = { eventStart, eventEnd, location };

    // --- Drive TO block ---
    const driveToStart = new Date(event.getStartTime().getTime() - driveMs);
    const driveToEnd = event.getStartTime();
    const driveToTitle = `Drive to ${event.getTitle()}`;
    const driveToDescription = createTag('drive-to', sourceId, metadata);
    const existingDriveTo = findManagedEvent(driveManagedEvents, 'drive-to', sourceId);

    if (existingDriveTo) {
      updateManagedEvent(existingDriveTo, driveToTitle, driveToStart, driveToEnd, driveToDescription, '3');
    } else {
      createManagedEvent(calendar, driveToTitle, driveToStart, driveToEnd, driveToDescription);
    }

    // --- Drive FROM block ---
    const driveFromStart = event.getEndTime();
    const driveFromEnd = new Date(event.getEndTime().getTime() + driveMs);
    const driveFromTitle = `Drive from ${event.getTitle()}`;
    const driveFromDescription = createTag('drive-from', sourceId, metadata);
    const existingDriveFrom = findManagedEvent(driveManagedEvents, 'drive-from', sourceId);

    if (existingDriveFrom) {
      updateManagedEvent(existingDriveFrom, driveFromTitle, driveFromStart, driveFromEnd, driveFromDescription, '3');
    } else {
      createManagedEvent(calendar, driveFromTitle, driveFromStart, driveFromEnd, driveFromDescription);
    }
  }

  cleanupOrphanedEvents(driveManagedEvents, processedSourceIds, ['drive-to', 'drive-from']);
}
