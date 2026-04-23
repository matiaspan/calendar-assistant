const TAG_MARKER = '--- Managed by Calendar Assistant ---';

/**
 * Create a tag string to embed in a managed event's description.
 * Optional metadata object stores extra key:value pairs (e.g. eventStart, location).
 */
function createTag(type, sourceEventId, metadata) {
  let tag = `${TAG_MARKER}\ntype:${type}\nsource:${sourceEventId}`;
  if (metadata) {
    for (const key in metadata) {
      tag += `\n${key}:${metadata[key]}`;
    }
  }
  return tag;
}

/**
 * Parse a tag from an event description.
 * Returns an object with type, source, and any extra metadata keys, or null.
 */
function parseTag(description) {
  if (!description || description.indexOf(TAG_MARKER) === -1) {
    return null;
  }

  const tagSection = description.substring(description.indexOf(TAG_MARKER));
  // Calendar may return descriptions as HTML (e.g. <br> in place of \n).
  // Normalize back to newlines so tag lines parse correctly.
  const normalized = tagSection
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(p|div)[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, '');
  const lines = normalized.split('\n');
  const result = {};

  for (const line of lines) {
    const colonIndex = line.indexOf(':');
    if (colonIndex > 0 && line !== TAG_MARKER) {
      result[line.substring(0, colonIndex)] = line.substring(colonIndex + 1).trim();
    }
  }

  return result.type && result.source ? result : null;
}

/**
 * Check if an event was created by Calendar Assistant.
 */
function isManagedEvent(event) {
  return parseTag(event.getDescription()) !== null;
}

// ---------------------------------------------------------------------------
// Event filtering
// ---------------------------------------------------------------------------

/**
 * Get all real (non-managed) events the user owns or has accepted.
 */
function getSourceEvents(calendar, startDate, endDate) {
  return calendar.getEvents(startDate, endDate).filter((event) => {
    if (event.isAllDayEvent()) return false;
    if (isManagedEvent(event)) return false;

    const status = event.getMyStatus();
    return (
      status === CalendarApp.GuestStatus.OWNER ||
      status === CalendarApp.GuestStatus.YES ||
      status === null
    );
  });
}

/**
 * Filter source events to only those with a physical location set.
 */
function getSourceEventsWithLocation(events) {
  return events.filter((event) => {
    const location = event.getLocation();
    return location && location.trim().length > 0;
  });
}

/**
 * Get all managed events (created by Calendar Assistant) in a date range.
 */
function getManagedEvents(calendar, startDate, endDate) {
  return calendar.getEvents(startDate, endDate).filter((event) => {
    return isManagedEvent(event);
  });
}

/**
 * Split managed events into {drive: [], buffer: []} by type.
 */
function categorizeManagedEvents(managedEvents) {
  const drive = [];
  const buffer = [];

  for (const event of managedEvents) {
    const tag = parseTag(event.getDescription());
    if (!tag) continue;

    if (tag.type === 'drive-to' || tag.type === 'drive-from') {
      drive.push(event);
    } else if (tag.type === 'buffer') {
      buffer.push(event);
    }
  }

  return { drive, buffer };
}

// ---------------------------------------------------------------------------
// Managed event CRUD
// ---------------------------------------------------------------------------

/**
 * Find a managed event by type and source ID.
 */
function findManagedEvent(managedEvents, type, sourceId) {
  for (const event of managedEvents) {
    const tag = parseTag(event.getDescription());
    if (tag && tag.type === type && tag.source === sourceId) {
      return event;
    }
  }
  return null;
}

/**
 * Create a new managed calendar event.
 */
function createManagedEvent(calendar, title, startTime, endTime, description, color) {
  const event = calendar.createEvent(title, startTime, endTime, {
    description: description,
  });
  event.setColor(color || CalendarApp.EventColor.MAUVE);
  event.removeAllReminders();
  Logger.log(`Created: ${title} (${startTime.toLocaleString()} - ${endTime.toLocaleString()})`);
  return event;
}

/**
 * Update an existing managed event if anything changed.
 */
function updateManagedEvent(event, title, startTime, endTime, description) {
  let changed = false;

  if (event.getTitle() !== title) {
    event.setTitle(title);
    changed = true;
  }
  if (
    event.getStartTime().getTime() !== startTime.getTime() ||
    event.getEndTime().getTime() !== endTime.getTime()
  ) {
    event.setTime(startTime, endTime);
    changed = true;
  }
  if (event.getDescription() !== description) {
    event.setDescription(description);
    changed = true;
  }

  if (changed) {
    Logger.log(`Updated: ${title}`);
  }
}

/**
 * Delete managed events whose source no longer exists.
 */
function cleanupOrphanedEvents(managedEvents, processedSourceIds, types) {
  for (const event of managedEvents) {
    try {
      const tag = parseTag(event.getDescription());
      if (tag && types.includes(tag.type) && !processedSourceIds[tag.source]) {
        Logger.log(`Deleting orphaned: ${event.getTitle()}`);
        event.deleteEvent();
      }
    } catch (e) {
      Logger.log(`Skipping already-deleted event: ${e.message}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Time helpers
// ---------------------------------------------------------------------------

function minutesToMs(minutes) {
  return minutes * 60 * 1000;
}
