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
 * Split managed events into {drive: [], buffer: [], busyMirror: []} by type.
 */
function categorizeManagedEvents(managedEvents) {
  const drive = [];
  const buffer = [];
  const busyMirror = [];

  for (const event of managedEvents) {
    const tag = parseTag(event.getDescription());
    if (!tag) continue;

    if (tag.type === 'drive-to' || tag.type === 'drive-from') {
      drive.push(event);
    } else if (tag.type === 'buffer') {
      buffer.push(event);
    } else if (tag.type === 'busy-mirror') {
      busyMirror.push(event);
    }
  }

  return { drive, buffer, busyMirror };
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
function updateManagedEvent(event, title, startTime, endTime, description, color) {
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
  if (color && event.getColor() !== color) {
    event.setColor(color);
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
// Color helpers
// ---------------------------------------------------------------------------

/**
 * Approximate hex values for the 11 per-event colors Apps Script exposes
 * via CalendarApp.EventColor. Used to find the closest event color to an
 * arbitrary calendar hex color (Calendar.getColor() returns hex from a
 * wider 24-color palette).
 */
function getEventColorHexMap_() {
  const m = {};
  m[CalendarApp.EventColor.PALE_BLUE] = '#7986cb';  // Lavender
  m[CalendarApp.EventColor.PALE_GREEN] = '#33b679'; // Sage
  m[CalendarApp.EventColor.MAUVE] = '#8e24aa';      // Grape
  m[CalendarApp.EventColor.PALE_RED] = '#e67c73';   // Flamingo
  m[CalendarApp.EventColor.YELLOW] = '#f6bf26';     // Banana
  m[CalendarApp.EventColor.ORANGE] = '#f4511e';     // Tangerine
  m[CalendarApp.EventColor.CYAN] = '#039be5';       // Peacock
  m[CalendarApp.EventColor.GRAY] = '#616161';       // Graphite
  m[CalendarApp.EventColor.BLUE] = '#3f51b5';       // Blueberry
  m[CalendarApp.EventColor.GREEN] = '#0b8043';      // Basil
  m[CalendarApp.EventColor.RED] = '#d50000';        // Tomato
  return m;
}

function hexToRgb_(hex) {
  const h = hex.replace('#', '').trim();
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16),
  };
}

/**
 * Return the CalendarApp.EventColor closest to the given hex color by
 * simple RGB-distance match. Falls back to GRAY on parse failure.
 */
function closestEventColor(hex) {
  if (!hex || typeof hex !== 'string' || hex.length < 7) {
    return CalendarApp.EventColor.GRAY;
  }

  let target;
  try {
    target = hexToRgb_(hex);
  } catch (e) {
    return CalendarApp.EventColor.GRAY;
  }

  const map = getEventColorHexMap_();
  let bestId = CalendarApp.EventColor.GRAY;
  let bestDist = Infinity;

  for (const id in map) {
    const c = hexToRgb_(map[id]);
    const dist =
      (c.r - target.r) * (c.r - target.r) +
      (c.g - target.g) * (c.g - target.g) +
      (c.b - target.b) * (c.b - target.b);
    if (dist < bestDist) {
      bestDist = dist;
      bestId = id;
    }
  }

  return bestId;
}

// ---------------------------------------------------------------------------
// Time helpers
// ---------------------------------------------------------------------------

function minutesToMs(minutes) {
  return minutes * 60 * 1000;
}
