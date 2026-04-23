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

/**
 * Unique per-instance ID for a calendar event.
 *
 * event.getId() returns the iCalUID, which is the SAME across every
 * instance of a recurring event. Using it alone as a managed-event
 * source key causes every instance of a recurring series to share one
 * drive block / busy mirror — each iteration of the main loop overwrites
 * the previous instance's block, so only the last instance ends up with
 * any managed blocks at all.
 *
 * Appending the instance start time makes each occurrence addressable
 * on its own.
 */
function getEventInstanceId(event) {
  return event.getId() + '_' + event.getStartTime().getTime();
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
  const chosen = color || CalendarApp.EventColor.MAUVE;
  try {
    event.setColor(String(chosen));
  } catch (e) {
    Logger.log(`Failed to set color on "${title}" (color=${chosen}): ${e.message}`);
  }
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
  if (color && String(color) !== String(event.getColor())) {
    try {
      event.setColor(String(color));
      changed = true;
    } catch (e) {
      Logger.log(`Skipping color update for "${title}" (color=${color}): ${e.message}`);
    }
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
 * Per-event color palette: the 11 named colors Apps Script exposes via
 * event.setColor(). Stored as [enumValue, hex] pairs so iteration order
 * is stable and we don't rely on object-key coercion of enum values
 * (which behaves inconsistently between runtimes).
 *
 * Hex values are the modern Google Calendar named-color palette
 * (Lavender, Sage, Grape, ...).
 */
function getEventColorPalette_() {
  return [
    [CalendarApp.EventColor.PALE_BLUE, '#7986cb'],  // Lavender
    [CalendarApp.EventColor.PALE_GREEN, '#33b679'], // Sage
    [CalendarApp.EventColor.MAUVE, '#8e24aa'],      // Grape
    [CalendarApp.EventColor.PALE_RED, '#e67c73'],   // Flamingo
    [CalendarApp.EventColor.YELLOW, '#f6bf26'],     // Banana
    [CalendarApp.EventColor.ORANGE, '#f4511e'],     // Tangerine
    [CalendarApp.EventColor.CYAN, '#039be5'],       // Peacock
    [CalendarApp.EventColor.GRAY, '#616161'],       // Graphite
    [CalendarApp.EventColor.BLUE, '#3f51b5'],       // Blueberry
    [CalendarApp.EventColor.GREEN, '#0b8043'],      // Basil
    [CalendarApp.EventColor.RED, '#d50000'],        // Tomato
  ];
}

function hexToRgb_(hex) {
  const h = String(hex || '').replace('#', '').trim();
  if (h.length < 6) return null;
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
  return { r: r, g: g, b: b };
}

/**
 * Return the CalendarApp.EventColor closest to the given hex color by
 * simple RGB-distance match. Falls back to GRAY on parse failure.
 */
function closestEventColor(hex) {
  const target = hexToRgb_(hex);
  const palette = getEventColorPalette_();
  if (!target) return CalendarApp.EventColor.GRAY;

  let bestId = CalendarApp.EventColor.GRAY;
  let bestDist = Infinity;

  for (let i = 0; i < palette.length; i++) {
    const id = palette[i][0];
    const c = hexToRgb_(palette[i][1]);
    if (!c) continue;
    const dr = c.r - target.r;
    const dg = c.g - target.g;
    const db = c.b - target.b;
    const dist = dr * dr + dg * dg + db * db;
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
