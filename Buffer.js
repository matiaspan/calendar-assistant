/**
 * Group sorted events into blocks of consecutive meetings.
 * Two events are consecutive if the gap between them is <= gapMs.
 */
function groupConsecutiveEvents(sortedEvents, gapMs) {
  if (sortedEvents.length === 0) return [];

  const blocks = [];
  let currentBlock = [sortedEvents[0]];

  for (let i = 1; i < sortedEvents.length; i++) {
    const prevEnd = currentBlock[currentBlock.length - 1].getEndTime().getTime();
    const currStart = sortedEvents[i].getStartTime().getTime();

    if (currStart - prevEnd <= gapMs) {
      currentBlock.push(sortedEvents[i]);
    } else {
      blocks.push(currentBlock);
      currentBlock = [sortedEvents[i]];
    }
  }

  blocks.push(currentBlock);
  return blocks;
}

/**
 * Process buffer blocks: insert a break before and after any 2+ hour stretch of meetings.
 */
function processBufferBlocks(calendar, sourceEvents, bufferManagedEvents, config) {
  const thresholdMs = minutesToMs(parseInt(config.MEETING_BLOCK_THRESHOLD_MINUTES, 10));
  const gapMs = minutesToMs(parseInt(config.CONSECUTIVE_GAP_MINUTES, 10));
  const bufferMinutes = parseInt(config.BUFFER_DURATION_MINUTES, 10);
  const bufferMs = minutesToMs(bufferMinutes);

  const sorted = sourceEvents.slice().sort(
    (a, b) => a.getStartTime().getTime() - b.getStartTime().getTime()
  );

  const blocks = groupConsecutiveEvents(sorted, gapMs);
  const processedSourceIds = {};
  let allBlocksProcessedOk = true;

  for (const block of blocks) {
    try {
      const blockStart = block[0].getStartTime().getTime();
      const blockEnd = block[block.length - 1].getEndTime().getTime();
      const blockDuration = blockEnd - blockStart;

      if (blockDuration < thresholdMs) continue;

      const firstId = block[0].getId();
      const lastId = block[block.length - 1].getId();
      const beforeSourceId = `block_before_${firstId}`;
      const afterSourceId = `block_after_${lastId}`;

      processedSourceIds[beforeSourceId] = true;
      processedSourceIds[afterSourceId] = true;

      // --- Buffer BEFORE the block ---
      const beforeStart = new Date(blockStart - bufferMs);
      const beforeEnd = new Date(blockStart);
      const beforeDescription = createTag('buffer', beforeSourceId);
      const existingBefore = findManagedEvent(bufferManagedEvents, 'buffer', beforeSourceId);

      if (existingBefore) {
        updateManagedEvent(existingBefore, 'Break', beforeStart, beforeEnd, beforeDescription, CalendarApp.EventColor.PALE_GREEN);
      } else {
        createManagedEvent(calendar, 'Break', beforeStart, beforeEnd, beforeDescription, CalendarApp.EventColor.PALE_GREEN);
      }

      // --- Buffer AFTER the block ---
      const afterStart = new Date(blockEnd);
      const afterEnd = new Date(blockEnd + bufferMs);
      const afterDescription = createTag('buffer', afterSourceId);
      const existingAfter = findManagedEvent(bufferManagedEvents, 'buffer', afterSourceId);

      if (existingAfter) {
        updateManagedEvent(existingAfter, 'Break', afterStart, afterEnd, afterDescription, CalendarApp.EventColor.PALE_GREEN);
      } else {
        createManagedEvent(calendar, 'Break', afterStart, afterEnd, afterDescription, CalendarApp.EventColor.PALE_GREEN);
      }
    } catch (e) {
      allBlocksProcessedOk = false;
      Logger.log(`Error processing buffer block: ${e.message}`);
    }
  }

  if (allBlocksProcessedOk) {
    cleanupOrphanedEvents(bufferManagedEvents, processedSourceIds, ['buffer']);
  } else {
    Logger.log(
      'Skipping buffer cleanup: one or more blocks errored during processing. ' +
        'Stale buffers will be cleaned on a subsequent successful run.'
    );
  }
}
