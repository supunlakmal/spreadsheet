const MINUTE = 60000;
const DAY = 86400000;

function addDays(date, days) {
  const next = new Date(date.getTime());
  next.setDate(next.getDate() + days);
  return next;
}

function addMonths(date, months) {
  const next = new Date(date.getTime());
  const day = next.getDate();
  next.setMonth(next.getMonth() + months);
  if (next.getDate() < day) {
    next.setDate(0);
  }
  return next;
}

function addYears(date, years) {
  const next = new Date(date.getTime());
  const month = next.getMonth();
  next.setFullYear(next.getFullYear() + years);
  if (next.getMonth() !== month) {
    next.setDate(0);
  }
  return next;
}

function getFirstOccurrence(start, rangeStart, rule) {
  let current = new Date(start.getTime());
  if (current >= rangeStart) return current;

  if (rule === "d" || rule === "w") {
    const step = rule === "d" ? 1 : 7;
    const diffDays = Math.floor((rangeStart.getTime() - current.getTime()) / DAY);
    const jumps = Math.floor(diffDays / step);
    current = addDays(current, jumps * step);
    while (current < rangeStart) {
      current = addDays(current, step);
    }
    return current;
  }

  if (rule === "m") {
    while (current < rangeStart) {
      current = addMonths(current, 1);
    }
    return current;
  }

  if (rule === "y") {
    while (current < rangeStart) {
      current = addYears(current, 1);
    }
  }

  return current;
}

function advanceOccurrence(date, rule) {
  if (rule === "d") return addDays(date, 1);
  if (rule === "w") return addDays(date, 7);
  if (rule === "m") return addMonths(date, 1);
  if (rule === "y") return addYears(date, 1);
  return addDays(date, 1);
}

export function expandEvents(events, rangeStart, rangeEnd) {
  const occurrences = [];
  const rangeStartMs = rangeStart.getTime();
  const rangeEndMs = rangeEnd.getTime();

  events.forEach((event, index) => {
    if (!Array.isArray(event)) return;
    const startMin = Number(event[0]);
    if (!Number.isFinite(startMin)) return;
    const duration = Number(event[1]) || 0;
    const title = String(event[2] || "Untitled");
    const colorIndex = Number(event[3]) || 0;
    const rule = event[4] || "";

    const startMs = startMin * MINUTE;
    const durationMs = duration * MINUTE;
    const isAllDay = duration === 0;

    if (!rule) {
      const endMs = isAllDay ? startMs : startMs + durationMs;
      if (startMs <= rangeEndMs && endMs >= rangeStartMs) {
        occurrences.push({
          start: startMs,
          end: endMs,
          title,
          colorIndex,
          rule: "",
          sourceIndex: index,
          isAllDay,
        });
      }
      return;
    }

    let current = getFirstOccurrence(new Date(startMs), rangeStart, rule);
    while (current.getTime() <= rangeEndMs) {
      const currentMs = current.getTime();
      const endMs = isAllDay ? currentMs : currentMs + durationMs;
      if (currentMs <= rangeEndMs && endMs >= rangeStartMs) {
        occurrences.push({
          start: currentMs,
          end: endMs,
          title,
          colorIndex,
          rule,
          sourceIndex: index,
          isAllDay,
        });
      }
      current = advanceOccurrence(current, rule);
    }
  });

  return occurrences;
}
