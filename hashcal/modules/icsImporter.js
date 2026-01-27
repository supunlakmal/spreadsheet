function unfoldLines(text) {
  const rawLines = text.replace(/\r\n/g, "\n").split("\n");
  const lines = [];
  rawLines.forEach((line) => {
    if (!line) return;
    if (line.startsWith(" ") || line.startsWith("\t")) {
      if (lines.length > 0) {
        lines[lines.length - 1] += line.slice(1);
      }
    } else {
      lines.push(line.trim());
    }
  });
  return lines;
}

function parseIcsDate(value) {
  if (!value) return null;
  const dateOnly = /^\d{8}$/.test(value);
  const hasTime = value.includes("T");

  if (dateOnly) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6)) - 1;
    const day = Number(value.slice(6, 8));
    return { date: new Date(year, month, day), isAllDay: true };
  }

  if (hasTime) {
    const match = value.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?(Z)?$/);
    if (!match) return { date: new Date(value), isAllDay: false };
    const [, y, m, d, hh, mm, ss = "00", z] = match;
    const year = Number(y);
    const month = Number(m) - 1;
    const day = Number(d);
    const hour = Number(hh);
    const minute = Number(mm);
    const second = Number(ss);
    if (z) {
      return { date: new Date(Date.UTC(year, month, day, hour, minute, second)), isAllDay: false };
    }
    return { date: new Date(year, month, day, hour, minute, second), isAllDay: false };
  }

  return { date: new Date(value), isAllDay: false };
}

function parseRRule(value) {
  if (!value) return "";
  const parts = value.split(";");
  const freqPart = parts.find((part) => part.startsWith("FREQ="));
  if (!freqPart) return "";
  const freq = freqPart.split("=")[1];
  if (!freq) return "";
  if (freq === "DAILY") return "d";
  if (freq === "WEEKLY") return "w";
  if (freq === "MONTHLY") return "m";
  if (freq === "YEARLY") return "y";
  return "";
}

export function parseIcs(text) {
  const lines = unfoldLines(text);
  const events = [];
  let current = null;

  lines.forEach((line) => {
    if (line === "BEGIN:VEVENT") {
      current = {
        title: "Untitled",
        start: null,
        end: null,
        rule: "",
        isAllDay: false,
      };
      return;
    }

    if (line === "END:VEVENT") {
      if (current && current.start) {
        events.push(current);
      }
      current = null;
      return;
    }

    if (!current) return;

    if (line.startsWith("SUMMARY")) {
      const [, value] = line.split(":");
      if (value) current.title = value.trim();
      return;
    }

    if (line.startsWith("DTSTART")) {
      const [, value] = line.split(":");
      const parsed = parseIcsDate(value ? value.trim() : "");
      if (parsed) {
        current.start = parsed.date;
        current.isAllDay = parsed.isAllDay;
      }
      return;
    }

    if (line.startsWith("DTEND")) {
      const [, value] = line.split(":");
      const parsed = parseIcsDate(value ? value.trim() : "");
      if (parsed) {
        current.end = parsed.date;
      }
      return;
    }

    if (line.startsWith("RRULE")) {
      const [, value] = line.split(":");
      current.rule = parseRRule(value ? value.trim() : "");
    }
  });

  return events;
}
