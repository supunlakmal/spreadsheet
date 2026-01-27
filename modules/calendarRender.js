function pad(num) {
  return String(num).padStart(2, "0");
}

export function formatDateKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

export function getMonthGridRange(viewDate, weekStartsOnMonday) {
  const year = viewDate.getFullYear();
  const month = viewDate.getMonth();
  const firstOfMonth = new Date(year, month, 1);
  const dayOfWeek = firstOfMonth.getDay();
  const offset = weekStartsOnMonday ? (dayOfWeek + 6) % 7 : dayOfWeek;
  const gridStart = new Date(year, month, 1 - offset);

  const dates = [];
  for (let i = 0; i < 42; i += 1) {
    const date = new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + i);
    dates.push(date);
  }
  const gridEnd = new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + 41, 23, 59, 59);
  return { start: gridStart, end: gridEnd, dates };
}

export function getWeekRange(viewDate, weekStartsOnMonday) {
  const anchor = new Date(viewDate.getFullYear(), viewDate.getMonth(), viewDate.getDate());
  const dayOfWeek = anchor.getDay();
  const offset = weekStartsOnMonday ? (dayOfWeek + 6) % 7 : dayOfWeek;
  const start = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - offset);
  const dates = [];
  for (let i = 0; i < 7; i += 1) {
    dates.push(new Date(start.getFullYear(), start.getMonth(), start.getDate() + i));
  }
  const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59);
  return { start, end, dates };
}

function formatWeekdayLabel(date) {
  return date.toLocaleDateString(undefined, { weekday: "short" });
}

export function renderWeekdayHeaders(container, weekStartsOnMonday, mode = "month", dates = []) {
  let labels = [];
  if (mode === "day" && dates[0]) {
    labels = [
      dates[0].toLocaleDateString(undefined, {
        weekday: "long",
        month: "short",
        day: "numeric",
      }),
    ];
  } else if (mode === "week" && dates.length) {
    labels = dates.map((date) => `${formatWeekdayLabel(date)} ${date.getDate()}`);
  } else {
    labels = weekStartsOnMonday
      ? ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
      : ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  }

  container.innerHTML = "";
  container.style.gridTemplateColumns = `repeat(${labels.length}, minmax(0, 1fr))`;
  labels.forEach((label) => {
    const cell = document.createElement("div");
    cell.className = "weekday";
    cell.textContent = label;
    container.appendChild(cell);
  });
}

export function renderCalendar({
  container,
  dates,
  currentMonth,
  selectedDate,
  eventsByDay,
  onSelectDay,
  onEventClick,
}) {
  container.innerHTML = "";
  const today = new Date();
  const todayKey = formatDateKey(today);
  const selectedKey = selectedDate ? formatDateKey(selectedDate) : null;

  dates.forEach((date) => {
    const key = formatDateKey(date);
    const dayCell = document.createElement("div");
    dayCell.className = "day-cell";
    dayCell.setAttribute("role", "gridcell");
    if (date.getMonth() !== currentMonth) dayCell.classList.add("is-outside");
    if (key === todayKey) dayCell.classList.add("is-today");
    if (selectedKey && key === selectedKey) dayCell.classList.add("is-selected");
    dayCell.dataset.date = key;

    const dayNumber = document.createElement("div");
    dayNumber.className = "day-number";
    dayNumber.textContent = String(date.getDate());
    dayCell.appendChild(dayNumber);

    const events = eventsByDay.get(key) || [];
    const maxEvents = 3;
    events.slice(0, maxEvents).forEach((event) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "event-chip";
      chip.style.background = event.color;
      chip.style.borderColor = event.color;

      const time = document.createElement("span");
      time.className = "time";
      time.textContent = event.isAllDay ? "All day" : event.timeLabel;
      const title = document.createElement("span");
      title.textContent = event.title;

      chip.appendChild(time);
      chip.appendChild(title);
      chip.addEventListener("click", (e) => {
        e.stopPropagation();
        onEventClick(event);
      });
      dayCell.appendChild(chip);
    });

    if (events.length > maxEvents) {
      const more = document.createElement("div");
      more.className = "event-more";
      more.textContent = `+${events.length - maxEvents} more`;
      dayCell.appendChild(more);
    }

    dayCell.addEventListener("click", () => onSelectDay(date));
    container.appendChild(dayCell);
  });
}

function formatHourLabel(hour) {
  const suffix = hour < 12 ? "AM" : "PM";
  const display = hour % 12 || 12;
  return `${display} ${suffix}`;
}

export function renderTimeGrid({ container, dates, occurrences, onSelectDay, onEventClick }) {
  const dayKeys = dates.map((date) => formatDateKey(date));
  const buckets = new Map();
  dayKeys.forEach((key) => {
    buckets.set(key, { allDay: [], hours: new Map() });
  });

  occurrences.forEach((occ) => {
    const key = formatDateKey(new Date(occ.start));
    const bucket = buckets.get(key);
    if (!bucket) return;
    if (occ.isAllDay) {
      bucket.allDay.push(occ);
      return;
    }
    const hour = new Date(occ.start).getHours();
    if (!bucket.hours.has(hour)) bucket.hours.set(hour, []);
    bucket.hours.get(hour).push(occ);
  });

  buckets.forEach((bucket) => {
    bucket.allDay.sort((a, b) => a.start - b.start);
    bucket.hours.forEach((list) => list.sort((a, b) => a.start - b.start));
  });

  container.innerHTML = "";
  container.style.gridTemplateColumns = `72px repeat(${dates.length}, minmax(0, 1fr))`;
  container.style.gridTemplateRows = `32px repeat(24, minmax(44px, 1fr))`;

  const allDayLabel = document.createElement("div");
  allDayLabel.className = "time-label all-day-label";
  allDayLabel.textContent = "All day";
  container.appendChild(allDayLabel);

  dates.forEach((date) => {
    const key = formatDateKey(date);
    const cell = document.createElement("div");
    cell.className = "time-cell all-day";
    cell.dataset.date = key;
    const bucket = buckets.get(key);
    const events = bucket ? bucket.allDay : [];
    events.slice(0, 2).forEach((event) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "time-event";
      chip.style.background = event.color;
      chip.textContent = event.title;
      chip.addEventListener("click", (e) => {
        e.stopPropagation();
        onEventClick(event);
      });
      cell.appendChild(chip);
    });
    if (events.length > 2) {
      const more = document.createElement("div");
      more.className = "event-more";
      more.textContent = `+${events.length - 2} more`;
      cell.appendChild(more);
    }
    cell.addEventListener("click", () => onSelectDay(date));
    container.appendChild(cell);
  });

  for (let hour = 0; hour < 24; hour += 1) {
    const label = document.createElement("div");
    label.className = "time-label";
    label.textContent = formatHourLabel(hour);
    container.appendChild(label);

    dates.forEach((date) => {
      const key = formatDateKey(date);
      const cell = document.createElement("div");
      cell.className = "time-cell";
      cell.dataset.date = key;
      cell.dataset.hour = String(hour);
      const bucket = buckets.get(key);
      const events = bucket && bucket.hours.get(hour) ? bucket.hours.get(hour) : [];
      events.slice(0, 3).forEach((event) => {
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "time-event";
        chip.style.background = event.color;
        chip.textContent = event.title;
        chip.addEventListener("click", (e) => {
          e.stopPropagation();
          onEventClick(event);
        });
        cell.appendChild(chip);
      });
      if (events.length > 3) {
        const more = document.createElement("div");
        more.className = "event-more";
        more.textContent = `+${events.length - 3} more`;
        cell.appendChild(more);
      }
      cell.addEventListener("click", () => onSelectDay(date));
      container.appendChild(cell);
    });
  }
}

export function renderYearView({ container, year, eventsByDay, selectedDate, weekStartsOnMonday, onSelectDay }) {
  container.innerHTML = "";
  container.style.gridTemplateColumns = "repeat(auto-fit, minmax(220px, 1fr))";
  container.style.gridTemplateRows = "auto";

  const labels = weekStartsOnMonday
    ? ["M", "T", "W", "T", "F", "S", "S"]
    : ["S", "M", "T", "W", "T", "F", "S"];

  for (let month = 0; month < 12; month += 1) {
    const wrapper = document.createElement("div");
    wrapper.className = "mini-month";

    const header = document.createElement("div");
    header.className = "mini-title";
    header.textContent = new Date(year, month, 1).toLocaleDateString(undefined, { month: "long" });
    wrapper.appendChild(header);

    const grid = document.createElement("div");
    grid.className = "mini-grid";

    labels.forEach((label) => {
      const cell = document.createElement("div");
      cell.className = "mini-weekday";
      cell.textContent = label;
      grid.appendChild(cell);
    });

    const { dates } = getMonthGridRange(new Date(year, month, 1), weekStartsOnMonday);
    dates.forEach((date) => {
      const key = formatDateKey(date);
      const dayCell = document.createElement("button");
      dayCell.type = "button";
      dayCell.className = "mini-day";
      if (date.getMonth() !== month) dayCell.classList.add("is-outside");
      if (selectedDate && key === formatDateKey(selectedDate)) dayCell.classList.add("is-selected");
      dayCell.textContent = String(date.getDate());
      if (eventsByDay.has(key)) {
        const dot = document.createElement("span");
        dot.className = "mini-dot";
        dayCell.appendChild(dot);
      }
      dayCell.addEventListener("click", () => onSelectDay(date));
      grid.appendChild(dayCell);
    });

    wrapper.appendChild(grid);
    container.appendChild(wrapper);
  }
}
