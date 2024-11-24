const CALENDAR = "primary";
const TIME_ZONE = "America/Los_Angeles";
//  America/Los_Angeles
//  America/Denver
//  America/Chicago
//  Europe/London
//  Europe/Berlin
const DAYS_IN_ADVANCE = 28;
//high numbered days in advance cause significant loading time slow down
const TIMESLOT_DURATION = 30;

type ConstraintEffect = "available" | "unavailable";

type ConstraintDayOfWeek = {
  type: "day-of-week";
  effect: ConstraintEffect;
  days: number[];
};
type ConstraintOverlappingEvent = {
  type: "overlapping-event";
  effect: ConstraintEffect;
};
type ConstraintTimeOfDay = {
  type: "time-of-day";
  effect: ConstraintEffect;
  hourStart: number;
  hourEnd: number;
};
type ConstraintMeetingTimeInDay = {
  type: "meeting-time-in-day";
  effect: ConstraintEffect;
  maximumHours: number;
};

type Constraint =
  | ConstraintDayOfWeek
  | ConstraintOverlappingEvent
  | ConstraintTimeOfDay
  | ConstraintMeetingTimeInDay;

const constraints: Constraint[] = [
  { type: "day-of-week", effect: "unavailable", days: [0, 6] },
  { type: "overlapping-event", effect: "unavailable" },
  { type: "time-of-day", effect: "unavailable", hourStart: 18, hourEnd: 9.5 },
  { type: "meeting-time-in-day", effect: "unavailable", maximumHours: 5 },
];

const TSDURMS = TIMESLOT_DURATION * 60000;

function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  return HtmlService.createHtmlOutputFromFile("dist/index")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getHoursWithDecimalMinutes(date: Date) {
  return date.getHours() + date.getMinutes() / 60;
}

function fetchAvailability(): {
  timeslots: string[];
  durationMinutes: number;
} {
  const nearestTimeslot = new Date(
    Math.floor(new Date().getTime() / TSDURMS) * TSDURMS,
  );
  const calendarId = CALENDAR;
  const now = nearestTimeslot;
  const end = new Date(
    Date.UTC(
      now.getUTCFullYear(),
      now.getUTCMonth(),
      now.getUTCDate() + DAYS_IN_ADVANCE,
    ),
  );

  const response = Calendar.Events!.list(calendarId, {
    timeMin: now.toISOString(),
    timeMax: end.toISOString(),
    eventTypes: ["default", "focusTime", "fromGmail", "outOfOffice"],
    singleEvents: true,
  }) as {
    items: {
      start: { date?: string; dateTime?: string; timeZone?: string };
      end: { date?: string; dateTime?: string; timeZone?: string };
      transparency?: "opaque" | "transparent";
    }[];
    timeZone: string;
  };

  const events = response.items.map(({ start, end, transparency }) => ({
    start: start.dateTime
      ? new Date(start.dateTime)
      : Utilities.parseDate(
          start.date!,
          start.timeZone || response.timeZone,
          "yyyy-MM-dd",
        ),
    end: end.dateTime
      ? new Date(end.dateTime)
      : Utilities.parseDate(
          end.date!,
          end.timeZone || response.timeZone,
          "yyyy-MM-dd",
        ),
    transparency: transparency,
  }));
  //get all timeslots between now and end date
  const timeslots = [];
  for (
    let t = nearestTimeslot.getTime();
    t + TSDURMS <= end.getTime();
    t += TSDURMS
  ) {
    const start = new Date(t);
    const end = new Date(t + TSDURMS);
    const startTZ = new Date(
      Utilities.formatDate(start, TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss"),
    );

    let desiredEffect = "available";

    for (const constraint of constraints) {
      let constraintMatchesTimeslot = false;
      switch (constraint.type) {
        case "day-of-week":
          constraintMatchesTimeslot =
            constraint.days.indexOf(startTZ.getDay()) >= 0;
          break;
        case "overlapping-event":
          constraintMatchesTimeslot = events.some(
            (event) =>
              event.start < end &&
              event.end > start &&
              event.transparency != "transparent",
          );
          break;
        case "time-of-day":
          if (constraint.hourStart <= constraint.hourEnd) {
            constraintMatchesTimeslot =
              getHoursWithDecimalMinutes(startTZ) >= constraint.hourStart &&
              getHoursWithDecimalMinutes(startTZ) < constraint.hourEnd;
          } else {
            constraintMatchesTimeslot = !(
              getHoursWithDecimalMinutes(startTZ) >= constraint.hourEnd &&
              getHoursWithDecimalMinutes(startTZ) < constraint.hourStart
            );
          }
          break;
      }

      if (constraintMatchesTimeslot) {
        desiredEffect = constraint.effect;
        break;
      }
    }

    if (desiredEffect == "available") {
      timeslots.push(start.toISOString());
    }
  }
  return { timeslots, durationMinutes: TIMESLOT_DURATION };
}

function bookTimeslot(
  timeslot: string,
  name: string,
  email: string,
  phone: string,
  note: string,
): string {
  Logger.log(`Booking timeslot: ${timeslot} for ${name}`);
  const calendarId = CALENDAR;
  const startTime = new Date(timeslot);
  if (isNaN(startTime.getTime())) {
    throw new Error("Invalid start time");
  }
  const endTime = new Date(startTime.getTime());
  endTime.setUTCMinutes(startTime.getUTCMinutes() + TIMESLOT_DURATION);

  Logger.log(`Timeslot start: ${startTime}, end: ${endTime}`);

  try {
    const possibleEvents = Calendar.Freebusy!.query({
      timeMin: startTime.toISOString(),
      timeMax: endTime.toISOString(),
      items: [{ id: calendarId }],
    });

    const busy = (possibleEvents as any).calendars[calendarId].busy;

    if (
      busy.some((event: { start: Date; end: Date }) => {
        const eventStart = new Date(event.start.toString());
        const eventEnd = new Date(event.end.toString());
        return eventStart <= endTime && eventEnd >= startTime;
      })
    ) {
      throw new Error("Timeslot not available");
    }

    const event = CalendarApp.getCalendarById(calendarId).createEvent(
      `Appointment with ${name}`,
      startTime,
      endTime,
      {
        description: `Phone: ${phone}\nNote: ${note}`,
        guests: email,
        sendInvites: true,
        status: "confirmed",
      },
    );
    Logger.log(`Event created: ${event.getId()}`);
    return `Timeslot booked successfully`;
  } catch (e) {
    const error = e as Error;
    Logger.log(`Failed to create event: ${error.message}`);
    throw new Error(`Failed to create event: ${error.message}`);
  }
}
