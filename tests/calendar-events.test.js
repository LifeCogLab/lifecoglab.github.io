const { rules, buildVisitEvents, buildConfirmedEvents, unbookedSoon, unbookedEmail, bookingsEmail, isoOf } = loadScript();
const data = { members: [{ email: "a@x", name: "Ann" }, { email: "b@x", name: "Ben" }, { email: "c@x", name: "Cat" }], availability: [], timeOff: [] };
const confirmed = (scan) => ({ id: "v", date: "2026-12-01", defaultFor: "2026-12-01", lab: { start: "09:00", end: "11:00" }, scan: { start: "11:00", end: "13:00", ...scan }, booking: "booked", confirmed: true });
const dniRes = (scan, staffed) => {
  const appointment = confirmed(scan);
  const doc = { date: appointment.date, visitKey: rules.visitKey(appointment), scan: staffed, lab: {} };
  return buildConfirmedEvents(rules, data, [appointment], [], [doc], new Date(2026, 10, 21)).find((event) => event.key.endsWith(":scan"));
};
check("a separate wanding/taping person is the Third scanner", dniRes({}, { lead: "a@x", support: "b@x", wanding: "c@x" }).description.split("\n"), ["Lead scanner: Ann", "Second scanner: Ben", "Third scanner: Cat"]);
check("a scanner who also wands isn't listed twice", dniRes({}, { lead: "a@x", support: "b@x", wanding: "a@x" }).title, "DNI Res (Ann, Ben)");
check("wanding needed but nobody doing it", dniRes({}, { lead: "a@x", support: "b@x" }).description.split("\n").pop(), "Third scanner needed for female participant");

const today = new Date(2026, 10, 21, 9, 0);
const slots = [
  { id: "c", date: "2026-11-24", defaultFor: "2026-11-24", lab: { start: "09:00", end: "11:00" }, scan: { start: "11:00", end: "13:00" }, booking: "rebooking" },
  { id: "d", date: "2026-11-23", defaultFor: "2026-11-23", lab: { start: "09:00", end: "11:00" }, scan: { start: "11:00", end: "13:00" }, booking: "booked", confirmed: true },
];
const ready = { ...rules, staffingState: () => "ready" };
const titles = Object.fromEntries(buildVisitEvents(ready, data, slots, [], today, []).map((event) => [isoOf(event.start).slice(5), event.title]));
check("confirmed visits are left off LCL Avail Visits", titles["11-23"], undefined);
check("a cancelled booking 3 days out", titles["11-24"], "⚠️⚠️⚠️ [BOOK NOW] [CANCELED APT] Rebooking: Participant Visit");
check("an open slot within a week", titles["11-25"], "⚠️⚠️ [BOOK ASAP] ✅ Open: Participant Visit");
check("an open slot within 2 weeks", titles["11-30"], "⚠️ [BOOK ASAP] ✅ Open: Participant Visit");
check("the 3-days-out email subject", unbookedEmail(rules, unbookedSoon(rules, slots, [], today)).subject, "CANCEL/Rebook DNI Res on Tue 11/24");

const booking = bookingsEmail(rules, [{ date: "2026-12-01", appointment: confirmed({}), by: "Jordan", screener: "S-123", note: "Thin veins" }]);
check("the new booking email has the Screener ID", booking.lines.some((line) => line.includes("S-123")), true);
check("and the booker's note", booking.lines.includes("  Note: Thin veins"), true);
