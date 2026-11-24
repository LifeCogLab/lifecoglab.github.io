load("scheduling.js");
const date = (iso) => iso;
const visit = (scanEnd = "12:30", scanStart = "11:00") => ({ date: "2026-12-02", booking: "booked", confirmed: true, lab: { start: "09:00", end: "11:00" }, scan: { start: scanStart, end: scanEnd, wandingRequired: true } });
const full = { lab: { dryLab: "dan" }, scan: { lead: "dan", support: "jo" }, managers: ["dan"] };
const managing = { lab: { dryLab: "amy" }, scan: { lead: "jo", support: "kim" }, managers: ["dan"] };
const entry = (appointment, doc) => ({ v: visitStaffEntry(appointment, doc, 1) });
const body = (changes, email) => staffEmailContent("Dan", changes[email], date).lines[2];

let changes = staffChanges({}, entry(visit(), full), "2026-11-24", () => true);
check("assigned lists each role once, with times and the short scan folded in", body(changes, "dan"),
  "You've been staffed for the visit on **==2026-12-02 (9:00 AM – 1:00 PM)==**. **Your assignment:** Dry Lab (9:00-11:00), Scanner (11:00-12:30 ⚠️ [SHORT SCAN]), & Lab Manager (Whole Visit).");

changes = staffChanges(entry(visit(), full), entry(visit(), managing), "2026-11-24", () => true);
check("a changed role says the assignment changed, not the times", body(changes, "dan"),
  "Your assignment has changed for the visit on **==2026-12-02 (9:00 AM – 1:00 PM)==**. **Before:** Dry Lab, Scanner, & Lab Manager. **Now:** Lab Manager (Whole Visit).");
check("no short scan notice once someone is off the scan", /SHORT SCAN/.test(body(changes, "dan")), false);
check("swapping lead and second scanner emails nobody", changes.jo, undefined);

changes = staffChanges(entry(visit(), full), entry(visit("13:00", "11:30"), full), "2026-11-24", () => true);
check("same roles at new times says the times changed", body(changes, "dan").startsWith("Your times have changed"), true);
check("new times flag an unusual start", /Scanner \(11:30-1:00 🕒 \[SEE TIME\] ⚠️ \[SHORT SCAN\]\)/.test(body(changes, "dan")), true);

changes = staffChanges(entry(visit(), full), {}, "2026-11-24", () => false);
check("a cancelled visit", body(changes, "dan"), "The visit on **==2026-12-02 (9:00 AM – 1:00 PM)==** has been **cancelled**. You were: Dry Lab, Scanner, & Lab Manager.");
changes = staffChanges(entry(visit(), full), {}, "2026-11-24", () => true);
check("taken off a visit", body(changes, "dan"), "You're no longer staffed for the visit on **==2026-12-02 (9:00 AM – 1:00 PM)==**. You were: Dry Lab, Scanner, & Lab Manager.");

const custom = staffEmailContent("Dan", [{ ...changes.dan[0], custom: "Line one\nLine two" }], date).lines;
check("a lab manager's own wording replaces the paragraph", custom.slice(2, 4), ["Line one", "Line two"]);

const reminder = staffEmailContent("Dan", [{ type: "reminder", key: "v", visit: visitStaffEntry(visit(), full, 1), roles: visitStaffEntry(visit(), full, 1).people.dan }], date);
check("a re-sent email reminds them of their assignment", reminder.lines[2], "A reminder that you're staffed for the visit on **==2026-12-02 (9:00 AM – 1:00 PM)==**. **Your assignment:** Dry Lab (9:00-11:00), Scanner (11:00-12:30 ⚠️ [SHORT SCAN]), & Lab Manager (Whole Visit).");
check("with its own subject", reminder.subject, "Study Visit: reminder that you're staffed for 2026-12-02");
