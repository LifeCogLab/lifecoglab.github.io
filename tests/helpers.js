// Loaded before each test file. check(label, actual, expected) prints pass or FAIL.
function check(label, actual, expected) {
  const same = JSON.stringify(actual) === JSON.stringify(expected);
  print(same ? `pass ${label}` : `FAIL ${label}\n  expected ${JSON.stringify(expected)}\n  got      ${JSON.stringify(actual)}`);
}

// The Apps Script's code, with the site's scheduling rules, as the script sees them.
function loadScript() {
  const RULES_SOURCE = readFile("scheduling.js");
  const script = new Function("RULES_SOURCE", `${readFile("apps-script/Code.gs")}\nreturn { loadRules, buildVisitEvents, buildConfirmedEvents, unbookedSoon, unbookedEmail, staffEmail, bookingsEmail, isoOf };`)(RULES_SOURCE);
  return { ...script, rules: script.loadRules() };
}
