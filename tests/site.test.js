// Small pieces of app.js that don't need a page.
const source = readFile("app.js");
eval(source.match(/function formatPhone[\s\S]*?\n}\n/)[0]);
check("a 10-digit phone number", formatPhone("213-555-1234"), "(213) 555-1234");
check("with the country code", formatPhone("+1 (213) 555 1234"), "(213) 555-1234");
check("an international number stays as typed", formatPhone("+44 20 7946 0958"), "+44 20 7946 0958");
