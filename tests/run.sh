#!/bin/bash
# Runs every tests/*.test.js with macOS's built-in JavaScript engine (no installs needed): ./tests/run.sh
cd "$(dirname "$0")/.."
jsc=/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc
failed=0
for test in tests/*.test.js; do
  if output=$("$jsc" tests/helpers.js "$test" 2>&1) && ! grep -q "^FAIL" <<<"$output"; then
    echo "ok   $test ($(grep -c '^pass' <<<"$output") checks)"
  else
    echo "FAIL $test"; grep -v '^pass' <<<"$output" | sed 's/^/     /'; failed=1
  fi
done
exit $failed
