#!/bin/bash
# Publishes everything in one go: the website (GitHub Pages), the Firestore rules, and the Apps Script, including a
# new version of its web app. Run it from this folder: ./deploy.sh "What changed"
# One-time setup is in README.md under "Publishing changes".
set -euo pipefail
cd "$(dirname "$0")"
message="${1:-Update the Researcher Portal}"

for tool in git clasp firebase python3; do
  command -v "$tool" >/dev/null || { echo "Missing $tool. See \"Publishing changes\" in README.md for the one-time setup."; exit 1; }
done
[ -f apps-script/.clasp.json ] || { echo "Missing apps-script/.clasp.json. See \"Publishing changes\" in README.md."; exit 1; }

echo "1/4 Website: committing and pushing to GitHub…"
git add -A
git diff --cached --quiet || git commit -q -m "$message"
git push -q

echo "2/4 Firestore rules…"
firebase deploy --only firestore:rules --project lab-schedules --non-interactive

echo "3/4 Apps Script: Code.gs, the manifest, and the scheduling rules bundled in…"
build=apps-script/build
rm -rf "$build" && mkdir -p "$build"
cp apps-script/Code.gs apps-script/appsscript.json "$build"/
python3 -I -c 'import json, sys; print("const RULES_SOURCE = " + json.dumps(open(sys.argv[1]).read()) + ";")' scheduling.js > "$build/Rules.js" 2>/dev/null
(cd apps-script && clasp push --force)

echo "4/4 Apps Script web app: new version…"
deployment=$(cd apps-script && clasp list-deployments | grep -v '@HEAD' | grep -o 'AKfy[A-Za-z0-9_-]*' | head -1)
if [ -n "$deployment" ]; then
  (cd apps-script && clasp update-deployment "$deployment" --description "$message" >/dev/null)
  echo "Done. The site updates on GitHub Pages within a minute or two; the script and its web app are updated now."
else
  echo "Done, but no web app deployment was found to update. In Apps Script: Deploy → New deployment → Web app."
fi
