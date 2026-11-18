/**
 * Launch arguments supported by the initial Outlook adaptation.
 * HTTPS Outlook URLs are accepted without the inherited Teams conversion.
 */
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const {
  PROTOCOL_SCHEMES,
  stripWebPrefix,
  resolveLaunchUrl,
} = require("../../app/mainAppWindow/protocolLink");
const options = require("../../app/config/options");
const pkg = require("../../package.json");

const config = {
  msOutlookProtocols: options.msOutlookProtocols.default,
  meetupJoinRegEx: options.meetupJoinRegEx.default,
  url: options.url.default,
};

test("opens Outlook HTTPS mail, calendar and OWA links unchanged", () => {
  for (const host of ["outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft", "outlook.live.com"]) {
    for (const route of ["mail/", "calendar/view/month", "owa/?path=/mail"]) {
      const url = `https://${host}/${route}`;
      assert.strictEqual(resolveLaunchUrl(["electron", "./app", url], config), url);
    }
  }
});

test("rejects unsupported schemes, hosts and paths", () => {
  for (const url of [
    "http://outlook.office.com/mail/",
    "https://outlook.office.com.evil.example/mail/",
    "https://evil.example/outlook.office.com/mail/",
    "https://outlook.office.com/mailbox/",
    "https://outlook.office.com/other/",
    "msteams:/meet/abc123",
    "web+msteams:/meet/abc123",
    "outlook:/mail/",
    "web+outlook:/mail/",
    "mailto:someone@example.com",
  ]) {
    assert.strictEqual(resolveLaunchUrl([url], config), null, url);
  }
});

test("opens the first supported link and ignores other arguments", () => {
  const first = "https://outlook.office.com/mail/";
  assert.strictEqual(resolveLaunchUrl(["--ozone-platform=wayland", first, "https://outlook.cloud.microsoft/calendar/"], config), first);
  assert.strictEqual(resolveLaunchUrl(["electron", "./app"], config), null);
});

test("processArgs hands launch arguments to resolveLaunchUrl", () => {
  const source = fs.readFileSync(path.join(__dirname, "../../app/mainAppWindow/index.js"), "utf8");
  assert.match(source, /resolveLaunchUrl\(args, config\)/);
});

test("stripWebPrefix normalizes only the registered Outlook scheme", () => {
  assert.strictEqual(stripWebPrefix("web+outlook:/mail/"), "outlook:/mail/");
  assert.strictEqual(stripWebPrefix("WEB+OUTLOOK:/mail/"), "OUTLOOK:/mail/");
  for (const value of ["web+msteams:/meet/abc123", "https://outlook.office.com/mail/", "", undefined, null]) {
    assert.strictEqual(stripWebPrefix(value), value);
  }
});

test("packaged protocols match the schemes registered at runtime", () => {
  assert.deepStrictEqual(pkg.build.protocols.schemes, PROTOCOL_SCHEMES);
  assert.deepStrictEqual(PROTOCOL_SCHEMES, ["outlook", "web+outlook"]);
});
