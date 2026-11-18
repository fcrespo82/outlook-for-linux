# Outlook for Linux

An unofficial Linux desktop client for Outlook on the web, forked from
[Teams for Linux](https://github.com/IsmaelMartinez/teams-for-linux).

This fork is an initial adaptation of the original Electron application. It keeps
the upstream architecture and build workflow, with changes focused on Outlook.
It is not affiliated with or endorsed by Microsoft.

## Current scope

- Outlook web mail with Microsoft account authentication.
- Outlook branding and a penguin icon derived from the upstream artwork.
- The original application and tray menus, adapted with Outlook actions.
- Inbox unread counts passed to the existing tray badge renderer.
- Back and forward buttons beside the Outlook name, outside the search field.
- The original **Choose App Icon** menu.

This is a work in progress. Tray appearance depends on the desktop environment.
Features inherited from Teams, including calls, presence, screen sharing, MQTT,
Graph integration and multiple account profiles, have not been validated for
Outlook. Their presence in the source or upstream documentation does not imply
Outlook support.

## Run from source

Clone this fork and switch to its Outlook branch:

```sh
git clone --branch outlook https://github.com/fcrespo82/outlook-for-linux.git
cd outlook-for-linux
npm start -- --user-data-dir="$HOME/.config/outlook-for-linux"
```

Node.js and npm are required. The original `prestart` script runs `npm ci` before
launching Electron. The explicit user data directory keeps the development
profile separate from other Electron applications. Reuse the same directory on
subsequent launches to keep your login and settings.

The default user-agent is Electron's native user-agent. Password sign-in was
available with this identity during testing; the inherited Chrome-only override
led to a blocked security-key/PIN sign-in flow with the tested work account.
Other authentication methods remain unverified.

## Build and checks

The fork uses the upstream scripts:

```sh
npm run pack        # Unpacked application
npm run dist:linux  # Linux packages
npm run lint
npm run test:unit
```

These commands are inherited build and validation entry points, not a guarantee
that every package format or inherited test currently passes for Outlook.
The Teams for Linux package repositories, Flathub and Snap Store listings
publish the upstream Teams application, not this Outlook fork.

## Upstream and development

The upstream source, history, credits and documentation remain available:

- [Teams for Linux repository](https://github.com/IsmaelMartinez/teams-for-linux)
- [Upstream documentation](https://ismaelmartinez.github.io/teams-for-linux/)
- [Local development documentation](docs-site/docs/development/contributing.md)
- [Configuration reference](docs-site/docs/configuration-generated.md)
- [Contributing](CONTRIBUTING.md) and [project history](HISTORY.md)

Upstream guides describe Teams behavior. Outlook-specific behavior must be
checked against this fork. Internal module and file names are kept where a
rename is unnecessary; references to upstream services and past releases are
not mechanically rewritten.

## License and artwork

[GPL-3.0-or-later](LICENSE.md). Original authorship and notices are retained.
The Outlook penguin artwork is adapted from the original Teams for Linux artwork.
Additional icons are from [Icon Duck](https://iconduck.com/sets/hugeicons-essential-free-icons)
(CC BY 4.0).
