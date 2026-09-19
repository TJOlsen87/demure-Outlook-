# Demure Outlook

*A dated reference for a dated app.*

Concise new-mail announcements for **NVDA and classic Microsoft Outlook on Windows**.
Hear the account, sender, and subject, without the message-body preview or its URLs:

> New mail for address from sender. Subject.

## Download and install

1. [Download the latest Demure Outlook add-on](https://github.com/TJOlsen87/demure-Outlook-/releases/latest).
2. Under Assets, download the `.nvda-addon` file. The source-code ZIP is not the installer.
3. Open the downloaded add-on, confirm installation in NVDA, and restart NVDA.
4. Keep classic Outlook running. Demure connects automatically in the background,
   retrying every 15 seconds while Outlook's automation interface is unavailable.

Requires NVDA 2026.1 or later. Tested with NVDA 2026.2 and classic Outlook on Windows.
**New Outlook and Outlook on the web are not supported.**

## Review the last five notifications

Press **Ctrl+Alt+Shift+O** from any application. The list opens with the newest
notification selected. Use **Up/Down** to review and **Escape** or **Close** to return.

- Each entry contains the same concise account, sender, and subject announcement.
- Reviewing an entry does not open or mark the email as read.
- History is kept only in memory and resets when NVDA restarts or the add-on reloads.
- Only mail received while Demure is connected is included; it does not backfill old mail.
- The open list is a snapshot. Close and reopen it to include newly arrived notifications.
- If no notifications have arrived, NVDA announces that the history is empty.
- Change the shortcut under **NVDA Preferences > Input gestures > Demure Outlook**.

## Prevent duplicate announcements

Demure's announcements are independent of Outlook's Windows notifications.
If you hear both Demure and Outlook's verbose preview, disable Outlook's notifications
in **Windows Settings > System > Notifications > Outlook**. Demure's new-mail event
connection and five-item history continue to work independently.

You can instead experiment with disabling only **Show notification banners** while
leaving Notification Center enabled. That preserves Outlook's original entries,
including any preview text. This configuration has not been verified with this add-on.
Do not enable Outlook notifications merely to use Demure's history.

If Outlook produces an extra sound, disable its new-mail sound in Outlook's Mail
options and Windows notification settings.

## Sound

When Outlook is in the background, the first new-mail event after 20 minutes without
new mail plays the Windows email notification sound. Subsequent events reset that
quiet-period timer. No sound is played by Demure while Outlook has focus.

## Privacy and limitations

Demure reads account, sender, and subject metadata for new-mail events. It does not
scan your Inbox, read email bodies, inspect attachments, or send mail data anywhere.
Control characters are removed and each metadata value is limited to 300 characters.
History is not written to disk. Spoken metadata and the history command are suppressed
while Windows is locked or NVDA is in secure mode.

Classic Outlook must expose its automation interface before Demure can connect.
Mail delivered while NVDA is stopped or disconnected is not recovered. If notifications
stop after closing and reopening Outlook, restart NVDA to reconnect. Outlook configuration,
account delivery behavior, and other add-ons can affect notification delivery.

## Build and test

The build uses only Python's standard library. From the repository root:

```text
python -m unittest discover -s tests -v
python build.py
```

The installer is written to `dist/`. Automated tests simulate NVDA, Outlook events,
and dialog behavior; they do not replace live screen-reader testing.

## Support

[Open an issue](https://github.com/TJOlsen87/demure-Outlook-/issues) with your NVDA and
classic Outlook versions, steps to reproduce, and the observed result. Remove private
mail details before sharing diagnostics.

See [CHANGELOG.md](CHANGELOG.md) for release notes.

A TJOlsen + AI enabled production.

## License

Copyright (C) 2026 TJOlsen. Licensed under the GNU General Public License,
version 2 or (at your option) any later version. See [LICENSE](LICENSE).
