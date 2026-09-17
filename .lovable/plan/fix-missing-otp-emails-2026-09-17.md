# Fix missing OTP emails

## Changes
- Make Inbox retrieve recent messages across the whole mailbox, including Inbox, Junk, and other folders.
- Merge and de-duplicate results, keep newest messages first, and stop silently hiding partial Microsoft fetch failures.
- Keep the existing 10-second refresh and OTP copy experience unchanged.

## Verification
- Confirm the app builds successfully and Inbox requests the broader Microsoft mail endpoint.
- Clarify when Microsoft reports that no message has actually reached the mailbox.
