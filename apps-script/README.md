# DecisionCAMP 2027 registration backend

This Google Apps Script is bound to the private **DecisionCAMP 2027 Registrations** Google Sheet.

## Setup

1. Create a Google Sheet named `DecisionCAMP 2027 Registrations`.
2. In the Sheet, open **Extensions → Apps Script**.
3. Replace the editor contents with `Code.gs` from this directory.
4. In **Project Settings**, enable the `appsscript.json` manifest and replace it with the manifest in this directory.
5. Run `setupRegistrationSystem` once and approve access to the Sheet, email sending, and the scheduled reminder trigger.
6. Choose **Deploy → New deployment → Web app**.
7. Set **Execute as** to the deploying account and **Who has access** to **Anyone**.
8. Copy the `/exec` URL into `registrationEndpoint` in `src/pages/index.astro`.

## Behavior

- Creates a `Registrations` worksheet with protected headers.
- Inserts or updates registrants by normalized email address.
- Sends an immediate confirmation with three calendar attachments for September 14–16, 2027, 9:00 AM–3:00 PM ET.
- Schedules a September 10 reminder. If the account's daily email quota is insufficient, remaining reminders are retried the next day.
- Stops reminder processing after September 17, 2027.

The web app endpoint is public by design, but the Sheet remains private. The handler validates and limits inputs, uses a honeypot field, and neutralizes spreadsheet-formula prefixes before storage.
