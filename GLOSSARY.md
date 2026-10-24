# Glossary

The vocabulary this project uses in code, tests, ADRs and issues. Terms are
canonical: use them exactly, and do not substitute synonyms.

## Accounts and auth

**Account** — A named personal Microsoft credential plus the Chrome profile that
owns it: an id, a session file, a profile directory and a CDP port. Accounts exist
only in personal mode. An account is not a user, not a tenant and not a token.

**Active account** — The account a request runs on when none is named
explicitly. Resolved per request by precedence: an explicit id from the CLI, then
environment overrides, then the registry's active entry, then the builtin layout.

**Auth mode** — Which family of credential and protocol a run uses. Exactly two:
`msal` and `personal`, selected by `M365_AUTH`. The mode is a property of the
process, not of an account.

**`msal`** — The enterprise auth mode. Azure AD sign-in through `@azure/msal-node`
with PKCE and TOTP, talking to the Graph-backed Copilot endpoint. Named after the
library that implements it, because that is the value `AuthMode` carries.

**`personal`** — The consumer auth mode. A Microsoft 365 Personal, Family or
Outlook.com account, whose credential is captured from or harvested out of a
running Chrome, talking to the consumer chathub endpoint. There is no refresh
token exposed to the proxy and no headless path; a windowed Chrome is required.

**Credential** — A token plus the metadata needed to use it: its account, the
account's scenario, and an estimated expiry. The `/healthz` response reports
`credential` in exactly this sense.

**Capture** — Reading a credential out of a Chrome page over CDP. Spends one turn.

**Harvest** — Reading a credential out of the socket the chat page opens when a
message is sent. Attaches to a running Chrome and spends one turn; the normal
mint path.

**Mint** — Obtaining a fresh credential for an account by capture or harvest,
subject to a minimum interval and an hourly cap.

**Record** — The persisted form of a credential for one account: the session
file's contents.

**Probe** — A capability observation written into a record, used to pick a
personal default model without spending a turn.

## Protocol

**Turn** — One user message sent to M365 and the reply streamed back. The unit
that throttling, metering and budgets count.

**Turn context** — Everything a turn needs from outside the transport: token,
socket URL, headers, payload defaults and capabilities. The transport receives it;
it never asks which mode produced it.

**Frame** — One WebSocket message from M365, typed `1`, `2` or `3`.

**Chathub** — The consumer endpoint. Distinct from the enterprise endpoint in URL,
headers, payload and available capabilities.

**Disengaged** — An M365 conversation state that ends a conversation permanently.
A Disengaged conversation stays Disengaged; retrying on it is pointless.

**Handoff** — Starting a fresh M365 conversation because the current one is
nearly full.

**Fencing** — The Markdown-fence format this project uses to express tool calls to
a model that has no native tool-call support.

## Limits

**Turn budget** — The local accounting of turns used against a daily cap, with a
hold when the cap is near.

**Metering** — M365's own reporting of paid-scenario usage, which is a property of
the model's scenario, not of the account.

**Degradation backoff** — A per-account wait imposed after repeated upstream
failures, so a failing account is not hammered.

**Priority access** — Per-model daily and weekly allowances that apply only to
certain premium models.
