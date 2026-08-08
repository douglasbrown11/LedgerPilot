# Ledger Pilot — AI Project Training

Read this file first when helping with Ledger Pilot. It records the product intent, current implementation, accounting conventions, and known gaps so an AI can work safely without rediscovering the project. Whenever you make changes to this project that make this document inaccurate or when you make significant changes in general, edit this file so it remains as accurate file describing the project 

## Product purpose

Ledger Pilot is accounting software for ClubGG poker club operators.

- **Club Owners** upload the standard ClubGG club Excel export. Ledger Pilot calculates player settlements using P&L, rake/tips, rakeback agreements, downlines, backed-player arrangements, fees, prior tabs, and ownership assignments.
- **Agents** need ClubGG player data converted into settlements without manually opening every player. The intended product creates data-collection tasks specifying clubs, players, fields, and dates; a future collector will read ClubGG and place the results into the accounting dashboard.
- The broader goal is to reduce the time owners and agents spend doing accounting and communicating results to players.

ClubGG terminology in this code often uses **fee**, **tips**, **rake**, **tipback**, and **rakeback** for closely related concepts. Do not rename or merge these concepts without confirming the source column and intended business meaning.

## Current user experience

The application currently starts with a role-selection splash:

- **Club Owner** opens the existing dashboard with `Fish Tank`, `My Clubs`, and `Tabs`.
- **Agent** opens a separate dashboard whose only navigation option is `Data Tasks`.
- `Choose role` returns to the splash.

Important: the current splash is text-only. The requested original Ledger Pilot lightning-logo splash asset has not been found in the repository or its Git history. Do not invent or replace that brand asset without user direction.

## Technology and commands

- React 19
- Vite 8
- JavaScript/JSX (no TypeScript)
- `xlsx` reads ClubGG workbooks.
- `exceljs` creates formatted exports.
- Browser `localStorage` is the only persistence layer.
- Firebase project `ledgerpilot-app` now has Email/Password Authentication and a production Firestore database in `australia-southeast2` (Melbourne), but the application is not connected to them yet.
- There is currently no application-level authentication, Firestore data sync, ClubGG integration, OCR pipeline, messaging service, or automated collector.

Run from `/Users/dougie/Desktop/LedgerPilot`:

```bash
npm install
npm run dev
npm run build
npm run lint
```

Always run `npm run build` after changing application code. Use lint when practical, but distinguish pre-existing warnings from new problems.

## Source map

- `src/main.jsx` — real entry point; imports `AppV10.jsx`.
- `src/AppV10.jsx` — the entire working product: UI, storage, Excel parsing/exporting, settlement engines, role splash, Data Tasks, and Tabs.
- `src/App.jsx`, `src/App.css`, `src/assets/hero.png`, and the React/Vite SVGs — unused Vite starter content.
- `src/index.css` — legacy global Vite styling. Most real product styling is inline in `AppV10.jsx` using palette variables.
- `README.md` — currently generic Vite documentation and not reliable product documentation.

`AppV10.jsx` is very large. Before editing it, locate the relevant named function with `rg` and inspect its surrounding code. Avoid broad rewrites that could change unrelated accounting behavior.

## Main code areas

Approximate locations can drift as code changes; search by function or constant name.

- `parseWorkbook` — reads a ClubGG workbook.
- `buildModel` — Club Owner/Fish Tank settlement engine.
- `RoleSplash` and `App` — role routing and main navigation.
- `AgentClubs`, `settleLine`, and `computeAgent` — the existing manual multi-club/agent accounting system shown under the owner's `My Clubs` area.
- `DataTasks` — Agent task generator; currently queues instructions only.
- `TabsLedger` — running balances and manual ledger entries.
- `download*` functions — formatted Excel reports.

## Club Owner data flow

1. User uploads `.xlsx` or `.xls`.
2. `parseWorkbook` selects a sheet containing `club overview`, otherwise the first sheet.
3. It finds a row whose first cell is `No.` and reads the period from a leading `Period` row.
4. It currently relies on fixed column positions:
   - super-agent ID/name: columns 1/2
   - agent ID/name: columns 3/4
   - role: column 6
   - member ID/name: columns 7/8
   - hands: column 10
   - fee/rake: column 11
   - P&L: column 19
5. Rows without a numeric row number, member ID, name, or activity are ignored.
6. `buildModel` classifies each player as an owner account, house-backed player, super-agent downline, or individual.
7. The UI displays settlements, reports, reconciliation, deal setup, and Excel downloads.

The parser is fragile if ClubGG changes its export layout. Preserve sample exports for tests before changing column mapping.

## Accounting conventions

### Fish Tank / Club Owner

- Normal player tipback: `fee × rakeback percentage`.
- Normal player settlement: `P&L + tipback - optional action cut`.
- Owner account position: `P&L + 100% fee back`.
- Action-backed player: calculate `gross = P&L + rakeback`, allocate the configured backing percentage to the backer, and settle the remainder with the player.
- Makeup-backed player: credit configured rakeback, offset the entering makeup balance, pay the player their profit percentage only on excess above makeup, and carry remaining makeup forward.
- Club profit: club rake/fee revenue minus all tipbacks and owner fee-back, then minus configured fees.
- Finalizing a period snapshots makeup balances so the same week is not applied twice.

### My Clubs / manual agent engine

Input values are currently `P&L` and `Tips` in club currency.

- `tipback = tips × player TB%`
- `net = P&L + tipback`
- the action/transaction percentage can apply to net, P&L, or gross (`P&L + tips`)
- `settlement = (net - action cut) × currency conversion`
- `margin = what the club pays the agent - what the agent settles with the player`

Clubs can use a custom formula with approved variables defined by `FORMULA_VARS`. This uses `new Function`; treat formula changes as sensitive accounting and security work.

### Tabs

The universal sign convention is:

- **Positive** = they owe you.
- **Negative** = you owe them.

Weekly settlements can be pushed into Tabs once per source/week. Tabs also accepts manual adjustments and optional P&L categories.

Do not change a sign convention without tracing every producer, report, and export.

## Browser storage

All data is local to the browser profile and origin. Clearing browser storage or changing the served origin/port may make saved data appear missing.

| Key | Purpose |
| --- | --- |
| `fishtank-config-v4` | Owner deals, theme, assignments, backed players, fees, finalized periods |
| `fishtank-lastweek-v4` | Last uploaded player rows, period, and weekly adjustments |
| `agentclubs-v3` | Clubs, players, agent deals, weeks, aliases, and accounts |
| `clubgg-data-tasks-v1` | Queued Agent data tasks |
| `tabs-v1` | Counterparties, ledger entries, and pushed-week markers |

Older `agentclubs-v1` and `agentclubs-v2` data is migrated into v3 on load.

There is no automatic backup or audit log. Changes affecting storage schemas must include backward-compatible migration and must never silently discard saved user data.

## Firebase foundation

- Project ID: `ledgerpilot-app`
- Email/Password Authentication is enabled.
- The default Firestore database uses Standard edition in Melbourne.
- `.firebaserc`, `firebase.json`, and `firestore.rules` contain the local deployment configuration.
- Current rules permit an authenticated user to access only `users/{their uid}` and its descendants; everything else is denied.
- No React Firebase SDK, sign-in page, user document creation, or migration from `localStorage` has been implemented yet.

## Agent Data Tasks: current state

`DataTasks` lets an agent select:

- one or more configured clubs;
- players belonging to those clubs;
- date presets or a custom date range;
- Hands, Rake, and/or P&L;
- an optional task name and note.

It validates the selection and saves a task with status `queued`. It does **not** run the task, open ClubGG, capture screenshots, use OCR, scrape data, or write results into a week.

The intended future ClubGG navigation described by the user is:

1. Start from the ClubGG dashboard.
2. Select the requested club.
3. Open the four-dot menu at bottom right.
4. Open `Members`.
5. Select the requested player.
6. Open `Custom`.
7. Select the requested period/history.
8. Read the player's rake and P&L data.
9. Insert those values into the appropriate Ledger Pilot club/player/week.

OCR means **Optical Character Recognition**: reading visible text and numbers from screenshots or screen regions. It may be one component of a collector, but none exists yet.

Before implementing automated ClubGG access, confirm the permitted method and current ClubGG terms. Prefer an official export/API if available. Keep source evidence (task, timestamps, raw captured values or source artifact, parsed values, and errors) so accounting outputs can be audited. Never automate gameplay or wagering.

## Likely next architecture

Keep collection separate from settlement math:

```text
Data Task
  -> Collector adapter (approved ClubGG access method)
  -> Raw evidence + normalized result
  -> Validation/error state
  -> Agent week entries
  -> Existing computeAgent settlement engine
  -> Dashboard, Tabs, reports, and future player messages
```

A task will need lifecycle states such as `queued`, `running`, `completed`, `partial`, and `failed`, plus per-player results and errors. Task execution should be idempotent so rerunning it does not duplicate week data.

Player communication should be a later service driven by finalized settlement records. It should use templates, delivery status, retries, and a message history. Do not send financial messages from draft or partially collected data.

## Working rules for an AI

1. Inspect the current code and `git status` before editing. Preserve unrelated user changes and generated/local files.
2. Confirm which portal is in scope: Club Owner, Agent Data Tasks, or the legacy `My Clubs` accounting engine.
3. Treat accounting formulas, signs, stored schemas, and Excel column positions as high-risk behavior.
4. Make focused changes; do not refactor the monolith while delivering a small UI request.
5. Never claim collection, OCR, messaging, auditing, authentication, or cloud persistence is implemented when it is not.
6. Do not fabricate the missing lightning logo. Ask for the original asset or exact source file.
7. Add tests or fixture-based checks when changing parsing or calculations. At minimum, verify representative positive, negative, zero, rakeback, action, makeup, and currency-conversion cases.
8. Run a production build and report what was actually verified.
9. Update this file whenever architecture, storage keys, core formulas, entry points, or major workflows change.

## Immediate product gaps

- Restore the intended lightning-logo splash after the original asset is supplied.
- Connect Agent tasks to a compliant collector.
- Normalize collected rake/P&L into `agentclubs-v3` week entries.
- Add task results, failures, reruns, provenance, and audit history.
- Add durable backend storage, accounts, and backups before production use.
- Add automated calculation and workbook-parser tests.
- Add player statement generation and controlled communication delivery.
- Break `AppV10.jsx` into focused modules once behavior is protected by tests.
