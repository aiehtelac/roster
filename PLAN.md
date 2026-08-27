# Roster App — Google Sheets / Auth / Workflow Plan

Planning doc from the 2026-08-27 design discussion. Step 1 (auth) implemented; steps 2–6 not started.

## Problem

Current workflow is 8 steps across 3 contexts (Sheets browser tab, local filesystem, Streamlit app):

1. Admin sets up the Sheet — *necessary*
2. Review submissions / free text in Sheets — *necessary, stays in Sheets*
3. Download CSV — *friction*
4. Upload CSV — *friction*
5. Upload previous month's xlsx — *friction, error-prone (wrong month)*
6. Upload Excel template — *friction*
7. Select config — *repetitive, same values re-entered each cycle*
8. Run — *necessary*

Goal: delete steps 3–6, reduce 7. Context-switching between three places is the main pain,
more than the raw step count.

Cadence matters: 3 people, one roster each, once every ~3 months. Nobody builds muscle memory,
so every session is effectively relearned from scratch. Smoothness matters *more* here than for
a daily tool.

## Target workflow

1. Log in → see only the rosters your role allows
2. Pick roster + month → staff data loads from that roster's spreadsheet automatically
3. Review/vet in Google Sheets (unchanged — see "Validation" below)
4. Import & lock → app snapshots the vetted sheet, solves against the snapshot
5. Run — template and previous month auto-resolved from GitHub
6. Download + approve & archive → becomes next month's previous-month input

## Decisions made

### Stay on Streamlit + Community Cloud
3 users, quarterly use, solves complete reliably at the 300s limit. Nowhere near platform limits.
Rejected: Dash/NiceGUI/Reflex (lateral moves, more code, same constraints); FastAPI+React
(disproportionate — the config UI alone would balloon ~10x and add a second language);
Apps Script (can't run OR-Tools).

Back pocket, only if solve time ever hurts: split the solver into a background job and let
Streamlit poll. Don't build until it does.

Known quirk: with quarterly use the app is always asleep and needs a ~30s cold wake. Tell users
so it doesn't read as broken.

### Auth: `streamlit-authenticator` (v0.4.2)
- Native `roles` support maps directly onto HO/MO/REG gating; `st.session_state['roles']` after login
- **No 2FA** — the library's 2FA routes OTP emails through a third-party service
  (streamlitauthenticator.com); not acceptable for staff data here
- Credentials sourced from a **separate, tightly-shared `Admins` spreadsheet**, not the roster data.
  It holds password hashes, so it must not sit alongside data that roster admins can open.
- Constructor takes a dict (not just a YAML path) — that's the hook for building credentials from
  the Sheet at runtime, so the allowlist is editable without a redeploy
- Set `auto_hash=False`, store bcrypt hashes generated via `stauth.Hasher.hash()`. With
  `auto_hash=True` the library can't persist the hash back to a Sheet, so plaintext would sit there.
- Accept non-persistent `failed_login_attempts` / `logged_in` (they'd need a Sheets write per failed
  login to survive restarts; not worth it)

`Admins` tab: `email | first_name | last_name | password_hash | roles`

### One spreadsheet + one Excel template per roster type
Registry in `st.secrets` or a repo JSON:

    HO  → {spreadsheet_id: ..., template: templates/HO.xlsx,  archive: results/HO/}
    MO  → {spreadsheet_id: ..., template: templates/MO.xlsx,  archive: results/MO/}
    REG → {spreadsheet_id: ..., template: templates/REG.xlsx, archive: results/REG/}

Also gives real least privilege: the service account is shared per-spreadsheet, so access is
enforced at the data layer, not just hidden in the UI.

### Sheets I/O
`gspread` + Google service account, JSON key in `st.secrets`. Service account shared as Editor on
each roster spreadsheet + the Admins spreadsheet.

### Templates and archiving via GitHub
Fine-grained PAT in `st.secrets`, scoped to this repo only. Templates in `templates/`.
Generated rosters committed to `results/<roster>/<YYYY-MM>.xlsx`.

**Archiving is an explicit "approve & archive" action, never automatic on solve** — because the
archived output becomes next month's input, a bad run archived by mistake silently poisons the
next cycle.

### Validation: stays in Google Sheets — SKIPPED as a build item
Considered and dropped. The sheet uses dropdown/fixed-form input, so token-level typos aren't a
real risk. The remaining work is reading free text and judging legitimacy of leave/blocks —
judgment, not something the machine can do. Sheets is the better tool for it anyway (freeze panes,
filters, comments to query submitters, revision history, concurrent viewers). Rebuilding a
spreadsheet inside `st.data_editor` would be the classic trap.

In-app editing is therefore limited to **small last-mile corrections only** (e.g. deleting an
invalid request), not bulk review.

Optional ideas that were discussed and are NOT in scope — revisit only if wanted later:
pre-solve feasibility triage (supply vs demand per day, to avoid 300s waits on infeasible input),
a per-day availability pressure map, per-person load summary, cross-month staff-list diff, and
surfacing `_print_soft_violations` output as a table instead of burying it in the solver log.

### Move the roster-month Google Form into Streamlit
A Form is append-only, so editing a month's dates writes a *new* row and "is March open?" becomes
"find the latest response and hope." Replace with a mutable config tab, one row per month:

    month | start_date | end_date | status | updated_by | updated_at

`updated_by` comes free from the login, which the Form can't reliably give.

### "Reset calendar to default" button
Two separable operations, usually done together at cycle start:
- **Rebuild** — regenerate date column headers from the month's start/end dates (structural)
- **Clear** — wipe submitted leave/blocks/requests (data)

Layout makes this clean: first `date_col_start` columns are metadata, everything right of that is
the date grid, headers on row 2.

Destructive — needs all three guards:
1. **Snapshot to a timestamped archive tab before touching anything** (non-negotiable; makes it reversible)
2. **Type the month name to arm the button** (not a bare click)
3. **Refuse or warn loudly if the month status is `open`** — the config tab gives this for free

## Security

The risk profile changes materially with this plan. Today the app holds no data — you bring a CSV,
it computes, you leave. After Sheets integration it holds a service-account key with standing
access to all three rosters' personnel data, including leave requests and free-text reasons.
**Community Cloud apps are on a public URL by default.**

Therefore: **restrict the app to invited viewers in Community Cloud settings before any service
account key lands in secrets.** That's the interim control; `streamlit-authenticator` adds
per-roster separation on top. Both — this is staff personal data.

This is why auth is step 1 rather than last.

## Build order

1. **Auth** — `streamlit-authenticator` + `Admins` spreadsheet + role gating. **CODE DONE**
   (`auth.py`, wired into `RosterApp.py`; hashes via `tools/hash_password.py`; secrets shape in
   `.streamlit/secrets.toml.example`). Still to do by hand, before any key lands in secrets:
   turn on Community Cloud viewer restriction, create the service account, create the `Admins`
   spreadsheet and share it with the service account as **Viewer**, fill in secrets.
2. **Per-roster registry + Sheets read** — replaces CSV download/upload. Biggest single win.
3. **GitHub template fetch + approve-and-archive** — replaces the other two uploads,
   auto-resolves previous month.
4. **Month config UI** — replaces the Google Form.
5. **Reset calendar button.**
6. **Saved per-roster config defaults** — turns "select config" into "confirm config".

## Codebase facts relevant to implementation

- `load_data` reads `pd.read_csv(csv_path, header=1)` — **headers on row 2**
  ([RosterScheduler_Combined.py:504](RosterScheduler_Combined.py#L504))
- `date_col_start`: MO 6, HO 4, REG 5. Columns before it are metadata
  (`Name`, `Team`, `Ward`, `Subspec`, `EligibleShifts`, `SpecialReq` varying by type), everything
  after is the date grid ([:97-200](RosterScheduler_Combined.py#L97-L200))
- `build_scheduler(roster_type, csv_path, output_dir, config, template_path, prev_month_path)` is
  **path-based**. So step 2 should write the Sheet-derived DataFrame to a **temp CSV** and pass the
  path — zero solver risk, exactly today's behaviour. Injecting a DataFrame directly is a later
  optional refactor.
- `_read_prev_month` ([:1023](RosterScheduler_Combined.py#L1023)) reads cumulative points/months
  from **last month's generated output workbook**. This is why archiving output closes the loop and
  removes an upload — the app's own output is next month's input.
- `_preflight_check` ([:562](RosterScheduler_Combined.py#L562)) is thin — only warns when zero
  CICU-eligible staff exist. Its output goes to stdout → the collapsed "Solver log" expander.
- Blank rows are silently dropped at [:511](RosterScheduler_Combined.py#L511), count printed to the
  same buried log.

## Open questions

1. **What currently reads the Google Form's response sheet?** If a staff-facing form or an Apps
   Script checks the open/closed flag to gate submissions, that consumer must be repointed at the
   new config tab — otherwise a stale "open" flag lets staff submit into a month already solved.
2. **What does "default" mean for the reset?** Completely blank cells, or pre-populated with
   recurring patterns (e.g. known `AUTOBLOCK` / `SUBSPEC`)?
3. **Does reset preserve the staff list**, or clear that too?
4. Spreadsheet IDs for HO/MO/REG + the Admins spreadsheet; where templates should live in the repo.

## Known limitations accepted

- Sheets API is last-write-wins — two admins editing simultaneously can clobber each other.
  Fine at this team size; worth a one-line note in the UI.
- Password management is manual: resets and rotation need a small helper script to generate bcrypt
  hashes. (Trade-off accepted when choosing `streamlit-authenticator` over Google OAuth.)
- The Admins spreadsheet's sharing settings are part of the security boundary.
- Local dev now needs live Sheets/GitHub access or mocks, not just a local CSV.
