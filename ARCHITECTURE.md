# OrderFlow Architecture

OrderFlow is a browser-based operations app for planning daily customer orders, exporting Excel workbooks, managing loading schedules, and tracking Manco's. The app is deployed as a Python web service and uses Supabase as the shared cloud state so multiple users can work on the same sessions from different computers.

## High-Level Structure

```text
Browser
  |
  | static HTML/CSS/JS
  v
public/
  |-- index.html              Main OrderFlow UI
  |-- app.js                  Main client-side application logic
  |-- styles.css              Main visual system
  |-- shortages/              Manco's UI
  |-- greenops/               Mirrored legacy entry point
  |
  | HTTP API calls
  v
app.py                        Main Python web server for Render/local use
  |
  | shared state
  v
api/_cloud_state_store.py     Supabase REST adapter
api/_team_state_store.py      Work session state store and merge logic
shortage_app.py               Manco's state, analytics, and exports
greenops_shortage_bridge.py   Converts OrderFlow orders into Manco's sessions
```

The Render deployment starts `app.py`, which serves the static frontend and handles the backend API routes in one process.

## Main Runtime Components

### Frontend

The main frontend lives in `public/`.

- `public/index.html` defines the app shell, session creation page, client cards, order workspace, Leverschema, Laadschema, Stock, Settings, and shared modals.
- `public/app.js` contains the main client-side state machine, email upload flow, client workspace handling, session sync, Excel export calls, CMR actions, Leverschema/Laadschema controls, and Stock upload/export logic.
- `public/styles.css` contains the green/white UI system used across the app.
- `public/shortages/` contains the Manco's UI.
- `public/greenops/` mirrors the main app for the older GreenOps entry path.

The frontend keeps temporary UI state in memory and mirrors some data to IndexedDB/localStorage for resilience, but the shared source of truth is the backend state stored in Supabase.

### Main Backend

`app.py` is the main backend when running locally or on Render.

It is responsible for:

- serving frontend files;
- login/logout/session cookie handling;
- work session creation and updates;
- email parsing;
- Excel exports;
- Leverschema export;
- Laadschema export;
- Stock parsing/export;
- CMR/Pakbon flows;
- saving OrderFlow orders into Manco's;
- removing deleted card orders from Manco's;
- Manco's analytics and export routes.

### API Modules

The `api/` folder contains reusable backend modules and alternative serverless-style handlers.

Important files:

- `api/_auth.py` handles login, signed session cookies, and environment-configured users.
- `api/_cloud_state_store.py` reads/writes JSON state to Supabase.
- `api/_team_state_store.py` stores the main shared OrderFlow state and merges updates from multiple browsers.
- `api/_shared.py` contains shared parsing/export helpers.
- `api/work_sessions.py` is the API-style handler for work session state.
- `api/orders_ingest.py` saves OrderFlow orders into Manco's.
- `api/orders_remove.py` removes an order from Manco's when it is deleted from a client card.
- `api/shortage.py` delegates Manco's API requests to `shortage_app.py`.

`vercel.json` maps API paths to these handlers for Vercel-style deployments, while Render uses `app.py` directly.

## Shared State Model

Supabase stores JSON documents in the configured state table, normally `orderflow_app_state`.

Main keys:

- `team_state:v1`
  Stores OrderFlow work sessions, client workspaces, Leverschema results, Laadschema data, and custom trucks.
- `shortage_sessions:v1`
  Stores Manco's order sessions created from OrderFlow orders or Manco's uploads.
- `shortage_day_sessions:v1`
  Stores Manco's day/session overview entries.

If Supabase is not configured, the app can fall back to local JSON files in `data/`, but production should use Supabase.

## Work Sessions

A work session is the daily planning container created on the first page.

Work sessions contain:

- `id`
- `date`
- `name`
- `createdBy`
- `createdAt`
- `updatedAt`
- `workspaces`

`workspaces` is keyed by canonical client name, for example:

- `Carrefour`
- `Colruyt`
- `Denemark`
- `HAVI`
- `NettoMD`

Each workspace stores the parsed preview, selected order index, selected delivery point, Leverschema inclusion indexes, and export settings.

Session creation is server-side through `POST /api/work_sessions`. This avoids creating local-only sessions in one browser. The backend also prevents duplicate sessions for the same date/name.

## Multi-User Synchronization

The browser periodically calls:

```text
GET /api/work_sessions
```

and writes changes through:

```text
PUT /api/work_sessions
```

The server merges incoming state instead of blindly replacing the whole document. This is important because multiple users may have different browser states open at the same time.

The merge logic in `api/_team_state_store.py`:

- preserves sessions created by other users;
- applies newer session updates using `updatedAt`;
- respects explicit deleted session IDs;
- merges Leverschema and Laadschema dictionaries;
- deduplicates empty duplicate sessions by date/name.

## Email Order Flow

1. The user opens a work session.
2. The user selects a client card.
3. The user uploads an `.eml` file.
4. `public/app.js` sends the email to:

```text
POST /api/parse
```

5. The backend parses the email and returns an order preview.
6. The frontend stores that preview inside the active session workspace.
7. The order is sent to Manco's via:

```text
POST /api/orders/ingest
```

The selected card client is sent with the ingest request. If the parser cannot identify a client from the email, `greenops_shortage_bridge.py` uses the selected card client as the source of truth. This prevents orders added under a known client card from appearing as `Unknown client` in Manco's.

## Manco's Integration

Manco's data is managed by `shortage_app.py`.

OrderFlow sends saved card orders to Manco's through `greenops_shortage_bridge.py`, which converts OrderFlow previews into Manco's preview sessions.

Each Manco's order session includes:

- client;
- delivery point;
- order reference;
- ordered quantity;
- delivered quantity;
- manco quantity;
- manco percentage;
- original OrderFlow metadata.

The link between OrderFlow and Manco's is mostly:

- `workSessionId`
- `greenopsReference`
- `greenopsCustomer`
- `greenopsFatrans`
- `deliveryPoint`

When an order is deleted from a client card, the frontend calls:

```text
POST /api/orders/remove
```

The backend removes the matching Manco's session so the Manco's page stays aligned with the card.

## Leverschema

Leverschema is managed inside the main OrderFlow app.

The frontend stores selected order indexes in each workspace. When the user saves to Leverschema, the result is stored in `leverschemaResults` under `team_state:v1`.

Excel export is handled through:

```text
POST /api/export_leverschema
```

Template-related logic lives mainly in `api/_shared.py` and `app.py`.

## Laadschema

Laadschema data is stored under:

- `laadschemaData`
- `laadschemaCustomTrucks`

inside `team_state:v1`.

The frontend edits and renders the loading schedule. Export is handled by:

```text
POST /api/export_laadschema
```

## Stock

The Stock module lets users upload Excel files with stock information.

Relevant columns are:

- `Item number`
- `Product name`
- `Physical Inventory HU`, shown/exported as `Quantity`
- `Expiration Date`, shown/exported as `THT`

Parsing and export are handled in `app.py` through:

```text
POST /api/stock/parse
POST /api/stock/export
```

## Authentication

Authentication is simple cookie-based authentication handled by `api/_auth.py`.

Users can be configured with:

```text
ORDERFLOW_USERS_JSON
```

The session cookie is signed using:

```text
GREENOPS_AUTH_SECRET
```

## Deployment

The current production deployment is designed for Render.

`render.yaml` defines:

- Python runtime;
- `pip install -r requirements.txt`;
- `python app.py`;
- required environment variables.

Required production environment variables:

```text
HOST=0.0.0.0
SUPABASE_URL=<project-url>
SUPABASE_SERVICE_ROLE_KEY=<legacy service_role key>
GREENOPS_AUTH_SECRET=<strong random secret>
ORDERFLOW_USERS_JSON=<JSON user list>
```

For the current Supabase REST approach, `SUPABASE_SERVICE_ROLE_KEY` should be the legacy `service_role` key, normally beginning with `eyJ...`.

## Important Design Decisions

- The server is the source of truth for sessions.
- The browser never assumes a session was created until the backend confirms it.
- Supabase stores coarse-grained JSON state documents rather than relational order rows.
- Manco's sessions are separate from OrderFlow sessions but linked by `workSessionId`.
- The selected client card is trusted when email recognition cannot identify the client.
- Deleting an order from a card also removes the matching Manco's entry.
- The `public/greenops/` files are mirrored for legacy compatibility and should be kept in sync with `public/`.

## Common Change Checklist

When changing core order/session behavior:

1. Update `public/app.js`.
2. Mirror relevant changes to `public/greenops/app.js`.
3. If styling changes, mirror `public/styles.css` to `public/greenops/styles.css`.
4. Check whether `app.py` and any matching `api/*.py` handler both need the same behavior.
5. Run syntax checks for changed Python and JavaScript files.
6. Commit and deploy through Render.
