# Server 
### This is excel based remote sync, for web its using web memory and in anodroid it uses sqllited

  Location: /Users/ganeshkumar/Documents/AMS/code/excel-api/

  Stack: TypeScript + Express + ExcelJS

  Endpoints:

  ┌────────┬─────────────────┬────────────────────────────────────────────────────────────┐
  │ Method │      Route      │                        Description                         | │        |                 |                                                            |                                               ├────────┼─────────────────┼────────────────────────────────────────────────────────────┤
  │ GET    │ /api/sheets     │ List all sheets (tables)                                   │
  ├────────┼─────────────────┼────────────────────────────────────────────────────────────┤
  │ GET    │ /api/:sheet     │ List rows (supports ?filter[field]=value&limit=N&offset=N) │
  ├────────┼─────────────────┼────────────────────────────────────────────────────────────┤                                                                         │ GET    │ /api/:sheet/:id │ Get row by UUID                                            │
  ├────────┼─────────────────┼────────────────────────────────────────────────────────────┤
  │ POST   │ /api/:sheet     │ Insert row (auto-assigns UUID id)                          │
  ├────────┼─────────────────┼────────────────────────────────────────────────────────────┤
  │ PUT    │ /api/:sheet/:id │ Update row (partial update)                                │
  ├────────┼─────────────────┼────────────────────────────────────────────────────────────┤
  │ DELETE │ /api/:sheet/:id │ Delete row                                                 |│        |                 |                                                            |                           
  └────────┴─────────────────┴────────────────────────────────────────────────────────────┘                                                                                                                                       

  Key design decisions:
  - Each Excel sheet = a table; posting to /api/users auto-creates a users sheet
  - Column headers are auto-created from the first POST body — no schema definition needed
  - Each row gets a UUID id automatically
  - data/database.xlsx is created on first write

  To start: npm run dev


❯ how to stop                                                          
  
  kill it by port:
  kill $(lsof -ti :3000)

  Add the uuid package and set the API URL env variable:

⏺ Bash(npm install uuid && npm install -D @types/uuid 2>&1 | tail -5)
  ⎿  npm warn deprecated uuid@7.0.3: uuid@10 and below is no longer supported.  For ESM codebases, update to uuid@latest.  For CommonJS codebases, use uuid@11 (but be aware this version will likely be deprecated in 2028).

     added 2 packages, changed 1 package, and audited 508 packages in 916ms
     … +25 lines (ctrl+o to expand)

⏺ Write(.env)
  ⎿  Wrote 1 line to .env
      1 EXPO_PUBLIC_API_URL=http://localhost:3000

⏺ Bash(npx tsc --noEmit 2>&1)
  ⎿  (No output)

⏺ Zero type errors. Here's the complete picture:

  ---

  Project: /Users/ganeshkumar/Documents/AMS/code/excel-sync-app/

  Architecture:
  Excel file  ←→  excel-api (Express server)  ←→  Expo app  ←→  SQLite (on device)
    (remote source of truth)                     (local cache + offline writes)

  How sync works:

  ┌───────────┬─────────────────────────┬─────────────────────────────────────────────────────────┐
  │ Direction │          When           │                      What happens                       │
  ├───────────┼─────────────────────────┼─────────────────────────────────────────────────────────┤
  │ Pull      │ Tap "Sync" or app opens │ Fetches all rows from Excel API → stores in SQLite      │
  ├───────────┼─────────────────────────┼─────────────────────────────────────────────────────────┤
  │ Push      │ Before every pull       │ Sends locally-created/updated/deleted rows to Excel API │
  ├───────────┼─────────────────────────┼─────────────────────────────────────────────────────────┤
  │ Offline   │ No network              │ App reads/writes SQLite; changes queued as pending_*    │
  └───────────┴─────────────────────────┴─────────────────────────────────────────────────────────┘

  Key files:
  - src/api/excel.api.ts — HTTP client for your Excel API
  - src/sync/sync.service.ts — push/pull logic, LocalDb CRUD
  - src/hooks/useSheet.ts — React hook: { rows, sync, insert, update, remove }
  - src/screens/SheetScreen.tsx — UI showing rows with Add/Delete/Sync
  - .env — set EXPO_PUBLIC_API_URL to your server's IP (use your machine's LAN IP for Android, e.g. http://192.168.1.x:3000)
  
  To run:
  # Terminal 1 — start Excel API
  cd excel-api && npm run dev
  
  # Terminal 2 — start Expo app
  cd excel-sync-app && npx expo start
  # press 'a' for Android emulator, scan QR for device

  ▎ For a real Android device: change .env to your machine's local IP — localhost won't work from the device.
