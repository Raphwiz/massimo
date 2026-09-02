# Local backend evaluation

Run the backend in one terminal:

```powershell
npm run server
```

Run the website in another terminal:

```powershell
npm run dev
```

The Vite development server forwards `/api` requests to the local API at `http://127.0.0.1:8787`.

For this evaluation build, calendar events, SMS activity, and Secretary-created admission applications are stored in `server/data/school-data.json`, so they remain after refreshing the website or restarting the frontend. SMS entries are recorded as **Prepared** only; no live SMS is sent. The API also includes an admission-enquiry endpoint ready for the next public admissions form.

Before production, replace the JSON file with an access-controlled database, add authentication and permissions, encrypted backups, audit logs, rate limiting, and approved M-Pesa/SMS integrations.
