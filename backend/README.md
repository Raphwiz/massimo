# Backend API

Run the backend in one terminal:

```powershell
npm run server
```

Run the website in another terminal:

```powershell
npm run dev
```

The Vite development server forwards `/api` requests to the local API at `http://127.0.0.1:8787`.

For local evaluation without PostgreSQL, data remains in `backend/data/school-data.json`. When `DATABASE_URL` is configured, PostgreSQL is the active persistence layer and the JSON file is imported only during the first database initialisation. SMS entries are recorded as **Prepared** only; no live SMS is sent. The API also includes an admission-enquiry endpoint ready for the next public admissions form.

See `../infra/DEPLOYMENT.md` for the VPS deployment process. Before accepting live payments, complete the approved M-Pesa, SMS, email, backup, and data-protection setup.
