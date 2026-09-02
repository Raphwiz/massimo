# St. Dorothy School Portal

Web-based school management portal for St. Dorothy School Massimo.

## Project layout

- `frontend/` — React/Vite public website and role-based portals.
- `backend/` — Node.js REST API, authentication, records, reports, and PostgreSQL persistence.
- `infra/` — VPS Docker and Nginx deployment files.
- `POSTGRESQL.md` — local database setup.
- `SECURITY.md` — security controls and deployment requirements.

## Local development

```powershell
npm install
npm run server
# In another terminal:
npm run dev
```

Open `http://127.0.0.1:5173`. For PostgreSQL, follow `POSTGRESQL.md`.

## VPS deployment

Follow `infra/DEPLOYMENT.md`. Never commit `.env`, `.env.production`, database credentials, M-Pesa credentials, or live SMS/email keys.
