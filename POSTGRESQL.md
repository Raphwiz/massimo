# PostgreSQL setup

The live St. Dorothy portal requires PostgreSQL. Its `DATABASE_URL` is the only database setting the server needs.

## Local development

1. Copy `.env.example` to `.env` and set a local connection string, for example:

   `DATABASE_URL=postgresql://st_dorothy_app:change-this-before-use@127.0.0.1:5432/st_dorothy_school`

2. Change `POSTGRES_PASSWORD` in `docker-compose.yml` and use the same value in `DATABASE_URL`.
3. Start PostgreSQL: `docker compose up -d postgres`.
4. Start the API: `npm run server`.

On the first PostgreSQL connection, the server creates `school_portal_state` and imports the current local prototype data from `server/data/school-data.json`. It will not overwrite an existing database.

## Production

- Use a managed PostgreSQL service with daily backups, point-in-time recovery, encryption at rest, and a private network connection where possible.
- Put the `DATABASE_URL` only in the hosting provider's encrypted environment-variable store.
- Set `NODE_ENV=production`, `DATABASE_SSL=true`, a precise `ALLOWED_ORIGINS` value, and unique initial administrator passwords.
- Do not expose port 5432 publicly. Only the portal API should connect to the database.
- Confirm a restore from backup before accepting live admissions or payments.

The JSON file remains available only as a local-development fallback when `DATABASE_URL` is absent. Production startup is blocked without PostgreSQL.
