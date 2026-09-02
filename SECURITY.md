# Security deployment checklist

The local preview is for evaluation. Before the portal is exposed on the internet, deploy it behind an HTTPS reverse proxy, set `NODE_ENV=production`, use a unique `INITIAL_PORTAL_PASSWORD`, and set `ALLOWED_ORIGINS` to the exact public portal URL.

Keep M-Pesa/Daraja keys in a managed secret store, never in source files or browser code. Replace the local JSON data file with a production database that has encrypted backups, access logging, retention rules, and tested recovery procedures. Restrict database and server access with a firewall, keep the host and dependencies patched, run malware scanning on uploaded files, and arrange an independent penetration test before go-live.

The portal now uses server-side sessions, role checks, parent-record ownership checks, brute-force throttling, restrictive CORS, browser security headers, upload signature checks, and forced replacement of temporary passwords. These measures reduce common application risks but do not replace secure hosting, monitoring, backups, and professional security testing.
