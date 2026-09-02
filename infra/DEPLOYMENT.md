# VPS deployment

This project is arranged as:

```text
frontend/       React/Vite browser application
backend/        Node.js API, PostgreSQL persistence, upload handling, reports
infra/          Docker Compose and Nginx deployment configuration
```

## Before the VPS

1. Choose a domain, for example `portal.stdorothyschool.ac.ke`, and point its DNS A record to the VPS IPv4 address.
2. Use Ubuntu 24.04 LTS with a non-root deployment user, SSH keys, firewall rules for ports 22, 80, and 443 only, and automatic security updates.
3. Install Docker Engine, Docker Compose plugin, Nginx, and Certbot on the VPS.
4. Clone the private Git repository to `/opt/st-dorothy-portal`.

## Configure secrets

On the VPS:

```bash
cd /opt/st-dorothy-portal
cp .env.production.example .env.production
chmod 600 .env.production
```

Edit `.env.production` with unique passwords and the actual school domain. Do not add Daraja, SMS, or email keys until those school-owned accounts are approved.

## Start the portal

```bash
docker compose --env-file .env.production -f infra/docker-compose.production.yml up -d --build
docker compose --env-file .env.production -f infra/docker-compose.production.yml ps
```

The public web container listens only on `127.0.0.1:8080`. PostgreSQL and the API are internal Docker services and are not publicly exposed.

## HTTPS reverse proxy

Copy `infra/nginx/vps-portal.conf` to `/etc/nginx/sites-available/st-dorothy-portal`, enable it, test Nginx, and request the certificate:

```bash
sudo ln -s /etc/nginx/sites-available/st-dorothy-portal /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
sudo certbot --nginx -d portal.stdorothyschool.ac.ke
```

Certbot updates Nginx with TLS. Confirm that HTTP redirects to HTTPS and that the certificate renews automatically.

## Verification before launch

- Open `/api/health` through the public domain.
- Create a test admission, update it, sign in to a parent account, and confirm permissions.
- Confirm PostgreSQL data survives `docker compose restart`.
- Test document upload/download via the public HTTPS domain.
- Create a database backup and restore it into a separate test database.
- Keep M-Pesa, SMS, and email in test mode until their providers confirm production access.

## Operations

Run a nightly encrypted backup outside the VPS for both PostgreSQL and uploaded documents. Review Docker logs and disk usage weekly. Rotate all deployment and integration secrets if an administrator leaves.

The current PostgreSQL implementation moves the portal off the local JSON file. Before a school-wide launch with heavy concurrent use or live payments, the next backend hardening milestone is a fully normalised PostgreSQL schema and database-backed session storage.
