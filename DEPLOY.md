# Deploying the Weekly Portfolio Kit to a server

This turns the local-only web app into something that runs on its own, 24/7, on a
server you control — reachable from anywhere over HTTPS, behind a password.

**Why this needs more than just "copy webapp.py to a server":**
- Flask's built-in server (what `python webapp.py` uses) explicitly warns it's not
  for production — no concurrency handling, no crash recovery. We use **gunicorn** instead.
- The original app has **no login** — fine on your own laptop, not fine on the open
  internet holding your live brokerage data. We added HTTP Basic Auth.
- `finalize.py` shells out to **LibreOffice**. A bare server won't have it installed;
  we bundle it into the Docker image so you don't have to think about it.
- "Runs on its own" = restarts itself if it crashes or the server reboots, which is
  what `restart: unless-stopped` in docker-compose gives you.

## 1. Get a server
Any small Linux VPS works — this app is lightweight (Flask + LibreOffice, no GPU/DB).
Cheap, reliable options: Hetzner Cloud (~€4/mo), DigitalOcean Droplet (~$6/mo),
AWS Lightsail (~$5/mo). Pick Ubuntu 22.04/24.04, smallest size (1 vCPU/1-2GB RAM is enough).

## 2. Point a domain at it (optional but recommended for HTTPS)
Add an A record: `portfolio.yourdomain.com` → your server's IP. If you don't have a
domain, you can skip nginx/TLS below and just access it as `http://SERVER_IP:8000`
over plain HTTP — fine if you're the only one ever hitting it, but credentials and
your portfolio data travel unencrypted, so a domain + HTTPS is worth the effort.

## 3. Install Docker on the server
```
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker $USER   # log out/in after this
```

## 4. Copy this kit to the server
```
scp -r portfolio-kit-demo-only/ you@your-server-ip:/opt/portfolio-kit
ssh you@your-server-ip
cd /opt/portfolio-kit
```

## 5. Set your login
```
cp .env.example .env
nano .env       # set WEBAPP_USER and a long random WEBAPP_PASS
```
The app **refuses to start** without these set — that's intentional, so it can never
accidentally run unauthenticated on a public IP.

## 6. Build and start it — always on
```
docker compose up -d --build
```
`restart: unless-stopped` in `docker-compose.yml` means: if the process crashes, or
the whole server reboots, Docker brings the container back up automatically. No cron
job, no manual restart, no systemd unit needed — Docker's own restart policy handles it.

Check it's up:
```
docker compose logs -f     # ctrl-C to stop watching, container keeps running
```

At this point the app is live at `http://127.0.0.1:8000` **on the server only** — the
compose file binds it to localhost on purpose so it isn't directly exposed.

## 7. Put HTTPS in front of it (nginx + Let's Encrypt)
```
sudo apt install -y nginx certbot python3-certbot-nginx
sudo cp nginx.conf.example /etc/nginx/sites-available/portfolio-kit
sudo nano /etc/nginx/sites-available/portfolio-kit   # replace portfolio.yourdomain.com
sudo ln -s /etc/nginx/sites-available/portfolio-kit /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
sudo certbot --nginx -d portfolio.yourdomain.com   # free cert, auto-renews
```
Now `https://portfolio.yourdomain.com` serves the app, browser will prompt for the
username/password you set in `.env`.

Also open the firewall for web traffic, close everything else:
```
sudo ufw allow OpenSSH
sudo ufw allow 80,443/tcp
sudo ufw enable
```

## Weekly use, once deployed
Same workflow as local, just at your URL instead of `127.0.0.1`:
1. Visit `https://portfolio.yourdomain.com`, log in.
2. Upload this week's ICICI/Kite files + screener export.
3. Set dates, click Run.
4. Download the outputs.

Your uploaded files and generated reviews persist across restarts/rebuilds — they
live in `./inputs` and `./outputs` on the server (mounted as Docker volumes), not
inside the container.

## Updating the kit later
```
cd /opt/portfolio-kit
# edit engine/tools files, or git pull if you version-control it
docker compose up -d --build
```
Inputs/outputs are untouched by rebuilds since they're mounted volumes.

## Optional: also auto-run the weekly build itself (not just keep the app up)
The app staying up doesn't build anything until you click Run — you still need to
drop this week's files in first, which is a manual download from your broker/Screener.
If you want a standing **schedule** on top of this (e.g. a Saturday 9am reminder or
an automatic run against whatever's already in `inputs/`), that's a separate cron
job inside the container — say the word and I'll wire that in too, but note it can
only build from files that are already there; it can't fetch your broker exports
for you.
