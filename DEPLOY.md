# Deploying the IPO tool (free & always-on)

The app runs a headless Chromium browser and each generation takes ~4–5 minutes,
so it needs an **always-on server with ~1 GB RAM** — not a serverless host.

The genuinely-free, always-on option is an **Oracle Cloud "Always Free" VM**.
Everything is packaged with Docker, so the same image also runs on Render/Railway/Fly
if you move later.

> Heads-up: from a datacenter IP, SEBI/BSE/NSE may throttle scraping more than from a
> home connection. And only **one generation runs at a time** — a second concurrent
> user gets a "busy" message. Both are fine for a handful of occasional users.

---

## Option A — Oracle Cloud Always Free VM (recommended: free + always on)

1. **Create the VM**
   - Sign up at <https://www.oracle.com/cloud/free/> (needs a card for verification; not charged).
   - Create a Compute instance → shape **VM.Standard.A1.Flex** (Ampere/ARM, Always Free:
     up to 4 OCPU / 24 GB) or **VM.Standard.E2.1.Micro** (Always Free, 1 GB).
     Give it **1–2 GB RAM** or more. Image: **Ubuntu 22.04**.
   - Download the SSH key when prompted.

2. **Open the web port**
   - In the instance's subnet **Security List**, add an **Ingress rule**:
     Source `0.0.0.0/0`, protocol TCP, destination port **80**.
   - Also allow it in the OS firewall:
     ```bash
     sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT
     sudo netfilter-persistent save
     ```

3. **Install Docker**
   ```bash
   sudo apt update && sudo apt install -y docker.io docker-compose-plugin
   sudo usermod -aG docker $USER && newgrp docker
   ```

4. **Copy the app and run it**
   ```bash
   # from your PC, upload the ipo_tool folder (or git clone it):
   scp -i <your-key> -r ipo_tool ubuntu@<server-ip>:~/

   # on the server:
   cd ~/ipo_tool
   docker compose up -d --build
   ```

5. **Open it:** `http://<server-ip>/`
   - `docker compose logs -f` to watch, `docker compose restart` to restart.
   - `restart: always` keeps it running across reboots.

### Optional: a real domain + HTTPS
Point a domain's A-record at the server IP, then put Caddy in front for automatic HTTPS,
or use a free Cloudflare Tunnel. Ask me and I'll add the config.

---

## Option B — Render (simplest, but the free tier sleeps)

Render's free web service is the easiest deploy, but it **spins down after ~15 min idle**
(first request after that is slow) and gives 512 MB RAM (tight for Chromium — a heavy run
can occasionally get killed). Fine for light, tolerant use.

1. Push `ipo_tool/` to a GitHub repo.
2. Render → **New → Web Service** → connect the repo.
3. Environment: **Docker** (it auto-detects the `Dockerfile`). No build/start command needed.
4. Instance type: Free. Deploy. Render gives you a public `https://…onrender.com` URL.

(Render injects `PORT` automatically; `serve.py` already reads it.)

---

## Running the Docker image anywhere

```bash
docker build -t ipo-tool .
docker run -d --restart always -p 80:8080 --shm-size 512m --memory 1500m ipo-tool
```

Generated files are written inside the container's `output/` and offered as a download.
They're ephemeral (cleared on redeploy), which is fine since users download immediately.
