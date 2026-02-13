#!/bin/bash
# =============================================
# SEEN Accounting - VPS Setup Script
# Run on fresh Ubuntu 22.04/24.04 LTS
# Usage: sudo bash setup-vps.sh
# =============================================

set -e

echo "🚀 SEEN Accounting - VPS Setup"
echo "================================"

# ==================== System Update ====================
echo "📦 Updating system..."
apt update && apt upgrade -y

# ==================== Install Node.js 20 LTS ====================
echo "📦 Installing Node.js 20 LTS..."
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt install -y nodejs
echo "✅ Node.js $(node -v) installed"

# ==================== Install PostgreSQL 16 ====================
echo "📦 Installing PostgreSQL 16..."
sh -c 'echo "deb http://apt.postgresql.org/pub/repos/apt $(lsb_release -cs)-pgdg main" > /etc/apt/sources.list.d/pgdg.list'
wget --quiet -O - https://www.postgresql.org/media/keys/ACCC4CF8.asc | apt-key add -
apt update
apt install -y postgresql-16
echo "✅ PostgreSQL installed"

# ==================== Install Redis ====================
echo "📦 Installing Redis..."
apt install -y redis-server
systemctl enable redis-server
systemctl start redis-server
echo "✅ Redis installed"

# ==================== Install Nginx ====================
echo "📦 Installing Nginx..."
apt install -y nginx
systemctl enable nginx
echo "✅ Nginx installed"

# ==================== Install PM2 ====================
echo "📦 Installing PM2..."
npm install -g pm2
pm2 startup systemd -u $SUDO_USER --hp /home/$SUDO_USER
echo "✅ PM2 installed"

# ==================== Install Certbot (SSL) ====================
echo "📦 Installing Certbot..."
apt install -y certbot python3-certbot-nginx
echo "✅ Certbot installed"

# ==================== Configure PostgreSQL ====================
echo "🔧 Configuring PostgreSQL..."
sudo -u postgres psql -c "CREATE USER seen_user WITH PASSWORD 'CHANGE_THIS_PASSWORD';"
sudo -u postgres psql -c "CREATE DATABASE seen_accounting OWNER seen_user;"
sudo -u postgres psql -c "GRANT ALL PRIVILEGES ON DATABASE seen_accounting TO seen_user;"
echo "✅ PostgreSQL configured"
echo "⚠️  IMPORTANT: Change the database password!"

# ==================== Configure Firewall ====================
echo "🔧 Configuring firewall..."
apt install -y ufw
ufw default deny incoming
ufw default allow outgoing
ufw allow 22/tcp    # SSH
ufw allow 80/tcp    # HTTP
ufw allow 443/tcp   # HTTPS
ufw --force enable
echo "✅ Firewall configured (SSH, HTTP, HTTPS)"

# ==================== Configure Fail2ban ====================
echo "🔧 Installing Fail2ban..."
apt install -y fail2ban
systemctl enable fail2ban
systemctl start fail2ban
echo "✅ Fail2ban installed"

# ==================== Setup Backup Script ====================
echo "🔧 Setting up daily backups..."
mkdir -p /backups

cat > /usr/local/bin/backup-seen.sh << 'BACKUP_EOF'
#!/bin/bash
BACKUP_DIR="/backups"
DATE=$(date +%Y%m%d_%H%M)
FILENAME="seen_accounting_${DATE}.sql.gz"

# Dump database
pg_dump -U seen_user seen_accounting | gzip > "${BACKUP_DIR}/${FILENAME}"

# Keep last 30 days
find ${BACKUP_DIR} -name "seen_accounting_*.sql.gz" -mtime +30 -delete

echo "Backup completed: ${FILENAME}"
BACKUP_EOF

chmod +x /usr/local/bin/backup-seen.sh

# Add to crontab (daily at 3 AM)
(crontab -l 2>/dev/null; echo "0 3 * * * /usr/local/bin/backup-seen.sh >> /var/log/seen-backup.log 2>&1") | crontab -

echo "✅ Daily backup configured (3:00 AM)"

# ==================== Nginx Config ====================
echo "🔧 Creating Nginx config..."

cat > /etc/nginx/sites-available/seen-api << 'NGINX_EOF'
server {
    listen 80;
    server_name _;  # Replace with your domain

    location / {
        proxy_pass http://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_cache_bypass $http_upgrade;
    }

    # Max upload size
    client_max_body_size 10M;
}
NGINX_EOF

ln -sf /etc/nginx/sites-available/seen-api /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl reload nginx
echo "✅ Nginx configured"

# ==================== Create App Directory ====================
echo "🔧 Creating app directory..."
mkdir -p /opt/seen-accounting
chown -R $SUDO_USER:$SUDO_USER /opt/seen-accounting
echo "✅ App directory: /opt/seen-accounting"

# ==================== Auto Security Updates ====================
echo "🔧 Enabling automatic security updates..."
apt install -y unattended-upgrades
dpkg-reconfigure -f noninteractive unattended-upgrades
echo "✅ Auto updates enabled"

# ==================== Summary ====================
echo ""
echo "============================================"
echo "🎉 VPS Setup Complete!"
echo "============================================"
echo ""
echo "Next steps:"
echo "1. Change PostgreSQL password:"
echo "   sudo -u postgres psql -c \"ALTER USER seen_user PASSWORD 'new_strong_password';\""
echo ""
echo "2. Clone your project:"
echo "   cd /opt/seen-accounting"
echo "   git clone <your-repo-url> ."
echo "   cd seen-accounting-api"
echo ""
echo "3. Install dependencies:"
echo "   npm install"
echo ""
echo "4. Setup environment:"
echo "   cp .env.example .env"
echo "   nano .env  # Edit with your values"
echo ""
echo "5. Run database migration:"
echo "   npx prisma migrate deploy"
echo "   npm run db:seed"
echo ""
echo "6. Start the app:"
echo "   pm2 start ecosystem.config.cjs"
echo "   pm2 save"
echo ""
echo "7. Setup SSL (replace domain.com):"
echo "   certbot --nginx -d domain.com"
echo ""
echo "============================================"
