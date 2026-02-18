#!/bin/bash
# =============================================================================
# One-time server setup for VisorQuotationWebApp on Hetzner
# Run this ONCE on the server before the first GitHub Actions deploy.
#
# Usage: ssh user@server 'bash -s' < server-setup.sh
# =============================================================================

set -e

APP_DIR="/var/www/visorquotation"
APP_PORT=5050
SERVICE_NAME="visorquotation"
APP_USER="www-data"

echo "=== Setting up VisorQuotation on port $APP_PORT ==="

# 1. Install .NET 8 Runtime
if ! command -v dotnet &> /dev/null; then
    echo "--- Installing .NET 8 Runtime ---"
    wget https://dot.net/v1/dotnet-install.sh -O /tmp/dotnet-install.sh
    chmod +x /tmp/dotnet-install.sh
    /tmp/dotnet-install.sh --channel 8.0 --runtime aspnetcore --install-dir /usr/share/dotnet
    ln -sf /usr/share/dotnet/dotnet /usr/bin/dotnet
    echo "dotnet installed: $(dotnet --version)"
else
    echo "dotnet already installed: $(dotnet --version)"
fi

# 2. Create application directory
echo "--- Creating application directory ---"
sudo mkdir -p "$APP_DIR"
sudo chown -R "$APP_USER":"$APP_USER" "$APP_DIR"

# 3. Create systemd service
echo "--- Creating systemd service ---"
sudo tee /etc/systemd/system/${SERVICE_NAME}.service > /dev/null <<EOF
[Unit]
Description=Visor Quotation Web App
After=network.target

[Service]
WorkingDirectory=${APP_DIR}
ExecStart=/usr/bin/dotnet ${APP_DIR}/VisorQuotationWebApp.dll --urls "http://0.0.0.0:${APP_PORT}"
Restart=always
RestartSec=10
SyslogIdentifier=${SERVICE_NAME}
User=${APP_USER}
Environment=ASPNETCORE_ENVIRONMENT=Production
Environment=DOTNET_PRINT_TELEMETRY_MESSAGE=false
Environment=PLAYWRIGHT_BROWSERS_PATH=${APP_DIR}/.playwright

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable ${SERVICE_NAME}

# 4. Open firewall port (if ufw is active)
if command -v ufw &> /dev/null && sudo ufw status | grep -q "active"; then
    echo "--- Opening port $APP_PORT in firewall ---"
    sudo ufw allow ${APP_PORT}/tcp
fi

echo ""
echo "=== Setup complete ==="
echo "  App directory:  $APP_DIR"
echo "  App port:       $APP_PORT"
echo "  Service name:   $SERVICE_NAME"
echo "  URL:            http://<server-ip>:$APP_PORT"
echo ""
echo "The app will start automatically after the first GitHub Actions deploy."
echo "Manual commands:"
echo "  sudo systemctl status  $SERVICE_NAME"
echo "  sudo systemctl restart $SERVICE_NAME"
echo "  sudo journalctl -u $SERVICE_NAME -f"
