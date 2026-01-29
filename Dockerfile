# Multi-stage build for lightweight production image
FROM node:18-alpine AS base

WORKDIR /app

# Copy only the necessary files for serving
COPY index.html styles.css script.js templates.json ./
COPY modules/ ./modules/
COPY demo/ ./demo/
COPY favicon.png logo.png ./

# Use a lightweight web server to serve static files
FROM caddy:2-alpine

WORKDIR /srv

# Copy the static files from the base stage
COPY --from=base /app /srv

# Copy a simple Caddyfile for serving
COPY Caddyfile /etc/caddy/Caddyfile

# Expose port 80 and 443 (for HTTPS support if needed)
EXPOSE 80 443

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD wget --no-verbose --tries=1 --spider http://localhost/ || exit 1

# Caddy will automatically start with the Caddyfile
CMD ["caddy", "run", "--config", "/etc/caddy/Caddyfile"]
