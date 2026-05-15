# ============================================================
# Multi-stage build for Office LLM Harness Static File Server
# Serves PowerPoint/Word/Excel/Outlook add-in task pane UI files
# ============================================================

# --- Stage 1: Build PowerPoint add-in static files ---
FROM node:20-alpine AS addin-builder
WORKDIR /build/addin

COPY src/powerpoint-addin/package.json src/powerpoint-addin/package-lock.json ./
RUN npm ci --omit=dev

COPY src/powerpoint-addin/webpack.config.js ./
COPY src/powerpoint-addin/tsconfig.json ./
COPY src/powerpoint-addin/assets/ ./assets/
COPY src/powerpoint-addin/src/ ./src/
COPY src/powerpoint-addin/manifest.xml ./

RUN npm run build

# --- Stage 2: Runtime image (nginx) ---
FROM nginx:alpine AS runtime
WORKDIR /usr/share/nginx/html

# Copy built static files
COPY --from=addin-builder /build/addin/dist .

# Copy custom nginx config
COPY docker/nginx.conf /etc/nginx/conf.d/default.conf

# Expose port
EXPOSE 80

# Health check
HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
    CMD wget -qO- http://localhost:80/health || exit 1

CMD ["nginx", "-g", "daemon off;"]
