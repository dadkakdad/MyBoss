FROM node:22-alpine AS deps

WORKDIR /app

# Copied before npm ci so a new deploy timestamp invalidates the dependency layer.
COPY package.json package-lock.json .cachebust ./
RUN echo "Cache bust: $(cat .cachebust)" && npm ci --omit=dev

FROM nginx:1.27-alpine

COPY nginx/default.conf /etc/nginx/conf.d/default.conf

# Cache bust: a new .cachebust invalidates every COPY below (CapRover may reuse layers).
COPY .cachebust /tmp/.cachebust
RUN echo "Cache bust: $(cat /tmp/.cachebust)"

COPY index.html app.bundle.js /usr/share/nginx/html/
COPY css /usr/share/nginx/html/css
COPY js /usr/share/nginx/html/js
COPY icons /usr/share/nginx/html/icons
COPY --from=deps /app/node_modules/xlsx/dist/xlsx.full.min.js /usr/share/nginx/html/node_modules/xlsx/dist/xlsx.full.min.js
COPY --from=deps /app/node_modules/pptxgenjs/dist/pptxgen.bundle.js /usr/share/nginx/html/node_modules/pptxgenjs/dist/pptxgen.bundle.js

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget -qO- http://127.0.0.1/ >/dev/null || exit 1
