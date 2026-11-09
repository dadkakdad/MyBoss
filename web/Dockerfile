FROM node:22-alpine

WORKDIR /app

COPY package.json package-lock.json ./
# The database client is installed in the image. The browser lockfile does not list it.
RUN npm ci --omit=dev && npm install --omit=dev pg@8.13.1

# A new deploy timestamp invalidates the source copy below.
COPY .cachebust /tmp/.cachebust
RUN echo "Cache bust: $(cat /tmp/.cachebust)"

COPY . .

ENV PORT=80
EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:'+(process.env.PORT||80)).then((r)=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

CMD ["node", "server.mjs"]
