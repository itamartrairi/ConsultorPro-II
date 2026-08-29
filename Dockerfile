# ---- Build stage ----
FROM node:20-slim AS build
WORKDIR /app

COPY package*.json ./
RUN npm install

COPY . .
RUN npm run build

# ---- Run stage ----
FROM node:20-slim AS run
WORKDIR /app
ENV NODE_ENV=production

# Only production dependencies
COPY package*.json ./
RUN npm install --omit=dev

# Built server + frontend assets from the build stage
COPY --from=build /app/dist ./dist

# Cloud Run injects PORT; the server must listen on it
ENV PORT=8080
EXPOSE 8080

CMD ["node", "dist/server.cjs"]
