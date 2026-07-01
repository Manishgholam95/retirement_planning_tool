# Stage 1: Install dependencies (sqlite3 needs python3/make/g++ to compile native bindings)
FROM node:22-alpine AS deps
RUN apk add --no-cache python3 make g++
WORKDIR /app
COPY package*.json ./
RUN npm install

# Stage 2: Build the Vite frontend
FROM deps AS builder
COPY . .
RUN npm run build

# Stage 3: Production runtime — no build tools, small image
FROM node:22-alpine AS runner
WORKDIR /app
# node_modules from stage 1 (includes tsx + compiled sqlite3.node, both built on Alpine)
COPY --from=deps /app/node_modules ./node_modules
# Built frontend served as static files
COPY --from=builder /app/dist ./dist
# Server source (tsx runs TypeScript directly, no pre-compile step needed)
COPY server ./server
COPY package*.json ./

ENV NODE_ENV=production
# DATABASE_PATH points to a persistent volume — set this in Railway/Render/Fly dashboard
ENV DATABASE_PATH=/data/database.sqlite

EXPOSE 3001
CMD ["npm", "run", "start"]
