# Use an official Node.js runtime as a parent image
FROM node:22.12-alpine AS builder

COPY src/sharepoint /app
COPY tsconfig.json /tsconfig.json

# Set the working directory inside the container
WORKDIR /app

RUN --mount=type=cache,target=/root/.npm npm install

# Compile TypeScript to JavaScript
RUN npx tsc

RUN --mount=type=cache,target=/root/.npm-production npm ci --ignore-scripts --omit-dev

FROM node:22-alpine AS release

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package.json /app/package.json
COPY --from=builder /app/package-lock.json /app/package-lock.json

ENV NODE_ENV=production

WORKDIR /app

RUN npm ci --ignore-scripts --omit-dev

ENTRYPOINT ["node", "dist/index.js"]
