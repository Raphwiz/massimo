FROM node:22-alpine

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY backend ./backend
COPY frontend/public/assets/st-dorothy-crest-refined.png ./frontend/public/assets/st-dorothy-crest-refined.png
COPY outputs ./outputs

ENV NODE_ENV=production
EXPOSE 8787
CMD ["node", "backend/server.js"]
