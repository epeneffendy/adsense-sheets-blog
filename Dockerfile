FROM node:20-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY package.json package-lock.json* ./

RUN npm ci || npm install

COPY . .

EXPOSE 3000

CMD ["npm", "run", "dev"]
