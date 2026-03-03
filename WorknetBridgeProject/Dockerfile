# syntax=docker/dockerfile:1

# Comments are provided throughout this file to help you get started.
# If you need more help, visit the Dockerfile reference guide at
# https://docs.docker.com/go/dockerfile-reference/

# Want to help us make this template better? Share your feedback here: https://forms.gle/ybq9Krt8jtBL3iCk7

ARG LATEST_NODE_VERSION=24.13.0

FROM node:20-bullseye

# Use production node environment by default.
ENV NODE_ENV=production

# Install system dependencies required for headed Chromium
RUN apt-get update && apt-get install -y \
  libnss3 \
  libatk-bridge2.0-0 \
  libatk1.0-0 \
  libcups2 \
  libdrm2 \
  libxkbcommon0 \
  libxcomposite1 \
  libxdamage1 \
  libxfixes3 \
  libxrandr2 \
  libgbm1 \
  libasound2 \
  libpangocairo-1.0-0 \
  libpango-1.0-0 \
  libcairo2 \
  libgtk-3-0 \
  fonts-liberation \
  xdg-utils \
  xvfb \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json ./

# Download dependencies as a separate step to take advantage of Docker's caching.
# Leverage a cache mount to /root/.npm to speed up subsequent builds.
# Leverage a bind mounts to package.json and package-lock.json to avoid having to copy them into
# into this layer.
RUN  --mount=type=bind,source=package.json,target=package.json \
    --mount=type=cache,target=/root/.npm \
    npm install --omit=dev

# Copy the rest of the source files into the image.
COPY . .

# Required for non-headless Chrome in Docker
ENV DISPLAY=:99

# Run the application.
CMD ["npm", "start"]
