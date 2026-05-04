## Build stage ##############################################################
FROM node:20-alpine AS build

WORKDIR /app

# Build-time configuration baked into the bundle and manifest.
ARG ADDIN_PUBLIC_URL=https://addin.postguard.eu/
ARG PKG_URL=https://postguard.eu/pkg
ARG CRYPTIFY_URL=https://fileshare.postguard.eu
ARG POSTGUARD_WEBSITE_URL=https://postguard.eu

ENV ADDIN_PUBLIC_URL=${ADDIN_PUBLIC_URL} \
    PKG_URL=${PKG_URL} \
    CRYPTIFY_URL=${CRYPTIFY_URL} \
    POSTGUARD_WEBSITE_URL=${POSTGUARD_WEBSITE_URL} \
    NODE_ENV=production

COPY package.json package-lock.json ./
RUN npm ci --no-audit --no-fund

COPY . .
RUN npx webpack --mode production

## Runtime stage ############################################################
FROM nginx:1.27-alpine AS runtime

COPY nginx/default.conf /etc/nginx/conf.d/default.conf
COPY --from=build /app/dist /usr/share/nginx/html

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
