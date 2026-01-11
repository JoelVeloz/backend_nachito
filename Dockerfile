# Usar la imagen base de Node.js en Alpine
FROM node:22.13.1-alpine3.20

# Establecer la zona horaria
ENV TZ=America/Guayaquil
RUN apk add --no-cache tzdata && \
    ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && \
    echo $TZ > /etc/timezone

# Establecer el directorio de trabajo
WORKDIR /app

# Copiar package.json y package-lock.json
COPY package*.json ./

# Instalar todas las dependencias (incluyendo devDependencies para el build)
RUN npm ci && npm cache clean --force

# Copiar el resto de la aplicación
COPY . .


# Construir la aplicación
RUN npm run build

# Limpiar dependencias de desarrollo y node_modules
RUN npm prune --production && npm cache clean --force

# Exponer el puerto
EXPOSE 4000

# Comando para ejecutar la aplicación
CMD ["node", "dist/main.js"]
