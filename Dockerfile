FROM node:20-alpine

WORKDIR /app

COPY . .

EXPOSE 4173

CMD ["node", "local-server.js", ".", "4173"]
