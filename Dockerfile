FROM node:22-alpine

WORKDIR /app

COPY package.json ./
RUN npm install --omit=dev

COPY . .

ENV PORT=3030

EXPOSE 3030

CMD ["npm", "start"]
