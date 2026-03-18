FROM node:18-slim


# Bật port 3000
EXPOSE 3000

# Tạo thư mục và cấp quyền
WORKDIR /usr/src/app

# Cài đặt thư viện Nodejs
COPY package*.json ./
RUN npm install

# Copy toàn bộ file hệ thống vào
COPY . .
RUN mkdir -p DAKHAOSAT && chmod -R 777 DAKHAOSAT public

# Lệnh chạy
CMD ["npm", "start"]
