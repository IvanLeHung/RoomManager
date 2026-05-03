const { PrismaClient } = require('@prisma/client');

// Khởi tạo Prisma với URL lấy từ biến môi trường
const prisma = new PrismaClient({
  datasources: {
    db: {
      url: process.env.DATABASE_URL,
    },
  },
});

module.exports = prisma;
