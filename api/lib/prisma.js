const { PrismaClient } = require('@prisma/client');

// Khởi tạo Prisma với URL lấy từ biến môi trường
const prisma = new PrismaClient();

module.exports = prisma;
