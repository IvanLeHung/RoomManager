const { Pool, neonConfig } = require('@neondatabase/serverless');
const { PrismaNeon } = require('@prisma/adapter-neon');
const { PrismaClient } = require('@prisma/client');
const ws = require('ws');

// Cấu hình WebSocket cho môi trường Serverless
neonConfig.webSocketConstructor = ws;

let prisma;

const initPrisma = () => {
  let connectionString = process.env.DATABASE_URL || '';
  
  // Loại bỏ mọi dấu ngoặc kép, nháy đơn và khoảng trắng thừa
  connectionString = connectionString.replace(/^["']|["']$/g, '').trim();
  
  if (!connectionString) {
    throw new Error("DATABASE_URL is missing or empty");
  }
  
  const pool = new Pool({ 
    connectionString,
    max: 1,
    ssl: true
  });
  const adapter = new PrismaNeon(pool);
  return new PrismaClient({ adapter });
};

if (process.env.NODE_ENV === 'production') {
  prisma = initPrisma();
} else {
  if (!global.prisma) {
    global.prisma = initPrisma();
  }
  prisma = global.prisma;
}

module.exports = prisma;
