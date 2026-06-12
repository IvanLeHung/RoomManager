const { Pool, neonConfig } = require('@neondatabase/serverless');
const { PrismaNeon } = require('@prisma/adapter-neon');
const { PrismaClient } = require('@prisma/client');
const ws = require('ws');

// Cấu hình WebSocket cho môi trường Serverless
neonConfig.webSocketConstructor = ws;
neonConfig.poolQueryViaFetch = true;

const createPrismaClient = () => {
  let connectionString = process.env.POSTGRES_PRISMA_URL || process.env.DATABASE_URL || '';
  
  // Loại bỏ mọi dấu ngoặc kép, nháy đơn và khoảng trắng thừa
  connectionString = connectionString.replace(/^["']|["']$/g, '').trim();
  if (connectionString) {
    const url = new URL(connectionString);
    url.searchParams.delete('channel_binding');
    if (!url.searchParams.has('sslmode')) url.searchParams.set('sslmode', 'require');
    connectionString = url.toString();
  }
  
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

let prisma;

if (process.env.NODE_ENV !== 'production') {
  if (!global.prisma) {
    global.prisma = createPrismaClient();
  }
  prisma = global.prisma;
}

module.exports = {
  prisma,
  createPrismaClient
};
