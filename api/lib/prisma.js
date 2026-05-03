const { Pool, neonConfig } = require('@neondatabase/serverless');
const { PrismaNeon } = require('@prisma/adapter-neon');
const { PrismaClient } = require('@prisma/client');
const ws = require('ws');

// Cấu hình WebSocket cho môi trường Serverless
neonConfig.webSocketConstructor = ws;

const connectionString = process.env.DATABASE_URL;

if (!connectionString) {
  throw new Error("DATABASE_URL is not defined in environment variables");
}

const pool = new Pool({ connectionString: connectionString });
const adapter = new PrismaNeon(pool);
const prisma = new PrismaClient({ adapter });

module.exports = prisma;
