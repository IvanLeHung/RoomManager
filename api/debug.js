const prisma = require('./lib/prisma');

module.exports = async (req, res) => {
  const dbUrl = process.env.DATABASE_URL || "KHÔNG TÌM THẤY BIẾN";
  
  try {
    const roomCount = await prisma.room.count();
    res.status(200).json({ 
      status: 'success', 
      dbUrlStart: dbUrl.substring(0, 15) + "...",
      roomCount 
    });
  } catch (error) {
    res.status(500).json({ 
      status: 'error', 
      dbUrlStart: dbUrl.substring(0, 15) + "...",
      message: error.message
    });
  } finally {
    try {
      await prisma.$disconnect();
    } catch (e) {}
  }
};
