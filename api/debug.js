const prisma = require('./lib/prisma');

module.exports = async (req, res) => {
  try {
    // Thử một truy vấn thực tế vào DB
    const roomCount = await prisma.room.count();
    res.status(200).json({ 
      status: 'success', 
      message: 'Kết nối Database thành công!',
      roomCount 
    });
  } catch (error) {
    res.status(500).json({ 
      status: 'error', 
      message: 'Lỗi kết nối Database!',
      details: error.message,
      stack: error.stack
    });
  }
};
