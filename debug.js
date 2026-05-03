module.exports = (req, res) => {
  const dbUrl = process.env.DATABASE_URL;
  res.status(200).json({
    status: 'ok',
    hasDbUrl: !!dbUrl,
    dbUrlPrefix: dbUrl ? dbUrl.substring(0, 15) + '...' : 'none',
    env: process.env.NODE_ENV,
    nodeVersion: process.version
  });
};
