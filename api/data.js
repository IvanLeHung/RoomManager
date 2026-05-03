const prisma = require('./lib/prisma');

module.exports = async (req, res) => {
  if (req.method === 'GET') {
    try {
      const [rooms, tenants, memberships, contracts, receipts, moveOutReports, contractRenewals] = await Promise.all([
        prisma.room.findMany(),
        prisma.tenant.findMany(),
        prisma.membership.findMany(),
        prisma.contract.findMany(),
        prisma.receipt.findMany(),
        prisma.moveOutReport.findMany(),
        prisma.contractRenewal.findMany(),
      ]);

      return res.status(200).json({
        rooms,
        tenants,
        memberships,
        contracts,
        receipts,
        moveOutReports,
        contractRenewals,
      });
    } catch (error) {
      console.error(error);
      return res.status(500).json({ error: 'Failed to fetch data' });
    }
  }

  if (req.method === 'POST') {
    const { type, payload } = req.body;
    
    try {
      // Bulk update/sync logic for specific types or full sync
      // For now, let's support a "full sync" for simplicity in migration
      if (type === 'full_sync') {
        const { rooms, tenants, memberships, contracts, receipts, moveOutReports, contractRenewals } = payload;

        await prisma.$transaction([
          // We use upsert for everything to ensure we don't lose data but also update existing
          ...rooms.map(item => prisma.room.upsert({ where: { id: item.id }, update: item, create: item })),
          ...tenants.map(item => prisma.tenant.upsert({ where: { id: item.id }, update: item, create: item })),
          ...memberships.map(item => prisma.membership.upsert({ where: { id: item.id }, update: item, create: item })),
          ...contracts.map(item => prisma.contract.upsert({ where: { id: item.id }, update: item, create: item })),
          ...receipts.map(item => prisma.receipt.upsert({ where: { id: item.id }, update: item, create: item })),
          ...moveOutReports.map(item => prisma.moveOutReport.upsert({ where: { id: item.id }, update: item, create: item })),
          ...contractRenewals.map(item => prisma.contractRenewal.upsert({ where: { id: item.id }, update: item, create: item })),
        ]);
        
        return res.status(200).json({ success: true });
      }

      return res.status(400).json({ error: 'Invalid sync type' });
    } catch (error) {
      console.error(error);
      return res.status(500).json({ error: 'Failed to save data' });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
