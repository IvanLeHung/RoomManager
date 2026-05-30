const prisma = require('./lib/prisma');

module.exports = async (req, res) => {
  if (req.method === 'GET') {
    try {
      const [rooms, tenants, memberships, contracts, receipts, moveOutReports, contractRenewals, roomTransfers, suppliers, expenseCategories, expensePayments] = await Promise.all([
        prisma.room.findMany(),
        prisma.tenant.findMany(),
        prisma.membership.findMany(),
        prisma.contract.findMany(),
        prisma.receipt.findMany(),
        prisma.moveOutReport.findMany(),
        prisma.contractRenewal.findMany(),
        prisma.roomTransfer.findMany(),
        prisma.supplier.findMany(),
        prisma.expenseCategory.findMany(),
        prisma.expensePayment.findMany(),
      ]);

      return res.status(200).json({
        rooms,
        tenants,
        memberships,
        contracts,
        receipts,
        moveOutReports,
        contractRenewals,
        roomTransfers,
        suppliers,
        expenseCategories,
        expensePayments,
      });
    } catch (error) {
      console.error(error);
      return res.status(500).json({ 
        error: 'Failed to fetch data', 
        details: error.message,
        stack: error.stack 
      });
    } finally {
      try {
        await prisma.$disconnect();
      } catch (e) {}
    }
  }

  if (req.method === 'POST') {
    const { type, payload } = req.body;
    
    try {
      if (type === 'full_sync') {
        const { rooms = [], tenants = [], memberships = [], contracts = [], receipts = [], moveOutReports = [], contractRenewals = [], roomTransfers = [] } = payload;

        const pick = (obj, keys) => {
          const res = {};
          keys.forEach(k => {
            if (obj[k] !== undefined) res[k] = obj[k];
          });
          return res;
        };

        const roomKeys = ['id', 'rent', 'deposit', 'cleaning', 'elevator', 'laundry', 'internet', 'electricPrice', 'waterPrice', 'initialElectric', 'initialWater', 'note', 'createdAt'];
        const tenantKeys = ['id', 'name', 'phone', 'cccd', 'cccdDate', 'cccdPlace', 'address', 'licensePlate', 'birthday', 'status', 'lastRoomId', 'note', 'createdAt'];
        const membershipKeys = ['id', 'contractId', 'tenantId', 'roomId', 'role', 'status', 'joinedDate', 'leftDate', 'createdAt'];
        const contractKeys = ['id', 'roomId', 'contractNo', 'startDate', 'endDate', 'signedDate', 'deposit', 'rent', 'paymentCycleDay', 'status', 'noticeDate', 'expectedMoveOutDate', 'actualEndDate', 'endedAt', 'previousEndDate', 'renewedAt', 'terms', 'renewalHistory', 'note', 'createdAt'];
        const receiptKeys = ['id', 'type', 'roomId', 'contractId', 'month', 'rent', 'fixedServices', 'electricOld', 'electricNew', 'electricUsed', 'electricAmount', 'waterOld', 'waterNew', 'waterUsed', 'waterAmount', 'other', 'total', 'paidAmount', 'debt', 'status', 'note', 'createdAt', 'savedAt'];
        const moveOutKeys = ['id', 'contractId', 'roomId', 'actualEndDate', 'electricOld', 'electricNew', 'electricUsed', 'electricAmount', 'waterOld', 'waterNew', 'waterUsed', 'waterAmount', 'depositUsed', 'unpaidRent', 'cleaningFee', 'damageFee', 'otherFee', 'totalIncurred', 'mustCollect', 'mustRefund', 'note', 'createdAt'];
        const renewalKeys = ['id', 'contractId', 'roomId', 'signedDate', 'oldEndDate', 'newStartDate', 'newEndDate', 'oldRent', 'newRent', 'oldDeposit', 'newDeposit', 'note', 'createdAt'];
        const transferKeys = ['id', 'tenantId', 'oldContractId', 'newContractId', 'oldRoomId', 'newRoomId', 'transferDate', 'oldRent', 'newRent', 'oldDeposit', 'newDeposit', 'note', 'createdAt'];

        const supplierKeys = ['id', 'name', 'group', 'defaultCategory', 'phone', 'email', 'address', 'bankName', 'bankAccount', 'bankOwner', 'note', 'createdAt', 'updatedAt'];
        const categoryKeys = ['id', 'name', 'description', 'createdAt', 'updatedAt'];
        const expenseKeys = ['id', 'supplierId', 'categoryId', 'expenseCode', 'recipientName', 'month', 'paymentDate', 'title', 'description', 'totalAmount', 'paidAmount', 'status', 'paymentMethod', 'attachmentUrl', 'note', 'createdAt', 'updatedAt'];

        const ids = (items) => items.filter(i => i && i.id).map(i => i.id);

        // full_sync is a snapshot: records missing from the payload should be removed too.
        // Run deletes only after all upserts succeed so a partial sync cannot wipe existing data.
        const deleteOperations = [
          prisma.expensePayment.deleteMany({ where: { id: { notIn: ids(payload.expensePayments || []) } } }),
          prisma.expenseCategory.deleteMany({ where: { id: { notIn: ids(payload.expenseCategories || []) } } }),
          prisma.supplier.deleteMany({ where: { id: { notIn: ids(payload.suppliers || []) } } }),
          prisma.roomTransfer.deleteMany({ where: { id: { notIn: ids(roomTransfers) } } }),
          prisma.contractRenewal.deleteMany({ where: { id: { notIn: ids(contractRenewals) } } }),
          prisma.moveOutReport.deleteMany({ where: { id: { notIn: ids(moveOutReports) } } }),
          prisma.receipt.deleteMany({ where: { id: { notIn: ids(receipts) } } }),
          prisma.membership.deleteMany({ where: { id: { notIn: ids(memberships) } } }),
          prisma.contract.deleteMany({ where: { id: { notIn: ids(contracts) } } }),
          prisma.tenant.deleteMany({ where: { id: { notIn: ids(tenants) } } }),
          prisma.room.deleteMany({ where: { id: { notIn: ids(rooms) } } }),
        ];

        // Sử dụng Promise.all thay cho $transaction để tránh lỗi "Unable to start a transaction" trên môi trường Serverless
        const operations = [
          ...rooms.filter(i => i && i.id).map(item => { const data = pick(item, roomKeys); return prisma.room.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...tenants.filter(i => i && i.id).map(item => { const data = pick(item, tenantKeys); return prisma.tenant.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...memberships.filter(i => i && i.id).map(item => { const data = pick(item, membershipKeys); return prisma.membership.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...contracts.filter(i => i && i.id).map(item => { const data = pick(item, contractKeys); return prisma.contract.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...receipts.filter(i => i && i.id).map(item => { const data = pick(item, receiptKeys); return prisma.receipt.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...moveOutReports.filter(i => i && i.id).map(item => { const data = pick(item, moveOutKeys); return prisma.moveOutReport.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...contractRenewals.filter(i => i && i.id).map(item => { const data = pick(item, renewalKeys); return prisma.contractRenewal.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...roomTransfers.filter(i => i && i.id).map(item => { const data = pick(item, transferKeys); return prisma.roomTransfer.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...(payload.suppliers || []).filter(i => i && i.id).map(item => { const data = pick(item, supplierKeys); return prisma.supplier.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...(payload.expenseCategories || []).filter(i => i && i.id).map(item => { const data = pick(item, categoryKeys); return prisma.expenseCategory.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...(payload.expensePayments || []).filter(i => i && i.id).map(item => { const data = pick(item, expenseKeys); return prisma.expensePayment.upsert({ where: { id: data.id }, update: data, create: data }); }),
        ];

        // Chạy các lệnh theo nhóm (Batch) với kích thước lớn hơn để đảm bảo hoàn thành trong giới hạn 10s của Vercel
        const BATCH_SIZE = 40;
        for (let i = 0; i < operations.length; i += BATCH_SIZE) {
          const batch = operations.slice(i, i + BATCH_SIZE);
          try {
            await Promise.all(batch);
          } catch (e) {
            console.error(`Lỗi ở batch ${i}:`, e.message);
            // Thử lại từng cái một nếu batch lớn thất bại (để cô lập lỗi)
            if (e.message.includes('terminated') || e.message.includes('timeout')) {
              for (const op of batch) {
                try { await op; } catch (inner) { console.error("Lỗi con:", inner.message); }
              }
            } else {
              throw e;
            }
          }
        }

        await Promise.all(deleteOperations);
        
        return res.status(200).json({ success: true });
      }

      return res.status(400).json({ error: 'Invalid sync type' });
    } catch (error) {
      console.error(error);
      return res.status(500).json({ 
        error: 'Failed to save data', 
        details: error.message,
        stack: error.stack 
      });
    } finally {
      // Đảm bảo đóng kết nối để tránh rò rỉ (leak) trong môi trường serverless
      try {
        await prisma.$disconnect();
      } catch (e) {
        console.error("Lỗi khi đóng prisma:", e);
      }
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
};
