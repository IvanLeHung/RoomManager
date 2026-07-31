const { createPrismaClient } = require('./lib/prisma');

async function retryNeonQuery(fn, retries = 2) {
  let lastError;
  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;
      const message = String(error?.message || '');
      const isConnectionError = /terminated|timeout|connection|fetch failed|socket/i.test(message);
      if (!isConnectionError || attempt === retries) break;
      await new Promise(resolve => setTimeout(resolve, 250 * (attempt + 1)));
    }
  }
  throw lastError;
}

async function ensureDatabaseShape(prisma) {
  const statements = [
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "licensePlate" TEXT',
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "fingerprintCode" TEXT',
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "fingerprintStatus" TEXT',
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "birthday" TEXT',
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "status" TEXT',
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "lastRoomId" TEXT',
    'ALTER TABLE "Tenant" ADD COLUMN IF NOT EXISTS "note" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "signedDate" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "noticeDate" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "expectedMoveOutDate" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "actualEndDate" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "endedAt" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "previousEndDate" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "renewedAt" TEXT',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "terms" JSONB',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "renewalHistory" JSONB',
    'ALTER TABLE "Contract" ADD COLUMN IF NOT EXISTS "note" TEXT',
    'ALTER TABLE "Receipt" ADD COLUMN IF NOT EXISTS "adjustmentDueAmount" INTEGER',
    'ALTER TABLE "Receipt" ADD COLUMN IF NOT EXISTS "adjustmentPaidAmount" INTEGER',
    'ALTER TABLE "Receipt" ADD COLUMN IF NOT EXISTS "adjustmentPaidDate" TEXT',
    'ALTER TABLE "Receipt" ADD COLUMN IF NOT EXISTS "adjustmentCreatedAt" TEXT',
    'ALTER TABLE "Receipt" ADD COLUMN IF NOT EXISTS "adjustmentReason" TEXT',
    'ALTER TABLE "MoveOutReport" ADD COLUMN IF NOT EXISTS "settlementMode" TEXT',
    'ALTER TABLE "MoveOutReport" ADD COLUMN IF NOT EXISTS "depositForfeited" INTEGER'
  ];

  for (const sql of statements) {
    await retryNeonQuery(() => prisma.$executeRawUnsafe(sql));
  }
}

module.exports = async (req, res) => {
  const prisma = createPrismaClient();

  if (req.method === 'GET') {
    try {
      await ensureDatabaseShape(prisma);
      const rooms = await retryNeonQuery(() => prisma.room.findMany());
      const tenants = await retryNeonQuery(() => prisma.tenant.findMany());
      const memberships = await retryNeonQuery(() => prisma.membership.findMany());
      const contracts = await retryNeonQuery(() => prisma.contract.findMany());
      const receipts = await retryNeonQuery(() => prisma.receipt.findMany());
      const moveOutReports = await retryNeonQuery(() => prisma.moveOutReport.findMany());
      const contractRenewals = await retryNeonQuery(() => prisma.contractRenewal.findMany());
      const roomTransfers = await retryNeonQuery(() => prisma.roomTransfer.findMany());
      const suppliers = await retryNeonQuery(() => prisma.supplier.findMany());
      const expenseCategories = await retryNeonQuery(() => prisma.expenseCategory.findMany());
      const expensePayments = await retryNeonQuery(() => prisma.expensePayment.findMany());

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
      await prisma.$disconnect().catch(() => {});
    }
  }

  if (req.method === 'POST') {
    const { type, payload } = req.body;
    
    try {
      if (type === 'full_sync') {
        await ensureDatabaseShape(prisma);
        const { rooms = [], tenants = [], memberships = [], contracts = [], receipts = [], moveOutReports = [], contractRenewals = [], roomTransfers = [] } = payload;

        const pick = (obj, keys) => {
          const res = {};
          keys.forEach(k => {
            if (obj[k] !== undefined) res[k] = obj[k];
          });
          return res;
        };

        const roomKeys = ['id', 'rent', 'deposit', 'cleaning', 'elevator', 'laundry', 'internet', 'electricPrice', 'waterPrice', 'initialElectric', 'initialWater', 'note', 'createdAt'];
        const tenantKeys = ['id', 'name', 'phone', 'cccd', 'cccdDate', 'cccdPlace', 'address', 'licensePlate', 'fingerprintCode', 'fingerprintStatus', 'birthday', 'status', 'lastRoomId', 'note', 'createdAt'];
        const membershipKeys = ['id', 'contractId', 'tenantId', 'roomId', 'role', 'status', 'joinedDate', 'leftDate', 'createdAt'];
        const contractKeys = ['id', 'roomId', 'contractNo', 'startDate', 'endDate', 'signedDate', 'deposit', 'rent', 'paymentCycleDay', 'status', 'noticeDate', 'expectedMoveOutDate', 'actualEndDate', 'endedAt', 'previousEndDate', 'renewedAt', 'terms', 'renewalHistory', 'note', 'createdAt'];
        const receiptKeys = ['id', 'type', 'roomId', 'contractId', 'month', 'rent', 'fixedServices', 'electricOld', 'electricNew', 'electricUsed', 'electricAmount', 'waterOld', 'waterNew', 'waterUsed', 'waterAmount', 'other', 'total', 'paidAmount', 'adjustmentDueAmount', 'adjustmentPaidAmount', 'adjustmentPaidDate', 'adjustmentCreatedAt', 'adjustmentReason', 'debt', 'status', 'note', 'createdAt', 'savedAt'];
        const moveOutKeys = ['id', 'contractId', 'roomId', 'actualEndDate', 'electricOld', 'electricNew', 'electricUsed', 'electricAmount', 'waterOld', 'waterNew', 'waterUsed', 'waterAmount', 'depositUsed', 'depositForfeited', 'settlementMode', 'unpaidRent', 'cleaningFee', 'damageFee', 'otherFee', 'totalIncurred', 'mustCollect', 'mustRefund', 'note', 'createdAt'];
        const renewalKeys = ['id', 'contractId', 'roomId', 'signedDate', 'oldEndDate', 'newStartDate', 'newEndDate', 'oldRent', 'newRent', 'oldDeposit', 'newDeposit', 'note', 'createdAt'];
        const transferKeys = ['id', 'tenantId', 'oldContractId', 'newContractId', 'oldRoomId', 'newRoomId', 'transferDate', 'oldRent', 'newRent', 'oldDeposit', 'newDeposit', 'note', 'createdAt'];

        const supplierKeys = ['id', 'name', 'group', 'defaultCategory', 'phone', 'email', 'address', 'bankName', 'bankAccount', 'bankOwner', 'note', 'createdAt', 'updatedAt'];
        const categoryKeys = ['id', 'name', 'description', 'createdAt', 'updatedAt'];
        const expenseKeys = ['id', 'supplierId', 'categoryId', 'expenseCode', 'recipientName', 'month', 'paymentDate', 'title', 'description', 'totalAmount', 'paidAmount', 'status', 'paymentMethod', 'attachmentUrl', 'note', 'createdAt', 'updatedAt'];

        const ids = (items) => items.filter(i => i && i.id).map(i => i.id);
        const nowIso = () => new Date().toISOString();
        const toInt = (value, fallback = 0) => Number.isFinite(Number(value)) ? Math.round(Number(value)) : fallback;
        const toFloat = (value, fallback = 0) => Number.isFinite(Number(value)) ? Number(value) : fallback;
        const prepareRoom = (item) => ({ rent: 0, deposit: 0, cleaning: 0, elevator: 0, laundry: 0, internet: 0, electricPrice: 3800, waterPrice: 32000, ...pick(item, roomKeys) });
        const prepareTenant = (item) => ({ name: 'Chưa cập nhật', status: 'active', ...pick(item, tenantKeys) });
        const prepareMembership = (item) => ({ role: 'primary', status: 'active', createdAt: nowIso(), ...pick(item, membershipKeys) });
        const prepareContract = (item) => ({ startDate: '', endDate: '', deposit: 0, rent: 0, status: 'active', createdAt: nowIso(), ...pick(item, contractKeys) });
        const prepareReceipt = (item) => {
          const data = { type: 'monthly', rent: 0, fixedServices: 0, electricOld: 0, electricNew: 0, electricUsed: 0, electricAmount: 0, waterOld: 0, waterNew: 0, waterUsed: 0, waterAmount: 0, other: 0, total: 0, paidAmount: 0, debt: 0, status: 'Chưa thanh toán', createdAt: nowIso(), ...pick(item, receiptKeys) };
          ['rent', 'fixedServices', 'other', 'total', 'paidAmount', 'adjustmentDueAmount', 'adjustmentPaidAmount', 'debt'].forEach(k => { if (data[k] !== undefined) data[k] = toInt(data[k]); });
          ['electricOld', 'electricNew', 'electricUsed', 'electricAmount', 'waterOld', 'waterNew', 'waterUsed', 'waterAmount'].forEach(k => { data[k] = toFloat(data[k]); });
          return data;
        };
        const prepareMoveOut = (item) => {
          const data = { electricOld: 0, electricNew: 0, electricUsed: 0, electricAmount: 0, waterOld: 0, waterNew: 0, waterUsed: 0, waterAmount: 0, depositUsed: 0, depositForfeited: 0, settlementMode: 'offset_deposit', unpaidRent: 0, cleaningFee: 0, damageFee: 0, mustCollect: 0, mustRefund: 0, createdAt: nowIso(), ...pick(item, moveOutKeys) };
          ['depositUsed', 'depositForfeited', 'unpaidRent', 'cleaningFee', 'damageFee', 'otherFee', 'mustCollect', 'mustRefund'].forEach(k => { if (data[k] !== undefined) data[k] = toInt(data[k]); });
          ['electricOld', 'electricNew', 'electricUsed', 'electricAmount', 'waterOld', 'waterNew', 'waterUsed', 'waterAmount', 'totalIncurred'].forEach(k => { if (data[k] !== undefined) data[k] = toFloat(data[k]); });
          return data;
        };
        const prepareRenewal = (item) => ({ signedDate: '', oldEndDate: '', newStartDate: '', newEndDate: '', oldRent: 0, newRent: 0, oldDeposit: 0, newDeposit: 0, createdAt: nowIso(), ...pick(item, renewalKeys) });
        const prepareTransfer = (item) => ({ transferDate: '', oldRent: 0, newRent: 0, oldDeposit: 0, newDeposit: 0, createdAt: nowIso(), ...pick(item, transferKeys) });
        const prepareSupplier = (item) => ({ name: 'Vãng lai', createdAt: nowIso(), updatedAt: nowIso(), ...pick(item, supplierKeys) });
        const prepareCategory = (item) => ({ name: 'Khác', createdAt: nowIso(), updatedAt: nowIso(), ...pick(item, categoryKeys) });
        const prepareExpense = (item) => ({ categoryId: 'cat_other', month: '', paymentDate: nowIso().slice(0, 10), title: 'Chi phí', totalAmount: 0, paidAmount: 0, status: 'unpaid', paymentMethod: 'transfer', createdAt: nowIso(), updatedAt: nowIso(), ...pick(item, expenseKeys) });

        const runInBatches = async (operations, batchSize = 12) => {
          for (let i = 0; i < operations.length; i += batchSize) {
            const batch = operations.slice(i, i + batchSize);
            await Promise.all(batch.map(op => op()));
          }
        };

        // full_sync is a snapshot: records missing from the payload should be removed too.
        // Deletes run only after all upserts succeed so a partial sync cannot wipe existing data.
        const deleteOperations = [
          () => prisma.expensePayment.deleteMany({ where: { id: { notIn: ids(payload.expensePayments || []) } } }),
          () => prisma.expenseCategory.deleteMany({ where: { id: { notIn: ids(payload.expenseCategories || []) } } }),
          () => prisma.supplier.deleteMany({ where: { id: { notIn: ids(payload.suppliers || []) } } }),
          () => prisma.roomTransfer.deleteMany({ where: { id: { notIn: ids(roomTransfers) } } }),
          () => prisma.contractRenewal.deleteMany({ where: { id: { notIn: ids(contractRenewals) } } }),
          () => prisma.moveOutReport.deleteMany({ where: { id: { notIn: ids(moveOutReports) } } }),
          () => prisma.receipt.deleteMany({ where: { id: { notIn: ids(receipts) } } }),
          () => prisma.membership.deleteMany({ where: { id: { notIn: ids(memberships) } } }),
          () => prisma.contract.deleteMany({ where: { id: { notIn: ids(contracts) } } }),
          () => prisma.tenant.deleteMany({ where: { id: { notIn: ids(tenants) } } }),
          () => prisma.room.deleteMany({ where: { id: { notIn: ids(rooms) } } }),
        ];

        const operations = [
          ...rooms.filter(i => i && i.id).map(item => () => { const data = prepareRoom(item); return prisma.room.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...tenants.filter(i => i && i.id).map(item => () => { const data = prepareTenant(item); return prisma.tenant.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...memberships.filter(i => i && i.id).map(item => () => { const data = prepareMembership(item); return prisma.membership.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...contracts.filter(i => i && i.id).map(item => () => { const data = prepareContract(item); return prisma.contract.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...receipts.filter(i => i && i.id).map(item => () => { const data = prepareReceipt(item); return prisma.receipt.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...moveOutReports.filter(i => i && i.id).map(item => () => { const data = prepareMoveOut(item); return prisma.moveOutReport.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...contractRenewals.filter(i => i && i.id).map(item => () => { const data = prepareRenewal(item); return prisma.contractRenewal.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...roomTransfers.filter(i => i && i.id).map(item => () => { const data = prepareTransfer(item); return prisma.roomTransfer.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...(payload.suppliers || []).filter(i => i && i.id).map(item => () => { const data = prepareSupplier(item); return prisma.supplier.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...(payload.expenseCategories || []).filter(i => i && i.id).map(item => () => { const data = prepareCategory(item); return prisma.expenseCategory.upsert({ where: { id: data.id }, update: data, create: data }); }),
          ...(payload.expensePayments || []).filter(i => i && i.id).map(item => () => { const data = prepareExpense(item); return prisma.expensePayment.upsert({ where: { id: data.id }, update: data, create: data }); }),
        ];

        await runInBatches(operations);
        await runInBatches(deleteOperations, 4);
        
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
      await prisma.$disconnect().catch(() => {});
    }
  }

  await prisma.$disconnect().catch(() => {});
  return res.status(405).json({ error: 'Method not allowed' });
};
