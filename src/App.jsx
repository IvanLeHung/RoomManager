import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import './styles.css';

const STORAGE_KEY = 'room_manager_qr_app_v1';
const BANK_KEY = 'room_manager_bank_v1';
const PIN_KEY = 'room_manager_pin_v1';

const DEFAULT_BANK = {
  bankName: 'MBBank',
  bankCode: 'MB',
  accountNo: '0123456789',
  accountName: 'DIỆM THỊ BÌNH',
};

const now = new Date();
const INITIAL_MONTH = `${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

const DEFAULT_DATA = {
  rooms: [
    { id: '201', rent: 3700000, deposit: 3700000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '202', rent: 6000000, deposit: 0, cleaning: 84375, elevator: 84375, laundry: 84375, internet: 84375, electricPrice: 3460, waterPrice: 29000, note: 'chủ nhà ở' },
    { id: '301', rent: 4000000, deposit: 4000000, cleaning: 50000, elevator: 50000, laundry: 50000, internet: 50000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '302', rent: 4200000, deposit: 4200000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '303', rent: 4700000, deposit: 4700000, cleaning: 50000, elevator: 50000, laundry: 50000, internet: 50000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '401', rent: 4300000, deposit: 4300000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '402', rent: 3600000, deposit: 3600000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '403', rent: 4700000, deposit: 4800000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '501', rent: 4000000, deposit: 4000000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '502', rent: 3800000, deposit: 3800000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '503', rent: 4500000, deposit: 4500000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '601', rent: 4500000, deposit: 4500000, cleaning: 50000, elevator: 50000, laundry: 50000, internet: 50000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '602', rent: 3800000, deposit: 3800000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '603', rent: 4400000, deposit: 4400000, cleaning: 50000, elevator: 50000, laundry: 50000, internet: 50000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '701', rent: 4400000, deposit: 4400000, cleaning: 50000, elevator: 50000, laundry: 50000, internet: 50000, electricPrice: 3800, waterPrice: 32000, note: 'Giảm 50% phí dịch vụ' },
    { id: '702', rent: 3500000, deposit: 3500000, cleaning: 100000, elevator: 100000, laundry: 100000, internet: 100000, electricPrice: 3800, waterPrice: 32000, note: '' },
    { id: '703', rent: 4000000, deposit: 4000000, cleaning: 50000, elevator: 50000, laundry: 50000, internet: 50000, electricPrice: 3800, waterPrice: 32000, note: '' },
  ],
  tenants: [],
  memberships: [],
  contracts: [],
  receipts: [],
  moveOutReports: [],
};

function safeRead(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw || raw === 'undefined' || raw === 'null') return fallback;
    return JSON.parse(raw);
  } catch (e) {
    console.error('Lỗi đọc dữ liệu:', e);
    return fallback;
  }
}

function uid(prefix) {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function diffMonths(d1, d2) {
  const start = parseDateFlexible(d1);
  const end = parseDateFlexible(d2);
  if (!start || !end) return 0;
  let months = (end.getFullYear() - start.getFullYear()) * 12;
  months -= start.getMonth();
  months += end.getMonth();
  return months <= 0 ? 0 : months;
}

function formatMoney(value) {
  return new Intl.NumberFormat('vi-VN', {
    style: 'currency',
    currency: 'VND',
  }).format(value);
}

function numberToWords(number) {
  if (number === 0) return 'Không đồng';
  const units = ['', ' nghìn', ' triệu', ' tỷ', ' nghìn tỷ', ' triệu tỷ'];
  const readThreeDigits = (num) => {
    const digits = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
    let a = Math.floor(num / 100);
    let b = Math.floor((num % 100) / 10);
    let c = num % 10;
    let res = '';
    if (a > 0) res += digits[a] + ' trăm ';
    if (b > 1) res += digits[b] + ' mươi ';
    if (b === 1) res += 'mười ';
    if (a > 0 && b === 0 && c > 0) res += 'lẻ ';
    if (c === 5 && b > 0) res += 'lăm ';
    else if (c === 1 && b > 1) res += 'mốt ';
    else if (c > 0 || (a === 0 && b === 0)) res += digits[c];
    return res;
  };
  let res = '';
  let unitIdx = 0;
  let temp = Math.abs(number);
  while (temp > 0) {
    let three = temp % 1000;
    if (three > 0) res = readThreeDigits(three) + units[unitIdx] + ' ' + res;
    temp = Math.floor(temp / 1000);
    unitIdx++;
  }
  return res.trim().charAt(0).toUpperCase() + res.trim().slice(1) + ' đồng';
}

function parseDateFlexible(value) {
  if (!value) return null;
  const text = String(value);
  if (text.includes('-')) return new Date(`${text}T00:00:00`);
  if (text.includes('/')) {
    const parts = text.split('/');
    if (parts.length === 3) return new Date(`${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}T00:00:00`);
  }
  return null;
}

function isValidDateString(value) {
  if (!value) return false;
  const date = new Date(value + 'T00:00:00');
  return !Number.isNaN(date.getTime());
}

function addMonthsToDate(value, months) {
  if (!isValidDateString(value)) return '';
  const date = new Date(value + 'T00:00:00');
  date.setMonth(date.getMonth() + months);
  return date.toISOString().slice(0, 10);
}

function getDaysUntil(dateStr) {
  const date = parseDateFlexible(dateStr);
  if (!date) return null;
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return Math.ceil((date.getTime() - now.getTime()) / 86400000);
}

function getRoomStatusInfo(data, roomId) {
  const contract = (data.contracts || []).find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice' || c.status === 'moving_out'));
  if (!contract) return { label: 'Trống', color: 'gray', contract: null };
  if (contract.status === 'notice') return { label: 'Báo chuyển', color: 'orange', contract };
  if (contract.status === 'moving_out') return { label: 'Đang tất toán', color: 'blue', contract };
  return { label: 'Đang ở', color: 'green', contract };
}

function getDashboardStats(data, currentMonth) {
  const totalRooms = data.rooms.length;
  const activeContracts = (data.contracts || []).filter(c => c.status === 'active' || c.status === 'notice');
  const occupiedRooms = activeContracts.length;
  const vacantRooms = totalRooms - occupiedRooms;
  
  const expiringContracts = activeContracts.filter(c => {
    const days = getDaysUntil(c.endDate);
    return days !== null && days <= 30;
  });

  const currentTenants = (data.memberships || []).filter(m => m.status === 'active').length;
  
  const monthReceipts = (data.receipts || []).filter(r => r.month === currentMonth && r.type === 'monthly');
  const unpaidReceipts = monthReceipts.filter(r => r.status === 'Chưa thanh toán' || r.status === 'Nợ một phần');
  
  const totalDebt = (data.receipts || []).reduce((sum, r) => {
    const debt = Number(r.total || 0) - Number(r.paidAmount || 0);
    return sum + (debt > 0 ? debt : 0);
  }, 0);

  const notifyingMoveOut = activeContracts.filter(c => c.status === 'notice');

  return {
    totalRooms,
    occupiedRooms,
    vacantRooms,
    expiringContracts,
    currentTenants,
    totalReceipts: monthReceipts.length,
    unpaidReceipts,
    totalDebt,
    notifyingMoveOut
  };
}

function getPrimaryTenantByContract(data, contractId) {
  if (!contractId) return null;
  const primaryMember = (data.memberships || []).find(m => m.contractId === contractId && m.role === 'primary');
  if (!primaryMember) return null;
  return (data.tenants || []).find(t => t.id === primaryMember.tenantId) || null;
}

function onlyDigits(value) {
  return String(value || '').split('').filter(c => c >= '0' && c <= '9').join('');
}

function parseLocaleNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  
  let str = String(value).trim();
  if (!str) return 0;

  // If there's only one separator and it's followed by 1 or 2 digits, treat it as decimal
  // Otherwise, if there are dots and a comma at the end, comma is decimal
  // Simple approach for this app: if it contains a comma, replace all dots then replace comma with dot.
  // If it contains only dots, check if it looks like a thousands separator or decimal.
  // Given room indices are usually > 1000 and have 1 decimal, we can assume:
  // if count of dots is 1 and it's near the end, it's decimal.
  
  // Revised simple & robust approach:
  // 1. Remove all spaces
  str = str.replace(/\s/g, '');
  
  // 2. If it has both , and . -> comma is almost always the decimal in VN or dot is decimal in Intl.
  // We'll treat the LAST one as the decimal separator.
  const lastDot = str.lastIndexOf('.');
  const lastComma = str.lastIndexOf(',');
  
  if (lastDot > lastComma) {
    // Dot is later, treat as decimal. Remove all other separators.
    return parseFloat(str.replace(/,/g, '')) || 0;
  } else if (lastComma > lastDot) {
    // Comma is later, treat as decimal. Remove all dots, then replace comma with dot.
    return parseFloat(str.replace(/\./g, '').replace(/,/g, '.')) || 0;
  }
  
  // Only one type or none. 
  return parseFloat(str) || 0;
}

function formatLocaleNumber(value) {
  if (value === null || value === undefined || value === '' || value === '?') return String(value || '—');
  const num = typeof value === 'number' ? value : parseFloat(String(value).replace(/,/g, '.'));
  if (isNaN(num)) return '—';
  return new Intl.NumberFormat('vi-VN').format(num);
}

function DecimalInput({ value, onChange, style, className }) {
  const [displayValue, setDisplayValue] = useState(String(value || ''));

  useEffect(() => {
    // Update display value when prop value changes from outside (e.g. from recalculateReceipt)
    // but only if it's not the same number to avoid cursor jumping
    const currentNum = parseLocaleNumber(displayValue);
    if (value !== currentNum) {
      setDisplayValue(String(value || ''));
    }
  }, [value]);

  const handleChange = (e) => {
    const val = e.target.value;
    // Allow digits, dots, commas, and spaces
    if (/^[0-9.,\s]*$/.test(val)) {
      setDisplayValue(val);
      const num = parseLocaleNumber(val);
      onChange(num);
    }
  };

  return (
    <input 
      type="text" 
      value={displayValue} 
      onChange={handleChange} 
      style={style} 
      className={className} 
    />
  );
}

function transferContent(receipt) {
  return `P${receipt.roomId} T${onlyDigits(receipt.month)}`;
}

function buildVietQrUrl(bankInfo, receipt) {
  if (!bankInfo.bankCode || !bankInfo.accountNo || !receipt) return '';
  const params = new URLSearchParams({
    amount: String(Math.round(Number(receipt.total || 0))),
    addInfo: transferContent(receipt),
    accountName: bankInfo.accountName || '',
  });
  return `https://img.vietqr.io/image/${bankInfo.bankCode}-${bankInfo.accountNo}-compact2.png?${params.toString()}`;
}

function fixedServiceTotal(room) {
  if (!room) return 0;
  return Number(room.cleaning || 0) + Number(room.elevator || 0) + Number(room.laundry || 0) + Number(room.internet || 0);
}

function getElectricOld(receipt) {
  return Number(receipt.electricOld ?? receipt.electricStart ?? 0);
}

function getElectricNew(receipt) {
  const old = getElectricOld(receipt);
  return Number(receipt.electricNew ?? receipt.electricEnd ?? old);
}

function getWaterOld(receipt) {
  return Number(receipt.waterOld ?? receipt.waterStart ?? 0);
}

function getWaterNew(receipt) {
  const old = getWaterOld(receipt);
  return Number(receipt.waterNew ?? receipt.waterEnd ?? old);
}

function receiptCode(receipt) {
  return "PT-" + receipt.roomId + "-" + onlyDigits(receipt.month);
}

function parseMonthValue(month) {
  if (!month) return 0;
  const [m, y] = month.split('/').map(Number);
  return y * 12 + m;
}

function getPreviousReceipt(receipts, roomId, contractId, currentMonth) {
  const currentValue = parseMonthValue(currentMonth);
  return (receipts || [])
    .filter(r =>
      r.roomId === roomId &&
      (contractId ? r.contractId === contractId : true) &&
      r.type === 'monthly' &&
      parseMonthValue(r.month) < currentValue
    )
    .sort((a, b) => parseMonthValue(b.month) - parseMonthValue(a.month))[0] || null;
}

function getPreviousReceiptByRoom(receipts, roomId, currentMonth) {
  const currentValue = parseMonthValue(currentMonth);
  return (receipts || [])
    .filter(r =>
      r.roomId === roomId &&
      r.type === 'monthly' &&
      parseMonthValue(r.month) < currentValue
    )
    .sort((a, b) => parseMonthValue(b.month) - parseMonthValue(a.month))[0] || null;
}

function createMonthlyReceipt(room, contract, previousReceipt, month) {
  const electricOld = previousReceipt
    ? Number(previousReceipt.electricNew ?? previousReceipt.electricEnd ?? 0)
    : Number(room.initialElectric || room.electricStart || 0);

  const waterOld = previousReceipt
    ? Number(previousReceipt.waterNew ?? previousReceipt.waterEnd ?? 0)
    : Number(room.initialWater || room.waterStart || 0);

  const rent = Number(contract?.rent || room.rent || 0);
  const fixedServices = fixedServiceTotal(room);

  return recalculateReceipt({
    id: uid("receipt"),
    type: "monthly",
    roomId: room.id,
    contractId: contract?.id || "",
    month,
    rent,
    fixedServices,
    electricOld,
    electricNew: electricOld,
    electricUsed: 0,
    electricAmount: 0,
    waterOld,
    waterNew: waterOld,
    waterUsed: 0,
    waterAmount: 0,
    other: 0,
    total: rent + fixedServices,
    paidAmount: 0,
    debt: rent + fixedServices,
    status: "Chưa thanh toán",
    note: "Vui lòng thanh toán trong vòng 5 ngày kể từ ngày nhận phiếu. Xin cảm ơn!",
    createdAt: new Date().toISOString()
  }, room);
}

function recalculateReceipt(receipt, room) {
  const electricOld = Number(receipt.electricOld ?? receipt.electricStart ?? 0);
  const electricNew = Number(receipt.electricNew ?? receipt.electricEnd ?? electricOld);

  const waterOld = Number(receipt.waterOld ?? receipt.waterStart ?? 0);
  const waterNew = Number(receipt.waterNew ?? receipt.waterEnd ?? waterOld);

  const electricUsed = Math.max(0, electricNew - electricOld);
  const waterUsed = Math.max(0, waterNew - waterOld);

  // Round used values to 2 decimal places to avoid floating point issues
  const electricUsedFixed = Math.round(electricUsed * 100) / 100;
  const waterUsedFixed = Math.round(waterUsed * 100) / 100;

  const electricAmount = electricUsedFixed * Number(room.electricPrice || 0);
  const waterAmount = waterUsedFixed * Number(room.waterPrice || 0);

  const total =
    Number(receipt.rent || 0) +
    Number(receipt.fixedServices || 0) +
    electricAmount +
    waterAmount +
    Number(receipt.other || 0);

  const paidAmount = Number(receipt.paidAmount || 0);
  const debt = Math.max(0, total - paidAmount);

  let status = "Chưa thanh toán";
  if (paidAmount >= total && total > 0) status = "Đã thanh toán";
  else if (paidAmount > 0 && paidAmount < total) status = "Nợ một phần";

  return {
    ...receipt,
    electricOld,
    electricNew,
    electricUsed: electricUsedFixed,
    electricAmount,
    waterOld,
    waterNew,
    waterUsed: waterUsedFixed,
    waterAmount,
    total,
    debt,
    status
  };
}

function App() {
  const [unlocked, setUnlocked] = useState(() => sessionStorage.getItem('room_app_unlocked') === 'true');
  if (!unlocked) return <LoginScreen onUnlock={() => setUnlocked(true)} />;
  return <AppMain />;
}

function LoginScreen({ onUnlock }) {
  const [pin, setPin] = useState('');
  const savedPin = safeRead(PIN_KEY, '1234');
  const handleSubmit = (e) => {
    e.preventDefault();
    if (pin === savedPin) {
      sessionStorage.setItem('room_app_unlocked', 'true');
      onUnlock();
    } else {
      alert('Mã PIN không chính xác!');
      setPin('');
    }
  };
  return (
    <div className="login-screen">
      <div className="login-card liquid-glass">
        <h1 style={{ fontSize: '28px', marginBottom: '8px' }}>🔐 Quản lý Phòng</h1>
        <p className="muted" style={{ marginBottom: '24px' }}>Vui lòng nhập mã PIN để tiếp tục</p>
        <form onSubmit={handleSubmit} className="stack" style={{ gap: '16px' }}>
          <input type="password" value={pin} onChange={(e) => setPin(e.target.value)} placeholder="Mã PIN" autoFocus style={{ width: '100%', textAlign: 'center', fontSize: '20px', height: '60px' }} />
          <button type="submit" className="primary-btn wide" style={{ height: '50px' }}>Mở khóa hệ thống</button>
        </form>
      </div>
    </div>
  );
}

function ContractPreview({ contract, room, tenants, bankInfo, report, type = 'main', onClose }) {
  const primaryTenant = tenants.find(t => t.role === 'primary') || tenants[0] || {};
  const today = new Date().toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const startDay = contract.startDate ? contract.startDate.split('-').reverse().join('/') : '.../.../....';
  const endDay = contract.endDate ? contract.endDate.split('-').reverse().join('/') : '................';
  const duration = diffMonths(contract.startDate, contract.endDate);
  const isMain = type === 'main';
  const isLiquidation = type === 'liquidation';
  const isRenewal = type === 'renewal';

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2" style={{ maxWidth: '950px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header no-print">
          <div>
            <h2 style={{ margin: 0 }}>{isLiquidation ? 'Biên bản tất toán' : isRenewal ? 'Phụ lục gia hạn' : 'Hợp đồng thuê'} P{room.id}</h2>
            <p className="muted small">Tài liệu đầy đủ 12 Điều khoản & 3 Phụ lục (Khổ A4)</p>
          </div>
          <div className="btn-group">
            <button className="primary-btn" onClick={() => window.print()}>🖨️ In tài liệu</button>
            <button className="secondary-btn" onClick={onClose}>Đóng</button>
          </div>
        </div>
        <div className="contract-paper">
          {(isMain || isRenewal) && (
            <>
              <div className="contract-header-text">
                <h1>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h1>
                <h2>Độc lập – Tự do – Hạnh phúc</h2>
                <p style={{ fontStyle: 'italic' }}>Hà Nội, ngày 01 tháng 03 năm 2026</p>
              </div>
              <h1 style={{ textAlign: 'center', margin: '30px 0 20px' }}>HỢP ĐỒNG THUÊ PHÒNG</h1>
              {isMain && (
                <div className="contract-section">
                  <p>- Căn cứ Bộ luật Dân sự số 91/2015/QH13 ngày 24/11/2015;</p>
                  <p>- Căn cứ Luật Thương mại số 36/2005/QH11 ngày 14/06/2005;</p>
                  <p>- Căn cứ nhu cầu và sự thỏa thuận của các Bên.</p>
                  <p style={{ marginTop: '10px' }}>Hôm nay, ngày 01 tháng 11 năm 2025, chúng tôi gồm:</p>
                </div>
              )}
              <div className="contract-section">
                <h3>BÊN CHO THUÊ (Bên A)</h3>
                <p>Bà: <b>DIỆM THỊ BÌNH</b></p>
                <p>CCCD Số: <b>019169000011</b> &nbsp;&nbsp;&nbsp; Ngày cấp: <b>07/04/2021</b></p>
                <p>Tại: <b>Cục Cảnh sát Quản lý hành chính về trật tự xã hội</b></p>
                <p>SĐT/Zalo: <b>056.201.1613</b></p>
                <p>Tài khoản ngân hàng: <b>{bankInfo.bankName} - {bankInfo.accountNo} - {bankInfo.accountName}</b></p>
                <h3>BÊN THUÊ (Bên B)</h3>
                <p>Ông/Bà: <b>{primaryTenant.name}</b></p>
                <p>CCCD số: <b>{primaryTenant.cccd || '................'}</b> &nbsp;&nbsp;&nbsp; Ngày cấp: <b>{primaryTenant.cccdDate || '................'}</b></p>
                <p>Cơ quan cấp: <b>{primaryTenant.cccdPlace || '................'}</b></p>
                <p>Nơi ĐKTT: <b>{primaryTenant.address || '................'}</b></p>
                <p>Địa chỉ liên hệ/SĐT/Zalo: <b>{primaryTenant.phone || '................'}</b></p>
                <p style={{ marginTop: '10px' }}>Bên A và Bên B sau đây gọi chung là "Các Bên".</p>
              </div>
              <div className="main-articles">
                <h3>ĐIỀU 1. TÀI SẢN THUÊ VÀ MỤC ĐÍCH SỬ DỤNG</h3>
                <p>1.1. Bên A cho Bên B thuê phòng số <b>{room.id}</b> tại Số 28, ngách 1, ngõ 162 Khương Đình, phường Khương Đình, TP Hà Nội (sau đây gọi “Phòng thuê”) để làm nơi ở.</p>
                <p>1.2. Số người ở tối đa: 03 người/phòng. Nếu thêm người, Bên B phải báo trước và được Bên A đồng ý bằng tin nhắn/văn bản; khách ở quá 03 ngày sẽ tính thêm 50.000 đồng/người và Bên B chịu chi phí phạt hành chính nếu khách chưa khai báo tạm trú nếu ở ≥15 ngày.</p>
                <p>1.3. Bên A cam kết Phòng thuê thuộc quyền sở hữu/sử dụng hợp pháp của Bên A; nếu phát sinh tranh chấp liên quan tài sản cho thuê, Bên A chịu trách nhiệm trước pháp luật.</p>
                <p>1.4. Mỗi phòng được bố trí tối đa 02 (hai) xe máy. Trường hợp có xe thứ 03, Bên B tự gửi ngoài. Tất cả xe để tại tầng 1 phải đăng ký biển số/loại xe với Bên A. Khi thay đổi xe, Bên B phải thông báo cập nhật trước khi đưa xe vào gửi.</p>
                <p>1.5. Không để qua đêm đối với xe không thuộc cư dân trong cùng tòa nhà.</p>
                <p>1.6. Nếu phát hiện vi phạm các quy định 1.5, Bên A có quyền yêu cầu di chuyển xe ngay và áp dụng mức phạt 300.000 đồng cho mỗi xe/mỗi lần (hoặc mỗi đêm) vi phạm.</p>
                <p>1.7. Nhằm đảm bảo PCCC tại tòa nhà, Bên A cấm tuyệt đối sạc mọi loại xe máy điện/xe đạp điện/scooter điện và pin, ắc-quy rời của các phương tiện này trong mọi khu vực tòa nhà (phòng, hành lang, cầu thang, khu kỹ thuật, bãi để xe, khu vực chung). Bên A được quyền ngắt nguồn, yêu cầu dừng sạc/di chuyển ngay; vi phạm bị phạt từ 300.000 – 500.000 đồng/lần, tái phạm có thể chấm dứt quyền gửi xe hoặc chấm dứt Hợp đồng. Mọi rủi ro, thiệt hại phát sinh do sạc trái quy định do Bên B tự chịu và bồi thường.</p>
                <h3>ĐIỀU 2. BÀN GIAO VÀ HIỆN TRẠNG</h3>
                <p>2.1. Thời điểm bàn giao: <b>{startDay}</b></p>
                <p>2.2. Hai Bên lập Phụ lục 01 – Biên bản bàn giao (kèm ảnh/video), ghi rõ thiết bị – hiện trạng – chỉ số điện/nước đầu kỳ – số thẻ/vân tay.</p>
                <p>2.3. Kể từ thời điểm bàn giao, Bên B có toàn quyền sử dụng Phòng thuê theo Hợp đồng.</p>
                <h3>ĐIỀU 3. THỜI HẠN THUÊ</h3>
                <p>3.1. Thời hạn thuê: <b>{duration}</b> tháng kể từ ngày bàn giao.</p>
                <p>3.2. Từ ngày <b>{startDay}</b> Đến hết ngày <b>{endDay}</b></p>
                <p>3.3. Hết thời hạn thuê, Hợp đồng tự động gia hạn theo tháng với điều khoản không đổi, trừ khi một Bên thông báo chấm dứt trước ≥30 ngày bằng văn bản/tin nhắn (Zalo/SMS).</p>
                <h3>ĐIỀU 4. TIỀN ĐẶT CỌC</h3>
                <p>4.1. Mức đặt cọc: 01 tháng tiền thuê, tương đương <b>{formatMoney(contract.deposit)}</b> (bằng chữ: <b>{numberToWords(contract.deposit)}</b>), nộp ngay khi ký.</p>
                <p>4.2. Không dùng cọc để trừ tiền thuê trừ khi Hai Bên đồng ý bằng văn bản/tin nhắn.</p>
                <p>4.3. Nếu Bên B đơn phương chấm dứt trái thỏa thuận báo trước theo Điều 9, mất cọc. Nếu Bên A đơn phương chấm dứt trái thỏa thuận báo trước, hoàn cọc và bồi thường thêm số tiền bằng đúng tiền cọc.</p>
                <p>4.4. Khi trả phòng, Hai Bên lập Biên bản kiểm tra hiện trạng. Bên A chỉ được khấu trừ cọc đối với hư hỏng do lỗi Bên B có bảng kê chi phí/hóa đơn. Bên B Hoàn cọc trong 07 (bảy) ngày làm việc kể từ khi nhận đủ chìa khóa/thẻ, bàn giao xong và quyết toán công nợ từ Bên B.</p>
                <h3>ĐIỀU 5. GIÁ THUÊ VÀ PHÍ DỊCH VỤ</h3>
                <p>5.1. Tiền thuê: <b>{formatMoney(contract.rent)}</b>/tháng (bằng chữ: <b>{numberToWords(contract.rent)}</b>).</p>
                <p>5.2. Phí dịch vụ (chưa gồm trong tiền thuê), thanh toán theo thực tế sử dụng:</p>
                <p>- Điện: 3.800đ/kWh, theo chỉ số công tơ (ghi đầu/cuối kỳ trong phiếu thu).</p>
                <p>- Nước: 32.000đ/m³, theo chỉ số công tơ (ghi đầu/cuối kỳ trong phiếu thu).</p>
                <p>- Mạng internet: <b>{formatMoney(room.internet)}</b>/phòng/tháng.</p>
                <p>- Vệ sinh rác & dịch vụ chung: <b>{formatMoney(room.cleaning)}</b>/phòng/tháng.</p>
                <p>- Phí bảo trì thang máy: <b>{formatMoney(room.elevator)}</b>/phòng/tháng.</p>
                <p>- Phí sử dụng máy giặt chung: <b>{formatMoney(room.laundry)}</b>/phòng/tháng.</p>
                <p>5.3. Khi đơn giá đầu vào (điện/nước/dịch vụ) thay đổi, Bên A báo trước ≥15 ngày và gửi bảng tính kèm chứng từ.</p>
                <h3>ĐIỀU 6. PHƯƠNG THỨC VÀ HẠN THANH TOÁN</h3>
                <p>6.1. Kỳ thanh toán: 01 (một) tháng/lần.</p>
                <p>6.2. Hạn thanh toán thống nhất: chậm nhất ngày 10 của tháng thuê.</p>
                <p>6.3. Hình thức: tiền mặt hoặc chuyển khoản VND vào tài khoản Bên A; thanh toán được coi là hoàn thành khi tiền ghi có vào tài khoản Bên A.</p>
                <p>6.4. Nợ quá hạn & chế tài:</p>
                <p>- Quá hạn {'>'}03 ngày: phạt 100.000 đ/kỳ;</p>
                <p>- Sau đó, mỗi 02 ngày quá hạn cộng thêm 50.000 đ, tối đa 500.000 đ/kỳ;</p>
                <p>- Tổng phạt không quá 1 tháng tiền thuê/kỳ.</p>
                <p>- Quá hạn đủ 10 ngày kể từ hạn, Bên A có quyền đơn phương chấm dứt Hợp đồng (Điều 9) sau khi gửi Thông báo chấm dứt tối thiểu 05 ngày.</p>
                <p>- Bên A không áp dụng biện pháp cắt điện, cắt nước, niêm phong phòng, xóa vân tay để cưỡng chế.</p>
                <h3>ĐIỀU 7. QUYỀN VÀ NGHĨA VỤ BÊN A</h3>
                <p>7.1. Quyền: Yêu cầu Bên B thanh toán đủ, đúng hạn; yêu cầu bồi thường hư hỏng do lỗi Bên B. Vào phòng 08:00–20:00, báo trước ≥24 giờ để kiểm tra/sửa chữa; trường hợp khẩn cấp (cháy nổ, rò rỉ, nguy hiểm…), được vào ngay và thông báo sau. Chấm dứt Hợp đồng theo Điều 9 khi Bên B vi phạm.</p>
                <p>7.2. Nghĩa vụ: Bàn giao đúng hạn, đúng hiện trạng; đảm bảo việc cho thuê hợp pháp. Sửa chữa kịp thời hư hỏng thuộc không gian chung. Đảm bảo quyền sử dụng liên tục, an toàn, riêng tư; không xâm phạm tài sản. Hỗ trợ khai báo tạm trú. Thông báo trước ≥30 ngày nếu muốn điều chỉnh giá thuê/nội quy hoặc chấm dứt. Bảo mật CCCD, số điện thoại, dữ liệu vân tay.</p>
                <h3>ĐIỀU 8. QUYỀN VÀ NGHĨA VỤ BÊN B</h3>
                <p>8.1. Quyền: Nhận bàn giao đúng thời gian/hiện trạng; sử dụng Phòng thuê để ở; yêu cầu Bên A sửa chữa sự cố không do lỗi mình; được gia hạn theo Điều 3; được hoàn cọc theo Điều 4; tháo dỡ tài sản cá nhân khi chấm dứt.</p>
                <p>8.2. Nghĩa vụ: Thanh toán đủ, đúng hạn tiền thuê và dịch vụ. Sử dụng đúng mục đích; không dùng làm nơi kinh doanh, kho tập kết hàng, hoặc mục đích trái pháp luật. Bảo quản tài sản; bồi thường hư hỏng do lỗi mình. Không đục phá kết cấu khi chưa được đồng ý. Trả phòng đúng nguyên trạng. Cung cấp hồ sơ (CCCD/hộ chiếu) trong 24 giờ để khai báo tạm trú.</p>
                <h3>ĐIỀU 9. GIA HẠN – CHẤM DỨT HỢP ĐỒNG</h3>
                <p>9.1. Hợp đồng chấm dứt khi hết hạn và có thông báo chấm dứt trước ≥30 ngày theo Điều 3; nếu không, gia hạn theo tháng.</p>
                <p>9.2. Hợp đồng chấm dứt trước hạn nếu căn nhà hư hỏng nặng do bất khả kháng (hỏa hoạn, thiên tai, dịch bệnh, cưỡng chế nhà nước…) khiến Bên B không thể tiếp tục ở bình thường.</p>
                <p>9.3. Một Bên vi phạm hợp đồng, Bên còn lại có quyền đơn phương chấm dứt sau khi thông báo nêu rõ lý do; Bên vi phạm bồi thường cho Bên còn lại số tiền tương đương tiền cọc.</p>
                <p>9.4. Nếu các Bên chấm dứt trước hạn không thuộc 9.1–9.3, Bên đơn phương phải bồi thường số tiền tương đương tiền cọc.</p>
                <p>9.5. Trường hợp Bên B muốn sang nhượng hợp đồng: phải đã ở ≥03 tháng, báo trước 20–30 ngày, được Bên A đồng ý; phí sang nhượng: 200.000 đ.</p>
                <h3>ĐIỀU 10. BẢO MẬT & DỮ LIỆU RA/VÀO</h3>
                <p>10.1. Bên A không được tự ý di chuyển/thu giữ tài sản của Bên B. Chỉ xử lý trong trường hợp khẩn cấp hoặc khi Bên B bỏ lại tài sản sau thời hạn thông báo.</p>
                <p>10.2. Bên A chỉ thu thập vân tay để quản lý ra/vào; không dùng cho mục đích khác. Xóa dữ liệu khi chấm dứt hợp đồng.</p>
                <h3>ĐIỀU 11. GIẢI QUYẾT TRANH CHẤP</h3>
                <p>Khi phát sinh bất đồng liên quan Hợp đồng, các Bên trước hết thương lượng. Nếu thương lượng không thành, tranh chấp sẽ do Tòa án nhân dân có thẩm quyền giải quyết. Hợp đồng và mọi tranh chấp phát sinh được điều chỉnh bởi pháp luật Việt Nam.</p>
                <h3>ĐIỀU 12. ĐIỀU KHOẢN CHUNG</h3>
                <p>Hợp đồng có hiệu lực từ ngày ký. Mọi sửa đổi phải lập bằng văn bản. Hợp đồng lập 02 bản có giá trị pháp lý như nhau.</p>
              </div>
              <div className="signature-row">
                <div><p><b>BÊN CHO THUÊ (Bên A)</b></p><div className="signature-space"></div><p><b>DIỆM THỊ BÌNH</b></p></div>
                <div><p><b>BÊN THUÊ (Bên B)</b></p><div className="signature-space"></div><p><b>{primaryTenant.name}</b></p></div>
              </div>
            </>
          )}
          {isMain && (
            <div className="appendix-container">
              <div className="appendix-page">
                <h1 style={{ textAlign: 'center', fontSize: '18px', fontWeight: 'bold' }}>PHỤ LỤC 01 – BIÊN BẢN BÀN GIAO</h1>
                <p>Ngày bàn giao: <b>{startDay}</b>; Phòng số: <b>{room.id}</b></p>
                <p>Chỉ số điện cũ: <b>{room.electricOld || room.electricStart || room.initialElectric || '.......'}</b> kWh; Chỉ số nước cũ: <b>{room.waterOld || room.waterStart || room.initialWater || '.......'}</b> m³</p>
                <table className="contract-table">
                  <thead>
                    <tr><th>STT</th><th>DANH MỤC THIẾT BỊ</th><th>SỐ LƯỢNG</th><th>HIỆN TRẠNG</th></tr>
                  </thead>
                  <tbody>
                    {[
                      'Điều hòa + Điều khiển', 'Giường', 'Tủ quần áo', 'Chăn', 'Ga', 'Gối', 'Đệm', 'Sofa', 'Bàn trà', 
                      'Bàn trang điểm + Ghế', 'Bóng điện chiếu sáng', 'Hút mùi', 'Tủ lạnh', 'Rèm vải', 'Modem Wi-Fi', 
                      'Tranh treo tường', 'Mặt nạ phòng độc', 'Chìa khóa phòng', 'Gương soi', 'Bình nóng lạnh', 'Lavabo'
                    ].map((item, i) => (
                      <tr key={i}>
                        <td style={{ textAlign: 'center' }}>{i + 1}</td>
                        <td>{item}</td>
                        <td style={{ textAlign: 'center' }}>01</td>
                        <td>☐ tốt &nbsp; ☐ xước &nbsp; ☐ hỏng</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="appendix-page">
                <h1 style={{ textAlign: 'center', fontSize: '18px', fontWeight: 'bold' }}>PHỤ LỤC 03 – NỘI QUY TÒA NHÀ</h1>
                <div style={{ fontSize: '11px', lineHeight: '1.4' }}>
                  {[
                    '1. Tuân thủ pháp luật: Nghiêm cấm đánh nhau, cờ bạc, mại dâm, ma túy...',
                    '2. PCCC: CẤM TUYỆT ĐỐI SẠC XE ĐIỆN trong mọi khu vực tòa nhà.',
                    '3. Khai báo tạm trú: Cung cấp giấy tờ trong 24 giờ để làm thủ tục.',
                    '4. Thoát nước: Không đổ rác, thức ăn thừa xuống bồn cầu/thoát sàn.',
                    '5. Vệ sinh: Không để rác hành lang, cầu thang. Bỏ rác đúng nơi quy định.',
                    '6. Tiếng ồn: Giữ trật tự sau 22:00. Tắt máy xe khi vào sau 22:30.',
                    '7. Ra vào: Đóng cửa cổng cẩn thận. Không mở cửa cho người lạ.',
                    '8. Kết cấu: Không tự ý khoan, đục, vẽ lên tường khi chưa đồng ý.',
                    '9. Bồi thường: Chịu trách nhiệm hư hại do mình gây ra.',
                    '10. Tài sản: Tự bảo quản tài sản cá nhân.',
                    '11. Trung thực: Không lấy đồ của người khác.',
                    '12. Vật nuôi: Không nuôi động vật gây mất vệ sinh/an toàn.',
                    '13. Kinh doanh: Không làm nơi buôn bán, kho hàng.',
                    '14. Vận chuyển: Không chở quá tải thang máy.',
                    '15. Kỹ thuật: Không tự ý vào phòng kỹ thuật, tủ điện.',
                    '16. Khách: Tiếp khách 06:00–22:30. Không để khách qua đêm không báo.',
                    '17. Ở ghép: Không cho thuê lại khi chưa có chấp thuận.',
                    '18. Cam kết: Tuân thủ nghiêm túc, tái phạm sẽ chấm dứt hợp đồng.'
                  ].map((rule, idx) => (
                    <p key={idx} style={{ marginBottom: '3px' }}><b>{rule.split(':')[0]}:</b> {rule.split(':')[1]}</p>
                  ))}
                </div>
                <h3 style={{ fontSize: '14px', fontWeight: 'bold', marginTop: '15px' }}>CHẾ TÀI XỬ PHẠT</h3>
                <table className="contract-table" style={{ fontSize: '10px' }}>
                  <thead>
                    <tr style={{ background: '#f0f0f0' }}><th>HÀNH VI VI PHẠM</th><th>MỨC PHẠT (VNĐ)</th></tr>
                  </thead>
                  <tbody>
                    {[
                      ['Sạc xe/pin điện trong phòng/hành lang', '500.000/lần'], 
                      ['Để xe không khai báo/qua đêm (xe khách)', '300.000/xe/đêm'], 
                      ['Vứt rác bừa bãi, ném rác qua cửa sổ', '200.000/lần'], 
                      ['Gây tắc cống (đổ thức ăn, tóc...)', '200.000 + phí sửa'],
                      ['Gây ồn ào sau 22:00', '100.000 - 300.000/lần'], 
                      ['Cải tạo (khoan, đục...) không phép', '500.000/lần']
                    ].map((item, idx) => (
                      <tr key={idx}>
                        <td>{item[0]}</td>
                        <td style={{ textAlign: 'right' }}>{item[1]}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div style={{ marginTop: '20px', textAlign: 'right' }}>
                  <p><b>ĐẠI DIỆN NGƯỜI THUÊ CAM KẾT</b></p>
                  <div className="signature-space" style={{ height: '50px' }}></div>
                  <p><b>{primaryTenant.name}</b></p>
                </div>
              </div>
            </div>
          )}
          {isLiquidation && report && (
            <div className="liquidation-document">
              <div className="contract-header-text">
                <h1>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h1>
                <h2>Độc lập – Tự do – Hạnh phúc</h2>
              </div>
              <h1 style={{ textAlign: 'center', margin: '30px 0' }}>PHỤ LỤC 02 – BIÊN BẢN KIỂM TRA KHI TRẢ PHÒNG</h1>
              <p>Hôm nay, ngày {report.actualEndDate.split('-').reverse().join('/')}, chúng tôi tiến hành chốt dọn đi cho phòng <b>{room.id}</b>:</p>
              <table className="contract-table" style={{ margin: '20px 0' }}>
                <tbody>
                  <tr><td>Chỉ số điện mới: <b>{report.electricNew || report.electricEnd}</b> kWh</td><td>Chỉ số nước mới: <b>{report.waterNew || report.waterEnd}</b> m³</td></tr>
                  <tr><td>Phát sinh điện: {report.electricUsed} kWh = <b>{formatMoney(report.electricAmount)}</b></td><td>Phát sinh nước: {report.waterUsed} m³ = <b>{formatMoney(report.waterAmount)}</b></td></tr>
                </tbody>
              </table>
              <h3 style={{ fontSize: '16px', fontWeight: 'bold' }}>QUYẾT TOÁN CÔNG NỢ</h3>
              <div className="stack" style={{ gap: '5px' }}>
                <div className="data-row"><span>Tiền cọc đối trừ</span><b>{formatMoney(report.depositUsed)}</b></div>
                <div className="data-row"><span>Tiền nhà chưa trả</span><b>{formatMoney(report.unpaidRent)}</b></div>
                <div className="data-row"><span>Phí vệ sinh / Hư hỏng</span><b>{formatMoney(report.cleaningFee + report.damageFee)}</b></div>
                <div className="data-row" style={{ borderTop: '2px solid #000', paddingTop: '10px', marginTop: '10px', display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '16px' }}><b>KẾT QUẢ TẤT TOÁN</b></span>
                  {report.mustCollect > 0 ? <b style={{ fontSize: '20px', color: 'var(--danger)' }}>Khách trả thêm: {formatMoney(report.mustCollect)}</b> : <b style={{ fontSize: '20px', color: 'var(--success)' }}>Chủ nhà hoàn: {formatMoney(report.mustRefund)}</b>}
                </div>
              </div>
              <div className="signature-row">
                <div><p><b>BÊN CHO THUÊ</b></p><div className="signature-space"></div><p><b>DIỆM THỊ BÌNH</b></p></div>
                <div><p><b>BÊN THUÊ</b></p><div className="signature-space"></div><p><b>{primaryTenant.name}</b></p></div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function AppMain() {
  const [data, setData] = useState(() => safeRead(STORAGE_KEY, DEFAULT_DATA));
  const [bankInfo, setBankInfo] = useState(() => safeRead(BANK_KEY, DEFAULT_BANK));
  const [tab, setTab] = useState('dashboard');
  const [query, setQuery] = useState('');
  const [paymentFilters, setPaymentFilters] = useState(null);
  const [selectedRoom, setSelectedRoom] = useState(null);
  const [detailTenant, setDetailTenant] = useState(null);
  const [newRentalRoom, setNewRentalRoom] = useState(null);
  const [settlingRoom, setSettlingRoom] = useState(null);
  const [viewingContract, setViewingContract] = useState(null);
  const [renewingContract, setRenewingContract] = useState(null);
  const [viewingReceipt, setViewingReceipt] = useState(null);
  const [printingReceipts, setPrintingReceipts] = useState(null);
  const [paymentReceipt, setPaymentReceipt] = useState(null);
  const [editingTenant, setEditingTenant] = useState(null);
  const [editingContract, setEditingContract] = useState(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSynced, setLastSynced] = useState(null);
  const fileInputRef = React.useRef(null);

  // Initial Fetch from Cloud
  useEffect(() => {
    fetch('/api/data')
      .then(res => res.ok ? res.json() : Promise.reject())
      .then(cloudData => {
        if (cloudData && !cloudData.error && cloudData.rooms?.length > 0) {
          setData(cloudData);
          setLastSynced(new Date());
        }
      })
      .catch(err => console.log("Hệ thống đang chạy chế độ Local hoặc chưa cấu hình Database."));
  }, []);

  // Cloud Sync Logic (Debounced)
  useEffect(() => {
    const timer = setTimeout(() => {
      if (!data || data === DEFAULT_DATA) return;
      
      setIsSyncing(true);
      fetch('/api/data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'full_sync', payload: data })
      })
      .then(res => {
        if (res.ok) {
          setLastSynced(new Date());
        }
        setIsSyncing(false);
      })
      .catch(err => {
        setIsSyncing(false);
        console.error("Lỗi đồng bộ Cloud:", err);
      });
      
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    }, 2000);

    return () => clearTimeout(timer);
  }, [data]);
  useEffect(() => {
    let changed = false;
    const newData = { ...data };
    
    // 1. Ensure memberships sync with contract status
    const updatedMemberships = (data.memberships || []).map(m => {
      const contract = (data.contracts || []).find(c => c.id === m.contractId);
      if (!contract || contract.status === 'ended' || contract.status === 'cancelled') {
        if (m.status === 'active') {
          changed = true;
          return { ...m, status: 'ended' };
        }
      }
      return m;
    });

    if (changed) {
      setData(prev => ({ ...prev, memberships: updatedMemberships }));
    }
  }, [data.contracts, data.memberships]);

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }, [data]);
  useEffect(() => { localStorage.setItem(BANK_KEY, JSON.stringify(bankInfo)); }, [bankInfo]);

  const filteredTenants = useMemo(() => {
    const q = query.toLowerCase();
    return (data.tenants || []).filter(t => t.name.toLowerCase().includes(q) || (t.phone && t.phone.includes(q)) || (t.cccd && t.cccd.includes(q)));
  }, [data.tenants, query]);

  function handleAction(type, roomOrTenant) {
    if (type === 'add_tenant') {
      const room = data.rooms.find(r => r.id === roomOrTenant.id);
      setNewRentalRoom(room);
    } else if (type === 'notice') {
      if (window.confirm(`Xác nhận báo chuyển cho phòng ${roomOrTenant.id}?`)) updateContractStatus(roomOrTenant.id, 'notice');
    } else if (type === 'cancel_notice') {
      if (window.confirm(`Hủy báo chuyển phòng ${roomOrTenant.id}?`)) updateContractStatus(roomOrTenant.id, 'active');
    } else if (type === 'moving_out') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const room = data.rooms.find(r => r.id === roomId);
      const activeContract = data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice'));
      if (activeContract) setSettlingRoom({ room, contract: activeContract });
    } else if (type === 'view_contract') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const contractId = roomOrTenant.contractId;
      const targetContract = contractId 
        ? data.contracts.find(c => c.id === contractId)
        : data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice' || c.status === 'moving_out'));
      
      if (targetContract) setViewingContract({ contract: targetContract, room: data.rooms.find(r => r.id === targetContract.roomId), type: 'main' });
      else alert('Không tìm thấy hợp đồng phù hợp.');
    } else if (type === 'edit_contract') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const contractId = roomOrTenant.contractId;
      const targetContract = contractId 
        ? data.contracts.find(c => c.id === contractId)
        : data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice' || c.status === 'moving_out'));
      
      if (targetContract) setEditingContract(targetContract);
      else alert('Không tìm thấy hợp đồng phù hợp.');
    } else if (type === 'view_history') {
      setTab('rental_history');
      setQuery(`P${roomOrTenant.id}`);
    } else if (type === 'view_payments') {
      setTab('payment_history');
      setQuery(`P${roomOrTenant.id}`);
    } else if (type === 'view_qr') {
      setViewingReceipt(roomOrTenant);
    } else if (type === 'pay_receipt') {
      setPaymentReceipt(roomOrTenant);
    } else if (type === 'detail_tenant') {
      setDetailTenant(roomOrTenant);
    } else if (type === 'edit_tenant') {
      setEditingTenant(roomOrTenant);
    } else if (type === 'renew_contract') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const activeContract = data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice'));
      if (activeContract) setRenewingContract(activeContract);
    } else if (type === 'view_tenants') {
      setQuery(`P${roomOrTenant.id}`);
      setTab('tenants');
    } else if (type === 'create_receipt_all') {
      setTab('receipts');
      // Trigger batch create logic if needed, but for now just navigation is fine
    } else if (type === 'view_payments_all') {
      setTab('payment_history');
      setQuery('');
    } else if (type === 'add_new_rental') {
      setTab('rooms');
      setQuery('');
    } else if (type === 'export_json') {
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `room_manager_backup_${new Date().toISOString().slice(0,10)}.json`;
      a.click();
    } else if (type === 'import_json') {
      fileInputRef.current.click();
    } else if (type === 'export_excel') {
      try {
        const wb = XLSX.utils.book_new();
        const now = new Date();
        const dateStr = now.toISOString().slice(0, 10);
        const stats = getDashboardStats(data, INITIAL_MONTH);

        // 1. Sheet Tong quan
        const overviewData = [
          ['Chỉ tiêu', 'Giá trị'],
          ['Tổng số phòng', stats.totalRooms],
          ['Phòng đang ở', stats.occupiedRooms],
          ['Phòng trống', stats.vacantRooms],
          ['Phòng báo chuyển', stats.notifyingMoveOut.length],
          ['Tổng người đang ở', stats.currentTenants],
          ['Tổng phiếu thu', data.receipts.length],
          ['Tổng phải thu', stats.totalDebt + data.receipts.reduce((s, r) => s + (r.paidAmount || 0), 0)],
          ['Tổng đã thu', data.receipts.reduce((s, r) => s + (r.paidAmount || 0), 0)],
          ['Tổng còn nợ', stats.totalDebt],
          ['Ngày xuất dữ liệu', now.toLocaleDateString('vi-VN')]
        ];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(overviewData), "Tong quan");

        // 2. Sheet Phong
        const roomsData = data.rooms.map(r => {
          const statusInfo = getRoomStatusInfo(data, r.id);
          const contract = statusInfo.contract;
          const tenant = contract ? getPrimaryTenantByContract(data, contract.id) : null;
          const memberCount = (data.memberships || []).filter(m => m.roomId === r.id && m.status === 'active').length;
          return {
            'Mã phòng': r.id,
            'Trạng thái': statusInfo.label,
            'Người đứng tên hiện tại': tenant ? tenant.name : '',
            'Số người đang ở': memberCount,
            'Giá thuê mặc định': r.rent,
            'Tiền cọc mặc định': r.deposit,
            'Phí vệ sinh': r.cleaning || 0,
            'Phí thang máy': r.elevator || 0,
            'Phí giặt': r.laundry || 0,
            'Internet': r.internet || 0,
            'Đơn giá điện': r.electricPrice,
            'Đơn giá nước': r.waterPrice,
            'Ngày hết hạn HĐ hiện tại': contract ? (contract.endDate || '') : '',
            'Ghi chú': r.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(roomsData), "Phong");

        // 3. Sheet Nguoi thue hien tai
        const currentTenantsData = data.tenants.filter(t => (data.memberships || []).some(m => m.tenantId === t.id && m.status === 'active')).map(t => {
          const m = data.memberships.find(m => m.tenantId === t.id && m.status === 'active');
          return {
            'Mã người thuê': t.id,
            'Phòng': m ? m.roomId : '',
            'Họ tên': t.name,
            'Vai trò': m?.role === 'primary' ? 'Người đứng tên' : 'Người ở cùng',
            'SĐT': t.phone,
            'CCCD': t.cccd,
            'Biển số xe': t.licensePlate || '',
            'Địa chỉ': t.address || '',
            'Ngày vào': m ? (m.joinedDate || '') : '',
            'Trạng thái': 'Đang ở',
            'Ghi chú': t.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(currentTenantsData), "Nguoi thue hien tai");

        // 4. Sheet Hop dong hien tai
        const currentContractsData = data.contracts.filter(c => ['active', 'notice', 'moving_out'].includes(c.status)).map(c => {
          const tenant = getPrimaryTenantByContract(data, c.id);
          return {
            'Mã hợp đồng': c.id,
            'Phòng': c.roomId,
            'Người đứng tên': tenant ? tenant.name : '',
            'SĐT': tenant ? tenant.phone : '',
            'Ngày ký': c.signedDate || '',
            'Ngày bắt đầu': c.startDate || '',
            'Ngày hết hạn': c.endDate || '',
            'Giá thuê': c.rent,
            'Tiền cọc': c.deposit,
            'Trạng thái': c.status === 'active' ? 'Đang hiệu lực' : c.status === 'notice' ? 'Báo chuyển' : 'Đang tất toán',
            'Ngày báo chuyển': c.noticeDate || '',
            'Ngày dự kiến dọn': c.expectedMoveOutDate || '',
            'Ghi chú': c.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(currentContractsData), "Hop dong hien tai");

        // 5. Sheet Lich su hop dong
        const pastContractsData = data.contracts.filter(c => c.status === 'ended').map(c => {
          const tenant = getPrimaryTenantByContract(data, c.id);
          return {
            'Mã hợp đồng': c.id,
            'Phòng': c.roomId,
            'Người đứng tên': tenant ? tenant.name : '',
            'Ngày bắt đầu': c.startDate || '',
            'Ngày hết hạn dự kiến': c.endDate || '',
            'Ngày kết thúc thực tế': c.actualEndDate || '',
            'Giá thuê': c.rent,
            'Tiền cọc': c.deposit,
            'Trạng thái': 'Đã kết thúc',
            'Ghi chú': c.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pastContractsData), "Lich su hop dong");

        // 6. Sheet Phieu thang
        const monthlyReceiptsData = data.receipts.filter(r => r.type === 'monthly').map(r => {
          const tenant = getPrimaryTenantByContract(data, r.contractId);
          return {
            'Mã phiếu': r.id,
            'Tháng': r.month,
            'Phòng': r.roomId,
            'Mã hợp đồng': r.contractId,
            'Người đứng tên': tenant ? tenant.name : '',
            'Tiền phòng': r.rent,
            'Dịch vụ cố định': r.fixedServices,
            'Điện cũ': getElectricOld(r),
            'Điện mới': getElectricNew(r),
            'Điện dùng': r.electricUsed,
            'Đơn giá điện': r.electricPrice || (data.rooms.find(rm => rm.id === r.roomId)?.electricPrice),
            'Tiền điện': r.electricAmount,
            'Nước cũ': getWaterOld(r),
            'Nước mới': getWaterNew(r),
            'Nước dùng': r.waterUsed,
            'Đơn giá nước': r.waterPrice || (data.rooms.find(rm => rm.id === r.roomId)?.waterPrice),
            'Tiền nước': r.waterAmount,
            'Phụ phí': r.other || 0,
            'Tổng tiền': r.total,
            'Đã thanh toán': r.paidAmount || 0,
            'Còn nợ': r.total - (r.paidAmount || 0),
            'Trạng thái': r.status,
            'Ngày tạo': r.createdAt ? new Date(r.createdAt).toLocaleDateString('vi-VN') : '',
            'Ngày lưu': r.savedAt ? new Date(r.savedAt).toLocaleDateString('vi-VN') : '',
            'Ghi chú': r.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthlyReceiptsData), "Phieu thang");

        // 7. Sheet Thanh toan
        const paymentsData = data.receipts.map(r => {
          const tenant = getPrimaryTenantByContract(data, r.contractId);
          return {
            'Mã phiếu': r.id,
            'Loại phiếu': r.type === 'monthly' ? 'Phiếu tháng' : 'Chốt trả phòng',
            'Tháng': r.month,
            'Phòng': r.roomId,
            'Người đứng tên': tenant ? tenant.name : '',
            'Tổng tiền': r.total,
            'Đã trả': r.paidAmount || 0,
            'Còn nợ': r.total - (r.paidAmount || 0),
            'Trạng thái': r.status,
            'Ngày thanh toán': r.paidDate || '',
            'Nội dung chuyển khoản': transferContent(r),
            'Ghi chú': r.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(paymentsData), "Thanh toan");

        // 8. Sheet Lich su tra phong
        const moveOutData = (data.moveOutReports || []).map(rep => ({
          'Mã biên bản': rep.id,
          'Phòng': rep.roomId,
          'Mã hợp đồng': rep.contractId,
          'Người đứng tên': getPrimaryTenantByContract(data, rep.contractId)?.name || '',
          'Ngày trả phòng': rep.actualEndDate || '',
          'Điện đầu': rep.electricStart,
          'Điện chốt': rep.electricEnd,
          'Tiền điện': rep.electricCharge,
          'Nước đầu': rep.waterStart,
          'Nước chốt': rep.waterEnd,
          'Tiền nước': rep.waterCharge,
          'Tiền phòng còn nợ': rep.rentCharge,
          'Phí hư hỏng': rep.damages || 0,
          'Phí vệ sinh': rep.cleaningFee || 0,
          'Phí khác': rep.otherCharges || 0,
          'Tiền cọc đối trừ': rep.depositUsed,
          'Tổng phát sinh': rep.totalIncurred,
          'Khách còn phải trả': rep.mustCollect,
          'Cần hoàn cọc': rep.mustRefund,
          'Ghi chú': rep.note || '',
          'Ngày tạo': rep.createdAt ? new Date(rep.createdAt).toLocaleDateString('vi-VN') : ''
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(moveOutData), "Lich su tra phong");

        // 9. Sheet Lich su nguoi thue
        const pastTenantsData = data.tenants.filter(t => (data.memberships || []).some(m => m.tenantId === t.id && m.status === 'moved_out')).map(t => {
          const m = data.memberships.find(m => m.tenantId === t.id && m.status === 'moved_out');
          return {
            'Mã người thuê': t.id,
            'Họ tên': t.name,
            'SĐT': t.phone,
            'CCCD': t.cccd,
            'Biển số xe': t.licensePlate || '',
            'Phòng cũ': m ? m.roomId : '',
            'Vai trò': m?.role === 'primary' ? 'Người đứng tên' : 'Người ở cùng',
            'Ngày vào': m ? (m.joinedDate || '') : '',
            'Ngày rời đi': m ? (m.leftDate || '') : '',
            'Mã hợp đồng': m ? (m.contractId || '') : '',
            'Trạng thái': 'Đã rời đi',
            'Ghi chú': t.note || ''
          };
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pastTenantsData), "Lich su nguoi thue");

        // 10. Sheet Cau hinh
        const configData = [
          ['Thông tin', 'Giá trị'],
          ['Tên chủ nhà', 'DIỆM THỊ BÌNH'],
          ['SĐT chủ nhà', '0123.456.789'],
          ['Ngân hàng', bankInfo.bankName],
          ['Số tài khoản', bankInfo.accountNo],
          ['Tên chủ tài khoản', bankInfo.accountName],
          ['Mã ngân hàng VietQR', bankInfo.bankCode],
          ['Ngày thu tiền hàng tháng', 5],
          ['Đơn giá điện mặc định', 3800],
          ['Đơn giá nước mặc định', 32000]
        ];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(configData), "Cau hinh");

        XLSX.writeFile(wb, `quan-ly-phong-${dateStr}.xlsx`);
        alert('Đã xuất file Excel thành công!');
      } catch (err) {
        console.error(err);
        alert('Không thể xuất Excel. Vui lòng thử lại.');
      }
    } else if (type === 'create_receipt') {
      const roomId = roomOrTenant.id;
      const activeContract = data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice'));
      if (!activeContract) return alert('Phòng trống hoặc không có hợp đồng.');
      const month = INITIAL_MONTH;
      const exists = (data.receipts || []).find(r => r.roomId === roomId && r.contractId === activeContract.id && r.month === month);
      if (exists) {
        if (!window.confirm(`Phòng ${roomId} đã có phiếu tháng ${month}. Bạn có muốn ghi đè?`)) return;
        setData(old => ({ ...old, receipts: old.receipts.filter(r => r.id !== exists.id) }));
      }
      const prev = getPreviousReceipt(data.receipts, roomId, activeContract.id, month);
      const newRec = createMonthlyReceipt(roomOrTenant, activeContract, prev, month);
      setData(old => ({ ...old, receipts: [...(old.receipts || []), newRec] }));
      alert(`Đã tạo phiếu tháng ${month} cho phòng ${roomId}`);
      setTab('receipts');
    } else if (type === 'delete_receipt') {
      if (window.confirm('Bạn có chắc chắn muốn xóa phiếu thu này?')) {
        setData(old => ({ ...old, receipts: old.receipts.filter(r => r.id !== roomOrTenant.id) }));
      }
    }
  }

  function updateContractStatus(roomId, status) {
    setData(old => ({ ...old, contracts: old.contracts.map(c => (c.roomId === roomId && (c.status === 'active' || c.status === 'notice')) ? { ...c, status } : c) }));
  }

  return (
    <div className="app">
      <header className="app-header no-print">
        <div><h1 className="app-title">Room Manager</h1><p className="app-subtitle">Hệ thống quản lý phòng trọ chuyên nghiệp</p></div>
        <div className="search-container" style={{ maxWidth: '400px' }}><input className="search" type="text" placeholder="Tìm nhanh..." value={query} onChange={e => setQuery(e.target.value)} /></div>
      </header>
      <div className="layout">
        <aside className="sidebar no-print">
          <button className={`nav-item ${tab === 'dashboard' ? 'active' : ''}`} onClick={() => setTab('dashboard')}>📊 Thống kê</button>
          <button className={`nav-item ${tab === 'rooms' ? 'active' : ''}`} onClick={() => setTab('rooms')}>🏠 Phòng trọ</button>
          <button className={`nav-item ${tab === 'tenants' ? 'active' : ''}`} onClick={() => setTab('tenants')}>👥 Người thuê</button>
          <button className={`nav-item ${tab === 'receipts' ? 'active' : ''}`} onClick={() => setTab('receipts')}>🧾 Phiếu tháng</button>
          <button className={`nav-item ${tab === 'rental_history' ? 'active' : ''}`} onClick={() => setTab('rental_history')}>📜 Lịch sử thuê</button>
          <button className={`nav-item ${tab === 'payment_history' ? 'active' : ''}`} onClick={() => setTab('payment_history')}>💰 Thanh toán</button>
          <button className={`nav-item ${tab === 'settings' ? 'active' : ''}`} onClick={() => setTab('settings')}>⚙️ Cài đặt</button>
        </aside>
        <div className="content content-fade">
          {tab === 'dashboard' && <Dashboard data={data} onRoomClick={(id) => { setTab('rooms'); setQuery(id); }} onAction={handleAction} />}
          {tab === 'rooms' && <RoomsTab data={data} onAction={handleAction} onSelect={setSelectedRoom} query={query} />}
          {tab === 'tenants' && <TenantsTab tenants={filteredTenants} data={data} onAction={(type, t) => { if (type==='detail') setDetailTenant(t); else handleAction(type, t); }} query={query} setQuery={setQuery} setData={setData} />}
          {tab === 'receipts' && (
            <ReceiptsTab 
              data={data} bankInfo={bankInfo}
              onUpdateReceipt={(updated) => setData(old => {
                const newReceipts = old.receipts.map(r => r.id === updated.id ? updated : r);
                // Đồng bộ chỉ số vào room
                const newRooms = old.rooms.map(room => {
                  if (room.id === updated.roomId && updated.type === 'monthly') {
                    return { ...room, electricOld: updated.electricOld, electricNew: updated.electricNew, waterOld: updated.waterOld, waterNew: updated.waterNew };
                  }
                  return room;
                });
                return { ...old, receipts: newReceipts, rooms: newRooms };
              })}
              onBatchCreate={(newReceipts) => setData(old => {
                const existingIds = new Set(newReceipts.map(nr => `${nr.roomId}-${nr.contractId}-${nr.month}`));
                const filteredOld = (old.receipts || []).filter(r => !existingIds.has(`${r.roomId}-${r.contractId}-${r.month}`));
                const mergedReceipts = [...filteredOld, ...newReceipts];
                
                // Đồng bộ chỉ số cho những phòng vừa tạo phiếu
                const newRooms = old.rooms.map(room => {
                  const latestR = newReceipts.find(nr => nr.roomId === room.id && nr.type === 'monthly');
                  if (latestR) {
                    return { ...room, electricOld: latestR.electricOld, electricNew: latestR.electricNew, waterOld: latestR.waterOld, waterNew: latestR.waterNew };
                  }
                  return room;
                });
                
                return { ...old, receipts: mergedReceipts, rooms: newRooms };
              })}
              onView={(r) => setViewingReceipt(r)} onPrintBatch={(receipts) => setPrintingReceipts(receipts)} onPay={(r) => setPaymentReceipt(r)}
              onGoToPayment={(filters) => {
                setPaymentFilters(filters);
                setTab('payment_history');
              }}
            />
          )}
          {tab === 'rental_history' && <RentalHistoryTab data={data} onAction={handleAction} />}
          {tab === 'payment_history' && (
            <PaymentHistoryTab 
              data={data} 
              bankInfo={bankInfo} 
              onAction={handleAction} 
              initialFilter={paymentFilters}
              onUpdateReceipt={(updated, deleteId) => {
                if (deleteId) {
                  setData(old => ({ ...old, receipts: old.receipts.filter(r => r.id !== deleteId) }));
                } else {
                  setData(old => ({ ...old, receipts: old.receipts.map(r => r.id === updated.id ? updated : r) }));
                }
              }} 
              onView={(r) => setViewingReceipt(r)} 
              onPay={(r) => setPaymentReceipt(r)} 
            />
          )}
          {tab === 'settings' && <SettingsTab data={data} setData={setData} bankInfo={bankInfo} setBankInfo={setBankInfo} onReset={() => { if(window.confirm('Xóa hết dữ liệu?')) setData(DEFAULT_DATA); }} />}
        </div>
      </div>
      {selectedRoom && (
        <RoomDetailModal 
          room={selectedRoom} 
          data={data} 
          onClose={() => setSelectedRoom(null)} 
          onAction={(type, arg) => { 
            handleAction(type, arg || selectedRoom); 
            setSelectedRoom(null); 
          }} 
        />
      )}
      {newRentalRoom && (
        <RentalFlowModal 
          room={newRentalRoom} 
          onClose={() => setNewRentalRoom(null)} 
          onSave={(result) => {
            const { tenant, contract, memberships } = result;
            setData(old => ({
              ...old,
              tenants: [...(old.tenants || []), tenant],
              contracts: [...(old.contracts || []), contract],
              memberships: [...(old.memberships || []), ...memberships]
            }));
            setNewRentalRoom(null);
          }}
        />
      )}
      {settlingRoom && (
        <SettlementModal room={settlingRoom.room} contract={settlingRoom.contract} data={data} onClose={() => setSettlingRoom(null)} onSave={(report) => {
            setData(old => {
              const updatedContracts = old.contracts.map(c => c.id === report.contractId ? { ...c, status: 'ended', actualEndDate: report.actualEndDate, endedAt: new Date().toISOString() } : c);
              const updatedMemberships = old.memberships.map(m => m.contractId === report.contractId ? { ...m, status: 'ended' } : m);
              const affectedTenantIds = old.memberships.filter(m => m.contractId === report.contractId).map(m => m.tenantId);
              const updatedTenants = old.tenants.map(t => affectedTenantIds.includes(t.id) ? { ...t, status: 'moved_out', lastRoomId: report.roomId } : t);
              
              // Tạo phiếu thu chốt trả phòng nếu còn nợ
              let newReceipts = [...(old.receipts || [])];
              if (report.mustCollect > 0) {
                newReceipts.push({
                  id: uid('receipt'),
                  roomId: report.roomId,
                  contractId: report.contractId,
                  type: 'move_out_settlement',
                  month: report.actualEndDate.split('-').slice(0, 2).reverse().join('/'),
                  total: report.mustCollect,
                  paidAmount: 0,
                  status: 'Chưa thanh toán',
                  note: 'Phiếu chốt trả phòng',
                  createdAt: new Date().toISOString()
                });
              }
              
              return { ...old, contracts: updatedContracts, memberships: updatedMemberships, tenants: updatedTenants, receipts: newReceipts, moveOutReports: [report, ...(old.moveOutReports || [])] };
            });
            setSettlingRoom(null);
          }}
        />
      )}
      {viewingContract && <ContractPreview {...viewingContract} tenants={data.tenants.filter(t => data.memberships.some(m => m.contractId === viewingContract.contract.id && m.tenantId === t.id))} bankInfo={bankInfo} onClose={() => setViewingContract(null)} />}
      {renewingContract && (
        <RenewalModal contract={renewingContract} data={data} onClose={() => setRenewingContract(null)} onSave={(form) => {
            setData(old => {
              const renewalRecord = {
                id: uid('renewal'),
                contractId: renewingContract.id,
                roomId: renewingContract.roomId,
                signedDate: form.signedDate,
                oldEndDate: renewingContract.endDate,
                newStartDate: form.newStartDate,
                newEndDate: form.newEndDate,
                oldRent: renewingContract.rent,
                newRent: form.keepPricing ? renewingContract.rent : Number(form.newRent),
                oldDeposit: renewingContract.deposit,
                newDeposit: form.keepPricing ? renewingContract.deposit : Number(form.newDeposit),
                note: form.note,
                createdAt: new Date().toISOString()
              };

              const updatedContracts = old.contracts.map(c => c.id === renewingContract.id ? { 
                ...c, 
                endDate: form.newEndDate,
                rent: form.keepPricing ? c.rent : Number(form.newRent),
                deposit: form.keepPricing ? c.deposit : Number(form.newDeposit),
                previousEndDate: renewingContract.endDate,
                renewedAt: new Date().toISOString(),
                renewalHistory: [renewalRecord, ...(c.renewalHistory || [])]
              } : c);
              
              return { 
                ...old, 
                contracts: updatedContracts, 
                contractRenewals: [renewalRecord, ...(old.contractRenewals || [])] 
              };
            });
            setRenewingContract(null);
          }}
        />
      )}
      {detailTenant && <TenantDetailModal tenant={detailTenant} data={data} onClose={() => setDetailTenant(null)} />}
      {editingTenant && (
        <EditTenantModal 
          tenant={editingTenant} 
          onClose={() => setEditingTenant(null)} 
          onSave={(updated) => {
            setData(old => {
              const updatedTenants = old.tenants.map(t => t.id === updated.id ? updated : t);
              // Cập nhật role trong membership nếu thay đổi
              const updatedMemberships = old.memberships.map(m => (m.tenantId === updated.id && m.status === 'active') ? { ...m, role: updated.role } : m);
              return { ...old, tenants: updatedTenants, memberships: updatedMemberships };
            });
            setEditingTenant(null);
          }} 
        />
      )}
      {editingContract && (
        <EditContractModal 
          contract={editingContract} 
          data={data}
          onClose={() => setEditingContract(null)} 
          onSave={(updated) => {
            setData(old => ({
              ...old,
              contracts: old.contracts.map(c => c.id === updated.id ? { ...updated, updatedAt: new Date().toISOString() } : c)
            }));
            setEditingContract(null);
          }} 
        />
      )}
      {viewingReceipt && (() => {
        const contextReceipts = (data.receipts || [])
          .filter(r => r.month === viewingReceipt.month && r.type === viewingReceipt.type)
          .sort((a, b) => a.roomId.localeCompare(b.roomId));
        
        const currentIndex = contextReceipts.findIndex(r => r.id === viewingReceipt.id);
        const prevReceipt = currentIndex > 0 ? contextReceipts[currentIndex - 1] : null;
        const nextReceipt = currentIndex < contextReceipts.length - 1 ? contextReceipts[currentIndex + 1] : null;

        return (
          <ReceiptModal 
            receipt={viewingReceipt} 
            room={data.rooms.find(r => r.id === viewingReceipt.roomId)} 
            data={data} 
            bankInfo={bankInfo} 
            onClose={() => setViewingReceipt(null)} 
            onPrev={prevReceipt ? () => setViewingReceipt(prevReceipt) : null}
            onNext={nextReceipt ? () => setViewingReceipt(nextReceipt) : null}
          />
        );
      })()}
      {printingReceipts && (
        <div className="print-overlay" style={{ position: 'fixed', inset: 0, background: 'white', zIndex: 9999, overflow: 'auto' }}>
          <div className="no-print" style={{ position: 'sticky', top: 0, padding: '20px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', gap: '10px' }}>
            <button className="primary-btn" onClick={() => window.print()}>🖨️ Bắt đầu in</button><button className="secondary-btn" onClick={() => setPrintingReceipts(null)}>Hủy bỏ</button>
          </div>
          <div className="print-container">
            {printingReceipts.map(r => <ReceiptItem key={r.id} receipt={r} room={data.rooms.find(rm => rm.id === r.roomId)} contract={data.contracts.find(c => c.id === r.contractId)} bankInfo={bankInfo} data={data} />)}
          </div>
        </div>
      )}
      {paymentReceipt && <PaymentModal receipt={paymentReceipt} onClose={() => setPaymentReceipt(null)} onSave={(updated) => { setData(old => ({ ...old, receipts: old.receipts.map(r => r.id === updated.id ? updated : r) })); setPaymentReceipt(null); }} />}
      <input 
        type="file" 
        ref={fileInputRef} 
        style={{ display: 'none' }} 
        accept=".json" 
        onChange={(e) => {
          const file = e.target.files[0];
          if (file) {
            const reader = new FileReader();
            reader.onload = (event) => {
              try {
                const imported = JSON.parse(event.target.result);
                if (window.confirm('Cảnh báo: Dữ liệu hiện tại sẽ bị ghi đè. Bạn có chắc chắn muốn nhập dữ liệu từ file này?')) {
                  setData(imported);
                  alert('Nhập dữ liệu thành công!');
                }
              } catch (err) {
                alert('Lỗi: File JSON không hợp lệ!');
              }
            };
            reader.readAsText(file);
          }
          e.target.value = ''; // Reset input
        }}
      />
    </div>
  );
}

function RentalHistoryTab({ data, onAction }) {
  const endedContracts = useMemo(() => (data.contracts || []).filter(c => c.status === 'ended').sort((a,b) => new Date(b.endedAt) - new Date(a.endedAt)), [data.contracts]);
  return (
    <div className="widget liquid-glass" style={{ padding: 0 }}>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Ngày trả</th><th>Phòng</th><th>Người đứng tên</th><th>Ngày bắt đầu</th><th>Thời hạn gốc</th><th>Tiền cọc</th><th>Thao tác</th></tr></thead>
          <tbody>
            {endedContracts.map(c => {
              const tenant = getPrimaryTenantByContract(data, c.id) || { name: 'N/A' };
              const report = (data.moveOutReports || []).find(r => r.contractId === c.id);
              return (
                <tr key={c.id}>
                  <td><b>{c.actualEndDate?.split('-').reverse().join('/') || 'N/A'}</b></td>
                  <td>P{c.roomId}</td>
                  <td>{tenant.name}</td>
                  <td>{c.startDate?.split('-').reverse().join('/')}</td>
                  <td>{c.endDate?.split('-').reverse().join('/')}</td>
                  <td>{formatMoney(c.deposit)}</td>
                  <td><button className="secondary-btn sm" onClick={() => onAction('view_contract', { roomId: c.roomId, contractId: c.id })}>Xem HĐ</button></td>
                </tr>
              );
            })}
            {endedContracts.length === 0 && <tr><td colSpan="7" style={{ textAlign: 'center', padding: '40px' }}>Chưa có lịch sử hợp đồng nào kết thúc.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PaymentHistoryTab({ data, bankInfo, onAction, onUpdateReceipt, onView, onPay, initialFilter }) {
  const [filter, setFilter] = useState({ roomId: '', status: 'all', type: 'all', month: '' });

  useEffect(() => {
    if (initialFilter) {
      setFilter(prev => ({ ...prev, ...initialFilter }));
    }
  }, [initialFilter]);

  const filteredReceipts = useMemo(() => {
    return (data.receipts || []).filter(r => {
      if (filter.roomId && r.roomId !== filter.roomId) return false;
      if (filter.status !== 'all' && r.status !== filter.status) return false;
      if (filter.type !== 'all' && r.type !== filter.type) return false;
      if (filter.month && r.month !== filter.month) return false;
      return true;
    }).sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));
  }, [data.receipts, filter]);

  const stats = useMemo(() => {
    const total = filteredReceipts.reduce((sum, r) => sum + r.total, 0);
    const paid = filteredReceipts.reduce((sum, r) => sum + (r.paidAmount || 0), 0);
    const debt = total - paid;
    return { total, paid, debt, count: filteredReceipts.length };
  }, [filteredReceipts]);

  const isFilteredByMonth = filter.month && filter.type === 'monthly';

  return (
    <div className="stack" style={{ gap: '16px' }}>
      {isFilteredByMonth && (
        <div className="widget liquid-glass" style={{ background: 'var(--primary-gradient)', color: 'white' }}>
          <h2 style={{ fontSize: '20px', margin: 0 }}>Thanh toán tháng {filter.month}</h2>
          <p style={{ opacity: 0.9, marginTop: '4px' }}>
            {stats.count} phiếu • Tổng <b>{formatMoney(stats.total)}</b> • Còn nợ <b style={{ color: '#fca5a5' }}>{formatMoney(stats.debt)}</b>
          </p>
        </div>
      )}

      <div className="widget liquid-glass no-print" style={{ padding: '16px' }}>
        <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap' }}>
          <label style={{ margin: 0 }}>Tháng <input type="month" value={filter.month ? `${filter.month.split('/')[1]}-${filter.month.split('/')[0]}` : ''} onChange={e => {
            if (!e.target.value) return setFilter({...filter, month: ''});
            const [y, m] = e.target.value.split('-'); 
            setFilter({...filter, month: `${m}/${y}`}); 
          }} /></label>
          <label style={{ margin: 0 }}>Phòng <select value={filter.roomId} onChange={e => setFilter({...filter, roomId: e.target.value})}><option value="">Tất cả</option>{data.rooms.map(r => <option key={r.id} value={r.id}>P{r.id}</option>)}</select></label>
          <label style={{ margin: 0 }}>Trạng thái <select value={filter.status} onChange={e => setFilter({...filter, status: e.target.value})}><option value="all">Tất cả</option><option value="Đã thanh toán">Đã thanh toán</option><option value="Chưa thanh toán">Chưa thanh toán</option><option value="Nợ một phần">Nợ một phần</option></select></label>
          <label style={{ margin: 0 }}>Loại phiếu <select value={filter.type} onChange={e => setFilter({...filter, type: e.target.value})}><option value="all">Tất cả</option><option value="monthly">Phiếu tháng</option><option value="move_out_settlement">Chốt trả phòng</option></select></label>
          <button className="secondary-btn" onClick={() => setFilter({ roomId: '', status: 'all', type: 'all', month: '' })}>🔄 Reset</button>
        </div>
      </div>
      <div className="widget liquid-glass" style={{ padding: 0 }}>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Ngày tạo</th><th>Phòng</th><th>Người đứng tên</th><th>Loại</th><th>Tổng tiền</th><th>Đã trả</th><th>Trạng thái</th><th>Hành động</th></tr></thead>
            <tbody>
              {filteredReceipts.map(r => {
                const tenant = getPrimaryTenantByContract(data, r.contractId) || { name: 'N/A' };
                return (
                  <tr key={r.id}>
                    <td>{new Date(r.createdAt).toLocaleDateString('vi-VN')}</td>
                    <td><b>P{r.roomId}</b></td>
                    <td>{tenant.name}</td>
                    <td><span style={{ fontSize: '12px' }}>{r.type === 'monthly' ? '📅 Phiếu tháng' : '🚪 Chốt trả phòng'}</span></td>
                    <td style={{ fontWeight: '700' }}>{formatMoney(r.total)}</td>
                    <td>{formatMoney(r.paidAmount || 0)}</td>
                    <td><span className={`status-badge-liquid ${r.status === 'Đã thanh toán' ? 'active' : r.status === 'Nợ một phần' ? 'notice' : 'debt'}`}>{r.status}</span></td>
                    <td>
                      <div style={{ display: 'flex', gap: '4px' }}>
                        <button className="primary-btn sm" onClick={() => onPay(r)}>💸 Thu</button>
                        <button className="secondary-btn sm" onClick={() => onView(r)}>📱 QR</button>
                        <button className="secondary-btn sm" style={{ color: '#ef4444' }} onClick={() => { if(window.confirm('Xóa phiếu này?')) onUpdateReceipt(null, r.id); }}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filteredReceipts.length === 0 && <tr><td colSpan="8" style={{ textAlign: 'center', padding: '40px' }}>Không có dữ liệu thanh toán phù hợp.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function Dashboard({ data, onRoomClick, onAction }) {
  const stats = getDashboardStats(data, INITIAL_MONTH);
  
  const recentRooms = [...data.rooms].slice(-5).reverse();
  const recentReceipts = [...(data.receipts || [])].slice(-5).reverse();
  const recentMoveOuts = (data.moveOutReports || []).slice(0, 5);

  const occupancyRate = stats.totalRooms > 0 ? (stats.occupiedRooms / stats.totalRooms) * 100 : 0;
  const paymentRate = stats.totalReceipts > 0 ? ((stats.totalReceipts - stats.unpaidReceipts.length) / stats.totalReceipts) * 100 : 0;

  return (
    <div className="dashboard-container stack">
      <div className="dashboard-header no-print">
        <div className="stack">
          <h2 style={{ fontSize: '24px' }}>Tổng quan hệ thống</h2>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', color: 'var(--text-muted)' }}>
            <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: isSyncing ? 'var(--warning)' : 'var(--success)' }}></span>
            {isSyncing ? 'Đang đồng bộ...' : lastSynced ? `Đã lưu Cloud: ${lastSynced.toLocaleTimeString()}` : 'Chế độ Local'}
          </div>
        </div>
        <div style={{ display: 'flex', gap: '12px' }}>
          <button className="primary-btn sm" onClick={() => onAction('create_receipt_all')}>⚡ Tạo phiếu tháng</button>
          <button className="secondary-btn sm" onClick={() => onAction('import_json')}>📤 Nhập JSON</button>
          <button className="secondary-btn sm" onClick={() => onAction('export_json')}>📥 Xuất JSON</button>
        </div>
      </div>

      <div className="stats-grid">
        <div className="stat-card-liquid" onClick={() => onRoomClick('')} style={{ cursor: 'pointer' }}>
          <p className="stat-label">Tổng số phòng</p>
          <p className="stat-value">{stats.totalRooms}</p>
          <div className="stat-progress-bg"><div className="stat-progress-bar" style={{ width: '100%' }}></div></div>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--success)' }}>
          <p className="stat-label">Phòng đang ở</p>
          <p className="stat-value">{stats.occupiedRooms}</p>
          <div className="stat-progress-bg"><div className="stat-progress-bar" style={{ width: `${occupancyRate}%`, background: 'var(--success)' }}></div></div>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--text-muted)' }}>
          <p className="stat-label">Phòng trống</p>
          <p className="stat-value">{stats.vacantRooms}</p>
          <div className="stat-progress-bg"><div className="stat-progress-bar" style={{ width: `${100 - occupancyRate}%`, background: 'var(--text-muted)' }}></div></div>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--warning)' }}>
          <p className="stat-label">Sắp hết hạn</p>
          <p className="stat-value">{stats.expiringContracts.length}</p>
          <p className="stat-note">Trong 30 ngày tới</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid #8b5cf6' }}>
          <p className="stat-label">Tổng người thuê</p>
          <p className="stat-value">{stats.currentTenants}</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid #06b6d4' }}>
          <p className="stat-label">Phiếu tháng {INITIAL_MONTH}</p>
          <p className="stat-value">{stats.totalReceipts}</p>
          <div className="stat-progress-bg"><div className="stat-progress-bar" style={{ width: `${paymentRate}%`, background: '#06b6d4' }}></div></div>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--danger)' }}>
          <p className="stat-label">Chưa thanh toán</p>
          <p className="stat-value">{stats.unpaidReceipts.length}</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid #f43f5e' }}>
          <p className="stat-label">Tổng nợ còn lại</p>
          <p className="stat-value" style={{ fontSize: '20px' }}>{formatMoney(stats.totalDebt)}</p>
        </div>
      </div>

      <div className="dashboard-grid-main">
        <div className="stack" style={{ gap: '24px' }}>
          {/* Cảnh báo */}
          <div className="widget liquid-glass">
            <h3 className="form-section-title">⚠️ Cảnh báo & Nhắc nhở</h3>
            <div className="alert-list stack" style={{ gap: '12px' }}>
              {stats.expiringContracts.length > 0 && (
                <div className="alert-item warning">
                  <span>🏠 {stats.expiringContracts.length} phòng sắp hết hạn HĐ: </span>
                  <b>{stats.expiringContracts.map(c => c.roomId).join(', ')}</b>
                </div>
              )}
              {stats.unpaidReceipts.length > 0 && (
                <div className="alert-item danger">
                  <span>💸 {stats.unpaidReceipts.length} phòng chưa đóng tiền tháng {INITIAL_MONTH}: </span>
                  <b>{stats.unpaidReceipts.map(r => r.roomId).join(', ')}</b>
                </div>
              )}
              {stats.notifyingMoveOut.length > 0 && (
                <div className="alert-item notice">
                  <span>🚪 {stats.notifyingMoveOut.length} phòng đang báo chuyển: </span>
                  <b>{stats.notifyingMoveOut.map(c => c.roomId).join(', ')}</b>
                </div>
              )}
              {stats.vacantRooms > 0 && (
                <div className="alert-item secondary">
                  <span>✨ Đang có {stats.vacantRooms} phòng trống sẵn sàng cho thuê.</span>
                </div>
              )}
              {stats.expiringContracts.length === 0 && stats.unpaidReceipts.length === 0 && stats.notifyingMoveOut.length === 0 && (
                <p className="muted center">Hiện tại không có cảnh báo nào.</p>
              )}
            </div>
          </div>

          {/* Vận hành nhanh */}
          <div className="widget liquid-glass">
            <h3 className="form-section-title">⚡ Vận hành nhanh</h3>
            <div className="quick-actions-grid">
              <button className="action-btn" onClick={() => onAction('create_receipt_all')}>
                <span className="icon">🧾</span>
                <span>Tạo phiếu tháng</span>
              </button>
              <button className="action-btn" onClick={() => onAction('add_new_rental')}>
                <span className="icon">🔑</span>
                <span>Thuê mới</span>
              </button>
              <button className="action-btn" onClick={() => onAction('view_payments_all')}>
                <span className="icon">💰</span>
                <span>Lịch sử thanh toán</span>
              </button>
              <button className="action-btn" style={{ background: 'var(--primary-gradient)', color: 'white' }} onClick={() => onAction('export_excel')}>
                <span className="icon">📊</span>
                <span>Xuất Excel</span>
              </button>
              <button className="action-btn" onClick={() => onAction('export_json')}>
                <span className="icon">🧩</span>
                <span>Xuất JSON backup</span>
              </button>
            </div>
          </div>
        </div>

        <div className="stack" style={{ gap: '24px' }}>
          {/* Danh sách nhanh */}
          <div className="widget liquid-glass">
            <h3 className="form-section-title">📅 Gần đây</h3>
            <div className="quick-tabs">
              <div className="quick-list-section">
                <p className="op-label uppercase" style={{ marginBottom: '10px' }}>5 phiếu thu mới nhất</p>
                <div className="mini-list stack" style={{ gap: '8px' }}>
                  {recentReceipts.map(r => (
                    <div key={r.id} className="mini-list-item">
                      <span>P{r.roomId} - {r.month}</span>
                      <b>{formatMoney(r.total)}</b>
                    </div>
                  ))}
                </div>
              </div>
              <div className="quick-list-section" style={{ marginTop: '20px' }}>
                <p className="op-label uppercase" style={{ marginBottom: '10px' }}>5 hợp đồng sắp hết hạn</p>
                <div className="mini-list stack" style={{ gap: '8px' }}>
                  {stats.expiringContracts.slice(0, 5).map(c => (
                    <div key={c.id} className="mini-list-item">
                      <span>P{c.roomId}</span>
                      <b className="danger">{c.endDate?.split('-').reverse().join('/')}</b>
                    </div>
                  ))}
                  {stats.expiringContracts.length === 0 && <p className="small muted">Không có hợp đồng nào sắp hết hạn.</p>}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function RoomsTab({ data, onAction, onSelect, query }) {
  const filteredRooms = useMemo(() => {
    const q = query.toLowerCase();
    return (data.rooms || []).filter(r => r.id.toLowerCase().includes(q));
  }, [data.rooms, query]);
  return (
    <div className="rooms-grid">
      {filteredRooms.map(room => {
        const { label, color, contract } = getRoomStatusInfo(data, room.id);
        const primaryTenant = contract ? getPrimaryTenantByContract(data, contract.id) : null;
        const currentReceipt = contract ? (data.receipts || []).find(r => r.roomId === room.id && r.contractId === contract.id && r.month === INITIAL_MONTH && r.type === 'monthly') : null;
        return (
          <div key={room.id} className="room-card-liquid" onClick={() => onSelect(room)}>
            <div className="room-header"><span className="room-id">P{room.id}</span><span className={`status-badge-liquid ${color === 'green' ? 'active' : color === 'gray' ? 'vacant' : color}`}>{label}</span></div>
            <div className="room-body">
              {primaryTenant ? (
                <div className="tenant-block-primary">
                  <p className="tenant-name-main">{primaryTenant.name}</p>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '8px' }}><p className="tenant-sub">{formatMoney(contract.rent)}</p>{currentReceipt && <span className={`status-badge-liquid ${currentReceipt.status === 'Đã thanh toán' ? 'active' : currentReceipt.status === 'Nợ một phần' ? 'notice' : 'debt'}`} style={{ fontSize: '10px' }}>{currentReceipt.status}</span>}</div>
                </div>
              ) : <p className="muted">Phòng đang trống</p>}
            </div>
            <div className="btn-group">
              {label === 'Trống' ? (
                <>
                  <button className="primary-btn wide" onClick={(e) => { e.stopPropagation(); onAction('add_tenant', room); }}>+ Thêm khách</button>
                  <button className="secondary-btn" title="Xem lịch sử" onClick={(e) => { e.stopPropagation(); onAction('view_history', room); }}>📜</button>
                </>
              ) : (
                <><button className="secondary-btn" title="Xem hợp đồng" onClick={(e) => { e.stopPropagation(); onAction('view_contract', room); }}>📄</button><button className="primary-btn wide" onClick={(e) => { e.stopPropagation(); onAction('create_receipt', room); }}>{currentReceipt ? 'Sửa phiếu' : 'Lập phiếu'}</button></>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function TenantsTab({ tenants, data, onAction, query, setQuery, setData }) {
  // Chỉ hiển thị những người đang có membership active
  const activeTenants = useMemo(() => {
    return tenants.filter(t => (data.memberships || []).some(m => m.tenantId === t.id && m.status === 'active'));
  }, [tenants, data.memberships]);

  return (
    <div className="widget stack" style={{ gap: '16px' }}>
      <div className="no-print" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h2 style={{ fontSize: '18px' }}>Khách đang ở ({activeTenants.length})</h2>
        <button className="secondary-btn" onClick={() => { setQuery(''); const d = localStorage.getItem(STORAGE_KEY); if(d) setData(JSON.parse(d)); }} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          🔄 Làm mới danh sách
        </button>
      </div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Người thuê</th><th>SĐT</th><th>CCCD</th><th>Trạng thái</th><th>Phòng</th><th>Vai trò</th><th>Thao tác</th></tr></thead>
          <tbody>
            {activeTenants.map(t => {
              const m = (data.memberships || []).find(m => m.tenantId === t.id && m.status === 'active');
              if (!m) return null; // Safety check
              return (
                <tr key={t.id}>
                  <td><button className="primary-btn sm" style={{ height: '32px', fontSize: '13px' }} onClick={() => onAction('detail', t)}>{t.name}</button></td>
                  <td>{t.phone}</td>
                  <td>{t.cccd}</td>
                  <td><span className="status-badge-liquid active">Đang ở</span></td>
                  <td>P{m.roomId}</td>
                  <td>{m?.role === 'primary' ? 'Đại diện' : 'Ở cùng'}</td>
                  <td>
                    <div style={{ display: 'flex', gap: '4px' }}>
                      <button className="secondary-btn sm" style={{ height: '32px', fontSize: '13px' }} onClick={() => onAction('view_contract', t)}>Xem HĐ</button>
                      <button className="secondary-btn sm" style={{ height: '32px', fontSize: '13px' }} onClick={() => onAction('edit_tenant', t)}>✏️ Sửa</button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {activeTenants.length === 0 && <tr><td colSpan="7" style={{ textAlign: 'center', padding: '40px' }}>Không tìm thấy khách đang ở nào.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ReceiptsTab({ data, bankInfo, onUpdateReceipt, onBatchCreate, onView, onPrintBatch, onPay, onGoToPayment }) {
  const [selectedMonth, setSelectedMonth] = useState(INITIAL_MONTH);
  const [activeTab, setActiveTab] = useState('entry');
  const [saveModal, setSaveModal] = useState(null);

  const monthlyReceipts = useMemo(() => (data.receipts || []).filter(r => r.month === selectedMonth && r.type === 'monthly'), [data.receipts, selectedMonth]);
  const isFinalized = monthlyReceipts.length > 0 && monthlyReceipts.every(r => r.isFinalized);

  function handleBatchCreate() {
    const activeContracts = (data.contracts || []).filter(c => c.status === 'active' || c.status === 'notice');
    if (activeContracts.length === 0) return alert('Không có hợp đồng nào đang hoạt động.');
    
    let createdCount = 0;
    let skippedCount = 0;
    const newReceipts = [];

    activeContracts.forEach(contract => {
      const room = data.rooms.find(r => r.id === contract.roomId);
      const exists = (data.receipts || []).find(r => r.roomId === room.id && r.contractId === contract.id && r.month === selectedMonth && r.type === 'monthly');
      
      if (!exists) {
        const prev = getPreviousReceipt(data.receipts, room.id, contract.id, selectedMonth);
        newReceipts.push(createMonthlyReceipt(room, contract, prev, selectedMonth));
        createdCount++;
      } else {
        skippedCount++;
      }
    });

    if (newReceipts.length > 0) {
      if (window.confirm(`Xác nhận tạo ${newReceipts.length} phiếu thu cho tháng ${selectedMonth}?`)) {
        onBatchCreate(newReceipts);
        alert(`Đã tạo ${createdCount} phiếu, bỏ qua ${skippedCount} phiếu đã tồn tại.`);
      }
    } else {
      alert(`Tất cả phòng đều đã có phiếu tháng ${selectedMonth}. Bỏ qua ${skippedCount} phiếu.`);
    }
  }

  function handleFinalizeMonth() {
    if (monthlyReceipts.length === 0) return alert('Không có phiếu nào để lưu.');
    
    const finalizedReceipts = monthlyReceipts.map(r => {
      const paid = r.paidAmount || 0;
      let status = 'Chưa thanh toán';
      if (paid >= r.total) status = 'Đã thanh toán';
      else if (paid > 0) status = 'Nợ một phần';

      return {
        ...r,
        isFinalized: true,
        savedAt: new Date().toISOString(),
        type: 'monthly',
        paidAmount: paid,
        status: status
      };
    });

    // Cập nhật lên store
    finalizedReceipts.forEach(r => onUpdateReceipt(r));

    // Tính toán thống kê cho modal
    const totalAmount = finalizedReceipts.reduce((sum, r) => sum + r.total, 0);
    const totalPaid = finalizedReceipts.reduce((sum, r) => sum + r.paidAmount, 0);
    const unpaidCount = finalizedReceipts.filter(r => r.status !== 'Đã thanh toán').length;

    setSaveModal({
      month: selectedMonth,
      count: finalizedReceipts.length,
      totalAmount,
      unpaidCount,
      totalDebt: totalAmount - totalPaid
    });
  }

  function handleUpdateWithWarning(updated) {
    const original = monthlyReceipts.find(r => r.id === updated.id);
    if (original?.isFinalized) {
      const isPaid = original.status === 'Đã thanh toán';
      const msg = isPaid 
        ? "Phiếu này ĐÃ THANH TOÁN. Thay đổi chỉ số có thể làm lệch lịch sử thu tiền. Bạn vẫn muốn cập nhật?"
        : "Phiếu tháng này đã được lưu chính thức. Nếu cập nhật chỉ số, tổng tiền và mã QR sẽ thay đổi. Bạn vẫn muốn cập nhật?";
      
      if (!window.confirm(msg)) return;
    }
    onUpdateReceipt(updated);
  }
  return (
    <div className="receipts-tab stack">
      <div className="widget liquid-glass no-print" style={{ padding: '16px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <label style={{ flexDirection: 'row', alignItems: 'center', gap: '8px', margin: 0 }}>Tháng: <input type="month" value={`${selectedMonth.split('/')[1]}-${selectedMonth.split('/')[0]}`} onChange={e => { const [y, m] = e.target.value.split('-'); setSelectedMonth(`${m}/${y}`); }} style={{ height: '40px' }} /></label>
            <button className="primary-btn" onClick={handleBatchCreate}>⚡ Tạo hàng loạt</button>
            <button className="secondary-btn" onClick={() => onPrintBatch(monthlyReceipts)}>🖨️ In tất cả ({monthlyReceipts.length})</button>
          </div>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <button className="primary-btn" style={{ backgroundColor: '#10b981', boxShadow: '0 4px 12px rgba(16,185,129,0.2)' }} onClick={handleFinalizeMonth}>💾 Lưu dữ liệu</button>
            <div className="btn-group" style={{ margin: 0 }}>
              <button className={`secondary-btn ${activeTab === 'entry' ? 'active-tab' : ''}`} onClick={() => setActiveTab('entry')}>Nhập chỉ số</button>
              <button className={`secondary-btn ${activeTab === 'history' ? 'active-tab' : ''}`} onClick={() => setActiveTab('history')}>Lịch sử & QR</button>
            </div>
          </div>
        </div>
      </div>

      {isFinalized && (
        <div className="widget liquid-glass no-print" style={{ background: 'rgba(16,185,129,0.1)', border: '1px solid #10b981', padding: '12px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <span style={{ fontSize: '20px' }}>✅</span>
            <div>
              <p style={{ fontWeight: 'bold', margin: 0, color: '#065f46' }}>Phiếu tháng {selectedMonth} đã được lưu.</p>
              <p className="small muted" style={{ margin: 0 }}>Bạn có thể sang Thanh toán để theo dõi thu tiền, in QR hoặc đánh dấu đã thu.</p>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '10px' }}>
             <button className="secondary-btn sm" onClick={() => onGoToPayment({ month: selectedMonth, type: 'monthly', status: 'all' })}>💰 Đi đến Thanh toán</button>
             <button className="secondary-btn sm" onClick={() => onPrintBatch(monthlyReceipts)}>🖨️ In tất cả</button>
          </div>
        </div>
      )}

      {saveModal && (
        <div className="modal" onClick={() => setSaveModal(null)}>
          <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '450px', textAlign: 'center' }} onClick={e => e.stopPropagation()}>
            <div style={{ fontSize: '48px', marginBottom: '16px' }}>✅</div>
            <h2 style={{ fontSize: '24px', marginBottom: '8px' }}>Đã lưu phiếu tháng</h2>
            <div style={{ background: 'rgba(255,255,255,0.05)', padding: '20px', borderRadius: '16px', margin: '20px 0', textAlign: 'left' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}><span className="muted">Tháng</span><b>{saveModal.month}</b></div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}><span className="muted">Số phiếu</span><b>{saveModal.count}</b></div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}><span className="muted">Tổng phải thu</span><b>{formatMoney(saveModal.totalAmount)}</b></div>
              <div style={{ height: '1px', background: 'rgba(255,255,255,0.1)', margin: '10px 0' }}></div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}><span className="muted">Phiếu chưa trả</span><b className="danger">{saveModal.unpaidCount}</b></div>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="muted">Tổng còn nợ</span><b className="danger">{formatMoney(saveModal.totalDebt)}</b></div>
            </div>
            
            <div className="stack" style={{ gap: '10px' }}>
              <button className="primary-btn wide" onClick={() => { setSaveModal(null); onGoToPayment({ month: selectedMonth, type: 'monthly', status: 'all' }); }}>💰 Đi đến Thanh toán</button>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button className="secondary-btn" style={{ flex: 1 }} onClick={() => onPrintBatch(monthlyReceipts)}>🖨️ In tất cả phiếu</button>
                <button className="secondary-btn" style={{ flex: 1 }} onClick={() => setSaveModal(null)}>📝 Ở lại nhập chỉ số</button>
              </div>
            </div>
          </div>
        </div>
      )}
      {activeTab === 'entry' ? (
        <div className="widget liquid-glass" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="table-wrap">
            <table className="entry-table">
              <thead>
                <tr>
                  <th>Phòng</th>
                  <th>Chủ phòng</th>
                  <th style={{ backgroundColor: '#fff7ed' }}>Điện: Chỉ số cũ</th>
                  <th style={{ backgroundColor: '#fff7ed' }}>Chỉ số mới</th>
                  <th style={{ backgroundColor: '#fff7ed' }}>Số sử dụng</th>
                  <th>Tiền điện</th>
                  <th style={{ backgroundColor: '#f0fdf4' }}>Nước: Chỉ số cũ</th>
                  <th style={{ backgroundColor: '#f0fdf4' }}>Chỉ số mới</th>
                  <th style={{ backgroundColor: '#f0fdf4' }}>Số sử dụng</th>
                  <th>Tiền nước</th>
                  <th>Phụ phí</th>
                  <th>Tổng tiền</th>
                  <th>Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {monthlyReceipts.sort((a,b) => a.roomId.localeCompare(b.roomId)).map(r => {
                  const room = data.rooms.find(rm => rm.id === r.roomId);
                  const m = data.memberships.find(ms => ms.contractId === r.contractId && ms.role === 'primary');
                  const tenant = m ? data.tenants.find(t => t.id === m.tenantId) : { name: 'N/A' };
                  
                  const eOld = r.electricOld ?? r.electricStart ?? 0;
                  const eNew = r.electricNew ?? r.electricEnd ?? eOld;
                  const wOld = r.waterOld ?? r.waterStart ?? 0;
                  const wNew = r.waterNew ?? r.waterEnd ?? wOld;

                  return (
                    <tr key={r.id}>
                      <td style={{ fontWeight: '800' }}>P{r.roomId}</td>
                      <td>{tenant.name}</td>
                      <td>
                        <DecimalInput 
                          value={eOld} 
                          onChange={val => handleUpdateWithWarning(recalculateReceipt({ ...r, electricOld: val }, room))} 
                          style={{ width: '80px', height: '32px', padding: '0 8px', border: '1px solid #e2e8f0', color: 'var(--text-muted)' }} 
                        />
                      </td>
                      <td>
                        <DecimalInput 
                          value={eNew} 
                          onChange={val => handleUpdateWithWarning(recalculateReceipt({ ...r, electricNew: val }, room))} 
                          style={{ width: '80px', height: '32px', padding: '0 8px', border: '1px solid var(--warning)' }} 
                        />
                      </td>
                      <td style={{ textAlign: 'center', fontWeight: 'bold', color: 'var(--warning)' }}>{r.electricUsed || 0}</td>
                      <td className="small">{formatMoney(r.electricAmount || 0)}</td>
                      
                      <td>
                        <DecimalInput 
                          value={wOld} 
                          onChange={val => handleUpdateWithWarning(recalculateReceipt({ ...r, waterOld: val }, room))} 
                          style={{ width: '80px', height: '32px', padding: '0 8px', border: '1px solid #e2e8f0', color: 'var(--text-muted)' }} 
                        />
                      </td>
                      <td>
                        <DecimalInput 
                          value={wNew} 
                          onChange={val => handleUpdateWithWarning(recalculateReceipt({ ...r, waterNew: val }, room))} 
                          style={{ width: '80px', height: '32px', padding: '0 8px', border: '1px solid var(--success)' }} 
                        />
                      </td>
                      <td style={{ textAlign: 'center', fontWeight: 'bold', color: 'var(--success)' }}>{r.waterUsed || 0}</td>
                      <td className="small">{formatMoney(r.waterAmount || 0)}</td>

                      <td>
                        <input 
                          type="number" 
                          value={r.other} 
                          onChange={e => handleUpdateWithWarning(recalculateReceipt({ ...r, other: e.target.value }, room))} 
                          style={{ width: '80px', height: '32px', padding: '0 8px' }} 
                        />
                      </td>
                      <td style={{ fontWeight: '700', color: 'var(--primary)' }}>{formatMoney(r.total)}</td>
                      <td>
                        <div style={{ display: 'flex', gap: '4px' }}>
                          <button className="secondary-btn sm" title="In phiếu" onClick={() => onPrintBatch([r])}>🖨️</button>
                          <button className="secondary-btn sm" title="Xem QR" onClick={() => onView(r)}>📱</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {monthlyReceipts.length === 0 && <tr><td colSpan="13" style={{ textAlign: 'center', padding: '40px' }}>Chưa có phiếu thu cho tháng {selectedMonth}. Hãy bấm "Tạo hàng loạt" để bắt đầu.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="widget liquid-glass" style={{ padding: 0 }}>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Phòng</th><th>Tháng</th><th>Tổng tiền</th><th>Đã trả</th><th>Còn nợ</th><th>Trạng thái</th><th>Hành động</th></tr></thead>
              <tbody>
                {monthlyReceipts.map(r => (
                  <tr key={r.id}>
                    <td><b>P{r.roomId}</b></td><td>{r.month}</td><td>{formatMoney(r.total)}</td><td>{formatMoney(r.paidAmount)}</td><td style={{ color: r.total - r.paidAmount > 0 ? 'var(--danger)' : 'inherit' }}>{formatMoney(r.total - r.paidAmount)}</td><td><span className={`status-badge-liquid ${r.status === 'Đã thanh toán' ? 'active' : r.status === 'Nợ một phần' ? 'notice' : 'debt'}`}>{r.status}</span></td>
                    <td><div style={{ display: 'flex', gap: '8px' }}><button className="primary-btn sm" onClick={() => onPay(r)}>💸 Thu tiền</button><button className="secondary-btn sm" onClick={() => onView(r)}>📄 Chi tiết</button></div></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function SettingsTab({ data, setData, bankInfo, setBankInfo, onReset }) {
  const [tempIndices, setTempIndices] = useState(data.rooms.map(r => ({ id: r.id, electric: r.initialElectric || r.electricStart || 0, water: r.initialWater || r.waterStart || 0 })));

  const handleUpdateIndices = () => {
    setData(old => ({
      ...old,
      rooms: old.rooms.map(r => {
        const found = tempIndices.find(ti => ti.id === r.id);
        if (found) {
          return { ...r, initialElectric: parseLocaleNumber(found.electric), initialWater: parseLocaleNumber(found.water), electricStart: parseLocaleNumber(found.electric), waterStart: parseLocaleNumber(found.water) };
        }
        return r;
      })
    }));
  };

  return (
    <div className="stack" style={{ gap: '24px' }}>
      <div className="widget liquid-glass" style={{ maxWidth: '800px' }}>
        <section>
          <h2 style={{ marginBottom: '16px' }}>🏦 Thông tin chuyển khoản</h2>
          <div className="form-grid-v2">
            <label>Ngân hàng <input value={bankInfo.bankName} onChange={e => setBankInfo({...bankInfo, bankName: e.target.value})} /></label>
            <label>Mã VietQR <input value={bankInfo.bankCode} onChange={e => setBankInfo({...bankInfo, bankCode: e.target.value})} /></label>
            <label>Số tài khoản <input value={bankInfo.accountNo} onChange={e => setBankInfo({...bankInfo, accountNo: e.target.value})} /></label>
            <label>Chủ tài khoản <input value={bankInfo.accountName} onChange={e => setBankInfo({...bankInfo, accountName: e.target.value})} /></label>
          </div>
        </section>
      </div>

      <div className="widget liquid-glass" style={{ maxWidth: '800px' }}>
        <section>
          <h2 style={{ marginBottom: '16px' }}>⚡ Cập nhật chỉ số đầu kỳ (Chỉ số cũ)</h2>
          <p className="muted small" style={{ marginBottom: '16px' }}>Dùng để khởi tạo chỉ số cho phiếu tháng đầu tiên nếu chưa có lịch sử.</p>
          <div className="table-wrap" style={{ maxHeight: '400px', overflowY: 'auto' }}>
            <table>
              <thead>
                <tr>
                  <th>Phòng</th>
                  <th>Điện cũ (đầu kỳ)</th>
                  <th>Nước cũ (đầu kỳ)</th>
                </tr>
              </thead>
              <tbody>
                {tempIndices.map((ti, idx) => (
                  <tr key={ti.id}>
                    <td><b>P{ti.id}</b></td>
                    <td>
                      <DecimalInput 
                        value={ti.electric} 
                        onChange={val => {
                          const newIndices = [...tempIndices];
                          newIndices[idx].electric = val;
                          setTempIndices(newIndices);
                        }} 
                        style={{ height: '36px', width: '100%' }}
                      />
                    </td>
                    <td>
                      <DecimalInput 
                        value={ti.water} 
                        onChange={val => {
                          const newIndices = [...tempIndices];
                          newIndices[idx].water = val;
                          setTempIndices(newIndices);
                        }} 
                        style={{ height: '36px', width: '100%' }}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button className="primary-btn wide" style={{ marginTop: '20px' }} onClick={handleUpdateIndices}>Lưu chỉ số đầu kỳ</button>
        </section>
      </div>

      <div className="widget liquid-glass" style={{ maxWidth: '800px' }}>
        <section>
          <h2 style={{ marginBottom: '16px' }}>⚙️ Hệ thống</h2>
          <button className="secondary-btn danger" onClick={onReset}>⚠️ Khôi phục dữ liệu gốc</button>
        </section>
      </div>
    </div>
  );
}

function RoomDetailModal({ room, data, onClose, onAction }) {
  const [activeTab, setActiveTab] = useState('overview');
  const { label, color, contract } = getRoomStatusInfo(data, room.id);
  const primaryTenant = contract ? getPrimaryTenantByContract(data, contract.id) : null;
  const currentReceipt = contract ? (data.receipts || []).find(r => r.roomId === room.id && r.contractId === contract.id && r.month === INITIAL_MONTH && r.type === 'monthly') : null;
  const allMembers = contract ? (data.memberships || []).filter(m => m.contractId === contract.id).map(m => ({ ...m, tenant: data.tenants.find(t => t.id === m.tenantId) })) : [];
  const roomReceipts = (data.receipts || []).filter(r => r.roomId === room.id).sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));

  const renderOverview = () => (
    <div className="op-grid">
      {/* Card 1: Người đứng tên */}
      <div className="op-card">
        <h3 className="op-card-title">👑 Người đứng tên</h3>
        {primaryTenant ? (
          <div className="stack" style={{ gap: '8px' }}>
            <p className="op-value">{primaryTenant.name}</p>
            <p className="op-label">📞 {primaryTenant.phone}</p>
            <p className="op-label">💳 CCCD: {primaryTenant.cccd}</p>
          </div>
        ) : <p className="muted small">Phòng đang trống</p>}
      </div>

      {/* Card 2: Hợp đồng hiện tại */}
      <div className="op-card">
        <h3 className="op-card-title">📄 Hợp đồng</h3>
        {contract ? (
          <div className="stack" style={{ gap: '8px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Hết hạn:</span><span className="op-value" style={{ fontSize: '13px' }}>{contract.endDate?.split('-').reverse().join('/')}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Tiền thuê:</span><span className="op-value" style={{ fontSize: '13px' }}>{formatMoney(contract.rent)}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Tiền cọc:</span><span className="op-value" style={{ fontSize: '13px' }}>{formatMoney(contract.deposit)}</span></div>
          </div>
        ) : <p className="muted small">Chưa có hợp đồng</p>}
      </div>

      {/* Card 3: Phiếu tháng hiện tại */}
      <div className="op-card" style={{ gridColumn: 'span 2' }}>
        <h3 className="op-card-title">🧾 Phiếu tháng {INITIAL_MONTH}</h3>
        {currentReceipt ? (
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <p className="op-value">{formatMoney(currentReceipt.total)}</p>
              <span className={`status-badge-liquid ${currentReceipt.status === 'Đã thanh toán' ? 'active' : currentReceipt.status === 'Nợ một phần' ? 'notice' : 'debt'}`} style={{ fontSize: '11px', marginTop: '4px' }}>{currentReceipt.status}</span>
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button className="secondary-btn sm" onClick={() => onAction('view_qr', currentReceipt)}>Xem QR</button>
              {currentReceipt.status !== 'Đã thanh toán' && <button className="primary-btn sm" onClick={() => onAction('pay_receipt', currentReceipt)}>Thu tiền</button>}
            </div>
          </div>
        ) : <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <p className="muted small">Chưa có phiếu tháng này</p>
              {contract && <button className="primary-btn sm" onClick={() => onAction('create_receipt')}>Tạo ngay</button>}
            </div>}
      </div>

      {/* Card 4: Điện nước gần nhất (Lấy từ phiếu tháng mới nhất hoặc chỉ số đầu) */}
      <div className="op-card" style={{ gridColumn: 'span 2' }}>
        <h3 className="op-card-title">⚡ Điện nước gần nhất</h3>
        <div style={{ display: 'flex', gap: '32px' }}>
          <div>
            <p className="op-label">Điện: Chỉ số cũ → Chỉ số mới</p>
            <p className="op-value" style={{ fontSize: '14px' }}>
              {formatLocaleNumber(room.electricOld || room.electricStart || 0)} → {formatLocaleNumber(room.electricNew || '—')}
            </p>
          </div>
          <div>
            <p className="op-label">Nước: Chỉ số cũ → Chỉ số mới</p>
            <p className="op-value" style={{ fontSize: '14px' }}>
              {formatLocaleNumber(room.waterOld || room.waterStart || 0)} → {formatLocaleNumber(room.waterNew || '—')}
            </p>
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '800px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <h2 style={{ fontSize: '28px' }}>Phòng {room.id}</h2>
              <span className={`status-badge-liquid ${color === 'green' ? 'active' : color === 'gray' ? 'vacant' : color}`}>{label}</span>
            </div>
            <p className="muted">{primaryTenant ? `${primaryTenant.name} • ${primaryTenant.phone}` : 'Chưa có hợp đồng hiện tại'}</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>

        <div className="detail-body-v2">
          <nav className="room-op-tabs">
            <button className={`tab-link ${activeTab === 'overview' ? 'active' : ''}`} onClick={() => setActiveTab('overview')}>Tổng quan</button>
            <button className={`tab-link ${activeTab === 'contract' ? 'active' : ''}`} onClick={() => setActiveTab('contract')}>Hợp đồng</button>
            <button className={`tab-link ${activeTab === 'residents' ? 'active' : ''}`} onClick={() => setActiveTab('residents')}>Người ở</button>
            <button className={`tab-link ${activeTab === 'payments' ? 'active' : ''}`} onClick={() => setActiveTab('payments')}>Thanh toán</button>
          </nav>

          <div className="tab-content" style={{ minHeight: '300px' }}>
            {activeTab === 'overview' && renderOverview()}
            {activeTab === 'contract' && (
              <div className="stack" style={{ gap: '16px' }}>
                {contract ? (
                  <div className="op-card">
                    <div className="op-grid">
                      <div className="op-item"><span className="op-label">Ngày bắt đầu</span><span className="op-value">{contract.startDate?.split('-').reverse().join('/')}</span></div>
                      <div className="op-item"><span className="op-label">Ngày hết hạn</span><span className="op-value">{contract.endDate?.split('-').reverse().join('/')}</span></div>
                      <div className="op-item"><span className="op-label">Tiền thuê</span><span className="op-value">{formatMoney(contract.rent)}</span></div>
                      <div className="op-item"><span className="op-label">Tiền cọc</span><span className="op-value">{formatMoney(contract.deposit)}</span></div>
                    </div>
                  </div>
                ) : <p className="muted center" style={{ padding: '40px' }}>Không có hợp đồng hoạt động.</p>}
              </div>
            )}
            {activeTab === 'residents' && (
              <div className="stack" style={{ gap: '12px' }}>
                {allMembers.map(m => (
                  <div key={m.id} className="op-card" style={{ padding: '12px 20px', flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <p className="op-value">{m.tenant?.name} {m.role === 'primary' && '👑'}</p>
                      <p className="op-label">{m.tenant?.phone} • {m.role === 'primary' ? 'Đại diện' : 'Ở cùng'}</p>
                    </div>
                    <button className="secondary-btn sm" onClick={() => onAction('detail_tenant', m.tenant)}>Chi tiết</button>
                  </div>
                ))}
                {contract && <button className="secondary-btn wide" style={{ borderStyle: 'dashed' }}>+ Thêm người ở cùng</button>}
                {!contract && <p className="muted center" style={{ padding: '40px' }}>Phòng đang trống.</p>}
              </div>
            )}
            {activeTab === 'payments' && (
              <div className="table-wrap" style={{ maxHeight: '400px', overflowY: 'auto' }}>
                <table>
                  <thead><tr><th>Tháng</th><th>Tổng tiền</th><th>Trạng thái</th><th>Thao tác</th></tr></thead>
                  <tbody>
                    {roomReceipts.map(r => (
                      <tr key={r.id}>
                        <td>{r.month}</td>
                        <td style={{ fontWeight: '700' }}>{formatMoney(r.total)}</td>
                        <td><span className={`status-badge-liquid ${r.status === 'Đã thanh toán' ? 'active' : r.status === 'Nợ một phần' ? 'notice' : 'debt'}`} style={{ fontSize: '10px' }}>{r.status}</span></td>
                        <td>
                          <div style={{ display: 'flex', gap: '4px' }}>
                            <button className="secondary-btn sm" onClick={() => onAction('view_qr', r)}>📱</button>
                            <button className="secondary-btn sm" style={{ color: '#ef4444' }} onClick={() => onAction('delete_receipt', r)}>🗑️</button>
                          </div>
                        </td>
                      </tr>
                    ))}
                    {roomReceipts.length === 0 && <tr><td colSpan="4" className="center muted" style={{ padding: '20px' }}>Chưa có dữ liệu thanh toán.</td></tr>}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          <div className="op-action-footer">
            {label === 'Trống' ? (
              <><button className="primary-btn" style={{ flex: 1 }} onClick={() => onAction('add_tenant')}>+ Thuê mới</button><button className="secondary-btn" style={{ flex: 1 }} onClick={() => onAction('view_history')}>📜 Xem lịch sử thuê</button></>
            ) : (
              <>
                <button className="primary-btn" onClick={() => onAction('create_receipt')}>⚡ Lập phiếu tháng</button>
                <button className="secondary-btn" onClick={() => onAction('edit_contract')}>✏️ Sửa HĐ</button>
                <button className="secondary-btn" onClick={() => onAction('renew_contract')}>🔄 Gia hạn</button>
                <button className="secondary-btn warning" style={{ flex: 1 }}>⚠️ Báo chuyển</button>
                <button className="secondary-btn danger" style={{ flex: 1 }} onClick={() => onAction('moving_out')}>💸 Tất toán</button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function EditContractModal({ contract, data, onClose, onSave }) {
  const [form, setForm] = useState({ ...contract });
  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '600px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Sửa hợp đồng: {contract.contractNo || contract.id}</h2>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2">
          <div className="form-grid-v2">
            <label>Số hợp đồng <input value={form.contractNo || ''} onChange={e => setForm({...form, contractNo: e.target.value})} /></label>
            <label>Ngày ký <input type="date" value={form.signedDate || ''} onChange={e => setForm({...form, signedDate: e.target.value})} /></label>
            <label>Ngày bắt đầu <input type="date" value={form.startDate} onChange={e => setForm({...form, startDate: e.target.value})} /></label>
            <label>Ngày hết hạn <input type="date" value={form.endDate} onChange={e => setForm({...form, endDate: e.target.value})} /></label>
            <label>Giá thuê <input type="number" value={form.rent} onChange={e => setForm({...form, rent: e.target.value})} /></label>
            <label>Tiền đặt cọc <input type="number" value={form.deposit} onChange={e => setForm({...form, deposit: e.target.value})} /></label>
            <label style={{ gridColumn: 'span 2' }}>Ghi chú <textarea value={form.note || ''} onChange={e => setForm({...form, note: e.target.value})} /></label>
          </div>
          <button className="primary-btn wide" style={{ marginTop: '24px' }} onClick={() => onSave(form)}>Cập nhật hợp đồng</button>
        </div>
      </div>
    </div>
  );
}

function RentalFlowModal({ room, onClose, onSave }) {
  const [step, setStep] = useState('form'); // 'form' | 'preview'
  const [form, setForm] = useState({
    // Tenant Info
    tenantName: '',
    tenantPhone: '',
    tenantCCCD: '',
    tenantCCCDDate: '',
    tenantCCCDPlace: '',
    tenantAddress: '',
    tenantVehicle: '',
    tenantBirthday: '',
    // Contract Info
    contractNo: `HĐ-${room.id}-${new Date().getFullYear()}${(new Date().getMonth()+1).toString().padStart(2,'0')}`,
    signedDate: new Date().toISOString().slice(0, 10),
    startDate: new Date().toISOString().slice(0, 10),
    endDate: addMonthsToDate(new Date().toISOString().slice(0, 10), 12),
    rent: room.rent,
    deposit: room.rent,
    paymentCycleDay: 10,
    // Terms
    electricPrice: room.electricPrice || 3800,
    waterPrice: room.waterPrice || 32000,
    services: {
      cleaning: room.cleaning || 0,
      elevator: room.elevator || 0,
      laundry: room.laundry || 0,
      internet: room.internet || 0
    },
    note: ''
  });

  const handleNext = () => {
    if (!form.tenantName || !form.startDate || !form.endDate) {
      alert('Vui lòng nhập đầy đủ tên khách, ngày bắt đầu và ngày kết thúc!');
      return;
    }
    setStep('preview');
  };

  const handleSave = (shouldPrint = false) => {
    const tenantId = uid('tenant');
    const contractId = uid('contract');
    
    const newTenant = {
      id: tenantId,
      name: form.tenantName,
      phone: form.tenantPhone,
      cccd: form.tenantCCCD,
      cccdDate: form.tenantCCCDDate,
      cccdPlace: form.tenantCCCDPlace,
      address: form.tenantAddress,
      vehicle: form.tenantVehicle,
      birthday: form.tenantBirthday,
      role: 'primary',
      status: 'active',
      createdAt: new Date().toISOString()
    };

    const newContract = {
      id: contractId,
      roomId: room.id,
      status: 'active',
      contractNo: form.contractNo,
      signedDate: form.signedDate,
      startDate: form.startDate,
      endDate: form.endDate,
      rent: Number(form.rent),
      deposit: Number(form.deposit),
      paymentCycleDay: form.paymentCycleDay,
      note: form.note,
      terms: {
        electricPrice: form.electricPrice,
        waterPrice: form.waterPrice,
        services: form.services
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    const memberships = [{
      id: uid('membership'),
      contractId,
      tenantId,
      roomId: room.id,
      role: 'primary',
      status: 'active',
      createdAt: new Date().toISOString()
    }];

    onSave({ tenant: newTenant, contract: newContract, memberships });
    
    if (shouldPrint) {
      setTimeout(() => window.print(), 500);
    }
  };

  if (step === 'preview') {
    return (
      <div className="modal">
        <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '1000px' }}>
          <div className="modal-header no-print">
            <div>
              <h2>Xem trước hợp đồng P{room.id}</h2>
              <p className="muted">Vui lòng kiểm tra kỹ thông tin trước khi lưu</p>
            </div>
            <div className="btn-group">
              <button className="secondary-btn" onClick={() => setStep('form')}>⬅️ Quay lại sửa</button>
              <button className="primary-btn" onClick={() => handleSave(false)}>💾 Lưu hợp đồng</button>
              <button className="primary-btn" style={{ background: 'var(--success)' }} onClick={() => handleSave(true)}>🖨️ Lưu & In</button>
            </div>
          </div>
          <div className="detail-body-v2" style={{ padding: 0 }}>
             {/* Reuse ContractPreview internal logic or just render the paper here */}
             <div className="contract-paper" style={{ boxShadow: 'none', border: '1px solid #eee' }}>
                <div className="contract-header-text">
                  <h1>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h1>
                  <h2>Độc lập – Tự do – Hạnh phúc</h2>
                </div>
                <h1 style={{ textAlign: 'center', margin: '30px 0 20px' }}>HỢP ĐỒNG THUÊ PHÒNG</h1>
                <div className="contract-section">
                  <p>Hôm nay, ngày {form.signedDate.split('-').reverse().join('/')}, chúng tôi gồm:</p>
                  <h3>BÊN CHO THUÊ (Bên A)</h3>
                  <p>Bà: <b>DIỆM THỊ BÌNH</b></p>
                  <p>CCCD Số: <b>019169000011</b></p>
                  <h3>BÊN THUÊ (Bên B)</h3>
                  <p>Ông/Bà: <b>{form.tenantName}</b></p>
                  <p>CCCD số: <b>{form.tenantCCCD || '................'}</b> &nbsp;&nbsp;&nbsp; Ngày cấp: <b>{form.tenantCCCDDate || '................'}</b></p>
                  <p>Nơi ĐKTT: <b>{form.tenantAddress || '................'}</b></p>
                  <p>SĐT/Zalo: <b>{form.tenantPhone || '................'}</b></p>
                </div>
                <div className="main-articles">
                  <p>1. Bên A cho Bên B thuê phòng số <b>{room.id}</b>.</p>
                  <p>2. Thời hạn thuê: Từ ngày <b>{form.startDate.split('-').reverse().join('/')}</b> Đến hết ngày <b>{form.endDate.split('-').reverse().join('/')}</b></p>
                  <p>3. Tiền thuê: <b>{formatMoney(form.rent)}</b>/tháng.</p>
                  <p>4. Tiền đặt cọc: <b>{formatMoney(form.deposit)}</b>.</p>
                  <p>5. Điện: {form.electricPrice}đ/kWh; Nước: {form.waterPrice}đ/m³.</p>
                  <p>6. Dịch vụ khác: Vệ sinh ({formatMoney(form.services.cleaning)}), Thang máy ({formatMoney(form.services.elevator)}), Giặt ({formatMoney(form.services.laundry)}), Internet ({formatMoney(form.services.internet)}).</p>
                </div>
                <div className="signature-row">
                  <div><p><b>BÊN CHO THUÊ</b></p><div className="signature-space"></div><p>DIỆM THỊ BÌNH</p></div>
                  <div><p><b>BÊN THUÊ</b></p><div className="signature-space"></div><p>{form.tenantName}</p></div>
                </div>
             </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '850px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>Tạo hợp đồng mới • Phòng {room.id}</h2>
            <p className="muted">Bước 1: Nhập thông tin người thuê và hợp đồng</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ gap: '24px' }}>
          <div className="settlement-grid" style={{ gridTemplateColumns: '1fr 1fr' }}>
            <div className="stack" style={{ gap: '20px' }}>
              <section>
                <h3 className="form-section-title">👤 Thông tin người thuê</h3>
                <div className="form-grid-v2">
                  <label style={{ gridColumn: 'span 2' }}>Họ và tên <input value={form.tenantName} onChange={e => setForm({...form, tenantName: e.target.value})} placeholder="Nguyễn Văn A" /></label>
                  <label>Số điện thoại <input value={form.tenantPhone} onChange={e => setForm({...form, tenantPhone: e.target.value})} placeholder="09xx..." /></label>
                  <label>Số CCCD <input value={form.tenantCCCD} onChange={e => setForm({...form, tenantCCCD: e.target.value})} /></label>
                  <label>Ngày cấp CCCD <input type="date" value={form.tenantCCCDDate} onChange={e => setForm({...form, tenantCCCDDate: e.target.value})} /></label>
                  <label>Nơi cấp <input value={form.tenantCCCDPlace} onChange={e => setForm({...form, tenantCCCDPlace: e.target.value})} /></label>
                  <label style={{ gridColumn: 'span 2' }}>Địa chỉ thường trú <input value={form.tenantAddress} onChange={e => setForm({...form, tenantAddress: e.target.value})} /></label>
                  <label>Ngày sinh <input type="date" value={form.tenantBirthday} onChange={e => setForm({...form, tenantBirthday: e.target.value})} /></label>
                  <label>Biển số xe <input value={form.tenantVehicle} onChange={e => setForm({...form, tenantVehicle: e.target.value})} /></label>
                </div>
              </section>
            </div>

            <div className="stack" style={{ gap: '20px' }}>
              <section>
                <h3 className="form-section-title">📄 Chi tiết hợp đồng</h3>
                <div className="form-grid-v2">
                  <label>Số hợp đồng <input value={form.contractNo} onChange={e => setForm({...form, contractNo: e.target.value})} /></label>
                  <label>Ngày ký <input type="date" value={form.signedDate} onChange={e => setForm({...form, signedDate: e.target.value})} /></label>
                  <label>Ngày bắt đầu <input type="date" value={form.startDate} onChange={e => setForm({...form, startDate: e.target.value})} /></label>
                  <label>Ngày hết hạn <input type="date" value={form.endDate} onChange={e => setForm({...form, endDate: e.target.value})} /></label>
                  <label>Giá thuê <input type="number" value={form.rent} onChange={e => setForm({...form, rent: e.target.value})} /></label>
                  <label>Tiền đặt cọc <input type="number" value={form.deposit} onChange={e => setForm({...form, deposit: e.target.value})} /></label>
                  <label>Ngày thu tiền <input type="number" value={form.paymentCycleDay} onChange={e => setForm({...form, paymentCycleDay: e.target.value})} /></label>
                </div>
              </section>

              <section>
                <h3 className="form-section-title">💡 Dịch vụ & Điều khoản</h3>
                <div className="form-grid-v2" style={{ gap: '10px' }}>
                  <label>Giá điện <input type="number" value={form.electricPrice} onChange={e => setForm({...form, electricPrice: e.target.value})} /></label>
                  <label>Giá nước <input type="number" value={form.waterPrice} onChange={e => setForm({...form, waterPrice: e.target.value})} /></label>
                  <label style={{ gridColumn: 'span 2' }}>Ghi chú thêm <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} style={{ minHeight: '60px' }} /></label>
                </div>
              </section>
            </div>
          </div>
          <div className="btn-group" style={{ marginTop: '12px' }}>
            <button className="secondary-btn" onClick={onClose} style={{ flex: 1 }}>Hủy bỏ</button>
            <button className="primary-btn" onClick={handleNext} style={{ flex: 2 }}>Tiếp theo: Xem trước hợp đồng ➡️</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function SettlementModal({ room, contract, data, onClose, onSave }) {
  const [form, setForm] = useState({
    actualEndDate: new Date().toISOString().split('T')[0],
    electricNew: room.electricNew || room.electricEnd || 0,
    waterNew: room.waterNew || room.waterEnd || 0,
    unpaidRent: 0,
    cleaningFee: 100000,
    damageFee: 0,
    otherFee: 0,
    note: ''
  });

  const primaryMembership = (data.memberships || []).find(m => m.contractId === contract.id && m.role === 'primary');
  const tenant = primaryMembership ? (data.tenants || []).find(t => t.id === primaryMembership.tenantId) : { name: 'N/A' };

  // Logic tính toán realtime
  const electricOld = Number(room.electricOld || room.electricStart || room.initialElectric || 0);
  const waterOld = Number(room.waterOld || room.waterStart || room.initialWater || 0);
  
  const electricUsed = Math.max(0, Number(form.electricNew) - electricOld);
  const waterUsed = Math.max(0, Number(form.waterNew) - waterOld);
  const electricAmount = electricUsed * Number(room.electricPrice || 3800);
  const waterAmount = waterUsed * Number(room.waterPrice || 32000);
  
  const totalIncurred = electricAmount + waterAmount + Number(form.unpaidRent) + 
                        Number(form.cleaningFee) + Number(form.damageFee) + Number(form.otherFee);
  
  const deposit = Number(contract.deposit || 0);
  const finalBalance = totalIncurred - deposit;

  const isRefund = finalBalance < 0;
  const isDebt = finalBalance > 0;
  const absBalance = Math.abs(finalBalance);

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '1000px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2 style={{ fontSize: '28px' }}>Tất toán / Trả phòng • Phòng {room.id}</h2>
            <p className="muted">Chốt công nợ, điện nước, phí phát sinh và hoàn cọc</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>

        <div className="detail-body-v2">
          <div className="settlement-grid">
            <div className="settlement-form">
              {/* Nhóm 1: Thông tin cơ bản */}
              <section>
                <h3 className="form-section-title">📅 Thông tin trả phòng</h3>
                <div className="form-grid-v2">
                  <label>Ngày trả phòng <input type="date" value={form.actualEndDate} onChange={e => setForm({...form, actualEndDate: e.target.value})} /></label>
                  <label>Người đứng tên <input value={tenant.name} readOnly style={{ background: '#f1f5f9' }} /></label>
                  <label style={{ gridColumn: 'span 2' }}>Tiền cọc đang giữ <input value={formatMoney(deposit)} readOnly style={{ background: '#f1f5f9', fontWeight: 'bold' }} /></label>
                </div>
              </section>

              {/* Nhóm 2: Điện nước */}
              <section>
                <h3 className="form-section-title">⚡ Chỉ số điện nước</h3>
                <div className="form-grid-v2">
                  <label>Điện: Chỉ số cũ <input value={electricOld} readOnly style={{ background: '#f1f5f9' }} /></label>
                  <label>Chỉ số mới <input type="number" value={form.electricNew} onChange={e => setForm({...form, electricNew: e.target.value})} /></label>
                  <label>Nước: Chỉ số cũ <input value={waterOld} readOnly style={{ background: '#f1f5f9' }} /></label>
                  <label>Chỉ số mới <input type="number" value={form.waterNew} onChange={e => setForm({...form, waterNew: e.target.value})} /></label>
                  <div style={{ gridColumn: 'span 2', display: 'flex', gap: '20px', fontSize: '13px', background: '#f8fafc', padding: '12px', borderRadius: '12px' }}>
                    <span>⚡ Tiêu thụ: <b>{electricUsed} kWh</b> = <b>{formatMoney(electricAmount)}</b></span>
                    <span>💧 Tiêu thụ: <b>{waterUsed} m³</b> = <b>{formatMoney(waterAmount)}</b></span>
                  </div>
                </div>
              </section>

              {/* Nhóm 3: Phí phát sinh */}
              <section>
                <h3 className="form-section-title">💸 Phí phát sinh & Ghi chú</h3>
                <div className="form-grid-v2">
                  <label>Tiền phòng còn nợ <input type="number" value={form.unpaidRent} onChange={e => setForm({...form, unpaidRent: e.target.value})} /></label>
                  <label>Phí vệ sinh trả phòng <input type="number" value={form.cleaningFee} onChange={e => setForm({...form, cleaningFee: e.target.value})} /></label>
                  <label>Chi phí hư hỏng <input type="number" value={form.damageFee} onChange={e => setForm({...form, damageFee: e.target.value})} /></label>
                  <label>Phí khác <input type="number" value={form.otherFee} onChange={e => setForm({...form, otherFee: e.target.value})} /></label>
                  <label style={{ gridColumn: 'span 2' }}>Ghi chú tất toán <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} placeholder="Nhập chi tiết các khoản hư hỏng hoặc lý do phí khác..." /></label>
                </div>
              </section>
            </div>

            <div className="settlement-summary-side">
              <div className="summary-card">
                <h3 className="summary-label-main">Kết quả tất toán</h3>
                
                <div className="summary-row"><span>Tổng phát sinh</span><b>{formatMoney(totalIncurred)}</b></div>
                <div className="summary-row"><span>Tiền cọc đối trừ</span><b style={{ color: 'var(--text-muted)' }}>- {formatMoney(deposit)}</b></div>
                
                <div className="summary-total">
                  <span className="summary-label-main">{isRefund ? 'Số tiền hoàn khách' : 'Khách cần trả thêm'}</span>
                  <p className="summary-amount" style={{ color: isRefund ? 'var(--success)' : isDebt ? 'var(--danger)' : 'var(--text-main)' }}>
                    {formatMoney(absBalance)}
                  </p>
                  <p style={{ fontSize: '13px', fontWeight: '500' }}>
                    {isRefund ? '✨ Cần hoàn trả tiền cọc cho khách' : isDebt ? '⚠️ Khách thuê cần đóng thêm tiền' : '✅ Công nợ đã được tất toán đủ'}
                  </p>
                </div>

                <div className="warning-box">
                  <b>Lưu ý sau khi hoàn tất:</b>
                  <ul>
                    <li>Hợp đồng sẽ chuyển sang "Đã kết thúc"</li>
                    <li>Phòng {room.id} sẽ trở về trạng thái trống</li>
                    <li>Lịch sử tất toán sẽ được lưu lại</li>
                  </ul>
                </div>

                <div className="stack" style={{ gap: '12px', marginTop: 'auto' }}>
                  <button className="primary-btn wide" onClick={() => {
                    if (window.confirm('Xác nhận hoàn tất mọi thủ tục trả phòng và chốt công nợ?')) {
                      onSave({
                        ...form,
                        electricUsed,
                        electricAmount,
                        waterUsed,
                        waterAmount,
                        totalIncurred,
                        depositUsed: deposit,
                        mustCollect: isDebt ? absBalance : 0,
                        mustRefund: isRefund ? absBalance : 0,
                        contractId: contract.id,
                        roomId: room.id
                      });
                    }
                  }}>Hoàn tất trả phòng</button>
                  <button className="secondary-btn wide" onClick={onClose}>Hủy bỏ</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function RenewalModal({ contract, data, onClose, onSave }) {
  const room = data.rooms.find(r => r.id === contract.roomId);
  const tenant = getPrimaryTenantByContract(data, contract.id) || { name: 'N/A', phone: 'N/A' };
  const [showAppendix, setShowAppendix] = useState(false);
  
  const [form, setForm] = useState({
    signedDate: new Date().toISOString().slice(0, 10),
    newStartDate: contract.endDate || new Date().toISOString().slice(0, 10),
    newEndDate: addMonthsToDate(contract.endDate, 12) || '',
    keepPricing: true,
    newRent: contract.rent,
    newDeposit: contract.deposit,
    note: ''
  });

  const isValid = form.newEndDate && (!contract.endDate || form.newEndDate > contract.endDate);
  const rentChanged = !form.keepPricing && Number(form.newRent) !== Number(contract.rent);
  const depositChanged = !form.keepPricing && Number(form.newDeposit) !== Number(contract.deposit);

  const handlePrintAppendix = () => {
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);
    
    const doc = iframe.contentWindow.document;
    doc.write(`
      <html>
        <head>
          <title>Phụ lục gia hạn - P${contract.roomId}</title>
          <style>
            @import url('https://fonts.googleapis.com/css2?family=Times+New+Roman&display=swap');
            body { font-family: 'Times New Roman', Times, serif; line-height: 1.5; color: black; margin: 0; padding: 0; }
            .appendix-print { width: 210mm; min-height: 297mm; padding: 20mm 25mm; margin: auto; box-sizing: border-box; background: white; }
            @page { size: A4; margin: 0; }
            h1, h2, h3 { text-align: center; margin: 10px 0; font-size: 16px; text-transform: uppercase; }
            .section { margin-top: 20px; }
            .section h4 { text-decoration: underline; margin-bottom: 5px; }
            .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 30px; }
            .grid div { text-align: center; }
            .signature-space { height: 80px; }
            table { width: 100%; border-collapse: collapse; margin: 15px 0; }
            table th, table td { border: 1px solid black; padding: 8px; text-align: left; font-size: 14px; }
            .muted { color: #555; }
            b { font-weight: bold; }
          </style>
        </head>
        <body>
          <div class="appendix-print">
            <div style="text-align: center; margin-bottom: 30px;">
              <h2 style="margin: 0;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h2>
              <h3 style="margin: 5px 0; text-transform: none;">Độc lập – Tự do – Hạnh phúc</h3>
              <div style="width: 150px; border-top: 1px solid black; margin: 10px auto;"></div>
            </div>

            <h1>PHỤ LỤC GIA HẠN HỢP ĐỒNG THUÊ PHÒNG</h1>
            <p style="text-align: center;">Số phụ lục: PL-${contract.contractNo || contract.id}-${form.signedDate.split('-').reverse().join('')}</p>
            <p style="text-align: center;">Kèm theo Hợp đồng thuê phòng số: ${contract.contractNo || contract.id}</p>

            <div class="section">
              <p>Hôm nay, ngày ${form.signedDate.split('-').reverse().join('/')}, tại Hà Nội, chúng tôi gồm:</p>
              
              <h4>BÊN CHO THUÊ — BÊN A</h4>
              <p>Bà: <b>DIỆM THỊ BÌNH</b></p>
              <p>Số điện thoại: <b>0123.456.789</b></p>
              <p>Địa chỉ: <b>Số 28, ngách 1, ngõ 162 Khương Đình, Thanh Xuân, Hà Nội</b></p>
              <p>Số tài khoản nhận tiền: <b>8847214661</b> - Ngân hàng: <b>BIDV</b></p>

              <h4>BÊN THUÊ — BÊN B</h4>
              <p>Ông/Bà: <b>${tenant.name}</b></p>
              <p>Số điện thoại: <b>${tenant.phone}</b></p>
              <p>CCCD/CMND: <b>${tenant.cccd || '................'}</b></p>
              <p>Đang thuê phòng: <b>${contract.roomId}</b></p>

              <p>Hai bên thống nhất ký phụ lục này để gia hạn thời hạn thuê phòng theo các nội dung sau:</p>
            </div>

            <div class="section">
              <h4>ĐIỀU 1. GIA HẠN THỜI HẠN THUÊ</h4>
              <p>Hai bên thống nhất gia hạn thời hạn thuê phòng ${contract.roomId} như sau:</p>
              <ul>
                <li>Ngày bắt đầu gia hạn: <b>${form.newStartDate.split('-').reverse().join('/')}</b></li>
                <li>Ngày hết hạn cũ: <b>${contract.endDate?.split('-').reverse().join('/') || '—'}</b></li>
                <li>Ngày hết hạn mới: <b>${form.newEndDate.split('-').reverse().join('/')}</b></li>
              </ul>
              <p>Sau thời hạn trên, nếu Bên B tiếp tục có nhu cầu thuê, hai bên sẽ thỏa thuận gia hạn tiếp hoặc ký hợp đồng/phụ lục mới.</p>
            </div>

            <div class="section">
              <h4>ĐIỀU 2. GIÁ THUÊ VÀ TIỀN CỌC</h4>
              <p>Kể từ ngày ${form.newStartDate.split('-').reverse().join('/')}, các khoản tiền được áp dụng như sau:</p>
              <table>
                <thead>
                  <tr><th>Nội dung</th><th>Trước gia hạn</th><th>Sau gia hạn</th></tr>
                </thead>
                <tbody>
                  <tr>
                    <td>Giá thuê phòng/tháng</td>
                    <td>${formatMoney(contract.rent)}</td>
                    <td><b>${formatMoney(form.keepPricing ? contract.rent : form.newRent)}</b></td>
                  </tr>
                  <tr>
                    <td>Tiền cọc</td>
                    <td>${formatMoney(contract.deposit)}</td>
                    <td><b>${formatMoney(form.keepPricing ? contract.deposit : form.newDeposit)}</b></td>
                  </tr>
                </tbody>
              </table>
              <p>Tiền thuê phòng được thanh toán theo chu kỳ hàng tháng, vào ngày <b>${contract.paymentCycleDay || 5}</b> hàng tháng.</p>
            </div>

            <div class="section">
              <h4>ĐIỀU 3. ĐIỆN, NƯỚC VÀ CÁC KHOẢN PHÍ DỊCH VỤ</h4>
              <p>Các khoản điện, nước và phí dịch vụ tiếp tục được áp dụng theo hợp đồng thuê phòng đã ký:</p>
              <ul>
                <li>Đơn giá điện: <b>${room.electricPrice}đ/kWh</b></li>
                <li>Đơn giá nước: <b>${formatMoney(room.waterPrice)}/m³</b></li>
                <li>Phí dịch vụ cố định: <b>${formatMoney((room.cleaning || 0) + (room.elevator || 0) + (room.laundry || 0) + (room.internet || 0))}</b></li>
              </ul>
            </div>

            <div class="section">
              <h4>ĐIỀU 4. HIỆU LỰC</h4>
              <p>Phụ lục này có hiệu lực kể từ ngày ${form.signedDate.split('-').reverse().join('/')}. Các nội dung khác không được sửa đổi trong phụ lục này vẫn tiếp tục thực hiện theo Hợp đồng gốc.</p>
            </div>

            <div class="grid">
              <div>
                <p><b>ĐẠI DIỆN BÊN A</b></p>
                <div class="signature-space"></div>
                <p><b>DIỆM THỊ BÌNH</b></p>
              </div>
              <div>
                <p><b>ĐẠI DIỆN BÊN B</b></p>
                <div class="signature-space"></div>
                <p><b>${tenant.name}</b></p>
              </div>
            </div>
          </div>
        </body>
      </html>
    `);
    doc.close();
    
    iframe.contentWindow.focus();
    setTimeout(() => {
      iframe.contentWindow.print();
      document.body.removeChild(iframe);
    }, 500);
  };


  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '700px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2 style={{ fontSize: '24px' }}>Gia hạn hợp đồng • Phòng {contract.roomId}</h2>
            <p className="muted">{tenant.name} • Hết hạn hiện tại: {contract.endDate?.split('-').reverse().join('/') || 'Chưa có'}</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>

        <div className="detail-body-v2 stack" style={{ gap: '24px' }}>
          {/* Thông tin hiện tại */}
          <section className="op-card" style={{ background: '#f8fafc', borderStyle: 'dashed' }}>
            <h3 className="op-card-title">📌 Thông tin hiện tại</h3>
            <div className="op-grid">
              <div className="op-item"><span className="op-label">Ngày bắt đầu</span><span className="op-value">{contract.startDate?.split('-').reverse().join('/')}</span></div>
              <div className="op-item"><span className="op-label">Hết hạn hiện tại</span><span className="op-value">{contract.endDate?.split('-').reverse().join('/') || 'N/A'}</span></div>
              <div className="op-item"><span className="op-label">Giá thuê</span><span className="op-value">{formatMoney(contract.rent)}</span></div>
              <div className="op-item"><span className="op-label">Tiền cọc</span><span className="op-value">{formatMoney(contract.deposit)}</span></div>
            </div>
          </section>

          {/* Form gia hạn */}
          <section className="stack" style={{ gap: '16px' }}>
            <h3 className="form-section-title">✍️ Chi tiết gia hạn</h3>
            <div className="form-grid-v2">
              <label>Ngày ký phụ lục <input type="date" value={form.signedDate} onChange={e => setForm({...form, signedDate: e.target.value})} /></label>
              <label>Bắt đầu gia hạn <input type="date" value={form.newStartDate} onChange={e => setForm({...form, newStartDate: e.target.value})} /></label>
              <label style={{ gridColumn: 'span 2' }}>
                Ngày hết hạn mới
                <input type="date" value={form.newEndDate} onChange={e => setForm({...form, newEndDate: e.target.value})} />
                {!contract.endDate && <p className="danger small" style={{ marginTop: '4px' }}>⚠ Chưa có ngày hết hạn cũ. Vui lòng chọn ngày mới.</p>}
                {contract.endDate && form.newEndDate <= contract.endDate && <p className="danger small" style={{ marginTop: '4px' }}>⚠ Ngày hết hạn mới phải sau ngày {contract.endDate.split('-').reverse().join('/')}</p>}
              </label>
            </div>

            <div className="op-card">
              <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', margin: 0 }}>
                <input type="checkbox" checked={form.keepPricing} onChange={e => setForm({...form, keepPricing: e.target.checked})} style={{ width: '18px', height: '18px' }} />
                <b>Giữ nguyên giá thuê và tiền cọc</b>
              </label>
              
              {!form.keepPricing && (
                <div className="form-grid-v2" style={{ marginTop: '12px' }}>
                  <label>
                    Giá thuê mới
                    <input type="number" value={form.newRent} onChange={e => setForm({...form, newRent: e.target.value})} />
                    {rentChanged && <span className="status-badge-liquid notice" style={{ fontSize: '10px' }}>Thay đổi giá</span>}
                  </label>
                  <label>
                    Tiền cọc mới
                    <input type="number" value={form.newDeposit} onChange={e => setForm({...form, newDeposit: e.target.value})} />
                    {depositChanged && <span className="status-badge-liquid notice" style={{ fontSize: '10px' }}>Thay đổi cọc</span>}
                  </label>
                </div>
              )}
            </div>

            <label>
              Ghi chú phụ lục gia hạn
              <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} placeholder="Ví dụ: Ưu đãi giảm giá 3 tháng đầu, cam kết không tăng giá trong 2 năm..." style={{ minHeight: '80px' }} />
            </label>
          </section>

          <div className="btn-group">
            <button className="secondary-btn" onClick={onClose}>Hủy</button>
            <button className="secondary-btn" onClick={() => setShowAppendix(!showAppendix)}>
              {showAppendix ? '👁️ Ẩn phụ lục' : '📄 Xem phụ lục'}
            </button>
            <button className="primary-btn" onClick={() => onSave(form)} disabled={!isValid} style={{ flex: 2 }}>
              🚀 Gia hạn hợp đồng
            </button>
          </div>

          {showAppendix && (
            <div className="appendix-preview-scroll" style={{ marginTop: '20px', borderTop: '1px solid rgba(255,255,255,0.1)', paddingTop: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
                <h3 className="form-section-title" style={{ margin: 0 }}>📄 Xem trước phụ lục (A4)</h3>
                <button className="secondary-btn sm" onClick={handlePrintAppendix}>🖨️ In phụ lục này</button>
              </div>
              <div className="appendix-container-scroll" style={{ maxHeight: '500px', overflowY: 'auto', background: 'rgba(0,0,0,0.2)', padding: '20px', borderRadius: '12px' }}>
                <div className="appendix-paper-a4" style={{ transform: 'scale(0.8)', transformOrigin: 'top center', margin: '0 auto', marginBottom: '-150px' }}>
                   <AppendixContent contract={contract} tenant={tenant} form={form} room={room} />
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function AppendixContent({ contract, tenant, form, room }) {
  return (
    <div className="appendix-content-v1" style={{ 
      width: '210mm', 
      minHeight: '297mm', 
      padding: '20mm 25mm', 
      background: 'white', 
      color: 'black', 
      fontFamily: '"Times New Roman", Times, serif',
      lineHeight: '1.5',
      boxSizing: 'border-box',
      textAlign: 'left'
    }}>
      <div style={{ textAlign: 'center', marginBottom: '30px' }}>
        <h2 style={{ fontSize: '16px', fontWeight: 'bold', margin: 0 }}>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h2>
        <h3 style={{ fontSize: '14px', margin: '5px 0', textTransform: 'none' }}>Độc lập – Tự do – Hạnh phúc</h3>
        <div style={{ width: '150px', borderTop: '1px solid black', margin: '10px auto' }}></div>
      </div>

      <h1 style={{ textAlign: 'center', fontSize: '18px', fontWeight: 'bold', textTransform: 'uppercase', marginBottom: '10px' }}>PHỤ LỤC GIA HẠN HỢP ĐỒNG THUÊ PHÒNG</h1>
      <p style={{ textAlign: 'center', fontSize: '13px', margin: '2px 0' }}>Số phụ lục: PL-{contract.contractNo || contract.id}-${form.signedDate.split('-').reverse().join('')}</p>
      <p style={{ textAlign: 'center', fontSize: '13px', marginBottom: '30px' }}>Kèm theo Hợp đồng thuê phòng số: {contract.contractNo || contract.id}</p>

      <div className="appendix-section">
        <p>Hôm nay, ngày {form.signedDate.split('-').reverse().join('/')}, tại Hà Nội, chúng tôi gồm:</p>
        
        <h4 style={{ fontWeight: 'bold', textDecoration: 'underline', marginTop: '15px', marginBottom: '5px' }}>BÊN CHO THUÊ — BÊN A</h4>
        <p>Bà: <b>DIỆM THỊ BÌNH</b></p>
        <p>Số điện thoại: <b>0123.456.789</b></p>
        <p>Địa chỉ: <b>Số 28, ngách 1, ngõ 162 Khương Đình, Thanh Xuân, Hà Nội</b></p>
        <p>Số tài khoản nhận tiền: <b>8847214661</b> - Ngân hàng: <b>BIDV</b></p>

        <h4 style={{ fontWeight: 'bold', textDecoration: 'underline', marginTop: '15px', marginBottom: '5px' }}>BÊN THUÊ — BÊN B</h4>
        <p>Ông/Bà: <b>{tenant.name}</b></p>
        <p>Số điện thoại: <b>{tenant.phone}</b></p>
        <p>CCCD/CMND: <b>{tenant.cccd || '................'}</b></p>
        <p>Đang thuê phòng: <b>{contract.roomId}</b></p>

        <p style={{ marginTop: '15px' }}>Hai bên thống nhất ký phụ lục này để gia hạn thời hạn thuê phòng theo các nội dung sau:</p>
      </div>

      <div className="appendix-section" style={{ marginTop: '20px' }}>
        <h4 style={{ fontWeight: 'bold' }}>ĐIỀU 1. GIA HẠN THỜI HẠN THUÊ</h4>
        <p>Hai bên thống nhất gia hạn thời hạn thuê phòng {contract.roomId} như sau:</p>
        <ul style={{ paddingLeft: '25px', margin: '10px 0' }}>
          <li>Ngày bắt đầu gia hạn: <b>{form.newStartDate.split('-').reverse().join('/')}</b></li>
          <li>Ngày hết hạn cũ: <b>{contract.endDate?.split('-').reverse().join('/') || '—'}</b></li>
          <li>Ngày hết hạn mới: <b>{form.newEndDate.split('-').reverse().join('/')}</b></li>
        </ul>
        <p>Sau thời hạn trên, nếu Bên B tiếp tục có nhu cầu thuê, hai bên sẽ thỏa thuận gia hạn tiếp hoặc ký hợp đồng/phụ lục mới.</p>
      </div>

      <div className="appendix-section" style={{ marginTop: '20px' }}>
        <h4 style={{ fontWeight: 'bold' }}>ĐIỀU 2. GIÁ THUÊ VÀ TIỀN CỌC</h4>
        <p>Kể từ ngày {form.newStartDate.split('-').reverse().join('/')}, các khoản tiền được áp dụng như sau:</p>
        <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '10px' }}>
          <thead>
            <tr>
              <th style={{ border: '1px solid black', padding: '8px', textAlign: 'left' }}>Nội dung</th>
              <th style={{ border: '1px solid black', padding: '8px', textAlign: 'left' }}>Trước gia hạn</th>
              <th style={{ border: '1px solid black', padding: '8px', textAlign: 'left' }}>Sau gia hạn</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style={{ border: '1px solid black', padding: '8px' }}>Giá thuê phòng/tháng</td>
              <td style={{ border: '1px solid black', padding: '8px' }}>{formatMoney(contract.rent)}</td>
              <td style={{ border: '1px solid black', padding: '8px' }}><b>{formatMoney(form.keepPricing ? contract.rent : form.newRent)}</b></td>
            </tr>
            <tr>
              <td style={{ border: '1px solid black', padding: '8px' }}>Tiền cọc</td>
              <td style={{ border: '1px solid black', padding: '8px' }}>{formatMoney(contract.deposit)}</td>
              <td style={{ border: '1px solid black', padding: '8px' }}><b>{formatMoney(form.keepPricing ? contract.deposit : form.newDeposit)}</b></td>
            </tr>
          </tbody>
        </table>
        <p style={{ marginTop: '10px' }}>Tiền thuê phòng được thanh toán theo chu kỳ hàng tháng, vào ngày <b>{contract.paymentCycleDay || 5}</b> hàng tháng.</p>
      </div>

      <div className="appendix-section" style={{ marginTop: '20px' }}>
        <h4 style={{ fontWeight: 'bold' }}>ĐIỀU 3. ĐIỆN, NƯỚC VÀ CÁC KHOẢN PHÍ DỊCH VỤ</h4>
        <p>Các khoản điện, nước và phí dịch vụ tiếp tục được áp dụng theo hợp đồng thuê phòng đã ký:</p>
        <ul style={{ paddingLeft: '25px', margin: '10px 0' }}>
          <li>Đơn giá điện: <b>{room.electricPrice}đ/kWh</b></li>
          <li>Đơn giá nước: <b>{formatMoney(room.waterPrice)}/m³</b></li>
          <li>Phí dịch vụ cố định: <b>{formatMoney((room.cleaning || 0) + (room.elevator || 0) + (room.laundry || 0) + (room.internet || 0))}</b></li>
        </ul>
      </div>

      <div className="appendix-section" style={{ marginTop: '20px' }}>
        <h4 style={{ fontWeight: 'bold' }}>ĐIỀU 4. HIỆU LỰC</h4>
        <p>Phụ lục này có hiệu lực kể từ ngày {form.signedDate.split('-').reverse().join('/')}. Các nội dung khác không được sửa đổi trong phụ lục này vẫn tiếp tục thực hiện theo Hợp đồng gốc.</p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', marginTop: '40px', textAlign: 'center' }}>
        <div>
          <p><b>ĐẠI DIỆN BÊN A</b></p>
          <div style={{ height: '70px' }}></div>
          <p><b>DIỆM THỊ BÌNH</b></p>
        </div>
        <div>
          <p><b>ĐẠI DIỆN BÊN B</b></p>
          <div style={{ height: '70px' }}></div>
          <p><b>{tenant.name}</b></p>
        </div>
      </div>
    </div>
  );
}

function EditTenantModal({ tenant, onClose, onSave }) {
  const [form, setForm] = useState({ ...tenant });
  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '500px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Sửa người thuê: {tenant.name}</h2>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2">
          <div className="form-grid-v2">
            <label style={{ gridColumn: 'span 2' }}>Họ tên <input value={form.name} onChange={e => setForm({...form, name: e.target.value})} /></label>
            <label>SĐT <input value={form.phone} onChange={e => setForm({...form, phone: e.target.value})} /></label>
            <label>CCCD <input value={form.cccd} onChange={e => setForm({...form, cccd: e.target.value})} /></label>
            <label>Biển số xe <input value={form.vehicle || ''} onChange={e => setForm({...form, vehicle: e.target.value})} /></label>
            <label>Vai trò 
              <select value={form.role} onChange={e => setForm({...form, role: e.target.value})}>
                <option value="primary">Người đứng tên</option>
                <option value="secondary">Người ở cùng</option>
              </select>
            </label>
            <label style={{ gridColumn: 'span 2' }}>Địa chỉ <input value={form.address || ''} onChange={e => setForm({...form, address: e.target.value})} /></label>
            <label style={{ gridColumn: 'span 2' }}>Ghi chú <textarea value={form.note || ''} onChange={e => setForm({...form, note: e.target.value})} /></label>
          </div>
          <button className="primary-btn wide" style={{ marginTop: '24px' }} onClick={() => onSave(form)}>Lưu thay đổi</button>
        </div>
      </div>
    </div>
  );
}

function ReceiptModal({ receipt, room, data, bankInfo, onClose, onPrev, onNext }) {
  if (!receipt || !room) return null;
  const isMonthly = receipt.type === 'monthly';
  const tenant = getPrimaryTenantByContract(data, receipt.contractId) || { name: '—', phone: '—' };

  function handlePrintReceipt() {
    const content = document.getElementById('printable-receipt');
    if (!content) return;
    
    // Create iframe
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);
    
    const doc = iframe.contentWindow.document;
    
    // Copy all styles
    const styles = document.querySelectorAll('link[rel="stylesheet"], style');
    doc.write('<html><head><title>In phiếu thu</title>');
    styles.forEach(s => doc.write(s.outerHTML));
    
    // Add A5 specific style for iframe
    doc.write(`
      <style>
        body { margin: 0; padding: 0; background: white; font-family: 'Inter', 'Be Vietnam Pro', sans-serif; }
        .printable-receipt { 
          width: 148mm !important; 
          margin: 0 !important;
          padding: 2mm 4mm !important; /* Extremely tight padding */
          box-shadow: none !important;
          border: none !important;
          font-size: 12px; /* Smaller base font */
        }
        @page { size: A5 portrait; margin: 0; }
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        
        /* Maximum Compression */
        .receipt-header-v4 { margin-bottom: 4px !important; }
        .receipt-header-v4 .brand { font-size: 14px !important; margin: 0 !important; }
        .receipt-header-v4 .sub { display: none !important; } /* Hide slogan to save space */
        .receipt-header-v4 h1 { font-size: 18px !important; margin: 0 !important; }
        
        .info-bar-v4 { gap: 4px !important; padding: 4px 10px !important; margin-bottom: 8px !important; border-radius: 8px !important; }
        .info-bar-v4 .value { font-size: 13px !important; }
        
        .receipt-body-v4 { gap: 10px !important; grid-template-columns: 1fr 210px !important; }
        .receipt-items-section h3 { margin-bottom: 8px !important; padding-bottom: 4px !important; }
        .charge-row-v4 { padding: 4px 0 !important; }
        .charge-row-v4 .name { font-size: 13px !important; }
        .charge-row-v4 .amount { font-size: 14px !important; }
        
        .payment-card-v4 { padding: 10px !important; gap: 8px !important; border-radius: 12px !important; }
        .qr-box-v4 { padding: 6px !important; gap: 2px !important; }
        .qr-box-v4 img { width: 110px !important; height: 110px !important; }
        .qr-box-v4 p { font-size: 9px !important; }
        
        .total-summary-v4 { padding: 10px !important; margin-top: 4px !important; border-radius: 10px !important; }
        .total-summary-v4 .total-amount { font-size: 18px !important; }
        
        .signatures-v4 { margin-top: 10px !important; margin-bottom: 10px !important; gap: 10px !important; }
        .signature-space-v4 { height: 30px !important; }
        .signature-box h4 { font-size: 12px !important; }
        
        .thank-you-v4 { padding-top: 10px !important; margin-top: 10px !important; font-size: 10px !important; }
        .tenant-info-v4 { padding: 8px !important; margin-top: 8px !important; font-size: 11px !important; }
      </style>
    `);
    
    doc.write('</head><body>');
    doc.write(content.outerHTML);
    doc.write('</body></html>');
    doc.close();
    
    iframe.contentWindow.focus();
    setTimeout(() => {
      iframe.contentWindow.print();
      document.body.removeChild(iframe);
    }, 500);
  }

  const handleCopyTransfer = () => {
    navigator.clipboard.writeText(transferContent(receipt));
    alert('Đã copy nội dung chuyển khoản!');
  };

  return (
    <div className="modal no-print-backdrop" onClick={onClose}>
      <div className="receipt-modal-v3 liquid-glass" onClick={e => e.stopPropagation()}>
        <div className="modal-header no-print" style={{ padding: '16px 24px', borderBottom: '1px solid rgba(255,255,255,0.1)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
            <div style={{ display: 'flex', gap: '4px' }}>
              <button className="secondary-btn" onClick={onPrev} disabled={!onPrev} style={{ borderRadius: '50%', width: '32px', height: '32px', padding: 0 }}>←</button>
              <button className="secondary-btn" onClick={onNext} disabled={!onNext} style={{ borderRadius: '50%', width: '32px', height: '32px', padding: 0 }}>→</button>
            </div>
            <div>
              <h2 style={{ fontSize: '18px', margin: 0 }}>Phiếu thu P{receipt.roomId}</h2>
              <p className="muted small">Tháng {receipt.month}</p>
            </div>
          </div>
          <button className="secondary-btn" onClick={onClose} style={{ borderRadius: '50%', width: '40px', height: '40px', padding: 0 }}>✕</button>
        </div>

        <div className="modal-body-v3 scrollable">
          <PrintableReceipt 
            receipt={receipt} 
            room={room} 
            tenant={tenant} 
            bankInfo={bankInfo} 
          />
        </div>

        <div className="modal-footer-v3 no-print" style={{ padding: '16px 24px', borderTop: '1px solid rgba(255,255,255,0.1)', display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
          <button className="secondary-btn" onClick={handleCopyTransfer}>📋 Copy nội dung CK</button>
          <button className="primary-btn" onClick={handlePrintReceipt}>🖨️ In phiếu thu</button>
        </div>
      </div>
    </div>
  );
}

function PrintableReceipt({ receipt, room, tenant, bankInfo }) {
  const isMonthly = receipt.type === 'monthly';
  const eOld = getElectricOld(receipt);
  const eNew = getElectricNew(receipt);
  const wOld = getWaterOld(receipt);
  const wNew = getWaterNew(receipt);
  
  const statusColor = receipt.status === 'Đã thanh toán' ? '#166534' : receipt.status === 'Nợ một phần' ? '#92400e' : '#991b1b';
  const statusBg = receipt.status === 'Đã thanh toán' ? '#dcfce7' : receipt.status === 'Nợ một phần' ? '#fef3c7' : '#fee2e2';

  return (
    <div id="printable-receipt" className="printable-receipt">
      <header className="receipt-header-v4">
        <div className="brand-box">
          <h3 className="brand">ROOM MANAGER</h3>
          <p className="sub">Hệ thống quản lý phòng trọ chuyên nghiệp</p>
        </div>
        <div className="status-section">
          <h1 className="title">{isMonthly ? 'PHIẾU THU TIỀN PHÒNG' : 'PHIẾU TẤT TOÁN'}</h1>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <span className="status-badge-liquid" style={{ background: statusBg, color: statusColor, fontSize: '11px' }}>
              {receipt.status.toUpperCase()}
            </span>
            <span className="code" style={{ fontSize: '12px', color: '#64748b' }}>{receiptCode(receipt)}</span>
          </div>
        </div>
      </header>

      <div className="info-bar-v4">
        <div className="item"><span className="label">Phòng</span><span className="value">P{receipt.roomId}</span></div>
        <div className="item"><span className="label">Tháng thu</span><span className="value">{receipt.month}</span></div>
        <div className="item"><span className="label">Ngày lập</span><span className="value">{new Date(receipt.createdAt).toLocaleDateString('vi-VN')}</span></div>
        <div className="item"><span className="label">Hạn trả</span><span className="value">Trong 5 ngày</span></div>
      </div>

      <div className="receipt-body-v4">
        <div className="receipt-items-section">
          <h3>CHI TIẾT KHOẢN THU</h3>
          
          {isMonthly ? (
            <>
              <div className="charge-row-v4">
                <div className="row-main"><span className="name">🏠 Tiền thuê phòng</span><span className="amount">{formatMoney(receipt.rent)}</span></div>
              </div>
              
              <div className="charge-row-v4">
                <div className="row-main"><span className="name">🛠️ Dịch vụ cố định</span><span className="amount">{formatMoney(receipt.fixedServices)}</span></div>
                <p className="details">(Rác, internet, dọn dẹp, thang máy...)</p>
              </div>

              <div className="charge-row-v4">
                <div className="row-main"><span className="name">⚡ Tiền điện</span><span className="amount">{formatMoney(receipt.electricAmount)}</span></div>
                <div className="details">
                  <span>CS cũ: <b>{eOld}</b></span>
                  <span>Mới: <b>{eNew}</b></span>
                  <span>Sử dụng: <b>{receipt.electricUsed}</b> kWh</span>
                  <span>Đơn giá: <b>{formatMoney(room.electricPrice)}</b></span>
                </div>
              </div>

              <div className="charge-row-v4">
                <div className="row-main"><span className="name">💧 Tiền nước</span><span className="amount">{formatMoney(receipt.waterAmount)}</span></div>
                <div className="details">
                  <span>CS cũ: <b>{wOld}</b></span>
                  <span>Mới: <b>{wNew}</b></span>
                  <span>Sử dụng: <b>{receipt.waterUsed}</b> m³</span>
                  <span>Đơn giá: <b>{formatMoney(room.waterPrice)}</b></span>
                </div>
              </div>

              {receipt.other > 0 && (
                <div className="charge-row-v4">
                  <div className="row-main"><span className="name">➕ Phụ phí khác</span><span className="amount">{formatMoney(receipt.other)}</span></div>
                </div>
              )}
            </>
          ) : (
            <div className="charge-row-v4">
              <div className="row-main"><span className="name">📝 Phí tất toán trả phòng</span><span className="amount">{formatMoney(receipt.total)}</span></div>
            </div>
          )}
        </div>

        <div className="payment-column-v4">
          <div className="payment-card-v4">
            <div className="qr-box-v4">
              <img src={buildVietQrUrl(bankInfo, receipt)} alt="QR VietQR" />
              <div style={{ textAlign: 'center' }}>
                <p style={{ fontSize: '11px', fontWeight: '700', margin: 0 }}>QUÉT MÃ THANH TOÁN</p>
                <p style={{ fontSize: '10px', color: '#64748b' }}>Tự động điền số tiền & nội dung</p>
              </div>
            </div>

            <div className="payment-details-v4">
              <div className="row"><span className="label">Ngân hàng</span><span className="val">{bankInfo.bankName}</span></div>
              <div className="row"><span className="label">Số tài khoản</span><span className="val">{bankInfo.accountNo}</span></div>
              <div className="row"><span className="label">Chủ tài khoản</span><span className="val">{bankInfo.accountName}</span></div>
              <div className="row"><span className="label">Nội dung CK</span><span className="val" style={{ color: '#1e40af' }}>{transferContent(receipt)}</span></div>
            </div>

            <div className="total-summary-v4">
              <p className="label">Tổng cộng cần trả</p>
              <p className="total-amount">{formatMoney(receipt.total)}</p>
            </div>
          </div>
        </div>
      </div>

      <footer className="receipt-footer-v4">
        <div style={{ marginBottom: '24px', padding: '16px', background: '#f8fafc', borderRadius: '12px', fontSize: '13px' }}>
          <p><b>Khách thuê:</b> {tenant.name} • <b>SĐT:</b> {tenant.phone}</p>
          {receipt.note && <p style={{ marginTop: '4px' }}><b>Ghi chú:</b> {receipt.note}</p>}
        </div>

        <div className="signatures-v4">
          <div className="signature-box">
            <h4>NGƯỜI THU TIỀN</h4>
            <p className="muted">(Ký và ghi rõ họ tên)</p>
            <div className="signature-space-v4"></div>
            <p><b>DIỆM THỊ BÌNH</b></p>
          </div>
          <div className="signature-box">
            <h4>NGƯỜI NỘP TIỀN</h4>
            <p className="muted">Ngày ..... tháng ..... năm 20...</p>
            <div className="signature-space-v4"></div>
            <p><b>{tenant.name}</b></p>
          </div>
        </div>

        <div className="thank-you-v4">
          Vui lòng thanh toán đúng hạn để đảm bảo quyền lợi dịch vụ. Trân trọng cảm ơn!
        </div>
      </footer>
    </div>
  );
}

function ReceiptItem({ receipt, room, contract, bankInfo, data }) {
  const tenant = getPrimaryTenantByContract(data, receipt.contractId) || { name: 'N/A' };
  const isMonthly = receipt.type === 'monthly';
  return (
    <div className="receipt-page" style={{ padding: '40px', background: 'white', color: 'black', fontFamily: 'serif', position: 'relative', borderBottom: '1px dashed #eee' }}>
      <div style={{ textAlign: 'center', marginBottom: '20px' }}><h1 style={{ fontSize: '20px', fontWeight: 'bold', margin: 0 }}>{isMonthly ? 'PHIẾU THU TIỀN PHÒNG' : 'PHIẾU CHỐT TẤT TOÁN'}</h1><p style={{ fontSize: '14px' }}>{isMonthly ? `Tháng ${receipt.month}` : 'Quyết toán trả phòng'}</p></div>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}><div><p>Phòng: <b>{receipt.roomId}</b></p><p>Khách thuê: <b>{tenant.name}</b></p><p>Ngày lập: {new Date(receipt.createdAt).toLocaleDateString('vi-VN')}</p></div><div style={{ textAlign: 'right' }}><p>Trạng thái: <b>{receipt.status}</b></p></div></div>
      <table className="contract-table" style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }}>
        <thead><tr style={{ background: '#f8fafc' }}><th style={{ border: '1px solid black', padding: '8px' }}>Nội dung</th><th style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>Chỉ số</th><th style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>Thành tiền</th></tr></thead>
        <tbody>
          {isMonthly ? (
            <>
              <tr><td style={{ border: '1px solid black', padding: '8px' }}>Tiền phòng</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.rent)}</td></tr>
              <tr><td style={{ border: '1px solid black', padding: '8px' }}>Dịch vụ cố định</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.fixedServices)}</td></tr>
              <tr>
                <td style={{ border: '1px solid black', padding: '8px' }}>Tiền điện</td>
                <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>
                  CS cũ: {receipt.electricOld} → CS mới: {receipt.electricNew}<br/>
                  (Sử dụng: {receipt.electricUsed} kWh)
                </td>
                <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.electricAmount)}</td>
              </tr>
              <tr>
                <td style={{ border: '1px solid black', padding: '8px' }}>Tiền nước</td>
                <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>
                  CS cũ: {receipt.waterOld} → CS mới: {receipt.waterNew}<br/>
                  (Sử dụng: {receipt.waterUsed} m³)
                </td>
                <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.waterAmount)}</td>
              </tr>
              {receipt.other > 0 && <tr><td style={{ border: '1px solid black', padding: '8px' }}>Phụ phí khác</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.other)}</td></tr>}
            </>
          ) : (
            <tr><td style={{ border: '1px solid black', padding: '8px' }}>Phí chốt tất toán trả phòng</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.total)}</td></tr>
          )}
          <tr style={{ fontWeight: 'bold' }}><td colSpan="2" style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>TỔNG CỘNG</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.total)}</td></tr>
        </tbody>
      </table>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}><div style={{ textAlign: 'center', width: '200px' }}><p><b>QUÉT MÃ THANH TOÁN</b></p><img src={buildVietQrUrl(bankInfo, receipt)} alt="QR" style={{ width: '120px', border: '1px solid #eee', padding: '5px' }} /><p style={{ fontSize: '10px' }}>{bankInfo.bankName} - {bankInfo.accountNo}</p></div><div style={{ textAlign: 'center', width: '200px' }}><p><b>CHỦ NHÀ KÝ TÊN</b></p><div style={{ height: '80px' }}></div><p><b>DIỆM THỊ BÌNH</b></p></div></div>
      <p style={{ fontStyle: 'italic', fontSize: '12px', marginTop: '20px', textAlign: 'center' }}>Quý khách vui lòng thanh toán trong vòng 5 ngày kể từ ngày nhận phiếu. Trân trọng!</p>
    </div>
  );
}

function PaymentModal({ receipt, onClose, onSave }) {
  const [paidAmount, setPaidAmount] = useState(receipt?.paidAmount || receipt?.total || 0);
  const [paidDate, setPaidDate] = useState(new Date().toISOString().split('T')[0]);
  
  if (!receipt) return null;

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '400px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Xác nhận thanh toán</h2>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack">
          <div style={{ background: '#f8fafc', padding: '16px', borderRadius: '12px' }}>
            <p>Phòng: <b>P{receipt.roomId}</b></p>
            <p>Tổng tiền: <b>{formatMoney(receipt.total)}</b></p>
            <p className="muted small">Tháng {receipt.month}</p>
          </div>
          <label>
            Số tiền khách trả 
            <input type="number" value={paidAmount} onChange={e => setPaidAmount(e.target.value)} />
          </label>
          <label>
            Ngày thanh toán 
            <input type="date" value={paidDate} onChange={e => setPaidDate(e.target.value)} />
          </label>
          <div style={{ marginTop: '10px' }}>
            {Number(paidAmount) >= receipt.total ? 
              <span className="status-badge-liquid active">Thanh toán đủ</span> : 
              Number(paidAmount) > 0 ? 
                <span className="status-badge-liquid notice">Thanh toán một phần</span> : 
                <span className="status-badge-liquid debt">Chưa trả tiền</span>
            }
          </div>
          <button className="primary-btn wide" onClick={() => { 
            let status = 'Chưa thanh toán'; 
            if (Number(paidAmount) >= receipt.total) status = 'Đã thanh toán'; 
            else if (Number(paidAmount) > 0) status = 'Nợ một phần'; 
            onSave({ ...receipt, paidAmount: Number(paidAmount), paidDate, status }); 
          }}>Xác nhận</button>
        </div>
      </div>
    </div>
  );
}

export default App;
