import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import './styles.css';

const STORAGE_KEY = 'room_manager_qr_app_v1';
const BANK_KEY = 'room_manager_bank_v1';
const PIN_KEY = 'room_manager_pin_v1';

const DEFAULT_BANK = {
  bankName: 'Ngân hàng TMCP Đầu tư và Phát triển Việt Nam',
  bankCode: 'BIDV',
  accountNo: '8847214661',
  accountName: 'HỘ KINH DOANH DIỆM THỊ BÌNH',
};

function getCurrentMonthLabel(date = new Date()) {
  return `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

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
  contractRenewals: [],
  roomTransfers: [],
  suppliers: [
    { id: 'sup_evn', name: 'Điện lực', group: 'Điện', defaultCategory: 'cat_elec' },
    { id: 'sup_water', name: 'Cấp nước', group: 'Nước', defaultCategory: 'cat_water' },
    { id: 'sup_vnpt', name: 'VNPT', group: 'Internet', defaultCategory: 'cat_internet' },
    { id: 'sup_fpt', name: 'FPT Telecom', group: 'Internet', defaultCategory: 'cat_internet' },
    { id: 'sup_viettel', name: 'Viettel Telecom', group: 'Internet', defaultCategory: 'cat_internet' },
    { id: 'sup_sctv', name: 'SCTV', group: 'Truyền hình / mạng', defaultCategory: 'cat_internet' },
    { id: 'sup_trash_private', name: 'Thu gom rác dân lập', group: 'Rác / môi trường', defaultCategory: 'cat_trash' },
    { id: 'sup_trash_city', name: 'Công ty môi trường đô thị', group: 'Rác / môi trường', defaultCategory: 'cat_trash' },
    { id: 'sup_clean_staff', name: 'Nhân viên vệ sinh', group: 'Vệ sinh', defaultCategory: 'cat_cleaning' },
    { id: 'sup_clean_service', name: 'Dịch vụ vệ sinh', group: 'Vệ sinh', defaultCategory: 'cat_cleaning' },
    { id: 'sup_security', name: 'Bảo vệ tòa nhà', group: 'Bảo vệ', defaultCategory: 'cat_guard' },
    { id: 'sup_elevator', name: 'Dịch vụ bảo trì thang máy', group: 'Thang máy', defaultCategory: 'cat_maintenance' },
    { id: 'sup_camera', name: 'Dịch vụ camera an ninh', group: 'Camera', defaultCategory: 'cat_camera' },
    { id: 'sup_pccc', name: 'Dịch vụ PCCC', group: 'PCCC', defaultCategory: 'cat_pccc' },
    { id: 'sup_elec_repair', name: 'Thợ điện', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_water_repair', name: 'Thợ nước', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_build_repair', name: 'Thợ hồ / xây dựng', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_paint_repair', name: 'Thợ sơn', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_lock_repair', name: 'Thợ khóa', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_glass_repair', name: 'Thợ nhôm kính', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_door_repair', name: 'Thợ cửa cuốn', group: 'Sửa chữa', defaultCategory: 'cat_repair' },
    { id: 'sup_ac_repair', name: 'Thợ sửa máy lạnh', group: 'Điện lạnh', defaultCategory: 'cat_repair' },
    { id: 'sup_ac_clean', name: 'Dịch vụ vệ sinh máy lạnh', group: 'Điện lạnh', defaultCategory: 'cat_cleaning' },
    { id: 'sup_furniture', name: 'Cửa hàng nội thất', group: 'Nội thất', defaultCategory: 'cat_material' },
    { id: 'sup_hardware', name: 'Cửa hàng điện nước', group: 'Vật tư', defaultCategory: 'cat_material' },
    { id: 'sup_materials', name: 'Cửa hàng vật liệu xây dựng', group: 'Vật tư', defaultCategory: 'cat_material' },
    { id: 'sup_metal', name: 'Cửa hàng kim khí', group: 'Vật tư', defaultCategory: 'cat_material' },
    { id: 'sup_electric_store', name: 'Cửa hàng thiết bị điện', group: 'Thiết bị', defaultCategory: 'cat_material' },
    { id: 'sup_water_store', name: 'Cửa hàng thiết bị nước', group: 'Thiết bị', defaultCategory: 'cat_material' },
    { id: 'sup_appliance', name: 'Cửa hàng điện máy', group: 'Gia dụng', defaultCategory: 'cat_material' },
    { id: 'sup_laundry', name: 'Dịch vụ giặt sấy', group: 'Giặt sấy', defaultCategory: 'cat_other' },
    { id: 'sup_pest', name: 'Dịch vụ diệt côn trùng', group: 'Diệt côn trùng', defaultCategory: 'cat_other' },
    { id: 'sup_garden', name: 'Dịch vụ chăm sóc cây xanh', group: 'Cây xanh', defaultCategory: 'cat_other' },
    { id: 'sup_legal', name: 'Dịch vụ pháp lý', group: 'Pháp lý', defaultCategory: 'cat_service' },
    { id: 'sup_acc', name: 'Dịch vụ kế toán', group: 'Kế toán', defaultCategory: 'cat_service' },
    { id: 'sup_bank_fee', name: 'Ngân hàng / phí chuyển khoản', group: 'Ngân hàng', defaultCategory: 'cat_bank' },
    { id: 'sup_tax', name: 'Cơ quan thuế', group: 'Thuế', defaultCategory: 'cat_tax' },
    { id: 'sup_landlord', name: 'Chủ nhà / bên cho thuê mặt bằng', group: 'Chủ nhà', defaultCategory: 'cat_rent' },
    { id: 'sup_mgmt', name: 'Ban quản lý tòa nhà', group: 'Quản lý', defaultCategory: 'cat_management' },
    { id: 'sup_other', name: 'Nhà cung cấp khác', group: 'Khác', defaultCategory: 'cat_other' },
  ],
  expenseCategories: [
    { id: 'cat_elec', name: 'Tiền điện' },
    { id: 'cat_water', name: 'Tiền nước' },
    { id: 'cat_internet', name: 'Internet / Truyền hình' },
    { id: 'cat_cleaning', name: 'Vệ sinh' },
    { id: 'cat_repair', name: 'Sửa chữa' },
    { id: 'cat_material', name: 'Vật tư / Nội thất' },
    { id: 'cat_rent', name: 'Tiền thuê mặt bằng' },
    { id: 'cat_trash', name: 'Phí rác' },
    { id: 'cat_guard', name: 'Bảo vệ' },
    { id: 'cat_maintenance', name: 'Bảo trì' },
    { id: 'cat_camera', name: 'Camera' },
    { id: 'cat_pccc', name: 'PCCC' },
    { id: 'cat_bank', name: 'Phí ngân hàng' },
    { id: 'cat_tax', name: 'Thuế / phí' },
    { id: 'cat_management', name: 'Phí quản lý' },
    { id: 'cat_service', name: 'Phí dịch vụ' },
    { id: 'cat_other', name: 'Chi phí khác' },
  ],
  expensePayments: [],
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

function getDataVersion(data) {
  if (!data) return 0;
  const dateFields = ['createdAt', 'updatedAt', 'savedAt', 'paidDate', 'signedDate', 'startDate', 'joinedDate', 'endedAt', 'renewedAt'];
  const collections = ['rooms', 'tenants', 'memberships', 'contracts', 'receipts', 'moveOutReports', 'contractRenewals', 'roomTransfers', 'suppliers', 'expenseCategories', 'expensePayments'];
  return collections.reduce((max, key) => {
    return Math.max(max, ...((data[key] || []).map(item => {
      return Math.max(0, ...dateFields.map(field => {
        const time = Date.parse(item?.[field] || '');
        return Number.isNaN(time) ? 0 : time;
      }));
    })));
  }, 0);
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
  if (value instanceof Date) return value;
  const text = String(value);
  if (text.includes('-')) return new Date(`${text}T00:00:00`);
  if (text.includes('/')) {
    const parts = text.split('/');
    if (parts.length === 3) return new Date(`${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}T00:00:00`);
  }
  return null;
}

const INVALID_CONTRACT_YEARS = [1900, 1970];

function isInvalidContractDate(value) {
  const date = parseDateFlexible(value);
  if (!date || Number.isNaN(date.getTime())) return true;
  return INVALID_CONTRACT_YEARS.includes(date.getFullYear());
}

function formatContractDate(value, fallback = 'Chưa cập nhật') {
  if (isInvalidContractDate(value)) return fallback;
  const date = parseDateFlexible(value);
  if (!date) return fallback;
  return date.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function isValidBusinessDate(value) {
  return !isInvalidContractDate(value);
}

function formatBusinessDate(value, fallback = 'Chưa cập nhật') {
  return formatContractDate(value, fallback);
}

function calculateRentalDuration(startDate, endDate) {
  if (isInvalidContractDate(startDate) || isInvalidContractDate(endDate)) return '—';
  const start = parseDateFlexible(startDate);
  const end = parseDateFlexible(endDate);
  if (!start || !end || end <= start) return '—';
  const months = diffMonths(startDate, endDate);
  return months > 0 ? `${months} tháng` : '—';
}

function isValidContractRange(startDate, endDate) {
  if (isInvalidContractDate(startDate) || isInvalidContractDate(endDate)) return false;
  return parseDateFlexible(endDate) > parseDateFlexible(startDate);
}

function parseVietnameseDateValue(value) {
  const match = String(value || '').match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if (!match) return '';
  const [, day, month, year] = match;
  return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
}

function normalizeSearchText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[*/_`]/g, '')
    .replace(/\s*\/\s*/g, '/')
    .toLowerCase();
}

function cleanOcrValue(value) {
  return String(value || '')
    .replace(/^[\s:：\-–—•]+/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripBilingualIdLabelTail(value) {
  return cleanOcrValue(value)
    .replace(/^[*/_`\s]*/, '')
    .replace(/^\/?\s*(full name|date of birth|date,\s*month,\s*year|date month year|sex|nationality|place of origin|place of residence|date of expiry|date of issue|personal identification|issued by|no\.?)\s*:?\s*/i, '')
    .replace(/^[*/_`\s:]*/, '')
    .trim();
}

function readIdCardField(lines, labels, stopLabels = []) {
  const normalizedLabels = labels.map(normalizeSearchText);
  const normalizedStops = [...labels, ...stopLabels].map(normalizeSearchText);
  for (let i = 0; i < lines.length; i += 1) {
    const rawLine = lines[i];
    const normalizedLine = normalizeSearchText(rawLine);
    const matchedLabel = normalizedLabels.find(label => normalizedLine.includes(label));
    if (!matchedLabel) continue;

    const labelIndex = normalizedLine.indexOf(matchedLabel);
    const markerlessLine = rawLine.replace(/[*/_`]/g, '');
    const valueAfterColon = markerlessLine.match(/[:：]\s*(.+)$/)?.[1] || '';
    const afterLabel = stripBilingualIdLabelTail(valueAfterColon || markerlessLine.slice(labelIndex + matchedLabel.length));
    if (afterLabel && !/^\/|^no\.?$/i.test(afterLabel)) return afterLabel;

    const collected = [];
    for (let j = i + 1; j < lines.length; j += 1) {
      const next = lines[j];
      const normalizedNext = normalizeSearchText(next);
      if (normalizedStops.some(stop => normalizedNext.includes(stop))) break;
      if (/^(mat truoc|mat sau|front|back|mrz|idvnm)/i.test(normalizedNext)) break;
      if (next.trim()) collected.push(next.trim());
      if (collected.length >= 3) break;
    }
    return cleanOcrValue(collected.join(', '));
  }
  return '';
}

function parseMrzName(text) {
  const mrzLines = String(text || '').split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => /^[A-Z0-9<]{15,}$/.test(line));
  const nameLine = [...mrzLines].reverse().find(line => line.includes('<<') && /[A-Z]/.test(line));
  if (!nameLine) return '';
  return nameLine
    .replace(/<+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/\b\w/g, char => char.toUpperCase());
}

function parseMrzBirthday(text) {
  const mrzLine = String(text || '').split(/\r?\n/).map(line => line.trim()).find(line => /^\d{6}\d?[MF]/i.test(line));
  if (!mrzLine) return '';
  const match = mrzLine.match(/^(\d{2})(\d{2})(\d{2})/);
  if (!match) return '';
  const [, yy, mm, dd] = match;
  const yearPrefix = Number(yy) > 30 ? '19' : '20';
  return `${yearPrefix}${yy}-${mm}-${dd}`;
}

function parseVietnameseIdCard(text) {
  const rawText = String(text || '');
  const lines = rawText.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  const stopLabels = [
    'Họ và tên', 'Full name', 'Ngày sinh', 'Date of birth', 'Giới tính', 'Sex',
    'Quốc tịch', 'Nationality', 'Quê quán', 'Place of origin', 'Nơi thường trú',
    'Place of residence', 'Có giá trị đến', 'Date of expiry', 'Đặc điểm nhận dạng',
    'Personal identification', 'Ngày cấp', 'Ngày, tháng, năm', 'Date, month, year', 'Date of issue', 'Người ký', 'Dòng mã', 'MRZ'
  ];
  const labeledId = readIdCardField(lines, ['Số CCCD', 'Số/No.', 'Số/No', 'Số', 'No.', 'No'], stopLabels);
  const standaloneId = rawText.match(/\b\d{12}\b/)?.[0] || '';
  const cccd = (labeledId.match(/\d{12}/)?.[0] || standaloneId).trim();
  const name = readIdCardField(lines, ['Họ và tên/Full name', 'Họ và tên', 'Full name'], stopLabels) || parseMrzName(rawText);
  const birthdayText = readIdCardField(lines, ['Ngày sinh/Date of birth', 'Ngày sinh', 'Date of birth'], stopLabels);
  const birthday = parseVietnameseDateValue(birthdayText) || parseMrzBirthday(rawText);
  const address = readIdCardField(lines, ['Nơi thường trú/Place of residence', 'Nơi thường trú', 'Place of residence'], stopLabels);
  const issueDate = parseVietnameseDateValue(readIdCardField(lines, ['Ngày, tháng, năm/Date, month, year', 'Ngày, tháng, năm', 'Date, month, year', 'Ngày cấp', 'Date of issue'], stopLabels));
  const issuePlace = readIdCardField(lines, ['Nơi cấp', 'Cơ quan cấp', 'Issued by'], stopLabels);
  const extraFields = [
    ['Giới tính', readIdCardField(lines, ['Giới tính/Sex', 'Giới tính', 'Sex'], stopLabels)],
    ['Quốc tịch', readIdCardField(lines, ['Quốc tịch/Nationality', 'Quốc tịch', 'Nationality'], stopLabels)],
    ['Quê quán', readIdCardField(lines, ['Quê quán/Place of origin', 'Quê quán', 'Place of origin'], stopLabels)],
    ['Có giá trị đến', readIdCardField(lines, ['Có giá trị đến/Date of expiry', 'Có giá trị đến', 'Date of expiry'], stopLabels)],
    ['Đặc điểm nhận dạng', readIdCardField(lines, ['Đặc điểm nhận dạng/Personal identification', 'Đặc điểm nhận dạng', 'Personal identification'], stopLabels)],
    ['Người ký', readIdCardField(lines, ['Người ký', 'Cục trưởng Cục Cảnh sát quản lý hành chính về trật tự xã hội'], stopLabels)]
  ].filter(([, value]) => value);
  const normalizedRawText = normalizeSearchText(rawText);
  const mrzIndex = normalizedRawText.indexOf('mrz') >= 0 ? normalizedRawText.indexOf('mrz') : normalizedRawText.indexOf('dong ma');
  const mrzText = mrzIndex >= 0
    ? rawText.slice(mrzIndex).split(/\r?\n/).slice(1).map(line => line.trim()).filter(Boolean).join('\n')
    : lines.filter(line => /^[A-Z0-9<]{15,}$/.test(line)).join('\n');
  const note = [
    ...extraFields.map(([label, value]) => `${label}: ${value}`),
    mrzText ? `MRZ:\n${mrzText}` : ''
  ].filter(Boolean).join('\n');

  return { cccd, name: cleanOcrValue(name), birthday, address, issueDate, issuePlace, note };
}

function formatDateForInput(value) {
  return isInvalidContractDate(value) ? '' : String(value).slice(0, 10);
}

function formatDisplayDate(value, fallback = '—') {
  if (!value) return fallback;
  const date = parseDateFlexible(value);
  if (!date || Number.isNaN(date.getTime())) return fallback;
  return date.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });
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

function addDaysToDate(value, days) {
  if (!isValidDateString(value)) return '';
  const date = new Date(value + 'T00:00:00');
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function getDaysUntil(dateStr) {
  const date = parseDateFlexible(dateStr);
  if (!date) return null;
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return Math.ceil((date.getTime() - now.getTime()) / 86400000);
}

function getLatestActiveMembershipForRoom(data, roomId) {
  return (data.memberships || [])
    .filter(m => m.roomId === roomId && m.status === 'active')
    .sort((a, b) => {
      const roleScore = (b.role === 'primary' ? 1 : 0) - (a.role === 'primary' ? 1 : 0);
      if (roleScore !== 0) return roleScore;
      const joinedDiff = (parseDateFlexible(b.joinedDate)?.getTime() || 0) - (parseDateFlexible(a.joinedDate)?.getTime() || 0);
      if (joinedDiff !== 0) return joinedDiff;
      return (parseDateFlexible(b.createdAt)?.getTime() || 0) - (parseDateFlexible(a.createdAt)?.getTime() || 0);
    })[0] || null;
}

function getCurrentContractForRoom(data, roomId) {
  const activeMembership = getLatestActiveMembershipForRoom(data, roomId);
  if (activeMembership?.contractId) {
    const membershipContract = (data.contracts || []).find(c => c.id === activeMembership.contractId);
    if (membershipContract) return membershipContract;
  }
  return (data.contracts || [])
    .filter(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice' || c.status === 'moving_out'))
    .sort((a, b) => {
      const startDiff = (parseDateFlexible(b.startDate)?.getTime() || 0) - (parseDateFlexible(a.startDate)?.getTime() || 0);
      if (startDiff !== 0) return startDiff;
      return (parseDateFlexible(b.createdAt)?.getTime() || 0) - (parseDateFlexible(a.createdAt)?.getTime() || 0);
    })[0] || null;
}

function getRoomStatusInfo(data, roomId) {
  const room = (data.rooms || []).find(r => r.id === roomId);
  const ownerOccupied = isOwnerOccupiedRoom(room);
  const contract = getCurrentContractForRoom(data, roomId);
  const hasActiveMembers = (data.memberships || []).some(m => m.roomId === roomId && m.status === 'active');
  if (!contract && hasActiveMembers) return { label: 'Đang ở', color: 'green', contract: null, ownerOccupied };
  if (!contract) return { label: 'Trống', color: 'gray', contract: null, ownerOccupied };
  if (contract.status === 'notice') return { label: 'Báo chuyển', color: 'orange', contract };
  if (contract.status === 'moving_out') return { label: 'Đang tất toán', color: 'blue', contract };
  return { label: 'Đang ở', color: 'green', contract, ownerOccupied };
}

function isOwnerOccupiedRoom(room) {
  if (!room) return false;
  return room.id === '202' || String(room.note || '').toLowerCase().includes('chủ nhà');
}

function getDashboardStats(data, currentMonth) {
  const totalRooms = data.rooms.length;
  const activeContracts = (data.contracts || []).filter(c => c.status === 'active' || c.status === 'notice');
  const occupiedRooms = (data.rooms || []).filter(r => getRoomStatusInfo(data, r.id).label !== 'Trống').length;
  const vacantRooms = totalRooms - occupiedRooms;
  
  const expiringContracts = activeContracts.filter(c => {
    const days = getDaysUntil(c.endDate);
    return days !== null && days <= 30;
  });

  const currentTenants = (data.memberships || []).filter(m => m.status === 'active').length;
  
  const monthReceipts = (data.receipts || []).filter(r => r.month === currentMonth && r.type === 'monthly');
  const unpaidReceipts = monthReceipts.filter(r => getReceiptPaymentState(r).debt > 0);
  
  const totalDebt = (data.receipts || []).reduce((sum, r) => {
    return sum + getReceiptPaymentState(r).debt;
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

function getTenantForReceipt(data, receipt) {
  if (!receipt) return null;
  const byContract = getPrimaryTenantByContract(data, receipt.contractId);
  if (byContract) return byContract;

  const activeRoomMember = (data.memberships || []).find(m =>
    m.roomId === receipt.roomId &&
    m.status === 'active' &&
    m.role === 'primary'
  ) || (data.memberships || []).find(m =>
    m.roomId === receipt.roomId &&
    m.status === 'active'
  );
  if (activeRoomMember) {
    const tenant = (data.tenants || []).find(t => t.id === activeRoomMember.tenantId);
    if (tenant) return tenant;
  }

  const contractMember = (data.memberships || []).find(m => m.contractId === receipt.contractId);
  return contractMember ? (data.tenants || []).find(t => t.id === contractMember.tenantId) || null : null;
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
  if (value === null || value === undefined || value === '' || value === '?') return '—';
  const num = typeof value === 'number' ? value : parseFloat(String(value).replace(/,/g, '.'));
  if (isNaN(num)) return '—';
  return new Intl.NumberFormat('vi-VN').format(num);
}

function DecimalInput({ value, onChange, style, className }) {
  const toDisplayValue = (val) => (val === null || val === undefined ? '' : String(val));
  const [displayValue, setDisplayValue] = useState(toDisplayValue(value));

  useEffect(() => {
    // Update display value when prop value changes from outside (e.g. from recalculateReceipt)
    // but only if it's not the same number to avoid cursor jumping
    const currentNum = parseLocaleNumber(displayValue);
    if (value !== currentNum) {
      setDisplayValue(toDisplayValue(value));
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
  const paymentState = getReceiptPaymentState(receipt);
  const amount = paymentState.debt > 0 ? paymentState.debt : Number(receipt.total || 0);
  const params = new URLSearchParams({
    amount: String(Math.round(amount)),
    addInfo: transferContent(receipt),
    accountName: bankInfo.accountName || '',
  });
  return `https://img.vietqr.io/image/${bankInfo.bankCode}-${bankInfo.accountNo}-compact2.png?${params.toString()}`;
}

function buildPaymentQrUrl(bankInfo, amount, addInfo) {
  if (!bankInfo?.bankCode || !bankInfo?.accountNo || !amount) return '';
  const params = new URLSearchParams({
    amount: String(Math.round(Number(amount || 0))),
    addInfo,
    accountName: bankInfo.accountName || '',
  });
  return `https://img.vietqr.io/image/${bankInfo.bankCode}-${bankInfo.accountNo}-compact2.png?${params.toString()}`;
}

function getContractOccupantCount(data, contract) {
  if (!contract) return 0;
  const byContract = (data.memberships || []).filter(m => m.contractId === contract.id && m.status === 'active').length;
  const byRoom = (data.memberships || []).filter(m => m.roomId === contract.roomId && m.status === 'active').length;
  return Math.max(byContract, byRoom);
}

function fixedServiceTotal(room, occupantCount = null) {
  if (!room) return 0;
  if (occupantCount !== null && occupantCount !== undefined) {
    if (Number(occupantCount) <= 0) return 0;
    return Number(occupantCount) <= 1 ? 200000 : 400000;
  }
  return Number(room.cleaning || 0) + Number(room.elevator || 0) + Number(room.laundry || 0) + Number(room.internet || 0);
}

function getExpensePaidAmount(expense) {
  return Number(expense?.paidAmount ?? expense?.amount ?? 0);
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

function receiptTypeLabel(type, withIcon = false) {
  if (type === 'monthly') return withIcon ? '📅 Phiếu tháng' : 'Phiếu tháng';
  if (type === 'renewal_adjustment') return withIcon ? '🔁 Điều chỉnh gia hạn' : 'Điều chỉnh gia hạn';
  return withIcon ? '🚪 Chốt trả phòng' : 'Chốt trả phòng';
}

function parseMonthValue(month) {
  if (!month) return 0;
  const [m, y] = month.split('/').map(Number);
  return y * 12 + m;
}

function monthFromDate(value) {
  const date = parseDateFlexible(value);
  if (!date) return '';
  return `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

function billingMonthParts(month) {
  const [m, y] = String(month || '').split('/').map(Number);
  return { month: m || 0, year: y || 0 };
}

function daysInBillingMonth(month) {
  const parts = billingMonthParts(month);
  if (!parts.month || !parts.year) return 30;
  return new Date(parts.year, parts.month, 0).getDate();
}

function billingMonthStart(month) {
  const parts = billingMonthParts(month);
  if (!parts.month || !parts.year) return null;
  return new Date(parts.year, parts.month - 1, 1);
}

function billingMonthEnd(month) {
  const parts = billingMonthParts(month);
  if (!parts.month || !parts.year) return null;
  return new Date(parts.year, parts.month, 0);
}

function formatDateInputValue(date) {
  if (!date || Number.isNaN(date.getTime())) return '';
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function addDays(value, days) {
  const date = parseDateFlexible(value);
  if (!date) return null;
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function clampDateToBillingMonth(value, month) {
  const date = parseDateFlexible(value);
  const start = billingMonthStart(month);
  const end = billingMonthEnd(month);
  if (!date || !start || !end) return null;
  if (date < start) return start;
  if (date > end) return end;
  return date;
}

function inclusiveDaysBetween(startValue, endValue) {
  const start = parseDateFlexible(startValue);
  const end = parseDateFlexible(endValue);
  if (!start || !end || end < start) return 0;
  const startUtc = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate());
  const endUtc = Date.UTC(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.floor((endUtc - startUtc) / 86400000) + 1;
}

function addMonthsToBillingMonth(month, delta) {
  if (!month) return '';
  const [m, y] = month.split('/').map(Number);
  if (!m || !y) return '';
  const date = new Date(y, m - 1 + delta, 1);
  return `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

function isSameBillingMonth(dateValue, month) {
  return monthFromDate(dateValue) === month;
}

function isTransferOldRoomBillingMonth(dateValue, month) {
  const transferMonth = monthFromDate(dateValue);
  return transferMonth === month;
}

function getTransferRoomLabel(receipt) {
  const transfer = receipt?.transferOldRoomUtility;
  if (!transfer) return `P${receipt?.roomId || ''}`;
  return `P${transfer.oldRoomId} → P${receipt.roomId}`;
}

function getPreviousReceiptByRoom(receipts, roomId, currentMonth, options = {}) {
  const currentValue = parseMonthValue(currentMonth);
  return (receipts || [])
    .filter(r =>
      r.roomId === roomId &&
      r.type === 'monthly' &&
      r.id !== options.excludeReceiptId &&
      (
        parseMonthValue(r.month) < currentValue ||
        (
          options.includeSameMonth &&
          parseMonthValue(r.month) === currentValue &&
          r.contractId !== options.excludeSameMonthContractId
        )
      )
    )
    .sort((a, b) => {
      const monthDiff = parseMonthValue(b.month) - parseMonthValue(a.month);
      if (monthDiff !== 0) return monthDiff;
      return (parseDateFlexible(b.createdAt)?.getTime() || 0) - (parseDateFlexible(a.createdAt)?.getTime() || 0);
    })[0] || null;
}

function getLatestMonthlyReceiptForRoom(data, roomId, contractId = '') {
  const receipts = (data.receipts || []).filter(r => r.roomId === roomId && r.type === 'monthly');
  const sortLatest = (rows) => [...rows].sort((a, b) => {
    const monthDiff = parseMonthValue(b.month) - parseMonthValue(a.month);
    if (monthDiff !== 0) return monthDiff;
    return (parseDateFlexible(b.createdAt)?.getTime() || 0) - (parseDateFlexible(a.createdAt)?.getTime() || 0);
  })[0] || null;
  if (contractId) {
    const byContract = sortLatest(receipts.filter(r => r.contractId === contractId));
    if (byContract) return byContract;
  }
  return sortLatest(receipts);
}

function getTransferBillingContext(data, contract, month) {
  const occupantCount = getContractOccupantCount(data, contract);
  const transfer = (data.roomTransfers || []).find(t =>
    isTransferOldRoomBillingMonth(t.transferDate, month) &&
    (t.oldContractId === contract?.id || t.newContractId === contract?.id)
  );
  if (!transfer) return { mode: 'normal', transfer: null, occupantCount };
  if (transfer.oldContractId === contract?.id) return { mode: 'transfer_old_room_skip', transfer, occupantCount };
  if (transfer.newContractId === contract?.id) return { mode: 'transfer_new_room', transfer, occupantCount, oldRoomUtility: buildTransferOldRoomUtility(data, transfer, month) };
  return { mode: 'normal', transfer: null, occupantCount };
}

function getNewContractPartialContext(contract, month) {
  if (!contract?.startDate || monthFromDate(contract.startDate) !== month) return null;
  const start = clampDateToBillingMonth(contract.startDate, month);
  const end = billingMonthEnd(month);
  const chargeDays = inclusiveDaysBetween(start, end);
  const monthDays = daysInBillingMonth(month);
  if (!chargeDays || chargeDays >= monthDays) return null;
  return {
    mode: 'new_contract_partial_month',
    startDate: formatDateInputValue(start),
    endDate: formatDateInputValue(end),
    chargeDays,
    monthDays
  };
}

function getMonthlyBillingContext(data, contract, month) {
  const transferContext = getTransferBillingContext(data, contract, month);
  if (transferContext.mode !== 'normal') return transferContext;
  const partialContext = getNewContractPartialContext(contract, month);
  if (partialContext) {
    return {
      ...partialContext,
      transfer: null,
      occupantCount: transferContext.occupantCount
    };
  }
  return transferContext;
}

function findTransferTargetForOldRoom(data, oldRoomId, month) {
  const transfer = (data.roomTransfers || []).find(t =>
    t.oldRoomId === oldRoomId &&
    isTransferOldRoomBillingMonth(t.transferDate, month)
  );
  if (!transfer) return null;
  const contract = (data.contracts || []).find(c => c.id === transfer.newContractId);
  const room = (data.rooms || []).find(r => r.id === transfer.newRoomId);
  return contract && room ? { transfer, contract, room } : null;
}

function getBillableContractsForMonth(data, month) {
  const activeMembershipContractIds = new Set((data.memberships || []).filter(m => m.status === 'active').map(m => m.contractId));
  const roomsWithActiveMembership = new Set((data.memberships || []).filter(m => m.status === 'active').map(m => m.roomId));
  const contracts = (data.contracts || []).filter(c => {
    if (!(c.status === 'active' || c.status === 'notice')) return false;
    if (roomsWithActiveMembership.has(c.roomId)) return activeMembershipContractIds.has(c.id);
    return true;
  });
  const byId = new Map(contracts.map(c => [c.id, c]));
  return [...byId.values()];
}

function buildTransferOldRoomUtility(data, transfer, month) {
  const oldRoom = (data.rooms || []).find(r => r.id === transfer.oldRoomId);
  if (!oldRoom) return null;
  const oldContract = (data.contracts || []).find(c => c.id === transfer.oldContractId);
  const oldBillingMonth = transfer.oldBillingMonth || monthFromDate(transfer.transferDate) || month;
  const oldRoomReceipt = (data.receipts || []).find(r =>
    r.roomId === transfer.oldRoomId &&
    r.contractId === transfer.oldContractId &&
    r.month === oldBillingMonth &&
    r.type === 'monthly'
  );
  if (oldRoomReceipt && Number(oldRoomReceipt.paidAmount || 0) >= Number(oldRoomReceipt.total || 0) && Number(oldRoomReceipt.total || 0) > 0) {
    return null;
  }

  const transferDate = parseDateFlexible(transfer.transferDate);
  const defaultOldStayTo = transferDate ? addDays(transferDate, -1) : null;
  const oldStayFromRaw = transfer.oldRoomStayFrom || oldContract?.startDate || billingMonthStart(oldBillingMonth);
  const oldStayToRaw = transfer.oldRoomStayTo || defaultOldStayTo || transfer.transferDate || billingMonthEnd(oldBillingMonth);
  const oldStayFrom = clampDateToBillingMonth(oldStayFromRaw, oldBillingMonth) || billingMonthStart(oldBillingMonth);
  const oldStayTo = clampDateToBillingMonth(oldStayToRaw, oldBillingMonth) || billingMonthEnd(oldBillingMonth);
  const oldRoomDays = inclusiveDaysBetween(oldStayFrom, oldStayTo);
  const oldMonthDays = daysInBillingMonth(oldBillingMonth);
  const oldMonthlyRent = Number(transfer.oldRoomMonthlyRent ?? transfer.oldRent ?? oldContract?.rent ?? oldRoom.rent ?? 0);
  const oldRentAmount = Number(transfer.oldRoomRentAmount ?? Math.round((oldMonthlyRent / oldMonthDays) * oldRoomDays));
  const oldOccupantCount = Number(transfer.occupantCount || getContractOccupantCount(data, oldContract) || 1);
  const oldMonthlyServiceFee = Number(transfer.oldRoomServiceFee ?? fixedServiceTotal(oldRoom, oldOccupantCount));
  const serviceMode = transfer.oldRoomServiceMode || 'included_in_transfer_receipt';
  const oldServiceAmount = Number(transfer.oldRoomServiceAmount ?? (
    serviceMode === 'monthly'
      ? oldMonthlyServiceFee
      : serviceMode === 'daily'
        ? Math.round((oldMonthlyServiceFee / oldMonthDays) * oldRoomDays)
        : 0
  ));
  const currentOldRoomReceipt = (data.receipts || []).find(r =>
    r.roomId === transfer.oldRoomId &&
    r.month === oldBillingMonth &&
    r.type === 'monthly'
  );
  if (currentOldRoomReceipt) {
    const electricOld = getElectricOld(currentOldRoomReceipt);
    const electricNew = getElectricNew(currentOldRoomReceipt);
    const waterOld = getWaterOld(currentOldRoomReceipt);
    const waterNew = getWaterNew(currentOldRoomReceipt);
    const electricUsed = Math.round(Math.max(0, electricNew - electricOld) * 100) / 100;
    const waterUsed = Math.round(Math.max(0, waterNew - waterOld) * 100) / 100;
    const electricAmount = electricUsed * Number(oldRoom.electricPrice || 0);
    const waterAmount = waterUsed * Number(oldRoom.waterPrice || 0);
    return {
      oldRoomId: transfer.oldRoomId,
      transferDate: transfer.transferDate,
      electricOld,
      electricNew,
      electricUsed,
      electricAmount,
      waterOld,
      waterNew,
      waterUsed,
      waterAmount,
      oldBillingMonth,
      oldStayFrom: formatDateInputValue(oldStayFrom),
      oldStayTo: formatDateInputValue(oldStayTo),
      oldRoomDays,
      oldMonthDays,
      oldMonthlyRent,
      oldRentAmount,
      oldMonthlyServiceFee,
      oldServiceAmount,
      oldServiceMode: serviceMode,
      total: oldRentAmount + oldServiceAmount + electricAmount + waterAmount
    };
  }
  const previousReceipt = getPreviousReceiptByRoom(data.receipts, transfer.oldRoomId, oldBillingMonth);
  const electricOld = previousReceipt
    ? Number(previousReceipt.electricNew ?? previousReceipt.electricEnd ?? 0)
    : Number(oldRoom.electricOld ?? oldRoom.electricStart ?? oldRoom.initialElectric ?? 0);
  const electricNew = Number(oldRoom.electricNew ?? oldRoom.electricEnd ?? electricOld);
  const waterOld = previousReceipt
    ? Number(previousReceipt.waterNew ?? previousReceipt.waterEnd ?? 0)
    : Number(oldRoom.waterOld ?? oldRoom.waterStart ?? oldRoom.initialWater ?? 0);
  const waterNew = Number(oldRoom.waterNew ?? oldRoom.waterEnd ?? waterOld);
  const electricUsed = Math.round(Math.max(0, electricNew - electricOld) * 100) / 100;
  const waterUsed = Math.round(Math.max(0, waterNew - waterOld) * 100) / 100;
  const electricAmount = electricUsed * Number(oldRoom.electricPrice || 0);
  const waterAmount = waterUsed * Number(oldRoom.waterPrice || 0);
  const total = oldRentAmount + oldServiceAmount + electricAmount + waterAmount;
  if (total <= 0) return null;
  return {
    oldRoomId: transfer.oldRoomId,
    transferDate: transfer.transferDate,
    oldBillingMonth,
    oldStayFrom: formatDateInputValue(oldStayFrom),
    oldStayTo: formatDateInputValue(oldStayTo),
    oldRoomDays,
    oldMonthDays,
    oldMonthlyRent,
    oldRentAmount,
    oldMonthlyServiceFee,
    oldServiceAmount,
    oldServiceMode: serviceMode,
    electricOld,
    electricNew,
    electricUsed,
    electricAmount,
    waterOld,
    waterNew,
    waterUsed,
    waterAmount,
    total
  };
}

function createMonthlyReceipt(room, contract, previousReceipt, month, billingContext = { mode: 'normal' }) {
  const electricOld = previousReceipt
    ? Number(previousReceipt.electricNew ?? previousReceipt.electricEnd ?? 0)
    : Number(room.electricNew ?? room.electricEnd ?? room.initialElectric ?? room.electricStart ?? 0);

  const waterOld = previousReceipt
    ? Number(previousReceipt.waterNew ?? previousReceipt.waterEnd ?? 0)
    : Number(room.waterNew ?? room.waterEnd ?? room.initialWater ?? room.waterStart ?? 0);

  const isOldTransferRoom = billingContext.mode === 'transfer_old_room' || billingContext.mode === 'transfer_old_room_skip';
  const isNewTransferSameMonth = billingContext.mode === 'transfer_new_room' && monthFromDate(billingContext.transfer?.transferDate) === month;
  const partialNewContract = billingContext.mode === 'new_contract_partial_month' ? billingContext : null;
  const newTransferStart = clampDateToBillingMonth(billingContext.transfer?.transferDate, month);
  const newTransferEnd = billingMonthEnd(month);
  const monthDays = daysInBillingMonth(month);
  const newRoomDays = isNewTransferSameMonth
    ? inclusiveDaysBetween(newTransferStart, newTransferEnd)
    : partialNewContract
      ? partialNewContract.chargeDays
      : monthDays;
  const monthlyRent = Number(contract?.rent || room.rent || 0);
  const monthlyFixedServices = fixedServiceTotal(room, billingContext.occupantCount);
  const rent = isOldTransferRoom ? 0 : (isNewTransferSameMonth || partialNewContract) ? Math.round((monthlyRent / monthDays) * newRoomDays) : monthlyRent;
  const fixedServices = isOldTransferRoom ? 0 : (isNewTransferSameMonth || partialNewContract) ? Math.round((monthlyFixedServices / monthDays) * newRoomDays) : monthlyFixedServices;
  const transferOldUtility = billingContext.mode === 'transfer_new_room' ? billingContext.oldRoomUtility : null;
  const other = Number(transferOldUtility?.total || 0);
  const transferNote = isOldTransferRoom
    ? `Phòng cũ đã chuyển sang P${billingContext.transfer?.newRoomId || ''} trong tháng ${month}: chỉ chốt điện nước, không tính tiền phòng và dịch vụ phòng cũ.`
    : billingContext.mode === 'transfer_new_room'
      ? `Phòng mới nhận khách từ P${billingContext.transfer?.oldRoomId || ''}: chi phí phòng cũ chưa thanh toán được tách riêng, chi phí phòng hiện tại tính theo tháng/phần tháng tương ứng.`
      : billingContext.mode === 'new_contract_partial_month'
        ? `Khách bắt đầu thuê từ ${formatDisplayDate(billingContext.startDate)}. Tiền phòng và dịch vụ tháng ${month} được tính theo ngày đến hết tháng.`
      : "Vui lòng thanh toán trong vòng 5 ngày kể từ ngày nhận phiếu. Xin cảm ơn!";

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
    other,
    total: rent + fixedServices + other,
    paidAmount: 0,
    debt: rent + fixedServices + other,
    status: "Chưa thanh toán",
    note: transferNote,
    billingMode: billingContext.mode,
    transferId: billingContext.transfer?.id || '',
    currentRoomChargeDays: newRoomDays,
    currentRoomMonthDays: monthDays,
    currentRoomMonthlyRent: monthlyRent,
    currentRoomMonthlyServiceFee: monthlyFixedServices,
    currentRoomChargeFrom: partialNewContract?.startDate || (isNewTransferSameMonth ? formatDateInputValue(newTransferStart) : ''),
    currentRoomChargeTo: partialNewContract?.endDate || (isNewTransferSameMonth ? formatDateInputValue(newTransferEnd) : ''),
    transferOldRoomUtility: transferOldUtility,
    createdAt: new Date().toISOString()
  }, room);
}

function enrichReceiptWithTransferUtility(receipt, data) {
  if (!receipt || receipt.type !== 'monthly') return receipt;
  const contract = (data.contracts || []).find(c => c.id === receipt.contractId);
  if (!contract) return receipt;
  const billingContext = getMonthlyBillingContext(data, contract, receipt.month);
  const room = (data.rooms || []).find(r => r.id === receipt.roomId);
  const isNewTransferSameMonth = billingContext.mode === 'transfer_new_room' && monthFromDate(billingContext.transfer?.transferDate) === receipt.month;
  const isPartialNewContract = billingContext.mode === 'new_contract_partial_month';
  const monthDays = daysInBillingMonth(receipt.month);
  const currentRoomChargeDays = isNewTransferSameMonth
    ? inclusiveDaysBetween(clampDateToBillingMonth(billingContext.transfer?.transferDate, receipt.month), billingMonthEnd(receipt.month))
    : isPartialNewContract
      ? billingContext.chargeDays
      : Number(receipt.currentRoomChargeDays || monthDays);
  const expectedMonthlyFixedServices = fixedServiceTotal(room, billingContext.occupantCount);
  const expectedFixedServices = (billingContext.mode === 'transfer_old_room' || billingContext.mode === 'transfer_old_room_skip')
    ? 0
    : (isNewTransferSameMonth || isPartialNewContract)
      ? Math.round((expectedMonthlyFixedServices / monthDays) * currentRoomChargeDays)
    : expectedMonthlyFixedServices;
  const expectedMonthlyRent = Number(contract?.rent || room?.rent || 0);
  const expectedRent = (billingContext.mode === 'transfer_old_room' || billingContext.mode === 'transfer_old_room_skip')
    ? 0
    : (isNewTransferSameMonth || isPartialNewContract)
      ? Math.round((expectedMonthlyRent / monthDays) * currentRoomChargeDays)
      : Number(receipt.rent || 0);
  const fixedDelta = expectedFixedServices - Number(receipt.fixedServices || 0);
  const rentDelta = expectedRent - Number(receipt.rent || 0);
  const transferOldUtility = billingContext.mode === 'transfer_new_room' ? billingContext.oldRoomUtility : null;
  const existingTransferTotal = receipt.transferId === billingContext.transfer?.id && receipt.transferOldRoomUtility
    ? Number(receipt.transferOldRoomUtility.total || 0)
    : 0;
  const transferExtra = transferOldUtility ? Number(transferOldUtility.total || 0) - existingTransferTotal : 0;
  const needsTransferMetadata = transferOldUtility && (
    receipt.transferOldRoomUtility !== transferOldUtility ||
    !receipt.transferOldRoomUtility?.oldRoomDays ||
    receipt.transferOldRoomUtility?.oldRentAmount !== transferOldUtility.oldRentAmount ||
    receipt.transferOldRoomUtility?.oldServiceAmount !== transferOldUtility.oldServiceAmount
  );
  if (rentDelta === 0 && fixedDelta === 0 && transferExtra === 0 && !needsTransferMetadata) return receipt;
  const total = Number(receipt.total || 0) + rentDelta + fixedDelta + transferExtra;
  const paidAmount = Number(receipt.paidAmount || 0);
  return {
    ...receipt,
    rent: expectedRent,
    fixedServices: expectedFixedServices,
    other: Number(receipt.other || 0) + transferExtra,
    total,
    debt: Math.max(0, total - paidAmount),
    status: paidAmount >= total && total > 0 ? 'Đã thanh toán' : paidAmount > 0 ? 'Nợ một phần' : receipt.status,
    billingMode: billingContext.mode,
    transferId: billingContext.transfer?.id || receipt.transferId || '',
    currentRoomChargeDays,
    currentRoomMonthDays: monthDays,
    currentRoomMonthlyRent: expectedMonthlyRent,
    currentRoomMonthlyServiceFee: expectedMonthlyFixedServices,
    currentRoomChargeFrom: billingContext.startDate || (isNewTransferSameMonth ? formatDateInputValue(clampDateToBillingMonth(billingContext.transfer?.transferDate, receipt.month)) : receipt.currentRoomChargeFrom || ''),
    currentRoomChargeTo: billingContext.endDate || (isNewTransferSameMonth ? formatDateInputValue(billingMonthEnd(receipt.month)) : receipt.currentRoomChargeTo || ''),
    transferOldRoomUtility: transferOldUtility || receipt.transferOldRoomUtility,
    note: receipt.note || (isPartialNewContract
      ? `Khách bắt đầu thuê từ ${formatDisplayDate(billingContext.startDate)}. Tiền phòng và dịch vụ tháng ${receipt.month} được tính theo ngày đến hết tháng.`
      : `Phòng mới nhận khách từ P${billingContext.transfer?.oldRoomId || ''}: chi phí phòng cũ chưa thanh toán được tách riêng trong phiếu này.`)
  };
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

function getReceiptPaymentState(receipt) {
  const total = Number(receipt?.total || 0);
  const paidAmount = Number(receipt?.paidAmount || 0);
  const explicitAdjustmentDue = Number(receipt?.adjustmentDueAmount || 0);
  const looksLikeUtilityCheckoutAdjustment =
    receipt?.type === 'monthly' &&
    receipt?.isFinalized &&
    paidAmount >= total &&
    total > 0 &&
    Number(receipt?.rent || 0) === 0 &&
    Number(receipt?.fixedServices || 0) === 0 &&
    (Number(receipt?.electricAmount || 0) + Number(receipt?.waterAmount || 0) + Number(receipt?.other || 0)) > 0 &&
    !receipt?.adjustmentPaidAmount;
  const adjustmentDue = explicitAdjustmentDue > 0 ? explicitAdjustmentDue : looksLikeUtilityCheckoutAdjustment ? total : 0;
  if (adjustmentDue > 0) {
    const adjustmentPaidAmount = Number(receipt?.adjustmentPaidAmount || 0);
    const adjustmentDebt = Math.max(0, adjustmentDue - adjustmentPaidAmount);
    if (receipt?.status === 'Đã hủy') return { status: 'Đã hủy', debt: adjustmentDebt, paidAmount: adjustmentPaidAmount, basePaidAmount: paidAmount, adjustmentDue, isPaid: false, isPartial: false, isAdjustment: true };
    if (adjustmentPaidAmount >= adjustmentDue) return { status: 'Đã thanh toán', debt: 0, paidAmount: adjustmentPaidAmount, basePaidAmount: paidAmount, adjustmentDue, isPaid: true, isPartial: false, isAdjustment: true };
    if (adjustmentPaidAmount > 0) return { status: 'Nợ một phần', debt: adjustmentDebt, paidAmount: adjustmentPaidAmount, basePaidAmount: paidAmount, adjustmentDue, isPaid: false, isPartial: true, isAdjustment: true };
    return { status: 'Chưa thanh toán', debt: adjustmentDebt, paidAmount: 0, basePaidAmount: paidAmount, adjustmentDue, isPaid: false, isPartial: false, isAdjustment: true };
  }
  const debt = Math.max(0, total - paidAmount);
  if (receipt?.status === 'Đã hủy') return { status: 'Đã hủy', debt, paidAmount, basePaidAmount: paidAmount, adjustmentDue: 0, isPaid: false, isPartial: false, isAdjustment: false };
  if (total > 0 && paidAmount >= total) return { status: 'Đã thanh toán', debt: 0, paidAmount, basePaidAmount: paidAmount, adjustmentDue: 0, isPaid: true, isPartial: false, isAdjustment: false };
  if (paidAmount > 0) return { status: 'Nợ một phần', debt, paidAmount, basePaidAmount: paidAmount, adjustmentDue: 0, isPaid: false, isPartial: true, isAdjustment: false };
  return { status: 'Chưa thanh toán', debt, paidAmount, basePaidAmount: paidAmount, adjustmentDue: 0, isPaid: false, isPartial: false, isAdjustment: false };
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

function printDocumentElement(elementId, title = 'Tài liệu', pageSize = 'A4') {
  const content = document.getElementById(elementId);
  if (!content) {
    alert('Chưa tìm thấy nội dung để in.');
    return;
  }

  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right = '0';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = '0';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;
  doc.open();
  doc.write(`
    <html>
      <head>
        <title>${title}</title>
        <style>
          @page { size: ${pageSize}; margin: 12mm 0 14mm; }
          html, body { margin: 0; padding: 0; background: white; color: black; }
          body { font-family: "Times New Roman", Times, serif; font-size: 13px; line-height: 1.38; }
          .contract-paper,
          .appendix-content-v1 {
            width: 210mm !important;
            max-width: none !important;
            min-height: 273mm !important;
            margin: 0 auto !important;
            padding: 2mm 20mm 6mm !important;
            box-sizing: border-box !important;
            box-shadow: none !important;
            background: white !important;
            color: black !important;
            transform: none !important;
          }
          .contract-header-text { text-align: center; margin-bottom: 18px; }
          h1 { font-size: 18px; text-align: center; margin: 12px 0 8px; }
          h2 { font-size: 15px; text-align: center; margin: 8px 0 5px; }
          h3, h4 { font-size: 13px; margin: 10px 0 5px; break-after: avoid; page-break-after: avoid; }
          p { margin: 3px 0; orphans: 3; widows: 3; }
          table { width: 100%; border-collapse: collapse; margin: 8px 0; page-break-inside: avoid; break-inside: avoid; }
          th, td { border: 1px solid #111; padding: 4px 6px; vertical-align: top; }
          .appendix-page { page-break-before: always; margin-top: 0 !important; border-top: 0 !important; padding-top: 10mm !important; }
          .appendix-section { break-inside: avoid; page-break-inside: avoid; }
          .signature-row,
          .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 36px; text-align: center; margin-top: 18px; break-inside: avoid; page-break-inside: avoid; }
          .signature-space { height: 58px; }
          * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        </style>
      </head>
      <body>${content.outerHTML}</body>
    </html>
  `);
  doc.close();

  setTimeout(() => {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    setTimeout(() => document.body.removeChild(iframe), 1500);
  }, 500);
}

function ContractPreview({ contract, room, tenants, bankInfo, report, type = 'main', onClose }) {
  const primaryTenant = tenants.find(t => t.role === 'primary') || tenants[0] || {};
  const signedDay = formatContractDate(contract.signedDate);
  const startDay = formatBusinessDate(contract.startDate);
  const endDay = formatBusinessDate(contract.endDate);
  const duration = calculateRentalDuration(contract.startDate, contract.endDate);
  const rentAmount = Number(contract.rent || room.rent || 0);
  const depositAmount = Number(contract.deposit || room.deposit || rentAmount || 0);
  const electricPrice = Number(contract.terms?.electricPrice || room.electricPrice || 3800);
  const waterPrice = Number(contract.terms?.waterPrice || room.waterPrice || 32000);
  const baseFixedServiceTotal = Number(room.internet || 0) + Number(room.cleaning || 0) + Number(room.elevator || 0) + Number(room.laundry || 0);
  const occupantCount = Math.max(1, tenants.length || 1);
  const appliedFixedServiceTotal = fixedServiceTotal(room, occupantCount);
  const fixedServiceDiscount = Math.max(0, baseFixedServiceTotal - appliedFixedServiceTotal);
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
            <button className="primary-btn" onClick={() => printDocumentElement('contract-preview-paper', `${isRenewal ? 'Phụ lục gia hạn' : isLiquidation ? 'Biên bản tất toán' : 'Hợp đồng thuê'} P${room.id}`)}>🖨️ In tài liệu</button>
            <button className="secondary-btn" onClick={onClose}>Đóng</button>
          </div>
        </div>
        <div id="contract-preview-paper" className="contract-paper">
          {(isMain || isRenewal) && (
            <>
              <div className="contract-header-text">
                <h1>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h1>
                <h2>Độc lập – Tự do – Hạnh phúc</h2>
                <p style={{ fontStyle: 'italic' }}>Hà Nội, ngày {signedDay}</p>
              </div>
              <h1 style={{ textAlign: 'center', margin: '30px 0 20px' }}>HỢP ĐỒNG THUÊ PHÒNG</h1>
              {isMain && (
                <div className="contract-section">
                  <p>- Căn cứ Bộ luật Dân sự số 91/2015/QH13 ngày 24/11/2015;</p>
                  <p>- Căn cứ Luật Thương mại số 36/2005/QH11 ngày 14/06/2005;</p>
                  <p>- Căn cứ nhu cầu và sự thỏa thuận của các Bên.</p>
                  <p style={{ marginTop: '10px' }}>Hôm nay, ngày {signedDay}, chúng tôi gồm:</p>
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
                <p>2.1. Ngày ký Hợp đồng: <b>{signedDay}</b>. Ngày bàn giao dự kiến/thực tế: <b>{startDay}</b>.</p>
                <p>2.2. Hai Bên lập Phụ lục 01 – Biên bản bàn giao kèm ảnh/video, ghi rõ từng thiết bị, tình trạng thiết bị, hiện trạng phòng, chỉ số điện/nước đầu kỳ, số chìa khóa/thẻ/vân tay đã bàn giao.</p>
                <p>2.3. Kể từ thời điểm bàn giao thực tế theo Phụ lục 01, Bên B có toàn quyền sử dụng Phòng thuê theo Hợp đồng.</p>
                <h3>ĐIỀU 3. THỜI HẠN THUÊ</h3>
                <p>3.1. Thời hạn thuê: <b>{duration}</b>.</p>
                <p>3.2. Thời hạn được tính từ ngày <b>{startDay}</b> đến hết ngày <b>{endDay}</b>. Nếu ngày bàn giao thực tế khác ngày bắt đầu thuê, Hai Bên ghi nhận bằng Phụ lục 01 hoặc tin nhắn xác nhận.</p>
                <p>3.3. Hết thời hạn thuê, Hợp đồng tự động gia hạn theo tháng với điều khoản không đổi, trừ khi một Bên thông báo chấm dứt trước ít nhất 30 ngày bằng văn bản/tin nhắn (Zalo/SMS).</p>
                <h3>ĐIỀU 4. TIỀN ĐẶT CỌC</h3>
                <p>4.1. Mức đặt cọc: <b>{formatMoney(depositAmount)}</b> (bằng chữ: <b>{numberToWords(depositAmount)}</b>), nộp ngay khi ký hoặc theo thỏa thuận thanh toán của Hai Bên.</p>
                <p>4.2. Không dùng cọc để trừ tiền thuê trừ khi Hai Bên đồng ý bằng văn bản/tin nhắn.</p>
                <p><b>4.3. Thời hạn thuê tối thiểu và xử lý tiền đặt cọc khi chấm dứt trước hạn</b></p>
                <p>4.3.1. Bên B cam kết thời gian thuê tối thiểu là <b>12 (mười hai) tháng liên tục</b>, tính từ ngày Bên A bàn giao phòng cho Bên B theo Hợp đồng này.</p>
                <p>4.3.2. Trường hợp Bên B tự ý trả phòng, ngừng thuê hoặc đơn phương chấm dứt Hợp đồng trước khi đủ 12 (mười hai) tháng vì nguyên nhân thuộc về Bên B, thì được xác định là vi phạm cam kết về thời hạn thuê tối thiểu. Khi đó, <b>toàn bộ số tiền đặt cọc thuộc về Bên A và Bên A không có nghĩa vụ hoàn trả cho Bên B</b>, kể cả trường hợp Bên B đã thông báo trước ít nhất 30 ngày.</p>
                <p>Việc thông báo trước chỉ nhằm thực hiện nghĩa vụ thông báo khi chấm dứt Hợp đồng và <b>không làm phát sinh quyền được hoàn trả tiền đặt cọc</b> của Bên B trong trường hợp này.</p>
                <p>4.3.3. Ngoài việc không được hoàn trả tiền đặt cọc, Bên B vẫn có trách nhiệm thanh toán đầy đủ tiền thuê phòng, tiền điện, nước, dịch vụ và các khoản công nợ khác phát sinh đến thời điểm bàn giao thực tế; đồng thời bồi thường chi phí sửa chữa, khắc phục các hư hỏng do lỗi của Bên B gây ra. Trường hợp các khoản phải thanh toán và thiệt hại thực tế vượt quá số tiền đặt cọc thì Bên B phải thanh toán phần chênh lệch còn thiếu cho Bên A.</p>
                <p>4.3.4. Quy định không hoàn trả tiền đặt cọc tại Khoản 4.3.2 không áp dụng trong các trường hợp sau:</p>
                <p>a) Hai Bên có thỏa thuận bằng văn bản về việc chấm dứt Hợp đồng và hoàn trả tiền đặt cọc;</p>
                <p>b) Bên B chấm dứt Hợp đồng do Bên A vi phạm nghiêm trọng nghĩa vụ theo Hợp đồng hoặc thuộc trường hợp Bên B có quyền đơn phương chấm dứt Hợp đồng theo quy định pháp luật;</p>
                <p>c) Xảy ra sự kiện bất khả kháng khiến mục đích thuê phòng không thể tiếp tục thực hiện và hai Bên có văn bản xác nhận hoặc cơ quan có thẩm quyền xác định;</p>
                <p>d) Các trường hợp khác do hai Bên thống nhất bằng văn bản.</p>
                <p>Trong các trường hợp trên, Bên A hoàn trả phần tiền đặt cọc còn lại cho Bên B sau khi khấu trừ các khoản công nợ, chi phí sửa chữa và nghĩa vụ tài chính hợp lệ trong thời hạn <b>07 ngày làm việc</b> kể từ ngày hai Bên hoàn tất bàn giao và đối soát công nợ.</p>
                <p>4.3.5. Trường hợp Bên B chấm dứt Hợp đồng sau khi đã thuê đủ 12 tháng, Bên B phải thông báo cho Bên A trước ít nhất 30 ngày bằng văn bản hoặc tin nhắn theo phương thức liên lạc đã thỏa thuận. Sau khi hoàn tất việc bàn giao phòng và thanh toán đầy đủ các nghĩa vụ, Bên A hoàn trả phần tiền đặt cọc còn lại cho Bên B theo Khoản 4.4 của Hợp đồng.</p>
                <p>4.4. Khi trả phòng, Hai Bên lập Biên bản kiểm tra hiện trạng. Bên A chỉ được khấu trừ cọc đối với công nợ, hư hỏng do lỗi Bên B hoặc chi phí có bảng kê/hóa đơn/chứng từ hợp lệ. Bên A hoàn cọc trong 07 (bảy) ngày làm việc kể từ khi nhận đủ chìa khóa/thẻ, bàn giao xong và quyết toán công nợ với Bên B.</p>
                <h3>ĐIỀU 5. GIÁ THUÊ VÀ PHÍ DỊCH VỤ</h3>
                <p>5.1. Tiền thuê: <b>{formatMoney(rentAmount)}</b>/tháng (bằng chữ: <b>{numberToWords(rentAmount)}</b>).</p>
                <p>5.2. Phí dịch vụ chưa gồm trong tiền thuê. Tổng phí dịch vụ cố định gốc: <b>{formatMoney(baseFixedServiceTotal)}</b>/tháng. Trường hợp Bên B ở một mình trong phòng được giảm 50% phí dịch vụ cố định; từ 02 người trở lên áp dụng đủ phí dịch vụ cố định.</p>
                <p>5.3. Mức phí dịch vụ cố định đang áp dụng cho Hợp đồng này: <b>{formatMoney(appliedFixedServiceTotal)}</b>/tháng theo số người ở hiện tại là <b>{occupantCount}</b> người{fixedServiceDiscount > 0 ? <>, đã giảm <b>{formatMoney(fixedServiceDiscount)}</b>/tháng.</> : <>.</>}</p>
                <p>- Điện: <b>{formatMoney(electricPrice)}</b>/kWh, theo chỉ số công tơ (ghi đầu/cuối kỳ trong phiếu thu).</p>
                <p>- Nước: <b>{formatMoney(waterPrice)}</b>/m³, theo chỉ số công tơ (ghi đầu/cuối kỳ trong phiếu thu).</p>
                <p>- Mạng internet: <b>{formatMoney(room.internet)}</b>/phòng/tháng.</p>
                <p>- Vệ sinh rác & dịch vụ chung: <b>{formatMoney(room.cleaning)}</b>/phòng/tháng.</p>
                <p>- Phí bảo trì thang máy: <b>{formatMoney(room.elevator)}</b>/phòng/tháng.</p>
                <p>- Phí sử dụng máy giặt chung: <b>{formatMoney(room.laundry)}</b>/phòng/tháng.</p>
                <p>5.4. Khi đơn giá đầu vào, số người ở hoặc phí dịch vụ thay đổi, Bên A báo trước ít nhất 15 ngày và gửi bảng tính kèm căn cứ/chứng từ nếu có.</p>
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
                <p>9.1. Hợp đồng chấm dứt khi hết hạn và có thông báo chấm dứt trước ít nhất 30 ngày theo Điều 3; nếu không, gia hạn theo tháng.</p>
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
                <p>Ngày ký hợp đồng: <b>{signedDay}</b>; Ngày bàn giao: <b>{startDay}</b>; Phòng số: <b>{room.id}</b></p>
                <p>Chỉ số điện cũ: <b>{room.electricOld || room.electricStart || room.initialElectric || '.......'}</b> kWh; Chỉ số nước cũ: <b>{room.waterOld || room.waterStart || room.initialWater || '.......'}</b> m³</p>
                <p>Hai Bên thống nhất chụp ảnh/quay video hiện trạng phòng, thiết bị, công tơ điện/nước và lưu kèm biên bản này. Ảnh/video là căn cứ đối chiếu khi trả phòng và hoàn cọc.</p>
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
                        <td>☐ Tốt &nbsp; ☐ Trầy/xước &nbsp; ☐ Cần kiểm tra &nbsp; ☐ Hỏng &nbsp; Ghi chú: ............</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="appendix-page">
                <h1 style={{ textAlign: 'center', fontSize: '18px', fontWeight: 'bold' }}>PHỤ LỤC 03 – NỘI QUY TÒA NHÀ</h1>
                <div style={{ fontSize: '11px', lineHeight: '1.4' }}>
                  {[
                    ['1. Tuân thủ pháp luật', 'Nghiêm cấm đánh nhau, cờ bạc, mại dâm, ma túy, tàng trữ hoặc sử dụng chất cấm và các hành vi vi phạm pháp luật khác.'],
                    ['2. PCCC', 'CẤM TUYỆT ĐỐI SẠC XE ĐIỆN, pin hoặc ắc-quy xe điện trong mọi khu vực tòa nhà, bao gồm phòng ở, hành lang, cầu thang, khu để xe và khu vực chung.'],
                    ['3. Khai báo tạm trú', 'Cung cấp giấy tờ trong 24 giờ kể từ khi vào ở để Bên A hỗ trợ làm thủ tục khai báo tạm trú theo quy định.'],
                    ['4. Thoát nước', 'Không đổ rác, thức ăn thừa, tóc, dầu mỡ hoặc vật cứng xuống bồn cầu, lavabo, chậu rửa và thoát sàn.'],
                    ['5. Vệ sinh', 'Không để rác ở hành lang, cầu thang hoặc khu vực chung. Bỏ rác đúng nơi quy định và giữ gìn vệ sinh phòng ở, khu sinh hoạt chung.'],
                    ['6. Tiếng ồn', 'Giữ trật tự trong khung giờ 22:00–06:00. Tắt máy xe khi vào nhà ban đêm, hạn chế nói chuyện lớn, mở nhạc, kéo đồ hoặc gây tiếng ồn sau 22:30.'],
                    ['7. Ra vào', 'Đóng cửa cổng/cửa chính cẩn thận, không mở cửa cho người lạ. Giờ tiếp khách thông thường từ 06:00–22:30; khách ở lại qua đêm phải báo trước và được Bên A đồng ý.'],
                    ['8. Kết cấu', 'Không tự ý khoan, đục, vẽ, dán, sửa chữa hoặc thay đổi kết cấu tường, trần, sàn, cửa, hệ thống điện nước khi chưa được Bên A đồng ý.'],
                    ['9. Bồi thường', 'Chịu trách nhiệm bồi thường hoặc thanh toán chi phí sửa chữa đối với hư hỏng, mất mát tài sản, thiết bị do lỗi của mình hoặc khách của mình gây ra.'],
                    ['10. Tài sản', 'Tự bảo quản tài sản cá nhân, tiền bạc, giấy tờ và phương tiện. Bên A không chịu trách nhiệm đối với mất mát do Bên B không khóa cửa hoặc tự quản lý sơ suất.'],
                    ['11. Trung thực', 'Không lấy, sử dụng hoặc di chuyển tài sản của người khác khi chưa được đồng ý. Mọi tranh chấp phát sinh phải báo ngay cho Bên A để phối hợp xử lý.'],
                    ['12. Vật nuôi', 'Không nuôi động vật gây mất vệ sinh, tiếng ồn, mùi hoặc nguy cơ mất an toàn. Trường hợp có vật nuôi phải được Bên A đồng ý trước bằng tin nhắn/văn bản.'],
                    ['13. Kinh doanh', 'Không sử dụng phòng làm nơi kinh doanh, kho hàng, tập kết hàng hóa, livestream bán hàng gây ảnh hưởng cư dân hoặc mục đích trái pháp luật.'],
                    ['14. Vận chuyển', 'Không chở quá tải thang máy, không kéo lê vật nặng gây hư hỏng sàn, tường, cửa hoặc thiết bị chung. Khi chuyển đồ lớn phải báo trước cho Bên A.'],
                    ['15. Kỹ thuật', 'Không tự ý vào phòng kỹ thuật, tủ điện, khu máy bơm, mái, kho hoặc khu vực có biển cấm. Không tự ý đấu nối, sửa chữa điện nước trong tòa nhà.'],
                    ['16. Khách', 'Tiếp khách trong khung giờ 06:00–22:30. Khách ở lại qua đêm phải báo trước và được Bên A đồng ý; Bên B chịu trách nhiệm về hành vi, tài sản và an ninh liên quan đến khách của mình.'],
                    ['17. Ở ghép', 'Không cho thuê lại, chuyển nhượng, cho người khác ở ghép hoặc thay đổi người ở khi chưa có chấp thuận của Bên A và chưa bổ sung thông tin cư trú.'],
                    ['18. Cam kết', 'Tuân thủ nghiêm túc nội quy tòa nhà, hợp đồng thuê phòng và các thông báo quản lý hợp lý của Bên A. Tái phạm nhiều lần hoặc vi phạm nghiêm trọng có thể bị chấm dứt hợp đồng theo thỏa thuận.']
                  ].map(([title, content], idx) => (
                    <p key={idx} style={{ marginBottom: '3px' }}><b>{title}:</b> {content}</p>
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
              <p>Hôm nay, ngày {formatDisplayDate(report.actualEndDate)}, chúng tôi tiến hành chốt dọn đi cho phòng <b>{room.id}</b>:</p>
              <table className="contract-table" style={{ margin: '20px 0' }}>
                <tbody>
                  <tr><td>Chỉ số điện mới: <b>{report.electricNew || report.electricEnd}</b> kWh</td><td>Chỉ số nước mới: <b>{report.waterNew || report.waterEnd}</b> m³</td></tr>
                  <tr><td>Phát sinh điện: {report.electricUsed} kWh = <b>{formatMoney(report.electricAmount)}</b></td><td>Phát sinh nước: {report.waterUsed} m³ = <b>{formatMoney(report.waterAmount)}</b></td></tr>
                </tbody>
              </table>
              <h3 style={{ fontSize: '16px', fontWeight: 'bold' }}>QUYẾT TOÁN CÔNG NỢ</h3>
              <div className="stack" style={{ gap: '5px' }}>
                <div className="data-row"><span>Tiền phòng phát sinh ({report.roomChargeDays || 0} ngày)</span><b>{formatMoney(report.proratedRent || 0)}</b></div>
                <div className="data-row"><span>Dịch vụ phát sinh ({report.roomChargeDays || 0} ngày)</span><b>{formatMoney(report.proratedFixedServices || 0)}</b></div>
                <div className="data-row"><span>Tiền điện</span><b>{formatMoney(report.electricAmount || 0)}</b></div>
                <div className="data-row"><span>Tiền nước</span><b>{formatMoney(report.waterAmount || 0)}</b></div>
                <div className="data-row"><span>Phí khác</span><b>{formatMoney(Number(report.unpaidRent || 0) + Number(report.cleaningFee || 0) + Number(report.damageFee || 0) + Number(report.otherFee || 0))}</b></div>
                <div className="data-row"><span>Tổng phát sinh</span><b>{formatMoney(report.totalIncurred || 0)}</b></div>
                <div className="data-row"><span>Tiền cọc đối trừ</span><b>{formatMoney(report.depositUsed)}</b></div>
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
  const [data, setData] = useState(() => {
    const loaded = safeRead(STORAGE_KEY, DEFAULT_DATA);
    return {
      ...DEFAULT_DATA,
      ...loaded,
      suppliers: (loaded.suppliers && loaded.suppliers.length > 20) 
        ? loaded.suppliers.map(s => {
            const def = DEFAULT_DATA.suppliers.find(d => d.id === s.id);
            return def ? { ...def, ...s, defaultCategory: s.defaultCategory || def.defaultCategory } : s;
          })
        : DEFAULT_DATA.suppliers,
      expenseCategories: loaded.expenseCategories || DEFAULT_DATA.expenseCategories,
      expensePayments: loaded.expensePayments || [],
      contractRenewals: loaded.contractRenewals || [],
      roomTransfers: loaded.roomTransfers || [],
    };
  });
  const [bankInfo, setBankInfo] = useState(() => {
    const saved = safeRead(BANK_KEY, DEFAULT_BANK);
    // Replace the old demo account while preserving any real account configured by the user.
    if (!saved?.accountNo || saved.accountNo === '0123456789') return DEFAULT_BANK;
    return { ...DEFAULT_BANK, ...saved };
  });
  const [tab, setTab] = useState('dashboard');
  const [query, setQuery] = useState('');
  const [paymentFilters, setPaymentFilters] = useState(null);
  const [selectedRoom, setSelectedRoom] = useState(null);
  const [detailTenant, setDetailTenant] = useState(null);
  const [newRentalRoom, setNewRentalRoom] = useState(null);
  const [addingRoommate, setAddingRoommate] = useState(null);
  const [settlingRoom, setSettlingRoom] = useState(null);
  const [transferringContract, setTransferringContract] = useState(null);
  const [viewingContract, setViewingContract] = useState(null);
  const [renewingContract, setRenewingContract] = useState(null);
  const [viewingAppendix, setViewingAppendix] = useState(null);
  const [viewingReceipt, setViewingReceipt] = useState(null);
  const [printingReceipts, setPrintingReceipts] = useState(null);
  const [paymentReceipt, setPaymentReceipt] = useState(null);
  const [editingTenant, setEditingTenant] = useState(null);
  const [editingContract, setEditingContract] = useState(null);
  const [viewingExpense, setViewingExpense] = useState(null);
  const [editingExpense, setEditingExpense] = useState(null);
  const [addingExpense, setAddingExpense] = useState(false);
  const [viewingSupplier, setViewingSupplier] = useState(null);
  const [editingSupplier, setEditingSupplier] = useState(null);
  const [addingSupplier, setAddingSupplier] = useState(false);
  const [roomOpsModal, setRoomOpsModal] = useState(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [hasLoadedCloud, setHasLoadedCloud] = useState(false);
  const [cloudEnabled, setCloudEnabled] = useState(true);
  const [lastSynced, setLastSynced] = useState(null);
  const fileInputRef = React.useRef(null);
  const cloudFailureRef = React.useRef(0);

  // Initial Fetch from Cloud
  useEffect(() => {
    async function initCloud() {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 8000);
      try {
        const res = await fetch('/api/data', { signal: controller.signal });
        if (!res.ok) {
          const body = await res.text().catch(() => '');
          throw new Error(`GET /api/data ${res.status}: ${body.slice(0, 400) || res.statusText}`);
        }
        const cloudData = await res.json();
        if (cloudData && !cloudData.error && Array.isArray(cloudData.rooms)) {
          // Cloud is authoritative after a successful fetch. Choosing the newer
          // local timestamp can resurrect records that were intentionally deleted
          // or restored directly on the server.
          setData(prev => ({
            ...prev,
            ...cloudData,
            suppliers: cloudData.suppliers || prev.suppliers || [],
            expenseCategories: cloudData.expenseCategories || prev.expenseCategories || DEFAULT_DATA.expenseCategories,
            expensePayments: cloudData.expensePayments || prev.expensePayments || [],
            contractRenewals: cloudData.contractRenewals || prev.contractRenewals || [],
            roomTransfers: cloudData.roomTransfers || prev.roomTransfers || [],
          }));
          setLastSynced(new Date());
          setCloudEnabled(true);
        }
      } catch (err) {
        setCloudEnabled(false);
        console.info("Cloud sync inactive; using local storage only.", err?.message || err);
      } finally {
        clearTimeout(timeout);
        setHasLoadedCloud(true);
      }
    }
    initCloud();
  }, []);

  // Cloud Sync Logic (Debounced)
  useEffect(() => {
    const timer = setTimeout(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      if (!data || data === DEFAULT_DATA || !hasLoadedCloud || !cloudEnabled) return;
      
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 8000);
      setIsSyncing(true);
      fetch('/api/data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'full_sync', payload: data }),
        signal: controller.signal
      })
      .then(async res => {
        if (res.ok) {
          cloudFailureRef.current = 0;
          setLastSynced(new Date());
        } else {
          const errData = await res.json().catch(async () => ({ raw: await res.text().catch(() => '') }));
          cloudFailureRef.current += 1;
          console.info(`Cloud sync failed (${res.status}); local data was saved.`, errData);
          if (res.status >= 500 || cloudFailureRef.current >= 2) {
            setCloudEnabled(false);
          }
        }
        clearTimeout(timeout);
        setIsSyncing(false);
      })
      .catch(err => {
        clearTimeout(timeout);
        cloudFailureRef.current += 1;
        setIsSyncing(false);
        console.info("Cloud sync unavailable; local data was saved.", err?.message || err);
        setCloudEnabled(false);
      });
    }, 2000);

    return () => clearTimeout(timer);
  }, [data, hasLoadedCloud, cloudEnabled]);
  useEffect(() => {
    if (!data || !data.contracts || !data.memberships) return;
    
    let changed = false;
    // 1. Ensure memberships sync with contract status
    const updatedMemberships = data.memberships.map(m => {
      const contract = data.contracts.find(c => c.id === m.contractId);
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
  }, [data?.contracts, data?.memberships]);

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }, [data]);
  useEffect(() => { localStorage.setItem(BANK_KEY, JSON.stringify(bankInfo)); }, [bankInfo]);

  const filteredTenants = useMemo(() => {
    const q = query.toLowerCase();
    return (data.tenants || []).filter(t => t.name.toLowerCase().includes(q) || (t.phone && t.phone.includes(q)) || (t.cccd && t.cccd.includes(q)));
  }, [data.tenants, query]);

  const confirmDanger = (title, consequence) => {
    if (!window.confirm(`${title}\n\n${consequence}`)) return false;
    return window.confirm('Xác nhận lần 2: thao tác này sẽ được ghi vào dữ liệu phòng. Bạn chắc chắn tiếp tục?');
  };

  function handleAction(type, roomOrTenant) {
    if (type === 'add_tenant') {
      if (!roomOrTenant?.id) {
        alert('Vui lòng chọn phòng cụ thể để thêm khách thuê.');
        return;
      }
      const room = data.rooms.find(r => r.id === roomOrTenant.id);
      if (!room) {
        alert('Không tìm thấy phòng cần tạo hợp đồng.');
        return;
      }
      const roomStatus = getRoomStatusInfo(data, room.id);
      if (roomStatus.ownerOccupied) {
        alert('Phòng chủ nhà ở không cần tạo hợp đồng thuê.');
        return;
      }
      if (roomStatus.contract && ['active', 'notice', 'moving_out'].includes(roomStatus.contract.status)) {
        alert(`Phòng ${room.id} đang có hợp đồng hiệu lực. Vui lòng tất toán/kết thúc hợp đồng hiện tại trước khi tạo hợp đồng mới.`);
        return;
      }
      setNewRentalRoom(room);
    } else if (type === 'add_roommate') {
      const roomId = roomOrTenant?.id || roomOrTenant?.roomId;
      const room = data.rooms.find(r => r.id === roomId);
      const activeContract = roomOrTenant?.contractId
        ? data.contracts.find(c => c.id === roomOrTenant.contractId)
        : data.contracts.find(c => c.roomId === roomId && ['active', 'notice', 'moving_out'].includes(c.status));
      if (!room || !activeContract) {
        alert('Phòng cần có hợp đồng hiệu lực trước khi thêm người ở cùng.');
        return;
      }
      setAddingRoommate({ room, contract: activeContract });
    } else if (type === 'notice') {
      if (window.confirm(`Xác nhận báo chuyển cho phòng ${roomOrTenant.id}?\n\nPhòng sẽ được đưa vào trạng thái sắp trống để theo dõi công nợ và lịch kiểm phòng.`)) updateContractStatus(roomOrTenant.id, 'notice');
    } else if (type === 'cancel_notice') {
      if (window.confirm(`Hủy báo chuyển phòng ${roomOrTenant.id}?`)) updateContractStatus(roomOrTenant.id, 'active');
    } else if (type === 'moving_out') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const room = data.rooms.find(r => r.id === roomId);
      const activeContract = roomOrTenant.contractId
        ? data.contracts.find(c => c.id === roomOrTenant.contractId)
        : getCurrentContractForRoom(data, roomId);
      if (activeContract && confirmDanger(
        `Bạn chắc chắn muốn tất toán/kết thúc thuê phòng ${roomId}?`,
        'Hệ thống sẽ chốt công nợ, cập nhật trạng thái hợp đồng, kết thúc cư trú và ghi lịch sử trả phòng.'
      )) setSettlingRoom({ room, contract: activeContract });
    } else if (type === 'view_contract') {
      const roomForContract = roomOrTenant?.roomId ? data.rooms.find(r => r.id === roomOrTenant.roomId) : data.rooms.find(r => r.id === roomOrTenant?.id);
      if (isOwnerOccupiedRoom(roomForContract)) {
        alert('Phòng chủ nhà ở không cần hợp đồng.');
        return;
      }
      const tenantMembership = roomOrTenant && !roomOrTenant.roomId
        ? (data.memberships || []).find(m => m.tenantId === roomOrTenant.id && m.status === 'active')
        : null;
      const roomId = roomOrTenant.roomId || tenantMembership?.roomId || roomOrTenant.id;
      const contractId = roomOrTenant.contractId || tenantMembership?.contractId;
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
      const activeMemberships = (data.memberships || []).filter(m => m.tenantId === roomOrTenant.id && m.status === 'active');
      const targetMembership = activeMemberships.find(m => m.contractId === roomOrTenant.contractId)
        || activeMemberships.find(m => m.roomId === roomOrTenant.roomId)
        || activeMemberships[0];
      setEditingTenant({
        ...roomOrTenant,
        role: targetMembership?.role || roomOrTenant.role || 'secondary',
        membershipId: targetMembership?.id || roomOrTenant.membershipId || '',
        contractId: targetMembership?.contractId || roomOrTenant.contractId || '',
        roomId: targetMembership?.roomId || roomOrTenant.roomId || ''
      });
    } else if (type === 'renew_contract') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const activeContract = roomOrTenant.contractId
        ? data.contracts.find(c => c.id === roomOrTenant.contractId)
        : getCurrentContractForRoom(data, roomId);
      if (activeContract) setRenewingContract(activeContract);
    } else if (type === 'print_appendix') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const contractId = roomOrTenant.contractId;
      const targetContract = contractId
        ? data.contracts.find(c => c.id === contractId)
        : data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice' || c.status === 'moving_out'));
      if (!targetContract) {
        alert('Không tìm thấy hợp đồng phù hợp.');
        return;
      }
      const renewalRows = [
        ...(data.contractRenewals || []),
        ...(targetContract.renewalHistory || [])
      ]
        .filter(r => r.contractId === targetContract.id)
        .filter((r, idx, arr) => arr.findIndex(x => x.id === r.id) === idx)
        .sort((a, b) => new Date(b.createdAt || b.signedDate || 0) - new Date(a.createdAt || a.signedDate || 0));
      const renewal = renewalRows[0];
      if (!renewal) {
        alert('Hợp đồng này chưa có phụ lục gia hạn để in.');
        return;
      }
      const oldRent = renewal.oldRent ?? targetContract.rent;
      const oldDeposit = renewal.oldDeposit ?? targetContract.deposit;
      const newRent = renewal.newRent ?? targetContract.rent;
      const newDeposit = renewal.newDeposit ?? targetContract.deposit;
      const oldEndDate = renewal.oldEndDate || targetContract.previousEndDate || targetContract.endDate;
      const appendixContract = {
        ...targetContract,
        endDate: oldEndDate,
        rent: oldRent,
        deposit: oldDeposit
      };
      const form = {
        signedDate: formatDateForInput(renewal.signedDate || renewal.createdAt || new Date().toISOString()),
        newStartDate: renewal.newStartDate || (isValidBusinessDate(oldEndDate) ? addDaysToDate(oldEndDate, 1) : ''),
        newEndDate: renewal.newEndDate || targetContract.endDate,
        keepPricing: Number(newRent || 0) === Number(oldRent || 0) && Number(newDeposit || 0) === Number(oldDeposit || 0),
        newRent,
        newDeposit,
        note: renewal.note || ''
      };
      setViewingAppendix({ contract: appendixContract, form });
    } else if (type === 'transfer_room') {
      const roomId = roomOrTenant.id || roomOrTenant.roomId;
      const activeContract = data.contracts.find(c => c.roomId === roomId && (c.status === 'active' || c.status === 'notice'));
      if (activeContract && window.confirm(`Đổi phòng cho P${roomId}?\n\nHệ thống sẽ kết thúc hợp đồng phòng cũ, tạo hợp đồng phòng mới và ghi lịch sử chuyển phòng.`)) setTransferringContract(activeContract);
    } else if (type === 'add_asset') {
      setRoomOpsModal({ type: 'asset', room: roomOrTenant });
    } else if (type === 'add_maintenance') {
      setRoomOpsModal({ type: 'maintenance', room: roomOrTenant });
    } else if (type === 'add_file') {
      setRoomOpsModal({ type: 'file', room: roomOrTenant });
    } else if (type === 'record_meter') {
      setRoomOpsModal({ type: 'meter', room: roomOrTenant });
    } else if (type === 'view_tenants') {
      setQuery(`P${roomOrTenant.id}`);
      setTab('tenants');
    } else if (type === 'view_room') {
      const room = data.rooms.find(r => r.id === roomOrTenant.id || r.id === roomOrTenant.roomId);
      if (room) setSelectedRoom(room);
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
        const stats = getDashboardStats(data, getCurrentMonthLabel());

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
            'Mã vân tay': t.fingerprintCode || '',
            'Trạng thái vân tay': t.fingerprintStatus || '',
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
          const tenant = getTenantForReceipt(data, r);
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
          const tenant = getTenantForReceipt(data, r);
          return {
            'Mã phiếu': r.id,
            'Loại phiếu': receiptTypeLabel(r.type),
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
          'Điện đầu': rep.electricOld,
          'Điện chốt': rep.electricNew,
          'Tiền điện': rep.electricAmount,
          'Nước đầu': rep.waterOld,
          'Nước chốt': rep.waterNew,
          'Tiền nước': rep.waterAmount,
          'Số ngày tính tiền phòng': rep.roomChargeDays || '',
          'Tiền phòng phát sinh': rep.proratedRent || 0,
          'Dịch vụ phát sinh': rep.proratedFixedServices || 0,
          'Tiền phòng/phí còn nợ khác': rep.unpaidRent,
          'Phí hư hỏng': rep.damageFee || 0,
          'Phí vệ sinh': rep.cleaningFee || 0,
          'Phí khác': rep.otherFee || 0,
          'Tiền cọc đối trừ': rep.depositUsed,
          'Tổng phát sinh': rep.totalIncurred,
          'Khách còn phải trả': rep.mustCollect,
          'Cần hoàn cọc': rep.mustRefund,
          'Cần xóa vân tay': (rep.accessRemovalTasks || []).filter(t => t.status === 'Cần xóa').map(t => `${t.name} (${t.fingerprintCode})`).join(', '),
          'Biển số xe cần kiểm tra': (rep.accessRemovalTasks || []).map(t => t.licensePlate).filter(Boolean).join(', '),
          'Ghi chú': rep.note || '',
          'Ngày tạo': rep.createdAt ? new Date(rep.createdAt).toLocaleDateString('vi-VN') : ''
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(moveOutData), "Lich su tra phong");

        // 9. Sheet Lich su nguoi thue
        const pastTenantsData = data.tenants.filter(t => (data.memberships || []).some(m => m.tenantId === t.id && (m.status === 'ended' || m.status === 'moved_out'))).map(t => {
          const m = data.memberships.find(m => m.tenantId === t.id && (m.status === 'ended' || m.status === 'moved_out'));
          return {
            'Mã người thuê': t.id,
            'Họ tên': t.name,
            'SĐT': t.phone,
            'CCCD': t.cccd,
            'Biển số xe': t.licensePlate || '',
            'Mã vân tay': t.fingerprintCode || '',
            'Trạng thái vân tay': t.fingerprintStatus || '',
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

        // 10. Sheet Lich su doi phong
        const transferData = (data.roomTransfers || []).map(tr => ({
          'Mã chuyển phòng': tr.id,
          'Ngày chuyển': tr.transferDate || '',
          'Người đứng tên': (data.tenants || []).find(t => t.id === tr.tenantId)?.name || '',
          'Phòng cũ': tr.oldRoomId,
          'Phòng mới': tr.newRoomId,
          'HĐ cũ': tr.oldContractId,
          'HĐ mới': tr.newContractId,
          'Giá thuê cũ': tr.oldRent,
          'Giá thuê mới': tr.newRent,
          'Cọc cũ': tr.oldDeposit,
          'Cọc mới': tr.newDeposit,
          'Ghi chú': tr.note || '',
          'Ngày tạo': tr.createdAt ? new Date(tr.createdAt).toLocaleDateString('vi-VN') : ''
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(transferData), "Lich su doi phong");

        // 11. Sheet Cau hinh
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
      const month = getCurrentMonthLabel();
      const transferTarget = findTransferTargetForOldRoom(data, roomId, month);
      const activeContract = getCurrentContractForRoom(data, roomId);
      const targetContract = transferTarget?.contract || activeContract;
      const targetRoom = transferTarget?.room || roomOrTenant;
      if (!targetContract) return alert('Phòng trống hoặc không có hợp đồng cần lập phiếu.');
      const exists = (data.receipts || []).find(r => r.roomId === targetRoom.id && r.contractId === targetContract.id && r.month === month && r.type === 'monthly');
      if (exists) {
        const tenant = getPrimaryTenantByContract(data, targetContract.id);
        if (!window.confirm(`Hợp đồng ${targetContract.contractNo || targetContract.id} của ${tenant?.name || `P${targetRoom.id}`} đã có phiếu tháng ${month}. Bạn muốn làm mới/ghi đè phiếu của chính hợp đồng này?`)) return;
      }
      const prev = getPreviousReceiptByRoom(data.receipts, targetRoom.id, month, {
        includeSameMonth: true,
        excludeReceiptId: exists?.id,
        excludeSameMonthContractId: targetContract.id
      });
      const billingContext = getMonthlyBillingContext(data, targetContract, month);
      if (billingContext.mode === 'transfer_old_room_skip') return alert(`Phòng ${roomId} đã chuyển sang P${billingContext.transfer?.newRoomId}. Hãy lập phiếu gộp tại phòng mới.`);
      const newRec = createMonthlyReceipt(targetRoom, targetContract, prev, month, billingContext);
      setData(old => ({
        ...old,
        receipts: [...(old.receipts || []).filter(r => r.id !== exists?.id), newRec]
      }));
      alert(`Đã tạo phiếu tháng ${month} cho ${transferTarget ? `P${transferTarget.transfer.oldRoomId} → P${targetRoom.id}` : `P${targetRoom.id}`}`);
      setTab('receipts');
    } else if (type === 'delete_receipt') {
      if (window.confirm('Bạn có chắc chắn muốn xóa phiếu thu này?')) {
        setData(old => ({ ...old, receipts: old.receipts.filter(r => r.id !== roomOrTenant.id) }));
      }
    }
  }

  function updateContractStatus(roomId, status) {
    setData(old => ({
      ...old,
      contracts: old.contracts.map(c => {
        if (c.roomId !== roomId || !(c.status === 'active' || c.status === 'notice')) return c;
        return {
          ...c,
          status,
          noticeDate: status === 'notice' ? (c.noticeDate || new Date().toISOString().slice(0, 10)) : '',
          expectedMoveOutDate: status === 'notice' ? c.expectedMoveOutDate : ''
        };
      })
    }));
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
          <button className={`nav-item ${tab === 'expenses' ? 'active' : ''}`} onClick={() => setTab('expenses')}>💸 Chi phí</button>
          <button className={`nav-item ${tab === 'financial_report' ? 'active' : ''}`} onClick={() => setTab('financial_report')}>📊 Báo cáo thu chi</button>
          <button className={`nav-item ${tab === 'settings' ? 'active' : ''}`} onClick={() => setTab('settings')}>⚙️ Cài đặt</button>
        </aside>
        <div className="content content-fade">
          {tab === 'dashboard' && (
            <Dashboard 
              data={data} 
              onRoomClick={(id) => { setTab('rooms'); setQuery(id); }} 
              onAction={handleAction} 
              isSyncing={isSyncing} 
              lastSynced={lastSynced} 
            />
          )}
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
              onDeleteReceipt={(receipt) => {
                const message = receipt.isFinalized
                  ? `Phiếu P${receipt.roomId} tháng ${receipt.month} đã lưu chính thức. Vẫn hủy phiếu này?`
                  : `Hủy phiếu P${receipt.roomId} tháng ${receipt.month}?`;
                if (window.confirm(message)) {
                  setData(old => ({ ...old, receipts: (old.receipts || []).filter(r => r.id !== receipt.id) }));
                }
              }}
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
          {tab === 'expenses' && (
            <ExpensesTab 
              data={data} 
              onAction={(type, item) => {
                if (type === 'create_expense') setAddingExpense(true);
                else if (type === 'edit_expense') setEditingExpense(item);
                else if (type === 'delete_expense') {
                  if (window.confirm('Xóa phiếu chi này?')) {
                    setData(old => ({ ...old, expensePayments: old.expensePayments.filter(e => e.id !== item.id) }));
                  }
                } else if (type === 'view_qr') setViewingExpense(item);
                else if (type === 'manage_suppliers') setTab('suppliers');
              }} 
            />
          )}
          {tab === 'suppliers' && (
            <SuppliersTab 
              data={data} 
              onAction={(type, item) => {
                if (type === 'add_supplier') setAddingSupplier(true);
                else if (type === 'edit_supplier') setEditingSupplier(item);
                else if (type === 'delete_supplier') {
                  if (window.confirm('Xóa nhà cung cấp này?')) {
                    setData(old => ({ ...old, suppliers: old.suppliers.filter(s => s.id !== item.id) }));
                  }
                } else if (type === 'back') setTab('expenses');
              }}
            />
          )}
          {tab === 'financial_report' && (
            <FinancialReportTab data={data} onAction={(type, arg) => {
              if (type === 'view_receipt') {
                setViewingReceipt(arg);
              } else if (type === 'view_expense') {
                setViewingExpense(arg);
              }
            }} />
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
            if (!['add_roommate', 'edit_tenant'].includes(type)) setSelectedRoom(null);
          }}
          onAddRoommate={(result) => {
            const { tenant, membership } = result;
            setData(old => ({
              ...old,
              tenants: [...(old.tenants || []).filter(t => t.id !== tenant.id), tenant],
              memberships: [...(old.memberships || []).filter(m => m.id !== membership.id), membership]
            }));
            alert(`Đã thêm ${tenant.name} vào phòng ${membership.roomId}.`);
            return true;
          }}
        />
      )}
      {newRentalRoom && (
        <RentalFlowModal 
          room={newRentalRoom} 
          onClose={() => setNewRentalRoom(null)} 
          onSave={(result) => {
            const { tenant, contract, memberships } = result;
            const existingActiveContract = (data.contracts || []).find(c => c.roomId === contract.roomId && ['active', 'notice', 'moving_out'].includes(c.status));
            if (existingActiveContract) {
              alert(`Phòng ${contract.roomId} đã có hợp đồng hiệu lực. Hệ thống không tạo hợp đồng trùng.`);
              return false;
            }
            setData(old => {
              const serviceConfig = contract.terms?.services || {};
              const updatedRooms = (old.rooms || []).map(r => r.id === contract.roomId ? {
                ...r,
                rent: Number(contract.rent || r.rent || 0),
                deposit: Number(contract.deposit || r.deposit || 0),
                electricPrice: Number(contract.terms?.electricPrice || r.electricPrice || 3800),
                waterPrice: Number(contract.terms?.waterPrice || r.waterPrice || 32000),
                cleaning: Number(serviceConfig.cleaning ?? r.cleaning ?? 0),
                elevator: Number(serviceConfig.elevator ?? r.elevator ?? 0),
                laundry: Number(serviceConfig.laundry ?? r.laundry ?? 0),
                internet: Number(serviceConfig.internet ?? r.internet ?? 0),
              } : r);
              return {
                ...old,
                rooms: updatedRooms,
                tenants: [...(old.tenants || []).filter(t => t.id !== tenant.id), tenant],
                contracts: [...(old.contracts || []).filter(c => c.id !== contract.id), contract],
                memberships: [...(old.memberships || []).filter(m => !memberships.some(newM => newM.id === m.id)), ...memberships]
              };
            });
            setNewRentalRoom(null);
            setSelectedRoom(null);
            setTab('rooms');
            alert(`Đã lưu hợp đồng ${contract.contractNo || contract.id} cho phòng ${contract.roomId}.`);
            return true;
          }}
        />
      )}
      {addingRoommate && (
        <RoommateModal
          room={addingRoommate.room}
          contract={addingRoommate.contract}
          onClose={() => setAddingRoommate(null)}
          onSave={(result) => {
            const { tenant, membership } = result;
            setData(old => ({
              ...old,
              tenants: [...(old.tenants || []).filter(t => t.id !== tenant.id), tenant],
              memberships: [...(old.memberships || []).filter(m => m.id !== membership.id), membership]
            }));
            setAddingRoommate(null);
            alert(`Đã thêm ${tenant.name} vào phòng ${membership.roomId}.`);
            return true;
          }}
        />
      )}
      {settlingRoom && (
        <SettlementModal room={settlingRoom.room} contract={settlingRoom.contract} data={data} bankInfo={bankInfo} onClose={() => setSettlingRoom(null)} onSave={(report) => {
            const affectedTenants = (data.memberships || [])
              .filter(m => m.contractId === report.contractId)
              .map(m => (data.tenants || []).find(t => t.id === m.tenantId))
              .filter(Boolean);
            setData(old => {
              const updatedRooms = old.rooms.map(r => r.id === report.roomId ? {
                ...r,
                electricOld: report.electricOld,
                electricNew: report.electricNew,
                waterOld: report.waterOld,
                waterNew: report.waterNew
              } : r);
              const updatedContracts = old.contracts.map(c => c.id === report.contractId ? { ...c, status: 'ended', actualEndDate: report.actualEndDate, endedAt: new Date().toISOString() } : c);
              const updatedMemberships = old.memberships.map(m => m.contractId === report.contractId ? { ...m, status: 'ended', leftDate: report.actualEndDate } : m);
              const affectedTenantIds = old.memberships.filter(m => m.contractId === report.contractId).map(m => m.tenantId);
              const updatedTenants = old.tenants.map(t => affectedTenantIds.includes(t.id) ? { ...t, status: 'moved_out', lastRoomId: report.roomId, fingerprintStatus: t.fingerprintCode ? 'Cần xóa' : (t.fingerprintStatus || 'Chưa đăng ký') } : t);
              const accessRemovalTasks = old.tenants
                .filter(t => affectedTenantIds.includes(t.id))
                .map(t => ({
                  tenantId: t.id,
                  name: t.name,
                  phone: t.phone,
                  fingerprintCode: t.fingerprintCode || '',
                  licensePlate: t.licensePlate || t.vehicle || '',
                  action: t.fingerprintCode ? 'Xóa vân tay khỏi máy chấm/khóa cửa' : 'Không có vân tay đã lưu',
                  status: t.fingerprintCode ? 'Cần xóa' : 'Không cần xử lý'
                }));
              
              // Tạo phiếu thu chốt trả phòng nếu khách còn phải thanh toán thêm
              let newReceipts = [...(old.receipts || [])];
              if (report.mustCollect > 0) {
                newReceipts.push({
                  id: uid('receipt'),
                  roomId: report.roomId,
                  contractId: report.contractId,
                  type: 'move_out_settlement',
                  month: report.actualEndDate.split('-').slice(0, 2).reverse().join('/'),
                  rent: report.proratedRent || 0,
                  fixedServices: report.proratedFixedServices || 0,
                  electricOld: report.electricOld,
                  electricNew: report.electricNew,
                  electricUsed: report.electricUsed,
                  electricAmount: report.electricAmount,
                  waterOld: report.waterOld,
                  waterNew: report.waterNew,
                  waterUsed: report.waterUsed,
                  waterAmount: report.waterAmount,
                  other: Number(report.unpaidRent || 0) + Number(report.cleaningFee || 0) + Number(report.damageFee || 0) + Number(report.otherFee || 0),
                  total: report.mustCollect,
                  paidAmount: 0,
                  debt: report.mustCollect,
                  status: 'Chưa thanh toán',
                  note: 'Phiếu chốt trả phòng',
                  createdAt: new Date().toISOString()
                });
              }

              let newExpensePayments = [...(old.expensePayments || [])];
              if (report.mustRefund > 0) {
                const [m, y] = report.actualEndDate.split('-').slice(0, 2).reverse();
                const prefix = `PC-${y}${m}`;
                const count = newExpensePayments.filter(e => e.expenseCode && e.expenseCode.startsWith(prefix)).length + 1;
                const expenseCode = `${prefix}-${String(count).padStart(3, '0')}`;
                newExpensePayments.push({
                  id: uid('exp'),
                  expenseCode,
                  type: 'deposit_refund',
                  source: 'move_out_settlement',
                  sourceReportId: report.id,
                  roomId: report.roomId,
                  contractId: report.contractId,
                  tenantId: report.tenantId,
                  supplierId: '',
                  recipientName: report.tenantName || 'Khách thuê',
                  recipientPhone: report.tenantPhone || '',
                  recipientBankName: report.refundBankName || '',
                  recipientBankAccount: report.refundBankAccount || '',
                  recipientBankOwner: report.refundBankOwner || report.tenantName || '',
                  recipientQrImageUrl: report.refundQrImageUrl || '',
                  categoryId: 'cat_other',
                  month: `${m}/${y}`,
                  paymentDate: report.actualEndDate,
                  title: `Hoàn cọc trả phòng P${report.roomId}`,
                  description: `Hoàn tiền cọc cho ${report.tenantName || 'khách thuê'} sau khi tất toán phòng ${report.roomId}.`,
                  totalAmount: report.mustRefund,
                  amount: report.mustRefund,
                  paidAmount: report.mustRefund,
                  status: 'paid',
                  paymentMethod: report.refundQrImageUrl || report.refundBankAccount ? 'transfer' : 'cash',
                  note: report.note || 'Tự tạo từ luồng hoàn tất trả phòng',
                  createdAt: new Date().toISOString()
                });
              }
              
              return { ...old, rooms: updatedRooms, contracts: updatedContracts, memberships: updatedMemberships, tenants: updatedTenants, receipts: newReceipts, expensePayments: newExpensePayments, moveOutReports: [{ ...report, accessRemovalTasks }, ...(old.moveOutReports || [])] };
            });
            const completionMessages = [];
            if (report.mustRefund > 0) {
              completionMessages.push(`Đã tạo phiếu chi hoàn cọc ${formatMoney(report.mustRefund)} cho ${report.tenantName || 'khách thuê'}.`);
            }
            const fingerprintNames = affectedTenants.filter(t => t.fingerprintCode).map(t => `${t.name} (${t.fingerprintCode})`);
            if (fingerprintNames.length > 0) {
              completionMessages.push(`Nhắc admin: cần xóa vân tay cho khách đã trả phòng:\n${fingerprintNames.join('\n')}`);
            }
            if (completionMessages.length > 0) alert(completionMessages.join('\n\n'));
            setSettlingRoom(null);
          }}
        />
      )}
      {viewingContract && <ContractPreview
        {...viewingContract}
        tenants={(data.memberships || [])
          .filter(m => m.contractId === viewingContract.contract.id)
          .map(m => {
            const tenant = (data.tenants || []).find(t => t.id === m.tenantId);
            return tenant ? { ...tenant, role: m.role, membershipStatus: m.status } : null;
          })
          .filter(Boolean)
          .sort((a, b) => (b.role === 'primary' ? 1 : 0) - (a.role === 'primary' ? 1 : 0))}
        bankInfo={bankInfo}
        onClose={() => setViewingContract(null)}
      />}
      {viewingAppendix && (
        <RenewalAppendixPreviewModal
          contract={viewingAppendix.contract}
          form={viewingAppendix.form}
          data={data}
          onClose={() => setViewingAppendix(null)}
        />
      )}
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

              const oldRent = Number(renewingContract.rent || 0);
              const newRent = Number(renewalRecord.newRent || 0);
              const oldDeposit = Number(renewingContract.deposit || 0);
              const newDeposit = Number(renewalRecord.newDeposit || 0);
              const rentDelta = newRent - oldRent;
              const depositDelta = newDeposit - oldDeposit;
              const rentIncrease = Math.max(0, rentDelta);
              const depositIncrease = Math.max(0, depositDelta);
              const rentRefund = Math.max(0, -rentDelta);
              const depositRefund = Math.max(0, -depositDelta);
              const amountToCollect = rentIncrease + depositIncrease;
              const amountToRefund = rentRefund + depositRefund;
              const renewalMonth = monthFromDate(form.newStartDate) || getCurrentMonthLabel();
              const primaryTenant = getPrimaryTenantByContract(old, renewingContract.id) || {};
              const adjustmentDetails = [
                rentDelta !== 0 ? `Giá thuê: ${formatMoney(oldRent)} → ${formatMoney(newRent)} (${rentDelta > 0 ? 'tăng' : 'giảm'} ${formatMoney(Math.abs(rentDelta))})` : '',
                depositDelta !== 0 ? `Tiền cọc: ${formatMoney(oldDeposit)} → ${formatMoney(newDeposit)} (${depositDelta > 0 ? 'tăng' : 'giảm'} ${formatMoney(Math.abs(depositDelta))})` : '',
                `Áp dụng từ ${formatDisplayDate(form.newStartDate)}`
              ].filter(Boolean).join('. ');

              const updatedReceipts = [...(old.receipts || [])];
              if (amountToCollect > 0) {
                updatedReceipts.unshift({
                  id: uid('receipt'),
                  roomId: renewingContract.roomId,
                  contractId: renewingContract.id,
                  type: 'renewal_adjustment',
                  month: renewalMonth,
                  rent: rentIncrease,
                  fixedServices: 0,
                  electricOld: 0,
                  electricNew: 0,
                  electricUsed: 0,
                  electricAmount: 0,
                  waterOld: 0,
                  waterNew: 0,
                  waterUsed: 0,
                  waterAmount: 0,
                  other: depositIncrease,
                  total: amountToCollect,
                  paidAmount: 0,
                  debt: amountToCollect,
                  status: 'Chưa thanh toán',
                  note: `Điều chỉnh gia hạn. ${adjustmentDetails}`,
                  createdAt: new Date().toISOString()
                });
              }

              const updatedExpenses = [...(old.expensePayments || [])];
              if (amountToRefund > 0) {
                const expensePrefix = `PC-${onlyDigits(renewalMonth)}-GH`;
                const expenseCount = updatedExpenses.filter(e => String(e.expenseCode || '').startsWith(expensePrefix)).length + 1;
                updatedExpenses.unshift({
                  id: uid('exp'),
                  expenseCode: `${expensePrefix}-${String(expenseCount).padStart(3, '0')}`,
                  recipientName: primaryTenant.name || 'Khách thuê',
                  supplierId: '',
                  categoryId: 'cat_other',
                  month: renewalMonth,
                  paymentDate: form.newStartDate,
                  title: `Hoàn chênh lệch gia hạn P${renewingContract.roomId}`,
                  description: adjustmentDetails,
                  totalAmount: amountToRefund,
                  paidAmount: 0,
                  status: 'unpaid',
                  paymentMethod: 'transfer',
                  note: form.note || 'Tự tạo khi giảm giá thuê hoặc tiền cọc lúc gia hạn',
                  createdAt: new Date().toISOString(),
                  updatedAt: new Date().toISOString()
                });
              }

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
                receipts: updatedReceipts,
                expensePayments: updatedExpenses,
                contractRenewals: [renewalRecord, ...(old.contractRenewals || [])] 
              };
            });
            setRenewingContract(null);
          }}
        />
      )}
      {transferringContract && (
        <TransferRoomModal
          contract={transferringContract}
          data={data}
          onClose={() => setTransferringContract(null)}
          onSave={(form) => {
            setData(old => {
              const targetRoom = old.rooms.find(r => r.id === form.newRoomId);
              if (!targetRoom) return old;

              const oldContract = transferringContract;
              const primaryMembership = (old.memberships || []).find(m => m.contractId === oldContract.id && m.role === 'primary');
              const activeMembers = (old.memberships || []).filter(m => m.contractId === oldContract.id && m.status === 'active');
              const newContractId = uid('contract');
              const transferDate = form.transferDate;
              const oldStayTo = form.oldRoomStayTo || formatDateInputValue(addDays(transferDate, -1)) || transferDate;
              const transferRecord = {
                id: uid('transfer'),
                tenantId: primaryMembership?.tenantId || '',
                oldContractId: oldContract.id,
                newContractId,
                oldRoomId: oldContract.roomId,
                newRoomId: targetRoom.id,
                transferDate,
                oldBillingMonth: monthFromDate(oldStayTo || transferDate),
                oldRoomStayFrom: form.oldRoomStayFrom || oldContract.startDate || '',
                oldRoomStayTo: oldStayTo,
                oldRoomServiceMode: form.oldRoomServiceMode || 'daily',
                occupantCount: activeMembers.length || 1,
                oldRent: Number(oldContract.rent || 0),
                newRent: Number(form.rent || 0),
                oldDeposit: Number(oldContract.deposit || 0),
                newDeposit: Number(form.deposit || 0),
                note: form.note,
                createdAt: new Date().toISOString()
              };

              const newContract = {
                ...oldContract,
                id: newContractId,
                roomId: targetRoom.id,
                contractNo: form.contractNo,
                signedDate: transferDate,
                startDate: transferDate,
                endDate: form.endDate,
                rent: Number(form.rent || 0),
                deposit: Number(form.deposit || 0),
                status: 'active',
                noticeDate: '',
                expectedMoveOutDate: '',
                actualEndDate: '',
                endedAt: '',
                previousEndDate: '',
                renewedAt: '',
                renewalHistory: [],
                note: form.note,
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString()
              };

              const updatedContracts = [
                ...old.contracts.map(c => c.id === oldContract.id ? {
                  ...c,
                  status: 'ended',
                  actualEndDate: transferDate,
                  endedAt: new Date().toISOString(),
                  note: [c.note, `Chuyển sang phòng ${targetRoom.id} ngày ${transferDate}`].filter(Boolean).join('\n')
                } : c),
                newContract
              ];

              const endedMemberships = old.memberships.map(m => m.contractId === oldContract.id ? { ...m, status: 'ended', leftDate: transferDate } : m);
              const newMemberships = activeMembers.map(m => ({
                ...m,
                id: uid('membership'),
                contractId: newContractId,
                roomId: targetRoom.id,
                status: 'active',
                joinedDate: transferDate,
                leftDate: ''
              }));

              return {
                ...old,
                contracts: updatedContracts,
                memberships: [...endedMemberships, ...newMemberships],
                roomTransfers: [transferRecord, ...(old.roomTransfers || [])]
              };
            });
            setTransferringContract(null);
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
              const { membershipId, contractId, roomId, role, ...tenantPatch } = updated;
              const updatedTenants = old.tenants.map(t => t.id === updated.id ? { ...tenantPatch, role } : t);
              const targetMembership = (old.memberships || []).find(m => m.id === updated.membershipId)
                || (old.memberships || []).find(m => m.tenantId === updated.id && m.status === 'active' && (!updated.contractId || m.contractId === updated.contractId))
                || (old.memberships || []).find(m => m.tenantId === updated.id && m.status === 'active');
              const targetContractId = targetMembership?.contractId || updated.contractId;
              const updatedMemberships = old.memberships.map(m => {
                if (updated.role === 'primary' && targetContractId && m.contractId === targetContractId && m.status === 'active') {
                  return { ...m, role: m.id === targetMembership?.id ? 'primary' : 'secondary' };
                }
                if (m.id === targetMembership?.id || (m.tenantId === updated.id && m.status === 'active' && (!targetContractId || m.contractId === targetContractId))) {
                  return { ...m, role: updated.role || m.role };
                }
                return m;
              });
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
              contracts: old.contracts.map(c => c.id === updated.id ? {
                ...updated,
                rent: Number(updated.rent),
                deposit: Number(updated.deposit),
                paymentCycleDay: updated.paymentCycleDay ? Number(updated.paymentCycleDay) : updated.paymentCycleDay,
                updatedAt: new Date().toISOString()
              } : c)
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
      {roomOpsModal && (
        <RoomOpsModal
          mode={roomOpsModal.type}
          room={roomOpsModal.room}
          onClose={() => setRoomOpsModal(null)}
          onSave={(payload) => {
            setData(old => ({
              ...old,
              rooms: (old.rooms || []).map(r => {
                if (r.id !== roomOpsModal.room.id) return r;
                if (roomOpsModal.type === 'asset') {
                  return { ...r, assets: [payload, ...(r.assets || [])] };
                }
                if (roomOpsModal.type === 'maintenance') {
                  return { ...r, maintenanceTickets: [payload, ...(r.maintenanceTickets || [])] };
                }
                if (roomOpsModal.type === 'file') {
                  return { ...r, attachments: [payload, ...(r.attachments || [])] };
                }
                return {
                  ...r,
                  electricOld: r.electricNew ?? r.electricOld ?? r.electricStart ?? 0,
                  electricNew: payload.electricNew,
                  waterOld: r.waterNew ?? r.waterOld ?? r.waterStart ?? 0,
                  waterNew: payload.waterNew,
                  lastMeterReading: payload
                };
              })
            }));
            setRoomOpsModal(null);
          }}
        />
      )}
      
      {(addingExpense || editingExpense) && (
        <ExpenseModal 
          expense={editingExpense} 
          data={data} 
          onClose={() => { setAddingExpense(false); setEditingExpense(null); }} 
          onSave={(updated) => {
            setData(old => {
              const list = old.expensePayments || [];
              const exists = list.find(e => e.id === updated.id);
              return { ...old, expensePayments: exists ? list.map(e => e.id === updated.id ? updated : e) : [...list, updated] };
            });
            setAddingExpense(false); setEditingExpense(null);
          }} 
        />
      )}

      {(addingSupplier || editingSupplier) && (
        <SupplierModal 
          supplier={editingSupplier} 
          onClose={() => { setAddingSupplier(false); setEditingSupplier(null); }} 
          onSave={(updated) => {
            setData(old => {
              const list = old.suppliers || [];
              const exists = list.find(s => s.id === updated.id);
              return { ...old, suppliers: exists ? list.map(s => s.id === updated.id ? updated : s) : [...list, updated] };
            });
            setAddingSupplier(false); setEditingSupplier(null);
          }} 
        />
      )}

      {viewingExpense && (
        <ExpenseQRModal 
          expense={viewingExpense} 
          supplier={data.suppliers.find(s => s.id === viewingExpense.supplierId)}
          onClose={() => setViewingExpense(null)} 
        />
      )}

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
                  localStorage.setItem(STORAGE_KEY, JSON.stringify(imported));
                  setCloudEnabled(true);
                  setHasLoadedCloud(true);
                  setIsSyncing(true);
                  fetch('/api/data', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ type: 'full_sync', payload: imported })
                  })
                    .then(async res => {
                      setIsSyncing(false);
                      if (!res.ok) {
                        const errData = await res.json().catch(() => ({}));
                        console.info('Import JSON saved locally but cloud sync failed.', errData);
                        const message = String(errData.details || errData.error || res.status).slice(0, 260);
                        alert(`Nhập JSON thành công trên máy này, nhưng chưa đồng bộ cloud: ${message}`);
                        return;
                      }
                      cloudFailureRef.current = 0;
                      setLastSynced(new Date());
                      alert('Nhập JSON thành công và đã đồng bộ lên cloud!');
                    })
                    .catch(err => {
                      setIsSyncing(false);
                      console.info('Import JSON saved locally but cloud sync failed.', err?.message || err);
                      alert(`Nhập JSON thành công trên máy này, nhưng chưa đồng bộ cloud: ${err?.message || 'Lỗi kết nối'}`);
                    });
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
          <thead><tr><th>Ngày trả</th><th>Phòng</th><th>Người đứng tên</th><th>Ngày bắt đầu</th><th>Thời hạn gốc</th><th>Tiền cọc</th><th>Ra vào</th><th>Thao tác</th></tr></thead>
          <tbody>
            {endedContracts.map(c => {
              const tenant = getPrimaryTenantByContract(data, c.id) || { name: 'N/A' };
              const report = (data.moveOutReports || []).find(r => r.contractId === c.id);
              return (
                <tr key={c.id}>
                  <td><b>{c.actualEndDate?.split('-').reverse().join('/') || 'N/A'}</b></td>
                  <td>P{c.roomId}</td>
                  <td>{tenant.name}</td>
                  <td>{formatBusinessDate(c.startDate)}</td>
                  <td>{formatBusinessDate(c.endDate)}</td>
                  <td>{formatMoney(c.deposit)}</td>
                  <td>{(report?.accessRemovalTasks || []).some(t => t.status === 'Cần xóa') ? <span className="status-badge-liquid debt">Cần xóa vân tay</span> : <span className="status-badge-liquid active">Đã kiểm tra</span>}</td>
                  <td><button className="secondary-btn sm" onClick={() => onAction('view_contract', { roomId: c.roomId, contractId: c.id })}>Xem HĐ</button></td>
                </tr>
              );
            })}
            {endedContracts.length === 0 && <tr><td colSpan="8" style={{ textAlign: 'center', padding: '40px' }}>Chưa có lịch sử hợp đồng nào kết thúc.</td></tr>}
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
          <label style={{ margin: 0 }}>Loại phiếu <select value={filter.type} onChange={e => setFilter({...filter, type: e.target.value})}><option value="all">Tất cả</option><option value="monthly">Phiếu tháng</option><option value="renewal_adjustment">Điều chỉnh gia hạn</option><option value="move_out_settlement">Chốt trả phòng</option></select></label>
          <button className="secondary-btn" onClick={() => setFilter({ roomId: '', status: 'all', type: 'all', month: '' })}>🔄 Reset</button>
        </div>
      </div>
      <div className="widget liquid-glass" style={{ padding: 0 }}>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Ngày tạo</th><th>Phòng</th><th>Người đứng tên</th><th>Loại</th><th>Tổng tiền</th><th>Đã trả</th><th>Trạng thái</th><th>Hành động</th></tr></thead>
            <tbody>
              {filteredReceipts.map(r => {
                const tenant = getTenantForReceipt(data, r) || { name: 'N/A' };
                const paymentState = getReceiptPaymentState(r);
                return (
                  <tr key={r.id}>
                    <td>{new Date(r.createdAt).toLocaleDateString('vi-VN')}</td>
                    <td><b>P{r.roomId}</b></td>
                    <td>{tenant.name}</td>
                    <td><span style={{ fontSize: '12px' }}>{receiptTypeLabel(r.type, true)}</span></td>
                    <td style={{ fontWeight: '700' }}>{formatMoney(r.total)}</td>
                    <td>{formatMoney(paymentState.paidAmount)}</td>
                    <td><span className={`status-badge-liquid ${paymentState.status === 'Đã thanh toán' ? 'active' : paymentState.status === 'Nợ một phần' ? 'notice' : 'debt'}`}>{paymentState.status}</span></td>
                    <td>
                      <div style={{ display: 'flex', gap: '4px' }}>
                        {paymentState.isPaid
                          ? <button className="secondary-btn sm paid-action" disabled>✓ Đã thu</button>
                          : <button className="primary-btn sm" onClick={() => onPay(r)}>💸 Thu</button>}
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

function Dashboard({ data, onRoomClick, onAction, isSyncing, lastSynced }) {
  const today = new Date();
  const [periodMode, setPeriodMode] = useState('month');
  const [periodMonth, setPeriodMonth] = useState(`${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`);
  const [periodQuarter, setPeriodQuarter] = useState(`${today.getFullYear()}-Q${Math.floor(today.getMonth() / 3) + 1}`);
  const [periodYear, setPeriodYear] = useState(String(today.getFullYear()));
  const [customRange, setCustomRange] = useState({
    start: `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`,
    end: today.toISOString().slice(0, 10)
  });

  const period = useMemo(() => {
    if (periodMode === 'month') {
      const [year, month] = periodMonth.split('-').map(Number);
      return { start: new Date(year, month - 1, 1), end: new Date(year, month, 0), label: `Tháng ${String(month).padStart(2, '0')}/${year}` };
    }
    if (periodMode === 'quarter') {
      const [yearText, qText] = periodQuarter.split('-Q');
      const year = Number(yearText);
      const quarter = Number(qText);
      const startMonth = (quarter - 1) * 3;
      return { start: new Date(year, startMonth, 1), end: new Date(year, startMonth + 3, 0), label: `Quý ${quarter}/${year}` };
    }
    if (periodMode === 'year') {
      const year = Number(periodYear);
      return { start: new Date(year, 0, 1), end: new Date(year, 11, 31), label: `Năm ${year}` };
    }
    return { start: parseDateFlexible(customRange.start) || new Date(), end: parseDateFlexible(customRange.end) || new Date(), label: `${formatDisplayDate(customRange.start)} - ${formatDisplayDate(customRange.end)}` };
  }, [periodMode, periodMonth, periodQuarter, periodYear, customRange]);

  const inPeriod = (value) => {
    const date = parseDateFlexible(value) || new Date(value || 0);
    if (!date || Number.isNaN(date.getTime())) return false;
    date.setHours(0, 0, 0, 0);
    const start = new Date(period.start); start.setHours(0, 0, 0, 0);
    const end = new Date(period.end); end.setHours(23, 59, 59, 999);
    return date >= start && date <= end;
  };

  const monthKey = (date) => `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
  const periodMonthKeys = [];
  const cursor = new Date(period.start.getFullYear(), period.start.getMonth(), 1);
  const endCursor = new Date(period.end.getFullYear(), period.end.getMonth(), 1);
  while (cursor <= endCursor) {
    periodMonthKeys.push(monthKey(cursor));
    cursor.setMonth(cursor.getMonth() + 1);
  }

  const currentMonth = getCurrentMonthLabel();
  const stats = getDashboardStats(data, currentMonth);
  const periodReceipts = (data.receipts || []).filter(r => periodMonthKeys.includes(r.month) || inPeriod(r.createdAt));
  const periodExpenses = (data.expensePayments || []).filter(e => periodMonthKeys.includes(e.month) || inPeriod(e.paymentDate || e.date || e.createdAt));
  const periodContracts = (data.contracts || []).filter(c => inPeriod(c.createdAt || c.startDate));
  const periodMoveOuts = (data.moveOutReports || []).filter(r => inPeriod(r.actualEndDate || r.createdAt));

  const totalRevenue = periodReceipts.reduce((sum, r) => sum + Number(r.total || 0), 0);
  const paidRevenue = periodReceipts.reduce((sum, r) => sum + Number(r.paidAmount || 0), 0);
  const unpaidRevenue = Math.max(0, totalRevenue - paidRevenue);
  const expenses = periodExpenses.reduce((sum, e) => sum + getExpensePaidAmount(e), 0);
  const netCash = paidRevenue - expenses;
  const collectionRate = totalRevenue > 0 ? (paidRevenue / totalRevenue) * 100 : 0;
  const occupancyRate = stats.totalRooms > 0 ? (stats.occupiedRooms / stats.totalRooms) * 100 : 0;
  const recentReceipts = [...periodReceipts].sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0)).slice(0, 5);
  const topDebtRooms = [...periodReceipts]
    .map(r => ({ ...r, debt: getReceiptPaymentState(r).debt }))
    .filter(r => r.debt > 0)
    .sort((a, b) => b.debt - a.debt)
    .slice(0, 5);
  const expiringAlerts = [...stats.expiringContracts]
    .map(contract => ({
      contract,
      roomId: contract.roomId,
      daysLeft: getDaysUntil(contract.endDate),
      tenant: getPrimaryTenantByContract(data, contract.id)
    }))
    .filter(item => item.daysLeft !== null)
    .sort((a, b) => a.daysLeft - b.daysLeft);
  const chartRows = periodMonthKeys.map(key => {
    const income = periodReceipts.filter(r => r.month === key).reduce((sum, r) => sum + Number(r.paidAmount || 0), 0);
    const out = periodExpenses.filter(e => e.month === key).reduce((sum, e) => sum + getExpensePaidAmount(e), 0);
    return { key, income, expense: out, net: income - out };
  });
  const utilityRows = periodMonthKeys.map(key => {
    const receipts = (data.receipts || []).filter(r => r.type === 'monthly' && r.month === key);
    return {
      key,
      roomCount: new Set(receipts.map(r => r.roomId).filter(Boolean)).size,
      electricUsed: receipts.reduce((sum, r) => sum + Number(r.electricUsed || 0), 0),
      electricAmount: receipts.reduce((sum, r) => sum + Number(r.electricAmount || 0), 0),
      waterUsed: receipts.reduce((sum, r) => sum + Number(r.waterUsed || 0), 0),
      waterAmount: receipts.reduce((sum, r) => sum + Number(r.waterAmount || 0), 0),
    };
  });
  const utilityTotals = utilityRows.reduce((totals, row) => ({
    electricUsed: totals.electricUsed + row.electricUsed,
    electricAmount: totals.electricAmount + row.electricAmount,
    waterUsed: totals.waterUsed + row.waterUsed,
    waterAmount: totals.waterAmount + row.waterAmount,
  }), { electricUsed: 0, electricAmount: 0, waterUsed: 0, waterAmount: 0 });
  const contractValueRows = (data.contracts || [])
    .filter(contract => contract.status !== 'ended' && isValidContractRange(contract.startDate, contract.endDate))
    .map(contract => {
      const room = (data.rooms || []).find(item => item.id === contract.roomId);
      const tenant = getPrimaryTenantByContract(data, contract.id);
      const contractReceipts = (data.receipts || []).filter(receipt => receipt.contractId === contract.id && receipt.type === 'monthly');
      const durationMonths = Math.max(1, diffMonths(contract.startDate, contract.endDate));
      const occupantCount = getContractOccupantCount(data, contract);
      const monthlyServices = room ? fixedServiceTotal(room, occupantCount || 1) : 0;
      const plannedRent = Number(contract.rent || room?.rent || 0) * durationMonths;
      const plannedServices = monthlyServices * durationMonths;
      const actualServices = contractReceipts.reduce((sum, receipt) => sum + Number(receipt.fixedServices || 0), 0);
      const electricAmount = contractReceipts.reduce((sum, receipt) => sum + Number(receipt.electricAmount || 0), 0);
      const waterAmount = contractReceipts.reduce((sum, receipt) => sum + Number(receipt.waterAmount || 0), 0);
      const otherAmount = contractReceipts.reduce((sum, receipt) => sum + Number(receipt.other || 0), 0);
      const billedAmount = contractReceipts.reduce((sum, receipt) => sum + Number(receipt.total || 0), 0);
      const paidAmount = contractReceipts.reduce((sum, receipt) => sum + Number(receipt.paidAmount || 0), 0);
      const start = parseDateFlexible(contract.startDate);
      const end = parseDateFlexible(contract.endDate);
      const totalDays = Math.max(1, Math.ceil((end - start) / 86400000));
      const elapsedDays = Math.ceil((today - start) / 86400000);
      const progress = Math.max(0, Math.min(100, elapsedDays / totalDays * 100));
      return {
        contract,
        tenant,
        durationMonths,
        plannedRent,
        plannedServices,
        plannedValue: plannedRent + plannedServices,
        actualServices,
        electricAmount,
        waterAmount,
        otherAmount,
        billedAmount,
        paidAmount,
        debt: Math.max(0, billedAmount - paidAmount),
        progress,
      };
    })
    .sort((a, b) => String(a.contract.roomId).localeCompare(String(b.contract.roomId), undefined, { numeric: true }));
  const chartMax = Math.max(1, ...chartRows.flatMap(r => [r.income, r.expense, Math.abs(r.net)]));

  return (
    <div className="dashboard-container stack">
      <div className="dashboard-header no-print">
        <div className="stack">
          <h2 style={{ fontSize: '24px' }}>Tổng quan hệ thống</h2>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', color: 'var(--text-muted)' }}>
            <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: isSyncing ? 'var(--warning)' : 'var(--success)' }}></span>
            {isSyncing ? 'Đang đồng bộ...' : lastSynced ? `Đã lưu: ${lastSynced.toLocaleTimeString()}` : 'Chế độ Local'}
          </div>
        </div>
        <div style={{ display: 'flex', gap: '12px' }}>
          <button className="primary-btn sm" onClick={() => onAction('create_receipt_all')}>⚡ Tạo phiếu tháng</button>
          <button className="secondary-btn sm" onClick={() => onAction('import_json')}>📤 Nhập JSON</button>
          <button className="secondary-btn sm" onClick={() => onAction('export_json')}>📥 Xuất JSON</button>
        </div>
      </div>

      <div className="dashboard-filter-bar">
        {['month', 'quarter', 'year', 'custom'].map(mode => <button key={mode} className={periodMode === mode ? 'active' : ''} onClick={() => setPeriodMode(mode)}>{mode === 'month' ? 'Tháng' : mode === 'quarter' ? 'Quý' : mode === 'year' ? 'Năm' : 'Tùy chỉnh'}</button>)}
        {periodMode === 'month' && <input type="month" value={periodMonth} onChange={e => setPeriodMonth(e.target.value)} />}
        {periodMode === 'quarter' && <select value={periodQuarter} onChange={e => setPeriodQuarter(e.target.value)}>{[0, 1, 2].flatMap(offset => [1, 2, 3, 4].map(q => `${today.getFullYear() - offset}-Q${q}`)).map(v => <option key={v} value={v}>{v.replace('-Q', ' - Quý ')}</option>)}</select>}
        {periodMode === 'year' && <select value={periodYear} onChange={e => setPeriodYear(e.target.value)}>{[0, 1, 2, 3].map(offset => <option key={today.getFullYear() - offset}>{today.getFullYear() - offset}</option>)}</select>}
        {periodMode === 'custom' && <><input type="date" value={customRange.start} onChange={e => setCustomRange({...customRange, start: e.target.value})} /><input type="date" value={customRange.end} onChange={e => setCustomRange({...customRange, end: e.target.value})} /></>}
        <span>{period.label}</span>
      </div>

      <div className="stats-grid dashboard-kpi-grid">
        <div className="stat-card-liquid kpi money"><p className="stat-label">Tổng doanh thu</p><p className="stat-value">{formatMoney(totalRevenue)}</p><p className="stat-note">{periodReceipts.length} phiếu</p></div>
        <div className="stat-card-liquid kpi success"><p className="stat-label">Đã thu</p><p className="stat-value">{formatMoney(paidRevenue)}</p><p className="stat-note">Tỷ lệ thu {collectionRate.toFixed(0)}%</p></div>
        <div className="stat-card-liquid kpi danger"><p className="stat-label">Chưa thu</p><p className="stat-value">{formatMoney(unpaidRevenue)}</p><p className="stat-note">Công nợ trong kỳ</p></div>
        <div className="stat-card-liquid kpi net"><p className="stat-label">Dòng tiền ròng</p><p className="stat-value">{formatMoney(netCash)}</p><p className="stat-note">Đã thu - chi phí</p></div>
        <div className="stat-card-liquid kpi rooms" onClick={() => onRoomClick('')}><p className="stat-label">Phòng đang ở</p><p className="stat-value">{stats.occupiedRooms}/{stats.totalRooms}</p><p className="stat-note">Lấp đầy {occupancyRate.toFixed(0)}%</p></div>
        <div className="stat-card-liquid kpi vacant"><p className="stat-label">Phòng trống</p><p className="stat-value">{stats.vacantRooms}</p><p className="stat-note">Sẵn sàng cho thuê</p></div>
        <div className="stat-card-liquid kpi tenant"><p className="stat-label">Thuê mới / Rời đi</p><p className="stat-value">{periodContracts.length} / {periodMoveOuts.length}</p><p className="stat-note">Trong kỳ</p></div>
        <div className="stat-card-liquid kpi warning"><p className="stat-label">HĐ sắp hết hạn</p><p className="stat-value">{stats.expiringContracts.length}</p><p className="stat-note">Trong 30 ngày tới</p></div>
      </div>

      <div className="dashboard-grid-main">
        <div className="stack" style={{ gap: '24px' }}>
          <div className="widget liquid-glass">
            <h3 className="form-section-title">📈 Doanh thu / Chi phí / Lợi nhuận</h3>
            <div className="mini-chart">
              {chartRows.map(row => (
                <div key={row.key} className="chart-row">
                  <span>{row.key}</span>
                  <div className="chart-bars">
                    <i className="income" style={{ width: `${Math.max(4, row.income / chartMax * 100)}%` }} title={`Thu ${formatMoney(row.income)}`}></i>
                    <i className="expense" style={{ width: `${Math.max(4, row.expense / chartMax * 100)}%` }} title={`Chi ${formatMoney(row.expense)}`}></i>
                    <i className={row.net >= 0 ? 'net-positive' : 'net-negative'} style={{ width: `${Math.max(4, Math.abs(row.net) / chartMax * 100)}%` }} title={`Ròng ${formatMoney(row.net)}`}></i>
                  </div>
                  <b>{formatMoney(row.net)}</b>
                </div>
              ))}
            </div>
            <div className="chart-legend"><span className="income"></span>Tiền vào <span className="expense"></span>Tiền ra <span className="net-positive"></span>Ròng</div>
          </div>

          <div className="widget liquid-glass">
            <h3 className="form-section-title">⚡💧 Điện nước theo tháng thu tiền</h3>
            <p className="small muted" style={{ marginBottom: '12px' }}>Tổng hợp từ phiếu thu tháng của toàn bộ các phòng trong kỳ đã chọn.</p>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Tháng thu</th>
                    <th>Số phòng</th>
                    <th>Số điện (kWh)</th>
                    <th>Tiền điện</th>
                    <th>Số nước (m³)</th>
                    <th>Tiền nước</th>
                  </tr>
                </thead>
                <tbody>
                  {utilityRows.map(row => (
                    <tr key={row.key}>
                      <td><b>{row.key}</b></td>
                      <td>{row.roomCount}</td>
                      <td style={{ fontWeight: '700', color: 'var(--warning)' }}>{formatLocaleNumber(Math.round(row.electricUsed * 100) / 100)}</td>
                      <td>{formatMoney(row.electricAmount)}</td>
                      <td style={{ fontWeight: '700', color: 'var(--success)' }}>{formatLocaleNumber(Math.round(row.waterUsed * 100) / 100)}</td>
                      <td>{formatMoney(row.waterAmount)}</td>
                    </tr>
                  ))}
                  <tr>
                    <td><b>Tổng kỳ</b></td>
                    <td>—</td>
                    <td><b>{formatLocaleNumber(Math.round(utilityTotals.electricUsed * 100) / 100)}</b></td>
                    <td><b>{formatMoney(utilityTotals.electricAmount)}</b></td>
                    <td><b>{formatLocaleNumber(Math.round(utilityTotals.waterUsed * 100) / 100)}</b></td>
                    <td><b>{formatMoney(utilityTotals.waterAmount)}</b></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <div className="widget liquid-glass">
            <h3 className="form-section-title">📑 Giá trị & tiến độ thực hiện hợp đồng</h3>
            <p className="small muted" style={{ marginBottom: '12px' }}>Giá trị dự kiến gồm tiền phòng và dịch vụ cố định trong thời hạn hợp đồng. Điện, nước và chi phí khác là số thực tế đã phát sinh trên phiếu thu.</p>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Phòng / Khách thuê</th>
                    <th>Thời hạn</th>
                    <th>Tiến độ</th>
                    <th>Tiền phòng dự kiến</th>
                    <th>Dịch vụ dự kiến</th>
                    <th>Giá trị hợp đồng dự kiến</th>
                    <th>Điện đã phát sinh</th>
                    <th>Nước đã phát sinh</th>
                    <th>Chi phí khác</th>
                    <th>Tổng đã lập phiếu</th>
                    <th>Đã thu</th>
                    <th>Còn phải thu</th>
                  </tr>
                </thead>
                <tbody>
                  {contractValueRows.map(row => (
                    <tr key={row.contract.id}>
                      <td>
                        <b>P{row.contract.roomId}</b><br />
                        <span className="small muted">{row.tenant?.name || 'Chưa cập nhật'}</span><br />
                        <span className="small muted">CCCD: {row.tenant?.cccd || 'Chưa cập nhật'}</span><br />
                        <span className="small muted">Số HĐ: {row.contract.contractNo || row.contract.id}</span>
                      </td>
                      <td><span className="small">{formatBusinessDate(row.contract.startDate)} → {formatBusinessDate(row.contract.endDate)}</span><br /><b>{row.durationMonths} tháng</b></td>
                      <td style={{ minWidth: '120px' }}>
                        <div style={{ height: '8px', borderRadius: '999px', background: 'rgba(148,163,184,.25)', overflow: 'hidden' }}>
                          <div style={{ width: `${row.progress}%`, height: '100%', background: 'var(--primary-gradient)' }}></div>
                        </div>
                        <b className="small">{row.progress.toFixed(0)}%</b>
                      </td>
                      <td>{formatMoney(row.plannedRent)}</td>
                      <td>{formatMoney(row.plannedServices)}<br /><span className="small muted">Đã lập: {formatMoney(row.actualServices)}</span></td>
                      <td><b>{formatMoney(row.plannedValue)}</b></td>
                      <td>{formatMoney(row.electricAmount)}</td>
                      <td>{formatMoney(row.waterAmount)}</td>
                      <td>{formatMoney(row.otherAmount)}</td>
                      <td><b>{formatMoney(row.billedAmount)}</b></td>
                      <td><b className="success">{formatMoney(row.paidAmount)}</b></td>
                      <td><b className={row.debt > 0 ? 'danger' : 'success'}>{formatMoney(row.debt)}</b></td>
                    </tr>
                  ))}
                  {contractValueRows.length === 0 && <tr><td colSpan="12" className="center muted" style={{ padding: '24px' }}>Chưa có hợp đồng đang thực hiện.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>

          {/* Cảnh báo */}
          <div className="widget liquid-glass">
            <h3 className="form-section-title">⚠️ Cảnh báo & Nhắc nhở</h3>
            <div className="alert-list stack" style={{ gap: '12px' }}>
              {stats.expiringContracts.length > 0 && (
                <div className="alert-item warning" style={{ alignItems: 'flex-start' }}>
                  <span>📄 <b>{expiringAlerts.length} hợp đồng sắp hết hạn cần xử lý</b></span>
                  <div className="stack" style={{ gap: '8px', marginTop: '8px', width: '100%' }}>
                    {expiringAlerts.map(item => (
                      <div key={item.contract.id} className="actionable" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '12px', paddingTop: '8px', borderTop: '1px solid rgba(146, 64, 14, 0.15)' }}>
                        <div className="alert-content">
                          <b>P{item.roomId} • {item.tenant?.name || 'Chưa rõ khách'} • {formatBusinessDate(item.contract.endDate)}</b>
                          <small>
                            {item.daysLeft < 0
                              ? `Đã hết hạn ${Math.abs(item.daysLeft)} ngày`
                              : item.daysLeft === 0
                                ? 'Hết hạn hôm nay'
                                : `Còn ${item.daysLeft} ngày`}
                          </small>
                        </div>
                        <div className="alert-actions">
                          <button className="secondary-btn sm" onClick={() => onAction('view_room', { id: item.roomId })}>Xem phòng</button>
                          <button className="primary-btn sm" onClick={() => onAction('renew_contract', { roomId: item.roomId, contractId: item.contract.id })}>Gia hạn</button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {stats.unpaidReceipts.length > 0 && (
                <div className="alert-item danger" style={{ alignItems: 'flex-start' }}>
                  <span>💸 <b>{stats.unpaidReceipts.length} phòng chưa đóng tiền tháng {currentMonth}</b></span>
                  <div className="stack" style={{ gap: '4px', marginTop: '8px', width: '100%' }}>
                    {stats.unpaidReceipts
                      .map(r => ({ ...r, debt: getReceiptPaymentState(r).debt }))
                      .sort((a, b) => Number(b.debt || 0) - Number(a.debt || 0))
                      .map(r => (
                        <div key={r.id} style={{ display: 'flex', justifyContent: 'space-between', gap: '12px' }}>
                          <span>• Phòng P{r.roomId}</span>
                          <b>{formatMoney(r.debt)}</b>
                        </div>
                      ))}
                  </div>
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
              {topDebtRooms.length > 0 && (
                <div className="alert-item danger" style={{ alignItems: 'flex-start' }}>
                  <span>📌 <b>Top nợ</b></span>
                  <div className="stack" style={{ gap: '4px', marginTop: '8px', width: '100%' }}>
                    {topDebtRooms.map((r, index) => (
                      <div key={r.id} style={{ display: 'flex', justifyContent: 'space-between', gap: '12px' }}>
                        <span>{index + 1}. Phòng P{r.roomId}</span>
                        <b>{formatMoney(r.debt)}</b>
                      </div>
                    ))}
                  </div>
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
          <div className="widget liquid-glass">
            <h3 className="form-section-title">💵 Dòng tiền</h3>
            <div className="cashflow-list">
              <div><span>Tiền vào</span><b className="success">{formatMoney(paidRevenue)}</b></div>
              <div><span>Tiền ra</span><b className="danger">{formatMoney(expenses)}</b></div>
              <div><span>Dòng tiền ròng</span><b className={netCash >= 0 ? 'success' : 'danger'}>{formatMoney(netCash)}</b></div>
              <div><span>Công nợ còn lại</span><b>{formatMoney(unpaidRevenue)}</b></div>
            </div>
          </div>

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
                      <b className="danger">{formatBusinessDate(c.endDate)}</b>
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
  const [caseFilter, setCaseFilter] = useState('all');
  const [selectedRoomId, setSelectedRoomId] = useState(null);
  const currentMonth = getCurrentMonthLabel();
  const roomCards = useMemo(() => {
    const q = query.toLowerCase();
    return (data.rooms || []).map(room => {
      const statusInfo = getRoomStatusInfo(data, room.id);
      const { label, color, contract, ownerOccupied } = statusInfo;
      const roomActiveMember = (data.memberships || []).find(m => m.roomId === room.id && m.status === 'active' && m.role === 'primary')
        || (data.memberships || []).find(m => m.roomId === room.id && m.status === 'active');
      const primaryTenant = contract
        ? (getPrimaryTenantByContract(data, contract.id) || (roomActiveMember ? (data.tenants || []).find(t => t.id === roomActiveMember.tenantId) : null))
        : roomActiveMember ? (data.tenants || []).find(t => t.id === roomActiveMember.tenantId) : null;
      const currentReceipt = contract ? (data.receipts || []).find(r => r.roomId === room.id && r.contractId === contract.id && r.month === currentMonth && r.type === 'monthly') : null;
      const debt = currentReceipt ? getReceiptPaymentState(currentReceipt).debt : 0;
      const caseType = ownerOccupied ? 'owner' : label === 'Trống' ? 'vacant' : debt > 0 ? 'debt' : 'occupied';
      return { room, label, color, contract, ownerOccupied, primaryTenant, currentReceipt, debt, caseType };
    }).filter(item => item.room.id.toLowerCase().includes(q));
  }, [data, query, currentMonth]);

  const caseOptions = [
    ['all', 'Tất cả', roomCards.length],
    ['occupied', 'Đang ở', roomCards.filter(r => r.caseType === 'occupied').length],
    ['debt', 'Đang nợ', roomCards.filter(r => r.caseType === 'debt').length],
    ['vacant', 'Trống', roomCards.filter(r => r.caseType === 'vacant').length],
    ['owner', 'Chủ nhà', roomCards.filter(r => r.caseType === 'owner').length],
  ];
  const visibleRooms = caseFilter === 'all' ? roomCards : roomCards.filter(r => r.caseType === caseFilter);
  const selectedCard = visibleRooms.find(r => r.room.id === selectedRoomId) || null;

  const renderActionButtons = ({ room, label, contract, ownerOccupied, currentReceipt }) => (
    label === 'Trống' ? (
      <>
        <button className="primary-btn wide" onClick={(e) => { e.stopPropagation(); onAction('add_tenant', room); }}>+ Thêm khách</button>
        <button className="secondary-btn" title="Xem lịch sử" onClick={(e) => { e.stopPropagation(); onAction('view_history', room); }}>📜</button>
      </>
    ) : (
      <>
        <button className="secondary-btn" title={ownerOccupied && !contract ? 'Không cần hợp đồng' : 'Xem hợp đồng'} onClick={(e) => { e.stopPropagation(); ownerOccupied && !contract ? setSelectedRoomId(room.id) : onAction('view_contract', room); }}>📄</button>
        {!ownerOccupied && <button className="primary-btn wide" onClick={(e) => { e.stopPropagation(); onAction('create_receipt', room); }}>{currentReceipt ? 'Sửa phiếu' : 'Lập phiếu'}</button>}
      </>
    )
  );

  return (
    <div className={`rooms-panel character-select ${selectedCard ? 'has-selection' : ''}`}>
      <div className="room-case-tabs">
        {caseOptions.map(([id, text, count]) => (
          <button key={id} className={caseFilter === id ? 'active' : ''} onClick={() => { setCaseFilter(id); setSelectedRoomId(null); }}>
            <span>{text}</span><b>{count}</b>
          </button>
        ))}
      </div>
      <div className="character-select-layout">
        <div className="room-roster">
          {visibleRooms.map(({ room, label, color, contract, ownerOccupied, primaryTenant, currentReceipt, debt, caseType }) => {
            const statusText = ownerOccupied ? 'Chủ nhà' : label;
            const receiptLabel = ownerOccupied ? 'Không lập phiếu' : debt > 0 ? `Nợ ${formatMoney(debt)}` : currentReceipt ? 'Đã thanh toán' : 'Chưa có phiếu';
            const isSelected = selectedRoomId === room.id;
            return (
              <div key={room.id} className={`room-card-liquid compact character-card ${caseType} ${isSelected ? 'selected' : ''}`} onClick={() => setSelectedRoomId(isSelected ? null : room.id)}>
                <div className="room-header">
                  <span className="room-id">P{room.id}</span>
                  <span className={`status-badge-liquid ${ownerOccupied ? 'notice' : color === 'green' ? 'active' : color === 'gray' ? 'vacant' : color}`}>{statusText}</span>
                </div>
                <div className="room-body">
                  {primaryTenant ? (
                    <div className="tenant-block-primary">
                      <p className="tenant-name-main">{primaryTenant.name}</p>
                      <div className="room-card-meta">
                        <span>{ownerOccupied && !contract ? 'Chủ nhà ở' : formatMoney(contract?.rent || room.rent || 0)}</span>
                        <b className={debt > 0 ? 'danger' : ownerOccupied ? 'muted' : 'success'}>{receiptLabel}</b>
                      </div>
                    </div>
                  ) : <p className="muted">Phòng đang trống</p>}
                </div>
                {!selectedCard && <div className="btn-group">{renderActionButtons({ room, label, contract, ownerOccupied, currentReceipt })}</div>}
              </div>
            );
          })}
          {!visibleRooms.length && <div className="empty-state-inline">Không có phòng phù hợp với bộ lọc này.</div>}
        </div>

        {selectedCard && (() => {
          const { room, label, color, contract, ownerOccupied, primaryTenant, currentReceipt, debt, caseType } = selectedCard;
          const activeMembers = (data.memberships || []).filter(m => m.roomId === room.id && m.status === 'active');
          const occupantCount = contract ? getContractOccupantCount(data, contract) : activeMembers.length;
          const fixedServices = fixedServiceTotal(room, occupantCount || (primaryTenant ? 1 : 0));
          const latestReceipt = currentReceipt || (data.receipts || []).filter(r => r.roomId === room.id).sort((a, b) => new Date(b.createdAt || b.savedAt || 0) - new Date(a.createdAt || a.savedAt || 0))[0];
          const statusText = ownerOccupied ? 'Chủ nhà' : label;
          const receiptLabel = ownerOccupied ? 'Không lập phiếu' : debt > 0 ? `Nợ ${formatMoney(debt)}` : currentReceipt ? 'Đã thanh toán' : 'Chưa có phiếu';
          return (
            <aside className={`character-detail ${caseType}`}>
              <div className="character-detail-hero">
                <div>
                  <span className="detail-eyebrow">Phòng được chọn</span>
                  <h2>P{room.id}</h2>
                  <p>{primaryTenant?.name || 'Phòng đang trống'}</p>
                </div>
                <span className={`status-badge-liquid ${ownerOccupied ? 'notice' : color === 'green' ? 'active' : color === 'gray' ? 'vacant' : color}`}>{statusText}</span>
              </div>

              <div className="detail-stat-grid">
                <div><span>Giá thuê</span><b>{formatMoney(contract?.rent || room.rent || 0)}</b></div>
                <div><span>Tiền cọc</span><b>{formatMoney(contract?.deposit || room.deposit || 0)}</b></div>
                <div><span>Công nợ</span><b className={debt > 0 ? 'danger' : 'success'}>{formatMoney(debt)}</b></div>
                <div><span>Thanh toán</span><b className={debt > 0 ? 'danger' : ownerOccupied ? 'muted' : 'success'}>{receiptLabel}</b></div>
              </div>

              <div className="detail-info-list">
                <div><span>Người thuê</span><b>{primaryTenant?.name || 'Chưa có khách'}</b></div>
                <div><span>Số điện thoại</span><b>{primaryTenant?.phone || '—'}</b></div>
                <div><span>Ngày vào</span><b>{formatDisplayDate(activeMembers[0]?.joinedDate || contract?.startDate || '')}</b></div>
                <div><span>Hết hạn HĐ</span><b>{contract ? formatBusinessDate(contract.endDate) : ownerOccupied ? 'Không cần HĐ' : '—'}</b></div>
                <div><span>Điện gần nhất</span><b>{latestReceipt ? `${formatLocaleNumber(getElectricOld(latestReceipt))} → ${formatLocaleNumber(getElectricNew(latestReceipt))}` : 'Chưa có'}</b></div>
                <div><span>Nước gần nhất</span><b>{latestReceipt ? `${formatLocaleNumber(getWaterOld(latestReceipt))} → ${formatLocaleNumber(getWaterNew(latestReceipt))}` : 'Chưa có'}</b></div>
                <div><span>Dịch vụ cố định</span><b>{formatMoney(fixedServices)}</b></div>
                <div><span>Ghi chú</span><b>{room.note || '—'}</b></div>
              </div>

              <div className="character-detail-actions">
                <button className="primary-btn" onClick={() => onSelect(room)}>Mở chi tiết</button>
                {!ownerOccupied && <button className="secondary-btn" onClick={() => onAction('create_receipt', room)}>{currentReceipt ? 'Sửa phiếu' : 'Lập phiếu'}</button>}
                {!ownerOccupied && debt > 0 && currentReceipt && <button className="secondary-btn" onClick={() => onAction('pay_receipt', currentReceipt)}>Thanh toán</button>}
                <button className="secondary-btn" onClick={() => onAction('view_history', room)}>Lịch sử thuê</button>
                {label === 'Trống' && <button className="primary-btn" onClick={() => onAction('add_tenant', room)}>Thêm khách</button>}
                {!ownerOccupied && contract && <button className="secondary-btn danger" onClick={() => onAction('moving_out', room)}>Trả phòng</button>}
              </div>
            </aside>
          );
        })()}
      </div>
    </div>
  );
}

function TenantsTab({ tenants, data, onAction, query, setQuery, setData }) {
  const [statusFilter, setStatusFilter] = useState('active');
  const [paymentFilter, setPaymentFilter] = useState('all');
  const [contractFilter, setContractFilter] = useState('all');
  const currentMonth = getCurrentMonthLabel();

  const tenantRows = useMemo(() => {
    return (data.memberships || []).map(m => {
      const tenant = (data.tenants || []).find(t => t.id === m.tenantId);
      if (!tenant) return null;
      const contract = (data.contracts || []).find(c => c.id === m.contractId);
      const room = (data.rooms || []).find(r => r.id === m.roomId);
      const receipt = contract ? (data.receipts || []).find(r => r.contractId === contract.id && r.month === currentMonth && r.type === 'monthly') : null;
      const debt = receipt ? getReceiptPaymentState(receipt).debt : 0;
      const contractDaysLeft = contract ? getDaysUntil(contract.endDate) : null;
      const contractState = !contract || isOwnerOccupiedRoom(room) ? 'none' : contractDaysLeft !== null && contractDaysLeft < 0 ? 'expired' : contractDaysLeft !== null && contractDaysLeft <= 30 ? 'expiring' : 'valid';
      const paymentState = debt > 0 ? 'debt' : receipt ? 'paid' : 'missing';
      return { tenant, membership: m, contract, room, receipt, debt, contractDaysLeft, contractState, paymentState };
    }).filter(Boolean).sort((a, b) => String(a.room?.id || a.membership.roomId).localeCompare(String(b.room?.id || b.membership.roomId), 'vi', { numeric: true }));
  }, [data, currentMonth]);

  const activeRows = tenantRows.filter(r => r.membership.status === 'active');
  const moveOutThisMonth = (data.moveOutReports || []).filter(r => String(r.actualEndDate || '').slice(0, 7) === new Date().toISOString().slice(0, 7)).length;
  const debtRows = activeRows.filter(r => r.debt > 0);
  const expiringRows = activeRows.filter(r => r.contractState === 'expiring' || r.contractState === 'expired');
  const normalizedQuery = String(query || '').trim().toLowerCase();

  const filteredRows = tenantRows.filter(row => {
    const roomId = String(row.room?.id || row.membership.roomId || '');
    const searchBlob = `${row.tenant.name || ''} ${row.tenant.phone || ''} ${row.tenant.cccd || ''} ${roomId}`.toLowerCase();
    if (normalizedQuery && !searchBlob.includes(normalizedQuery)) return false;
    if (statusFilter === 'active' && row.membership.status !== 'active') return false;
    if (statusFilter === 'notice' && row.membership.status !== 'notice' && row.contract?.status !== 'notice') return false;
    if (statusFilter === 'left' && row.membership.status !== 'ended') return false;
    if (paymentFilter === 'paid' && row.paymentState !== 'paid') return false;
    if (paymentFilter === 'debt' && row.paymentState !== 'debt') return false;
    if (paymentFilter === 'missing' && row.paymentState !== 'missing') return false;
    if (contractFilter === 'expiring' && row.contractState !== 'expiring') return false;
    if (contractFilter === 'expired' && row.contractState !== 'expired') return false;
    return true;
  });

  return (
    <div className="tenant-crm stack">
      <div className="tenant-crm-header no-print">
        <div>
          <h2>👥 Người thuê</h2>
          <p className="muted small">Quản lý vòng đời khách thuê: vào ở, hợp đồng, thanh toán và rời đi.</p>
        </div>
        <div className="btn-group" style={{ marginTop: 0 }}>
          <button className="primary-btn" onClick={() => onAction('add_tenant')}>+ Thêm khách</button>
          <button className="secondary-btn" onClick={() => setQuery('')}>🔎 Xóa tìm kiếm</button>
          <button className="secondary-btn" onClick={() => onAction('export_excel')}>📊 Xuất Excel</button>
        </div>
      </div>

      <div className="tenant-stats-grid">
        <div className="tenant-stat-card active"><b>{activeRows.length}</b><span>Đang thuê</span><small>+{tenantRows.filter(r => String(r.membership.createdAt || '').slice(0, 7) === new Date().toISOString().slice(0, 7)).length} tháng này</small></div>
        <div className="tenant-stat-card notice"><b>{moveOutThisMonth}</b><span>Rời đi</span><small>{currentMonth}</small></div>
        <div className="tenant-stat-card debt"><b>{debtRows.length}</b><span>Đang nợ</span><small>{formatMoney(debtRows.reduce((s, r) => s + r.debt, 0))}</small></div>
        <div className="tenant-stat-card warning"><b>{expiringRows.length}</b><span>HĐ cảnh báo</span><small>Hết hạn / 30 ngày</small></div>
      </div>

      <div className="tenant-filter-bar no-print">
        <label>Trạng thái
          <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
            <option value="active">Đang ở</option>
            <option value="notice">Sắp rời</option>
            <option value="left">Đã rời</option>
            <option value="all">Tất cả</option>
          </select>
        </label>
        <label>Thanh toán
          <select value={paymentFilter} onChange={e => setPaymentFilter(e.target.value)}>
            <option value="all">Tất cả</option>
            <option value="paid">Đã đóng</option>
            <option value="debt">Đang nợ</option>
            <option value="missing">Chưa có phiếu</option>
          </select>
        </label>
        <label>Hợp đồng
          <select value={contractFilter} onChange={e => setContractFilter(e.target.value)}>
            <option value="all">Tất cả</option>
            <option value="expiring">Sắp hết hạn</option>
            <option value="expired">Hết hạn</option>
          </select>
        </label>
        <button className="secondary-btn sm" onClick={() => { setStatusFilter('active'); setPaymentFilter('all'); setContractFilter('all'); }}>Reset lọc</button>
      </div>

      <div className="widget liquid-glass" style={{ padding: 0 }}>
        <div className="table-wrap">
        <table>
          <thead><tr><th>Người thuê</th><th>Phòng</th><th>Hợp đồng</th><th>Thanh toán</th><th>Ngày vào</th><th>Trạng thái</th><th>Thao tác</th></tr></thead>
          <tbody>
            {filteredRows.map(({ tenant: t, membership: m, contract, room, receipt, debt, contractDaysLeft, contractState, paymentState }) => {
              const contractText = contractState === 'none' ? 'Không áp dụng' : contractDaysLeft === null ? 'Chưa cập nhật' : contractDaysLeft < 0 ? 'Hết hạn' : `Còn ${contractDaysLeft} ngày`;
              const contractBadge = contractState === 'valid' ? 'active' : contractState === 'expiring' ? 'notice' : contractState === 'expired' ? 'debt' : 'vacant';
              const paymentBadge = paymentState === 'paid' ? 'active' : paymentState === 'debt' ? 'debt' : 'notice';
              const membershipBadge = m.status === 'active' ? 'active' : m.status === 'notice' ? 'notice' : 'vacant';
              const membershipLabel = m.status === 'active' ? (m.role === 'primary' ? 'Đại diện' : 'Ở cùng') : m.status === 'notice' ? 'Sắp rời' : 'Đã rời';
              return (
                <tr key={`${t.id}-${m.id || m.roomId}`}>
                  <td>
                    <button className="tenant-identity" onClick={() => onAction('detail', t)}>
                      <b>👤 {t.name}</b>
                      <span>📱 {t.phone || 'Chưa có SĐT'}</span>
                      <span>🪪 {t.cccd || 'Chưa có CCCD'}</span>
                    </button>
                  </td>
                  <td><button className="secondary-btn sm" onClick={() => onAction('view_room', room || { id: m.roomId })}>P{m.roomId}</button></td>
                  <td>
                    <div className="tenant-contract-cell">
                      <span className={`status-badge-liquid ${contractBadge}`}>{contractText}</span>
                      {contract && <small>{formatBusinessDate(contract.startDate)} → {formatBusinessDate(contract.endDate)}</small>}
                    </div>
                  </td>
                  <td>
                    <button className={`tenant-payment-pill ${paymentBadge}`} onClick={() => receipt ? onAction('view_qr', receipt) : onAction('create_receipt', room || { id: m.roomId })}>
                      {paymentState === 'paid' ? `Đã đóng ${receipt?.month || ''}` : paymentState === 'debt' ? `Nợ ${formatMoney(debt)}` : 'Chưa có phiếu'}
                    </button>
                  </td>
                  <td>{formatDisplayDate(m.joinedDate || contract?.startDate || String(m.createdAt || '').slice(0, 10))}</td>
                  <td><span className={`status-badge-liquid ${membershipBadge}`}>{membershipLabel}</span></td>
                  <td>
                    <div style={{ display: 'flex', gap: '4px' }}>
                      {contract && <button className="secondary-btn sm" onClick={() => onAction('view_contract', { roomId: m.roomId, contractId: contract.id })}>HĐ</button>}
                      <button className="secondary-btn sm" onClick={() => onAction('edit_tenant', t)}>Sửa</button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {filteredRows.length === 0 && <tr><td colSpan="7" style={{ textAlign: 'center', padding: '40px' }}>Không tìm thấy khách phù hợp.</td></tr>}
          </tbody>
        </table>
        </div>
      </div>
    </div>
  );
}

function ReceiptsTab({ data, bankInfo, onUpdateReceipt, onBatchCreate, onView, onPrintBatch, onPay, onDeleteReceipt, onGoToPayment }) {
  const [selectedMonth, setSelectedMonth] = useState(() => getCurrentMonthLabel());
  const [activeTab, setActiveTab] = useState('entry');
  const [saveModal, setSaveModal] = useState(null);

  const monthlyReceipts = useMemo(() => (data.receipts || []).filter(r => r.month === selectedMonth && r.type === 'monthly'), [data.receipts, selectedMonth]);
  const isFinalized = monthlyReceipts.length > 0 && monthlyReceipts.every(r => r.isFinalized);

  function handleBatchCreate() {
    const billableContracts = getBillableContractsForMonth(data, selectedMonth);
    if (billableContracts.length === 0) return alert('Không có hợp đồng nào cần lập phiếu trong tháng này.');
    
    let createdCount = 0;
    let skippedCount = 0;
    const newReceipts = [];

    billableContracts.forEach(contract => {
      const room = data.rooms.find(r => r.id === contract.roomId);
      if (!room) return;
      const exists = (data.receipts || []).find(r => r.roomId === room.id && r.contractId === contract.id && r.month === selectedMonth && r.type === 'monthly');
      
      if (!exists) {
        const prev = getPreviousReceiptByRoom(data.receipts, room.id, selectedMonth, {
          includeSameMonth: true,
          excludeSameMonthContractId: contract.id
        });
        const billingContext = getMonthlyBillingContext(data, contract, selectedMonth);
        newReceipts.push(createMonthlyReceipt(room, contract, prev, selectedMonth, billingContext));
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

  function handleRefreshMonthReadings() {
    const refreshedReceipts = monthlyReceipts
      .filter(r => !r.isFinalized)
      .map(receipt => {
        const room = data.rooms.find(r => r.id === receipt.roomId);
        const contract = data.contracts.find(c => c.id === receipt.contractId);
        if (!room || !contract) return null;
        const previousReceipt = getPreviousReceiptByRoom(
          (data.receipts || []).filter(r => r.id !== receipt.id),
          receipt.roomId,
          selectedMonth,
          {
            includeSameMonth: true,
            excludeSameMonthContractId: receipt.contractId
          }
        );
        const billingContext = getMonthlyBillingContext(data, contract, selectedMonth);
        return {
          ...createMonthlyReceipt(room, contract, previousReceipt, selectedMonth, billingContext),
          id: receipt.id,
          paidAmount: receipt.paidAmount || 0,
          status: receipt.status || 'Chưa thanh toán',
          note: receipt.note || ''
        };
      })
      .filter(Boolean);

    if (refreshedReceipts.length === 0) {
      alert('Không có phiếu chưa lưu nào để làm mới chỉ số.');
      return;
    }

    onBatchCreate(refreshedReceipts);
    alert(`Đã làm mới chỉ số cho ${refreshedReceipts.length} phiếu tháng ${selectedMonth}.`);
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
      const originalState = getReceiptPaymentState(original);
      if (originalState.isPaid) {
        updated = {
          ...updated,
          adjustmentDueAmount: Number(updated.total || 0),
          adjustmentPaidAmount: Number(updated.adjustmentPaidAmount || 0),
          adjustmentCreatedAt: updated.adjustmentCreatedAt || new Date().toISOString(),
          adjustmentReason: updated.adjustmentReason || 'Phát sinh/chỉnh lại điện nước sau khi phiếu đã thanh toán'
        };
      }
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
            <button className="secondary-btn" onClick={handleRefreshMonthReadings}>🔄 Làm mới chỉ số</button>
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
                  const tenant = getTenantForReceipt(data, r) || { name: 'N/A' };
                  
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
                          <button className="secondary-btn sm" title="Hủy phiếu" style={{ color: '#ef4444' }} onClick={() => onDeleteReceipt(r)}>✕</button>
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
                {monthlyReceipts.map(r => {
                  const paymentState = getReceiptPaymentState(r);
                  return (
                    <tr key={r.id}>
                      <td><b>P{r.roomId}</b></td><td>{r.month}</td><td>{formatMoney(r.total)}</td><td>{formatMoney(paymentState.paidAmount)}</td><td style={{ color: paymentState.debt > 0 ? 'var(--danger)' : 'inherit' }}>{formatMoney(paymentState.debt)}</td><td><span className={`status-badge-liquid ${paymentState.status === 'Đã thanh toán' ? 'active' : paymentState.status === 'Nợ một phần' ? 'notice' : 'debt'}`}>{paymentState.status}</span></td>
                      <td><div style={{ display: 'flex', gap: '8px' }}>{paymentState.isPaid ? <button className="secondary-btn sm paid-action" disabled>✓ Đã thu</button> : <button className="primary-btn sm" onClick={() => onPay(r)}>💸 Thu tiền</button>}<button className="secondary-btn sm" onClick={() => onView(r)}>📄 Chi tiết</button></div></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function SettingsTab({ data, setData, bankInfo, setBankInfo, onReset }) {
  const buildTempIndices = (rooms) => rooms.map(r => ({
    id: r.id,
    electric: r.initialElectric ?? r.electricStart ?? r.electricNew ?? 0,
    water: r.initialWater ?? r.waterStart ?? r.waterNew ?? 0
  }));
  const [tempIndices, setTempIndices] = useState(() => buildTempIndices(data.rooms));
  const [savedMessage, setSavedMessage] = useState('');

  useEffect(() => {
    setTempIndices(buildTempIndices(data.rooms));
  }, [data.rooms]);

  const handleUpdateIndices = () => {
    setData(old => ({
      ...old,
      rooms: old.rooms.map(r => {
        const found = tempIndices.find(ti => ti.id === r.id);
        if (found) {
          const electric = parseLocaleNumber(found.electric);
          const water = parseLocaleNumber(found.water);
          return {
            ...r,
            initialElectric: electric,
            initialWater: water,
            electricStart: electric,
            waterStart: water,
            electricOld: electric,
            waterOld: water,
            electricNew: r.electricNew ?? electric,
            waterNew: r.waterNew ?? water
          };
        }
        return r;
      })
    }));
    setSavedMessage(`Đã lưu chỉ số đầu kỳ lúc ${new Date().toLocaleTimeString('vi-VN')}`);
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
          {savedMessage && <p className="small" style={{ marginTop: '10px', color: 'var(--success)', fontWeight: 700 }}>{savedMessage}</p>}
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

function RoomDetailModal({ room, data, onClose, onAction, onAddRoommate }) {
  const [activeTab, setActiveTab] = useState('overview');
  const [historyFilter, setHistoryFilter] = useState('all');
  const [showMoreActions, setShowMoreActions] = useState(false);
  const [showRoommateForm, setShowRoommateForm] = useState(false);
  const currentMonth = getCurrentMonthLabel();
  const { label, color, contract, ownerOccupied } = getRoomStatusInfo(data, room.id);
  const roomActiveMembers = (data.memberships || []).filter(m => m.roomId === room.id && m.status === 'active');
  const roomPrimaryMember = roomActiveMembers.find(m => m.role === 'primary') || roomActiveMembers[0];
  const primaryTenant = contract
    ? (getPrimaryTenantByContract(data, contract.id) || (roomPrimaryMember ? (data.tenants || []).find(t => t.id === roomPrimaryMember.tenantId) : null))
    : roomPrimaryMember ? (data.tenants || []).find(t => t.id === roomPrimaryMember.tenantId) : null;
  const currentReceipt = contract ? (data.receipts || []).find(r => r.roomId === room.id && r.contractId === contract.id && r.month === currentMonth && r.type === 'monthly') : null;
  const allMembers = contract
    ? (data.memberships || []).filter(m => m.contractId === contract.id).map(m => ({ ...m, tenant: data.tenants.find(t => t.id === m.tenantId) }))
    : (data.memberships || []).filter(m => m.roomId === room.id).map(m => ({ ...m, tenant: data.tenants.find(t => t.id === m.tenantId) }));
  const activeMembers = allMembers.filter(m => m.status === 'active');
  const roomReceipts = (data.receipts || []).filter(r => r.roomId === room.id).sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));
  const roomTransfers = (data.roomTransfers || []).filter(t => t.oldRoomId === room.id || t.newRoomId === room.id);
  const roomMoveOuts = (data.moveOutReports || []).filter(r => r.roomId === room.id);
  const contractRenewals = contract ? (data.contractRenewals || []).filter(r => r.contractId === contract.id) : [];
  const roomAssets = room.assets || [];
  const maintenanceItems = room.maintenanceTickets || [];
  const attachments = room.attachments || contract?.attachments || [];
  const paymentDay = Number(contract?.paymentCycleDay || 5);
  const dueDate = new Date();
  dueDate.setDate(paymentDay);
  dueDate.setHours(0, 0, 0, 0);
  const receiptPaid = Number(currentReceipt?.paidAmount || 0);
  const receiptTotal = Number(currentReceipt?.total || 0);
  const receiptDebt = Math.max(0, receiptTotal - receiptPaid);
  const isReceiptOverdue = currentReceipt && receiptDebt > 0 && new Date() > dueDate;
  const latestReceipt = roomReceipts.find(r => r.type === 'monthly') || null;
  const electricOld = latestReceipt ? getElectricOld(latestReceipt) : (room.electricOld ?? room.electricStart ?? room.initialElectric ?? 0);
  const electricNew = latestReceipt ? getElectricNew(latestReceipt) : (room.electricNew ?? electricOld);
  const waterOld = latestReceipt ? getWaterOld(latestReceipt) : (room.waterOld ?? room.waterStart ?? room.initialWater ?? 0);
  const waterNew = latestReceipt ? getWaterNew(latestReceipt) : (room.waterNew ?? waterOld);
  const tenantInitial = primaryTenant?.name ? primaryTenant.name.trim().charAt(0).toUpperCase() : 'P';
  const formatDate = formatContractDate;
  const dateDiffDays = (value) => {
    if (isInvalidContractDate(value)) return null;
    const date = parseDateFlexible(value);
    if (!date) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.ceil((date.getTime() - today.getTime()) / 86400000);
  };
  const contractStartValid = !contract || ownerOccupied || !isInvalidContractDate(contract.startDate);
  const contractEndValid = !contract || ownerOccupied || !isInvalidContractDate(contract.endDate);
  const contractSignedValid = !contract || ownerOccupied || !isInvalidContractDate(contract.signedDate);
  const contractDateValid = !contract || ownerOccupied || isValidContractRange(contract.startDate, contract.endDate);
  const contractRequiredDatesValid = ownerOccupied || !contract || (contractSignedValid && contractDateValid);
  const contractDuration = contract ? calculateRentalDuration(contract.startDate, contract.endDate) : '—';
  const contractDaysLeft = dateDiffDays(contract?.endDate);
  const contractStatus = ownerOccupied ? 'Chủ nhà ở - không cần hợp đồng' : !contract ? 'Chưa có hợp đồng' : !contractRequiredDatesValid ? 'Thiếu ngày hợp lệ' : contractDaysLeft < 0 ? 'Đã hết hạn' : contractDaysLeft <= 30 ? `Sắp hết hạn (${contractDaysLeft} ngày)` : 'Còn hạn';
  const contractStatusClass = ownerOccupied ? 'active' : !contract ? 'vacant' : !contractRequiredDatesValid ? 'notice' : contractDaysLeft !== null && contractDaysLeft < 0 ? 'debt' : contractDaysLeft !== null && contractDaysLeft <= 30 ? 'notice' : 'active';
  const receiptStatus = !currentReceipt ? 'Chưa có phiếu' : isReceiptOverdue ? 'Quá hạn' : currentReceipt.status;
  const receiptStatusClass = receiptStatus === 'Đã thanh toán' ? 'active' : receiptStatus === 'Nợ một phần' ? 'notice' : receiptStatus === 'Quá hạn' ? 'debt' : 'debt';
  const occupantLimit = room.maxPeople || room.capacity || '—';
  const formatRenewalRange = (renewal) => {
    if (isInvalidContractDate(renewal.oldEndDate)) return `Gia hạn đến ${formatDate(renewal.newEndDate)}`;
    return `${formatDate(renewal.oldEndDate)} -> ${formatDate(renewal.newEndDate)}`;
  };
  const activityItems = [
    ...roomReceipts.slice(0, 5).map(r => ({ id: r.id, type: 'receipt', date: r.savedAt || r.createdAt, actor: 'Hệ thống', title: `Phiếu ${r.month}`, detail: `${formatMoney(r.total)} - ${r.status}` })),
    ...roomReceipts.filter(r => Number(r.paidAmount || 0) > 0).slice(0, 5).map(r => ({ id: `pay-${r.id}`, type: 'payment', date: r.paidAt || r.savedAt || r.createdAt, actor: r.collector || 'Admin', title: 'Ghi nhận thanh toán', detail: `Đã thu ${formatMoney(r.paidAmount || 0)} cho phiếu ${r.month}` })),
    ...contractRenewals.map(r => ({ id: r.id, type: 'contract', date: r.createdAt, actor: 'Admin', title: 'Gia hạn hợp đồng', detail: formatRenewalRange(r) })),
    ...roomTransfers.map(t => ({ id: t.id, type: 'room', date: t.createdAt, actor: 'Admin', title: 'Đổi phòng', detail: `P${t.oldRoomId} -> P${t.newRoomId}` })),
    ...roomMoveOuts.map(r => ({ id: r.id, type: 'room', date: r.createdAt, actor: 'Admin', title: 'Tất toán', detail: formatDisplayDate(r.actualEndDate) })),
    ...roomAssets.map(a => ({ id: a.id || a.name, type: 'asset', date: a.createdAt || a.handoverDate, actor: 'Admin', title: 'Tài sản phòng', detail: `${a.name} - ${a.currentStatus || a.status || 'Đang dùng'}` })),
    ...maintenanceItems.map(m => ({ id: m.id || m.title, type: 'maintenance', date: m.createdAt || m.createdDate, actor: m.assignee || 'Admin', title: 'Bảo trì', detail: `${m.title} - ${m.status || 'Mới tạo'}` })),
    ...attachments.map(f => ({ id: f.id || f.name, type: 'file', date: f.createdAt || f.uploadedAt, actor: f.uploader || 'Admin', title: 'Tệp đính kèm', detail: `${f.group || f.type || 'Khác'} - ${f.name}` }))
  ].sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0)).slice(0, 8);
  const filteredActivityItems = historyFilter === 'all' ? activityItems : activityItems.filter(item => item.type === historyFilter);
  const electricUsage = Math.max(0, electricNew - electricOld);
  const waterUsage = Math.max(0, waterNew - waterOld);
  const electricValue = electricUsage === 0 ? 'Chưa ghi chỉ số mới' : `${formatLocaleNumber(electricOld)} -> ${formatLocaleNumber(electricNew)}`;
  const waterValue = waterUsage === 0 ? 'Chưa ghi chỉ số mới' : `${formatLocaleNumber(waterOld)} -> ${formatLocaleNumber(waterNew)}`;
  const electricNote = electricUsage === 0 ? `Chỉ số gần nhất: ${formatLocaleNumber(electricNew)}` : `Dùng ${formatLocaleNumber(electricUsage)} kWh`;
  const waterNote = waterUsage === 0 ? `Chỉ số gần nhất: ${formatLocaleNumber(waterNew)}` : `Dùng ${formatLocaleNumber(waterUsage)} m3`;
  const electricMeterNote = electricUsage === 0 ? 'Chưa ghi chỉ số mới' : `Tiêu thụ ${formatLocaleNumber(electricUsage)} kWh`;
  const waterMeterNote = waterUsage === 0 ? 'Chưa ghi chỉ số mới' : `Tiêu thụ ${formatLocaleNumber(waterUsage)} m3`;
  const tabs = [
    ['overview', 'Tổng quan', ''],
    ['contract', 'Hợp đồng', !contractRequiredDatesValid ? '!' : ''],
    ['residents', 'Người ở', activeMembers.length || ''],
    ['payments', 'Thanh toán', receiptDebt > 0 ? 'Nợ' : ''],
    ['meters', 'Điện nước', ''],
    ['assets', 'Tài sản', roomAssets.length || ''],
    ['maintenance', 'Bảo trì', maintenanceItems.length || ''],
    ['history', 'Lịch sử', activityItems.length || ''],
    ['files', 'Tệp', attachments.length || '']
  ];

  const renderMetric = (labelText, value, note) => (
    <div className="op-item">
      <span className="op-label">{labelText}</span>
      <span className="op-value">{value}</span>
      {note && <span className="small muted">{note}</span>}
    </div>
  );

  const renderOverview = () => (
    <div className="stack" style={{ gap: '16px' }}>
      <div className="op-grid">
        <div className="op-card" style={{ gridColumn: 'span 2', background: '#f8fafc' }}>
          <h3 className="op-card-title">Tổng quan nhanh</h3>
          <div className="op-grid">
            {renderMetric('Tiền phải thu', currentReceipt ? formatMoney(receiptTotal) : 'Chưa có phiếu', null)}
            {renderMetric('Đã thu', currentReceipt ? formatMoney(receiptPaid) : formatMoney(0), currentReceipt ? receiptStatus : 'Chưa ghi nhận')}
            {renderMetric('Còn nợ', formatMoney(receiptDebt), receiptDebt > 0 ? `Hạn thu ngày ${paymentDay}` : 'Không còn nợ')}
            {renderMetric('Hợp đồng đến', contract ? formatDate(contract.endDate) : '—', contractStatus)}
            {renderMetric('Số người ở', `${activeMembers.length}/${occupantLimit}`, 'Đang cư trú')}
            {renderMetric('Trạng thái cư trú', label, primaryTenant ? primaryTenant.name : 'Phòng trống')}
          </div>
        </div>
      </div>
      <div className="op-grid">
      <div className="op-card">
        <h3 className="op-card-title">Người đại diện</h3>
        {primaryTenant ? (
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <div style={{ width: '44px', height: '44px', borderRadius: '50%', background: '#dbeafe', color: '#1d4ed8', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800 }}>{tenantInitial}</div>
            <div className="stack" style={{ gap: '4px' }}>
              <p className="op-value">{primaryTenant.name}</p>
              <a className="op-label" href={`tel:${primaryTenant.phone || ''}`}>{primaryTenant.phone || 'Chưa có SĐT'}</a>
              <p className="op-label">CCCD: {primaryTenant.cccd || '—'}</p>
            </div>
          </div>
        ) : <p className="muted small">Phòng đang trống</p>}
      </div>

      <div className="op-card">
        <h3 className="op-card-title">Hợp đồng</h3>
        {contract ? (
          <div className="stack" style={{ gap: '8px' }}>
            <span className={`status-badge-liquid ${contractStatusClass}`} style={{ alignSelf: 'flex-start' }}>{contractStatus}</span>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Bắt đầu</span><span className="op-value" style={{ fontSize: '13px' }}>{formatDate(contract.startDate)}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Hết hạn</span><span className="op-value" style={{ fontSize: '13px' }}>{formatDate(contract.endDate)}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Tiền thuê</span><span className="op-value" style={{ fontSize: '13px' }}>{formatMoney(contract.rent)}</span></div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}><span className="op-label">Chu kỳ thu</span><span className="op-value" style={{ fontSize: '13px' }}>Ngày {paymentDay}</span></div>
          </div>
        ) : <p className="muted small">Chưa có hợp đồng</p>}
      </div>

      <div className="op-card" style={{ gridColumn: 'span 2' }}>
        <h3 className="op-card-title">Công nợ tháng {currentMonth}</h3>
        {currentReceipt ? (
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <p className="op-value">{formatMoney(currentReceipt.total)}</p>
              <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', marginTop: '6px' }}>
                <span className={`status-badge-liquid ${receiptStatusClass}`}>{receiptStatus}</span>
                <span className="small muted">Còn nợ {formatMoney(receiptDebt)}</span>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button className="primary-btn sm" onClick={() => onAction('pay_receipt', currentReceipt)}>Ghi nhận</button>
              <button className="secondary-btn sm" onClick={() => onAction('view_qr', currentReceipt)}>Xem QR</button>
            </div>
          </div>
        ) : <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <p className="muted small">Chưa có phiếu tháng này</p>
              {contract && <button className="primary-btn sm" onClick={() => onAction('create_receipt')}>Tạo ngay</button>}
            </div>}
      </div>

      <div className="op-card" style={{ gridColumn: 'span 2' }}>
        <h3 className="op-card-title">Điện nước gần nhất</h3>
        <div className="op-grid">
          {renderMetric('Điện', electricValue, electricNote)}
          {renderMetric('Nước', waterValue, waterNote)}
        </div>
      </div>
      </div>
    </div>
  );

  const renderReceiptBreakdown = (receipt) => {
    if (!receipt) return <p className="muted center" style={{ padding: '24px' }}>Chưa có phiếu tháng hiện tại.</p>;
    const rows = [
      ['Tiền phòng', '1 tháng', contract?.rent || room.rent || 0, receipt.rent || 0],
      ['Dịch vụ cố định', '1', 'Theo số người', receipt.fixedServices || 0],
      ['Tiền điện', `${formatLocaleNumber(receipt.electricUsed || 0)} số`, room.electricPrice || 0, receipt.electricAmount || 0],
      ['Tiền nước', `${formatLocaleNumber(receipt.waterUsed || 0)} số`, room.waterPrice || 0, receipt.waterAmount || 0],
      ['Phí phát sinh', '1', 'Khác', receipt.other || 0]
    ];
    return (
      <div className="table-wrap">
        <table>
          <thead><tr><th>Khoản thu</th><th>Số lượng</th><th>Đơn giá</th><th>Thành tiền</th></tr></thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx}><td>{row[0]}</td><td>{row[1]}</td><td>{typeof row[2] === 'number' ? formatMoney(row[2]) : row[2]}</td><td style={{ fontWeight: 700 }}>{formatMoney(row[3])}</td></tr>
            ))}
            <tr><td colSpan="3" style={{ textAlign: 'right', fontWeight: 800 }}>Tổng phải thu</td><td style={{ fontWeight: 800, color: 'var(--primary)' }}>{formatMoney(receipt.total || 0)}</td></tr>
          </tbody>
        </table>
      </div>
    );
  };

  const renderEmptyOps = (title, detail, actionLabel, actionHandler) => (
    <div className="op-card" style={{ alignItems: 'center', textAlign: 'center', padding: '32px' }}>
      <p className="op-value">{title}</p>
      <p className="op-label">{detail}</p>
      {actionLabel && <button className="primary-btn sm" onClick={actionHandler}>{actionLabel}</button>}
    </div>
  );

  const renderTabContent = () => {
    if (activeTab === 'overview') return renderOverview();
    if (activeTab === 'contract') return (
      <div className="stack" style={{ gap: '16px' }}>
        {ownerOccupied ? (
          <>
            <div className="op-card">
              <h3 className="op-card-title">Thông tin cư trú</h3>
              <div className="op-grid">
                {renderMetric('Loại phòng', 'Chủ nhà ở')}
                {renderMetric('Trạng thái', 'Đang ở')}
                {renderMetric('Hợp đồng', 'Không áp dụng')}
                {renderMetric('Số người ở', `${activeMembers.length}/${occupantLimit}`)}
                {renderMetric('Ghi chú', room.note || 'Chủ nhà ở, không cần hợp đồng thuê')}
              </div>
            </div>
            <div className="warning-box" style={{ background: '#ecfdf5', borderColor: '#10b981', color: '#065f46' }}>
              Phòng {room.id} là phòng chủ nhà ở nên không yêu cầu ngày ký, ngày bắt đầu, ngày hết hạn hoặc phụ lục hợp đồng.
            </div>
          </>
        ) : contract ? (
          <>
            <div className="op-card">
              <h3 className="op-card-title">Thông tin hợp đồng</h3>
              <div className="op-grid">
                {renderMetric('Số hợp đồng', contract.contractNo || contract.id)}
                {renderMetric('Trạng thái', contractStatus)}
                {renderMetric('Ngày ký', formatDate(contract.signedDate), 'Ngày cố định dùng để in hợp đồng')}
                {renderMetric('Ngày bắt đầu', formatDate(contract.startDate))}
                {renderMetric('Ngày hết hạn', formatDate(contract.endDate))}
                {renderMetric('Thời hạn thuê', contractDuration)}
                {renderMetric('Chu kỳ thanh toán', `Ngày ${paymentDay} hằng tháng`)}
                {renderMetric('Tiền thuê', formatMoney(contract.rent))}
                {renderMetric('Tiền cọc', formatMoney(contract.deposit))}
              </div>
            </div>
            {!contractRequiredDatesValid && (
              <div className="warning-box">Hợp đồng thiếu ngày ký / ngày bắt đầu / ngày hết hạn hợp lệ. Vui lòng cập nhật để in hợp đồng, tính đúng thời hạn thuê và phụ lục gia hạn.</div>
            )}
            <div className="op-card">
              <h3 className="op-card-title">Lịch sử gia hạn</h3>
              {contractRenewals.length ? contractRenewals.map(r => (
                <div key={r.id} className="op-item"><span className="op-value">{formatRenewalRange(r)}</span><span className="op-label">{r.note || 'Không có ghi chú'}</span></div>
              )) : <p className="muted small">Chưa có lịch sử gia hạn.</p>}
            </div>
            <div className="op-action-footer" style={{ marginTop: 0 }}>
              <button className="secondary-btn" onClick={() => onAction('edit_contract')}>Sửa hợp đồng</button>
              <button className="secondary-btn" onClick={() => onAction('renew_contract')}>Gia hạn</button>
              <button className="secondary-btn" onClick={() => onAction('view_contract')}>In hợp đồng</button>
              {contractRenewals.length > 0 && <button className="secondary-btn" onClick={() => onAction('print_appendix')}>In phụ lục</button>}
              <button className="secondary-btn danger" onClick={() => onAction('moving_out')}>Kết thúc thuê</button>
            </div>
          </>
        ) : renderEmptyOps('Chưa có hợp đồng', 'Phòng trống hoặc chưa tạo hợp đồng.', '+ Tạo hợp đồng', () => onAction('add_tenant'))}
      </div>
    );
    if (activeTab === 'residents') return (
      <div className="stack" style={{ gap: '12px' }}>
        {['primary', 'member'].map(group => {
          const rows = allMembers.filter(m => group === 'primary' ? m.role === 'primary' : m.role !== 'primary');
          if (!rows.length) return null;
          return (
            <div key={group} className="op-card">
              <h3 className="op-card-title">{group === 'primary' ? 'Người đứng tên' : 'Người ở cùng'}</h3>
              {rows.map(m => (
                <div key={m.id} className="op-row">
                  <div>
                    <p className="op-value">{m.tenant?.name || 'N/A'}</p>
                    <p className="op-label">{m.tenant?.phone || '—'} • CCCD {m.tenant?.cccd || '—'} • Vào {formatDate(m.joinedDate || contract?.startDate)}{m.leftDate ? ` • Ra ${formatDisplayDate(m.leftDate)}` : ''}</p>
                    <p className="op-label">Xe máy: <b>{m.tenant?.licensePlate || m.tenant?.vehicle || 'Chưa cập nhật'}</b> • Vân tay: <b>{m.tenant?.fingerprintCode || 'Chưa đăng ký'}</b> • {m.tenant?.fingerprintStatus || 'Chưa đăng ký'}</p>
                  </div>
                  <div className="row-actions">
                    {m.tenant?.fingerprintStatus === 'Cần xóa' && <span className="status-badge-liquid debt">Cần xóa vân tay</span>}
                    <span className={`status-badge-liquid ${m.status === 'active' ? 'active' : 'vacant'}`}>{m.status === 'active' ? 'Đang ở' : 'Đã rời đi'}</span>
                    {m.tenant && <button className="secondary-btn sm" onClick={() => onAction('edit_tenant', m.tenant)}>Sửa</button>}
                  </div>
                </div>
              ))}
            </div>
          );
        })}
        {contract && (
          <button
            type="button"
            className="secondary-btn wide"
            style={{ borderStyle: 'dashed' }}
            onClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setShowRoommateForm(true);
            }}
          >
            + Thêm người ở
          </button>
        )}
        {!allMembers.length && renderEmptyOps('Chưa có người ở', 'Phòng đang trống.', '+ Thuê mới', () => onAction('add_tenant'))}
      </div>
    );
    if (activeTab === 'payments') return (
      <div className="stack" style={{ gap: '16px' }}>
        <div className="op-card">
          <h3 className="op-card-title">Chi tiết phiếu tháng {currentMonth}</h3>
          {renderReceiptBreakdown(currentReceipt)}
        </div>
        <div className="op-card" style={{ padding: 0 }}>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Phiếu</th><th>Tổng tiền</th><th>Đã thu</th><th>Còn nợ</th><th>Hạn thanh toán</th><th>Trạng thái</th><th>Thao tác</th></tr></thead>
              <tbody>
                {roomReceipts.map(r => {
                  const paymentState = getReceiptPaymentState(r);
                  const debt = paymentState.debt;
                  const status = paymentState.status === 'Nợ một phần' ? 'Thanh toán một phần' : paymentState.status;
                  const dueText = `${String(paymentDay).padStart(2, '0')}/${String(r.month || getCurrentMonthLabel()).padStart(7, '0')}`;
                  return (
                    <tr key={r.id}>
                      <td>Phiếu {r.month}</td>
                      <td>{formatMoney(r.total || 0)}</td>
                      <td>{formatMoney(paymentState.paidAmount || 0)}</td>
                      <td style={{ color: debt > 0 ? 'var(--danger)' : 'var(--success)', fontWeight: 700 }}>{formatMoney(debt)}</td>
                      <td>{dueText}</td>
                      <td><span className={`status-badge-liquid ${status === 'Đã thanh toán' ? 'active' : status === 'Thanh toán một phần' ? 'notice' : 'debt'}`}>{status}</span></td>
                      <td><div className="row-actions"><button className="primary-btn sm" onClick={() => onAction('pay_receipt', r)}>Ghi nhận</button><button className="secondary-btn sm" onClick={() => onAction('view_qr', r)}>QR</button><button className="secondary-btn sm" onClick={() => onAction('view_qr', r)}>Chi tiết</button><button className="secondary-btn sm" onClick={() => window.print()}>Biên nhận</button></div></td>
                    </tr>
                  );
                })}
                {!roomReceipts.length && <tr><td colSpan="7" className="center muted" style={{ padding: '20px' }}>Chưa có dữ liệu thanh toán.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
    if (activeTab === 'meters') return (
      <div className="stack" style={{ gap: '16px' }}>
        <div className="op-card">
          <h3 className="op-card-title">Chỉ số hiện tại</h3>
          <div className="op-grid">
            {renderMetric('Điện', `${formatLocaleNumber(electricOld)} -> ${formatLocaleNumber(electricNew)}`, electricMeterNote)}
            {renderMetric('Nước', `${formatLocaleNumber(waterOld)} -> ${formatLocaleNumber(waterNew)}`, waterMeterNote)}
            {renderMetric('Đơn giá điện', formatMoney(room.electricPrice || 0))}
            {renderMetric('Đơn giá nước', formatMoney(room.waterPrice || 0))}
          </div>
          <button className="primary-btn sm" style={{ alignSelf: 'flex-start' }} onClick={() => onAction('record_meter')}>+ Ghi chỉ số điện nước</button>
        </div>
        <div className="op-card" style={{ padding: 0 }}>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Tháng</th><th>Chỉ số cũ</th><th>Chỉ số mới</th><th>Tiêu thụ</th><th>Đơn giá</th><th>Thành tiền</th><th>Ngày ghi</th><th>Người ghi</th></tr></thead>
              <tbody>
                {roomReceipts.filter(r => r.type === 'monthly').map(r => (
                  <React.Fragment key={r.id}>
                    <tr><td>Điện {r.month}</td><td>{formatLocaleNumber(getElectricOld(r))}</td><td>{formatLocaleNumber(getElectricNew(r))}</td><td>{formatLocaleNumber(r.electricUsed || 0)}</td><td>{formatMoney(room.electricPrice || 0)}</td><td>{formatMoney(r.electricAmount || 0)}</td><td>{formatDisplayDate((r.savedAt || r.createdAt || '').slice(0, 10))}</td><td>{r.reader || 'Admin'}</td></tr>
                    <tr><td>Nước {r.month}</td><td>{formatLocaleNumber(getWaterOld(r))}</td><td>{formatLocaleNumber(getWaterNew(r))}</td><td>{formatLocaleNumber(r.waterUsed || 0)}</td><td>{formatMoney(room.waterPrice || 0)}</td><td>{formatMoney(r.waterAmount || 0)}</td><td>{formatDisplayDate((r.savedAt || r.createdAt || '').slice(0, 10))}</td><td>{r.reader || 'Admin'}</td></tr>
                  </React.Fragment>
                ))}
                {!roomReceipts.length && <tr><td colSpan="8" className="center muted" style={{ padding: '20px' }}>Chưa có lịch sử chỉ số.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
    if (activeTab === 'assets') return roomAssets.length ? (
      <div className="stack" style={{ gap: '12px' }}>
        <div className="op-action-footer" style={{ marginTop: 0 }}><button className="primary-btn sm" onClick={() => onAction('add_asset')}>+ Thêm tài sản</button><button className="secondary-btn sm">Nhập từ mẫu</button><button className="secondary-btn sm" onClick={() => window.print()}>In bàn giao</button></div>
        <div className="table-wrap"><table><thead><tr><th>Tài sản</th><th>Loại</th><th>SL</th><th>Bàn giao</th><th>Hiện tại</th><th>Giá trị</th><th>Ghi chú</th><th>Thao tác</th></tr></thead><tbody>{roomAssets.map((asset, idx) => <tr key={asset.id || idx}><td>{asset.name}</td><td>{asset.category || asset.type || '—'}</td><td>{asset.quantity || 1}</td><td>{asset.handoverStatus || asset.status || 'Tốt'}</td><td>{asset.currentStatus || asset.status || 'Đang dùng'}</td><td>{formatMoney(asset.value || 0)}</td><td>{asset.note || '—'}</td><td><div className="row-actions"><button className="secondary-btn sm">Sửa</button><button className="secondary-btn sm">Thêm ảnh</button></div></td></tr>)}</tbody></table></div>
      </div>
    ) : renderEmptyOps('Chưa có danh sách tài sản', 'Thêm giường, tủ, máy lạnh, đồng hồ điện/nước hoặc ảnh bàn giao phòng.', '+ Thêm tài sản', () => onAction('add_asset'));
    if (activeTab === 'maintenance') return maintenanceItems.length ? (
      <div className="stack" style={{ gap: '12px' }}>
        <button className="primary-btn sm" style={{ alignSelf: 'flex-start' }} onClick={() => onAction('add_maintenance')}>+ Tạo yêu cầu bảo trì</button>
        {maintenanceItems.map((item, idx) => <div key={item.id || idx} className="op-card"><div className="op-row"><div><p className="op-value">{item.title}</p><p className="op-label">{item.category || 'Khác'} • {item.priority || 'Trung bình'} • {item.reporter || 'Khách thuê'} báo • {item.assignee || 'Chưa phân công'}</p><p className="op-label">Chi phí {formatMoney(item.cost || 0)} • {item.costOwner || 'Chưa rõ'} chịu phí • {item.note || 'Không có ghi chú'}</p></div><span className={`status-badge-liquid ${item.status === 'Hoàn tất' ? 'active' : item.status === 'Đã hủy' ? 'vacant' : 'notice'}`}>{item.status || 'Mới tạo'}</span></div></div>)}
      </div>
    ) : renderEmptyOps('Chưa có sự cố bảo trì', 'Các sự cố điện, nước, khóa cửa và thiết bị sẽ được ghi tại đây.', '+ Tạo yêu cầu bảo trì', () => onAction('add_maintenance'));
    if (activeTab === 'history') return activityItems.length ? (
      <div className="stack" style={{ gap: '12px' }}>
        <div className="btn-group" style={{ margin: 0, flexWrap: 'wrap' }}>
          {[
            ['all', 'Tất cả'],
            ['receipt', 'Phiếu tháng'],
            ['payment', 'Thanh toán'],
            ['contract', 'Hợp đồng'],
            ['room', 'Phòng'],
            ['resident', 'Người ở'],
            ['meter', 'Điện nước'],
            ['asset', 'Tài sản'],
            ['maintenance', 'Bảo trì'],
            ['file', 'Tệp']
          ].map(([id, text]) => <button key={id} className={`secondary-btn sm ${historyFilter === id ? 'active-tab' : ''}`} onClick={() => setHistoryFilter(id)}>{text}</button>)}
        </div>
        {filteredActivityItems.map(item => <div key={`${item.title}-${item.id}`} className="op-card"><p className="op-value">{item.title}</p><p className="op-label">{formatDisplayDate(String(item.date || '').slice(0, 10))} • {item.actor} • {item.detail}</p></div>)}
        {!filteredActivityItems.length && renderEmptyOps('Chưa có lịch sử phù hợp', 'Thử chọn bộ lọc khác để xem thêm hoạt động.')}
      </div>
    ) : renderEmptyOps('Chưa có lịch sử hoạt động', 'Các lần lập phiếu, đổi phòng, gia hạn và tất toán sẽ hiển thị tại đây.');
    if (activeTab === 'files') return attachments.length ? (
      <div className="stack" style={{ gap: '12px' }}><button className="primary-btn sm" style={{ alignSelf: 'flex-start' }} onClick={() => onAction('add_file')}>+ Tải tệp lên</button>{attachments.map((file, idx) => <div key={file.id || idx} className="op-card"><div className="op-row"><div><p className="op-value">{file.name || `Tệp ${idx + 1}`}</p><p className="op-label">{file.group || file.type || 'Khác'} • {file.size || '—'} • {file.uploader || 'Admin'} • {formatDisplayDate(file.uploadedAt || String(file.createdAt || '').slice(0, 10))}</p></div><div className="row-actions"><button className="secondary-btn sm">Xem</button><button className="secondary-btn sm">Tải xuống</button><button className="secondary-btn sm danger">Xóa</button></div></div></div>)}</div>
    ) : renderEmptyOps('Chưa có tệp đính kèm', 'Có thể tải hợp đồng, CCCD, ảnh đồng hồ và ảnh bàn giao phòng.', '+ Tải tệp lên', () => onAction('add_file'));
    return null;
  };

  return (
    <>
      <div className="modal" onClick={onClose}>
        <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '980px' }} onClick={e => e.stopPropagation()}>
          <div className="modal-header">
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <h2 style={{ fontSize: '28px' }}>Phòng {room.id}</h2>
                <span className={`status-badge-liquid ${color === 'green' ? 'active' : color === 'gray' ? 'vacant' : color}`}>{label}</span>
                {receiptDebt > 0 && <span className="status-badge-liquid debt">Nợ {formatMoney(receiptDebt)}</span>}
              </div>
              <p className="muted">{primaryTenant ? `${primaryTenant.name} • ${primaryTenant.phone}` : ownerOccupied ? 'Chủ nhà ở, không cần hợp đồng' : 'Chưa có hợp đồng hiện tại'}</p>
            </div>
            <button className="secondary-btn" onClick={onClose}>✕</button>
          </div>

        <div className="detail-body-v2">
          <nav className="room-op-tabs">
            {tabs.map(([id, text, badge]) => (
              <button key={id} className={`tab-link ${activeTab === id ? 'active' : ''}`} onClick={() => setActiveTab(id)}>
                {text}{badge !== '' && <span className={`tab-badge ${badge === '!' || badge === 'Nợ' ? 'warn' : ''}`}>{badge}</span>}
              </button>
            ))}
          </nav>

          <div className="tab-content" style={{ minHeight: '300px' }}>
            {renderTabContent()}
          </div>

          <div className="op-action-footer">
            {label === 'Trống' ? (
              <><button className="primary-btn" style={{ flex: 1 }} onClick={() => onAction('add_tenant')}>+ Thuê mới</button><button className="secondary-btn" style={{ flex: 1 }} onClick={() => onAction('view_history')}>📜 Xem lịch sử thuê</button></>
            ) : (
              <>
                {!ownerOccupied && <button className="primary-btn" onClick={() => onAction('create_receipt')}>⚡ Lập phiếu tháng</button>}
                {currentReceipt && <button className="primary-btn" style={{ background: 'var(--success)' }} onClick={() => onAction('pay_receipt', currentReceipt)}>Ghi nhận thanh toán</button>}
                {currentReceipt && <button className="secondary-btn" onClick={() => onAction('view_qr', currentReceipt)}>Xem QR</button>}
                <div style={{ position: 'relative' }}>
                  <button className="secondary-btn" onClick={() => setShowMoreActions(!showMoreActions)}>⋯ Thao tác khác</button>
                  {showMoreActions && (
                    <div className="more-actions-menu">
                      {!ownerOccupied && <button onClick={() => { setShowMoreActions(false); onAction('edit_contract'); }}>Sửa hợp đồng</button>}
                      {!ownerOccupied && <button onClick={() => { setShowMoreActions(false); onAction('renew_contract'); }}>Gia hạn hợp đồng</button>}
                      {!ownerOccupied && <button className="warning" onClick={() => { setShowMoreActions(false); onAction('transfer_room'); }}>Đổi phòng</button>}
                      {!ownerOccupied && <button className="warning" onClick={() => { setShowMoreActions(false); onAction(label === 'Báo chuyển' ? 'cancel_notice' : 'notice'); }}>{label === 'Báo chuyển' ? 'Hủy báo chuyển' : 'Báo chuyển'}</button>}
                      {!ownerOccupied && <button onClick={() => { setShowMoreActions(false); onAction('view_contract'); }}>Xuất PDF</button>}
                      {!ownerOccupied && <button onClick={() => { setShowMoreActions(false); onAction('view_contract'); }}>In hợp đồng</button>}
                      {!ownerOccupied && <button onClick={() => { setShowMoreActions(false); onAction('print_appendix'); }}>In phụ lục</button>}
                      {!ownerOccupied && <button className="danger" onClick={() => { setShowMoreActions(false); onAction('moving_out'); }}>Tất toán</button>}
                      {!ownerOccupied && <button className="danger" onClick={() => { setShowMoreActions(false); onAction('moving_out'); }}>Kết thúc thuê</button>}
                      {ownerOccupied && <button onClick={() => { setShowMoreActions(false); setActiveTab('residents'); }}>Quản lý người ở</button>}
                      {ownerOccupied && <button onClick={() => { setShowMoreActions(false); setActiveTab('history'); }}>Xem lịch sử</button>}
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
      {showRoommateForm && contract && (
        <RoommateModal
          room={room}
          contract={contract}
          onClose={() => setShowRoommateForm(false)}
          onSave={(result) => {
            const saved = onAddRoommate ? onAddRoommate(result) : false;
            if (saved !== false) {
              setShowRoommateForm(false);
              setActiveTab('residents');
            }
            return saved;
          }}
        />
      )}
    </>
  );
}

function EditContractModal({ contract, data, onClose, onSave }) {
  const [form, setForm] = useState({
    ...contract,
    signedDate: formatDateForInput(contract.signedDate),
    startDate: formatDateForInput(contract.startDate),
    endDate: formatDateForInput(contract.endDate)
  });
  const errors = {};
  if (!form.signedDate) errors.signedDate = 'Vui lòng nhập ngày ký hợp đồng';
  if (!form.startDate) errors.startDate = 'Vui lòng nhập ngày bắt đầu hợp đồng';
  if (!form.endDate) errors.endDate = 'Vui lòng nhập ngày hết hạn hợp đồng';
  if (form.startDate && form.endDate && form.endDate <= form.startDate) errors.endDate = 'Ngày hết hạn phải sau ngày bắt đầu';
  const dateWarning = Object.keys(errors).length > 0;
  const handleSave = () => {
    if (dateWarning) {
      alert(Object.values(errors).join('\n'));
      return;
    }
    onSave(form);
  };
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
            <label>Ngày ký <input type="date" value={form.signedDate || ''} onChange={e => setForm({...form, signedDate: e.target.value})} /><span className="small muted">Ngày ký cố định, không tự cập nhật theo hôm nay.</span></label>
            <label>Ngày bắt đầu <input type="date" value={form.startDate} onChange={e => setForm({...form, startDate: e.target.value})} /></label>
            <label>Ngày hết hạn <input type="date" value={form.endDate} onChange={e => setForm({...form, endDate: e.target.value})} /></label>
            <label>Giá thuê <input type="number" value={form.rent} onChange={e => setForm({...form, rent: e.target.value})} /></label>
            <label>Tiền đặt cọc <input type="number" value={form.deposit} onChange={e => setForm({...form, deposit: e.target.value})} /></label>
            <label style={{ gridColumn: 'span 2' }}>Ghi chú <textarea value={form.note || ''} onChange={e => setForm({...form, note: e.target.value})} /></label>
          </div>
          {dateWarning && <div className="warning-box" style={{ marginTop: '16px' }}>Hợp đồng thiếu ngày ký / ngày bắt đầu / ngày hết hạn hợp lệ. Vui lòng cập nhật để in hợp đồng, tính đúng thời hạn thuê và phụ lục gia hạn.</div>}
          <button className="primary-btn wide" style={{ marginTop: '24px' }} onClick={handleSave}>Cập nhật hợp đồng</button>
        </div>
      </div>
    </div>
  );
}

function RentalFlowModal({ room, onClose, onSave }) {
  const [step, setStep] = useState('form'); // 'form' | 'preview'
  const [isSaving, setIsSaving] = useState(false);
  const [form, setForm] = useState({
    // Tenant Info
    tenantName: '',
    tenantPhone: '',
    tenantCCCD: '',
    tenantCCCDDate: '',
    tenantCCCDPlace: '',
    tenantAddress: '',
    tenantVehicle: '',
    tenantFingerprintCode: '',
    tenantFingerprintStatus: 'Chưa đăng ký',
    tenantBirthday: '',
    tenantNote: '',
    tenantIdPaste: '',
    // Contract Info
    contractNo: `HĐ-${room.id}-${new Date().getFullYear()}${(new Date().getMonth()+1).toString().padStart(2,'0')}`,
    signedDate: new Date().toISOString().slice(0, 10),
    startDate: new Date().toISOString().slice(0, 10),
    endDate: '',
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
  const previewBaseFixedServiceTotal = Number(form.services.cleaning || 0) + Number(form.services.elevator || 0) + Number(form.services.laundry || 0) + Number(form.services.internet || 0);
  const previewAppliedFixedServiceTotal = fixedServiceTotal(room, 1);
  const previewFixedServiceDiscount = Math.max(0, previewBaseFixedServiceTotal - previewAppliedFixedServiceTotal);

  const applyPastedIdInfo = (text) => {
    const parsed = parseVietnameseIdCard(text);

    setForm(prev => ({
      ...prev,
      tenantIdPaste: text,
      tenantCCCD: parsed.cccd || prev.tenantCCCD,
      tenantName: parsed.name || prev.tenantName,
      tenantBirthday: parsed.birthday || prev.tenantBirthday,
      tenantAddress: parsed.address || prev.tenantAddress,
      tenantCCCDDate: parsed.issueDate || prev.tenantCCCDDate,
      tenantCCCDPlace: parsed.issuePlace || prev.tenantCCCDPlace,
      tenantNote: parsed.note || prev.tenantNote
    }));
  };

  const handleNext = () => {
    if (!form.tenantName || !form.signedDate || !form.startDate || !form.endDate) {
      alert('Vui lòng nhập đầy đủ tên khách, ngày ký, ngày bắt đầu và ngày kết thúc!');
      return;
    }
    if (form.endDate <= form.startDate) {
      alert('Ngày hết hạn phải sau ngày bắt đầu hợp đồng!');
      return;
    }
    setStep('preview');
  };

  const handleSave = (shouldPrint = false) => {
    if (isSaving) return;
    const tenantId = uid('tenant');
    const contractId = uid('contract');
    setIsSaving(true);
    
    const newTenant = {
      id: tenantId,
      name: form.tenantName,
      phone: form.tenantPhone,
      cccd: form.tenantCCCD,
      cccdDate: form.tenantCCCDDate,
      cccdPlace: form.tenantCCCDPlace,
      address: form.tenantAddress,
      licensePlate: form.tenantVehicle,
      fingerprintCode: form.tenantFingerprintCode,
      fingerprintStatus: form.tenantFingerprintStatus,
      birthday: form.tenantBirthday,
      note: form.tenantNote,
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
      paymentCycleDay: Number(form.paymentCycleDay),
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
      joinedDate: form.startDate,
      leftDate: '',
      createdAt: new Date().toISOString()
    }];

    try {
      const saved = onSave({ tenant: newTenant, contract: newContract, memberships });
      if (saved === false) {
        setIsSaving(false);
        return;
      }
      if (shouldPrint) {
        setTimeout(() => window.print(), 500);
      }
    } catch (error) {
      console.error(error);
      setIsSaving(false);
      alert('Không lưu được hợp đồng. Vui lòng thử lại hoặc kiểm tra console.');
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
              <button className="primary-btn" onClick={() => handleSave(false)} disabled={isSaving}>{isSaving ? 'Đang lưu...' : '💾 Lưu hợp đồng'}</button>
              <button className="primary-btn" style={{ background: 'var(--success)' }} onClick={() => handleSave(true)} disabled={isSaving}>{isSaving ? 'Đang lưu...' : '🖨️ Lưu & In'}</button>
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
                  <p>6. Dịch vụ cố định gốc: <b>{formatMoney(previewBaseFixedServiceTotal)}</b>/tháng. Khách ở một mình được giảm 50% phí dịch vụ cố định, mức áp dụng hiện tại: <b>{formatMoney(previewAppliedFixedServiceTotal)}</b>/tháng{previewFixedServiceDiscount > 0 ? <>, giảm <b>{formatMoney(previewFixedServiceDiscount)}</b>/tháng.</> : <>.</>}</p>
                  <p>7. Chi tiết dịch vụ: Vệ sinh ({formatMoney(form.services.cleaning)}), Thang máy ({formatMoney(form.services.elevator)}), Giặt ({formatMoney(form.services.laundry)}), Internet ({formatMoney(form.services.internet)}).</p>
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
      <div className="detail-modal-v2 liquid-glass rental-contract-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>Tạo hợp đồng mới • Phòng {room.id}</h2>
            <p className="muted">Bước 1: Nhập thông tin người thuê và hợp đồng</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 rental-contract-body">
          <div className="rental-contract-layout">
            <div className="rental-column">
              <section className="rental-form-section">
                <h3 className="form-section-title">👤 Thông tin người thuê</h3>
                <div className="paste-id-card">
                  <label>Dán nhanh thông tin CCCD
                    <textarea
                      value={form.tenantIdPaste}
                      onChange={e => applyPastedIdInfo(e.target.value)}
                      placeholder={'Dán nguyên nội dung CCCD ở đây, hệ thống tự điền: họ tên, số CCCD, ngày sinh, thường trú, ngày cấp...'}
                    />
                  </label>
                </div>
                <div className="rental-subsection-title">CCCD mặt trước</div>
                <div className="rental-form-grid tenant-id-grid">
                  <label className="span-2">Họ và tên <input value={form.tenantName} onChange={e => setForm({...form, tenantName: e.target.value})} placeholder="Nguyễn Văn A" /></label>
                  <label>Số CCCD <input value={form.tenantCCCD} onChange={e => setForm({...form, tenantCCCD: e.target.value})} /></label>
                  <label>Ngày sinh <input type="date" value={form.tenantBirthday} onChange={e => setForm({...form, tenantBirthday: e.target.value})} /></label>
                  <label className="span-2">Địa chỉ thường trú <input value={form.tenantAddress} onChange={e => setForm({...form, tenantAddress: e.target.value})} /></label>
                </div>
                <div className="rental-subsection-title">CCCD mặt sau</div>
                <div className="rental-form-grid tenant-id-grid">
                  <label>Ngày cấp CCCD <input type="date" value={form.tenantCCCDDate} onChange={e => setForm({...form, tenantCCCDDate: e.target.value})} /></label>
                  <label>Nơi cấp <input value={form.tenantCCCDPlace} onChange={e => setForm({...form, tenantCCCDPlace: e.target.value})} /></label>
                  <label className="span-2">Ghi chú CCCD / MRZ <textarea value={form.tenantNote} onChange={e => setForm({...form, tenantNote: e.target.value})} placeholder="Giới tính, quốc tịch, quê quán, hạn CCCD, đặc điểm nhận dạng, MRZ..." /></label>
                </div>
                <div className="rental-subsection-title">Liên hệ & quản lý ra vào</div>
                <div className="rental-form-grid tenant-ops-grid">
                  <label>Số điện thoại <input value={form.tenantPhone} onChange={e => setForm({...form, tenantPhone: e.target.value})} placeholder="09xx..." /></label>
                  <label>Biển số xe <input value={form.tenantVehicle} onChange={e => setForm({...form, tenantVehicle: e.target.value})} /></label>
                  <label>Mã vân tay <input value={form.tenantFingerprintCode} onChange={e => setForm({...form, tenantFingerprintCode: e.target.value})} placeholder="VD: F502-01" /></label>
                  <label>Trạng thái vân tay
                    <select value={form.tenantFingerprintStatus} onChange={e => setForm({...form, tenantFingerprintStatus: e.target.value})}>
                      <option>Chưa đăng ký</option>
                      <option>Đã đăng ký</option>
                      <option>Cần xóa</option>
                      <option>Đã xóa</option>
                    </select>
                  </label>
                </div>
              </section>
            </div>

            <div className="rental-column">
              <section className="rental-form-section">
                <h3 className="form-section-title">📄 Chi tiết hợp đồng</h3>
                <div className="rental-form-grid contract-info">
                  <label>Số hợp đồng <input value={form.contractNo} onChange={e => setForm({...form, contractNo: e.target.value})} /></label>
                  <label>Ngày ký <input type="date" value={form.signedDate} onChange={e => setForm({...form, signedDate: e.target.value})} /><span className="small muted">Ngày ký cố định, không tự cập nhật theo hôm nay.</span></label>
                  <label>Ngày bắt đầu <input type="date" value={form.startDate} onChange={e => setForm({...form, startDate: e.target.value})} /></label>
                  <label>Ngày hết hạn <input type="date" value={form.endDate} onChange={e => setForm({...form, endDate: e.target.value})} /></label>
                  <label>Giá thuê <input type="number" value={form.rent} onChange={e => setForm({...form, rent: e.target.value})} /></label>
                  <label>Tiền đặt cọc <input type="number" value={form.deposit} onChange={e => setForm({...form, deposit: e.target.value})} /></label>
                  <label>Ngày thu tiền <input type="number" value={form.paymentCycleDay} onChange={e => setForm({...form, paymentCycleDay: e.target.value})} /></label>
                </div>
              </section>

              <section className="rental-form-section">
                <h3 className="form-section-title">💡 Dịch vụ & Điều khoản</h3>
                <div className="rental-form-grid service-info">
                  <label>Giá điện <input type="number" value={form.electricPrice} onChange={e => setForm({...form, electricPrice: e.target.value})} /></label>
                  <label>Giá nước <input type="number" value={form.waterPrice} onChange={e => setForm({...form, waterPrice: e.target.value})} /></label>
                  <label className="span-2">Ghi chú thêm <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} style={{ minHeight: '82px' }} /></label>
                </div>
              </section>
            </div>
          </div>
          <div className="rental-contract-footer">
            <button className="secondary-btn" onClick={onClose} style={{ flex: 1 }}>Hủy bỏ</button>
            <button className="primary-btn" onClick={handleNext} style={{ flex: 2 }}>Tiếp theo: Xem trước hợp đồng ➡️</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function TransferRoomModal({ contract, data, onClose, onSave }) {
  const oldRoom = data.rooms.find(r => r.id === contract.roomId);
  const tenant = getPrimaryTenantByContract(data, contract.id) || { name: 'N/A', phone: '' };
  const vacantRooms = (data.rooms || []).filter(r => r.id !== contract.roomId && getRoomStatusInfo(data, r.id).label === 'Trống');
  const firstRoom = vacantRooms[0] || null;
  const [form, setForm] = useState({
    newRoomId: firstRoom?.id || '',
    transferDate: new Date().toISOString().slice(0, 10),
    oldRoomStayFrom: formatDateForInput(contract.startDate) || new Date().toISOString().slice(0, 10),
    oldRoomStayTo: formatDateInputValue(addDays(new Date().toISOString().slice(0, 10), -1)) || new Date().toISOString().slice(0, 10),
    oldRoomServiceMode: 'included_in_transfer_receipt',
    endDate: contract.endDate || addMonthsToDate(new Date().toISOString().slice(0, 10), 12),
    rent: firstRoom?.rent ?? contract.rent,
    deposit: contract.deposit,
    contractNo: `HĐ-${firstRoom?.id || 'ROOM'}-${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, '0')}`,
    note: `Chuyển từ phòng ${contract.roomId}`
  });

  const selectedRoom = data.rooms.find(r => r.id === form.newRoomId);
  const oldRoomDaysPreview = inclusiveDaysBetween(form.oldRoomStayFrom, form.oldRoomStayTo);
  const oldRoomMonth = monthFromDate(form.oldRoomStayTo || form.transferDate);
  const oldRoomMonthDays = daysInBillingMonth(oldRoomMonth);
  const oldRoomRentPreview = Math.round((Number(contract.rent || oldRoom?.rent || 0) / oldRoomMonthDays) * oldRoomDaysPreview);
  const oldRoomServiceMonthlyPreview = fixedServiceTotal(oldRoom, getContractOccupantCount(data, contract) || 1);
  const oldRoomServicePreview = form.oldRoomServiceMode === 'monthly'
    ? oldRoomServiceMonthlyPreview
    : form.oldRoomServiceMode === 'daily'
      ? Math.round((oldRoomServiceMonthlyPreview / oldRoomMonthDays) * oldRoomDaysPreview)
      : 0;

  const handleRoomChange = (roomId) => {
    const room = data.rooms.find(r => r.id === roomId);
    setForm(prev => ({
      ...prev,
      newRoomId: roomId,
      rent: room?.rent ?? prev.rent,
      contractNo: `HĐ-${roomId}-${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, '0')}`
    }));
  };

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '720px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>Đổi phòng • {tenant.name}</h2>
            <p className="muted">Kết thúc hợp đồng phòng cũ và tạo hợp đồng mới cho phòng trống.</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ gap: '20px' }}>
          <div className="op-card">
            <div className="op-grid">
              <div className="op-item"><span className="op-label">Phòng hiện tại</span><span className="op-value">P{oldRoom?.id}</span></div>
              <div className="op-item"><span className="op-label">Người đứng tên</span><span className="op-value">{tenant.name}</span></div>
              <div className="op-item"><span className="op-label">Giá thuê cũ</span><span className="op-value">{formatMoney(contract.rent)}</span></div>
              <div className="op-item"><span className="op-label">Cọc đang giữ</span><span className="op-value">{formatMoney(contract.deposit)}</span></div>
            </div>
          </div>

          {vacantRooms.length === 0 ? (
            <div className="warning-box">Hiện không có phòng trống để chuyển. Cần tất toán hoặc tạo phòng trống trước.</div>
          ) : (
            <>
              <div className="form-grid-v2">
                <label>Phòng mới
                  <select value={form.newRoomId} onChange={e => handleRoomChange(e.target.value)}>
                    {vacantRooms.map(r => <option key={r.id} value={r.id}>P{r.id} - {formatMoney(r.rent)}</option>)}
                  </select>
                </label>
                <label>Ngày chuyển phòng
                  <input type="date" value={form.transferDate} onChange={e => setForm({...form, transferDate: e.target.value, oldRoomStayTo: formatDateInputValue(addDays(e.target.value, -1)) || form.oldRoomStayTo})} />
                </label>
                <label>Phòng cũ ở từ ngày
                  <input type="date" value={form.oldRoomStayFrom} onChange={e => setForm({...form, oldRoomStayFrom: e.target.value})} />
                </label>
                <label>Phòng cũ tính đến ngày
                  <input type="date" value={form.oldRoomStayTo} onChange={e => setForm({...form, oldRoomStayTo: e.target.value})} />
                </label>
                <label>Số hợp đồng mới
                  <input value={form.contractNo} onChange={e => setForm({...form, contractNo: e.target.value})} />
                </label>
                <label>Ngày hết hạn mới
                  <input type="date" value={form.endDate} onChange={e => setForm({...form, endDate: e.target.value})} />
                </label>
                <label>Giá thuê mới
                  <input type="number" value={form.rent} onChange={e => setForm({...form, rent: e.target.value})} />
                </label>
                <label>Tiền cọc áp dụng
                  <input type="number" value={form.deposit} onChange={e => setForm({...form, deposit: e.target.value})} />
                </label>
                <label>Cách tính dịch vụ phòng cũ
                  <select value={form.oldRoomServiceMode} onChange={e => setForm({...form, oldRoomServiceMode: e.target.value})}>
                    <option value="included_in_transfer_receipt">Thu chung một lần trong phiếu đổi phòng</option>
                    <option value="daily">Tính theo ngày ở thực tế</option>
                    <option value="monthly">Thu trọn tháng</option>
                  </select>
                </label>
              </div>
              <div className="warning-box">
                Phòng cũ P{contract.roomId}: {formatDisplayDate(form.oldRoomStayFrom)} - {formatDisplayDate(form.oldRoomStayTo)} ({oldRoomDaysPreview} ngày). Dự kiến tiền phòng cũ {formatMoney(oldRoomRentPreview)}, dịch vụ phòng cũ {formatMoney(oldRoomServicePreview)}. Nếu chọn thu chung, phiếu gộp sẽ hiển thị dịch vụ tháng một lần theo số người.
              </div>
              <label>Ghi chú
                <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} />
              </label>

              <div className="warning-box">
                Sau khi xác nhận, phòng {contract.roomId} sẽ chuyển về trống, toàn bộ người đang ở sẽ sang P{selectedRoom?.id}, và phiếu tháng mới sẽ dùng chỉ số đầu kỳ của phòng mới.
              </div>
            </>
          )}

          <div className="op-action-footer">
            <button className="secondary-btn" onClick={onClose}>Hủy</button>
            <button className="primary-btn" disabled={vacantRooms.length === 0} onClick={() => {
              if (!form.newRoomId) return alert('Vui lòng chọn phòng mới.');
              if (!form.transferDate || !form.endDate) return alert('Vui lòng nhập ngày chuyển và ngày hết hạn mới.');
              if (!form.oldRoomStayFrom || !form.oldRoomStayTo) return alert('Vui lòng nhập thời gian ở phòng cũ.');
              if (parseDateFlexible(form.oldRoomStayTo) < parseDateFlexible(form.oldRoomStayFrom)) return alert('Ngày kết thúc ở phòng cũ phải sau ngày bắt đầu.');
              if (form.endDate < form.transferDate) return alert('Ngày hết hạn mới phải sau ngày chuyển phòng.');
              if (window.confirm(`Xác nhận chuyển ${tenant.name} từ P${contract.roomId} sang P${form.newRoomId}?`)) {
                onSave(form);
              }
            }}>Xác nhận đổi phòng</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function SettlementModal({ room, contract, data, bankInfo, onClose, onSave }) {
  const latestMonthlyReceipt = getLatestMonthlyReceiptForRoom(data, room.id, contract.id);
  const meterElectricOld = latestMonthlyReceipt
    ? getElectricNew(latestMonthlyReceipt)
    : Number(room.electricNew ?? room.electricEnd ?? room.electricOld ?? room.electricStart ?? room.initialElectric ?? 0);
  const meterWaterOld = latestMonthlyReceipt
    ? getWaterNew(latestMonthlyReceipt)
    : Number(room.waterNew ?? room.waterEnd ?? room.waterOld ?? room.waterStart ?? room.initialWater ?? 0);
  const meterSourceText = latestMonthlyReceipt
    ? `Lấy từ phiếu tháng ${latestMonthlyReceipt.month} của P${latestMonthlyReceipt.roomId}: điện ${formatLocaleNumber(meterElectricOld)}, nước ${formatLocaleNumber(meterWaterOld)}.`
    : 'Chưa có phiếu tháng trước đó, hệ thống dùng chỉ số đang lưu trong phòng.';
  const [form, setForm] = useState({
    actualEndDate: new Date().toISOString().split('T')[0],
    settlementMode: 'offset_deposit',
    electricNew: meterElectricOld,
    waterNew: meterWaterOld,
    unpaidRent: 0,
    cleaningFee: 100000,
    damageFee: 0,
    otherFee: 0,
    refundBankName: '',
    refundBankAccount: '',
    refundBankOwner: '',
    refundQrImageUrl: '',
    note: ''
  });

  const primaryMembership = (data.memberships || []).find(m => m.contractId === contract.id && m.role === 'primary');
  const tenant = primaryMembership ? (data.tenants || []).find(t => t.id === primaryMembership.tenantId) : { name: 'N/A' };
  const settlementMembers = (data.memberships || [])
    .filter(m => m.contractId === contract.id)
    .map(m => ({ ...m, tenant: (data.tenants || []).find(t => t.id === m.tenantId) }))
    .filter(m => m.tenant);
  const fingerprintRemovalList = settlementMembers.filter(m => m.tenant.fingerprintCode);

  // Logic tính toán realtime
  const electricOld = Number(meterElectricOld || 0);
  const waterOld = Number(meterWaterOld || 0);
  
  const electricUsed = Math.max(0, Number(form.electricNew) - electricOld);
  const waterUsed = Math.max(0, Number(form.waterNew) - waterOld);
  const electricAmount = electricUsed * Number(room.electricPrice || 3800);
  const waterAmount = waterUsed * Number(room.waterPrice || 32000);
  const endDate = new Date(form.actualEndDate);
  const roomChargeDays = Number.isNaN(endDate.getTime()) ? 0 : endDate.getDate();
  const billingMonthDays = Number.isNaN(endDate.getTime()) ? 30 : new Date(endDate.getFullYear(), endDate.getMonth() + 1, 0).getDate();
  const monthlyRent = Number(contract.rent || room.rent || 0);
  const occupantCount = getContractOccupantCount(data, contract);
  const monthlyFixedServices = fixedServiceTotal(room, occupantCount);
  const proratedRent = Math.round((monthlyRent / billingMonthDays) * roomChargeDays);
  const proratedFixedServices = Math.round((monthlyFixedServices / billingMonthDays) * roomChargeDays);
  const dailyRent = monthlyRent / billingMonthDays;
  const dailyFixedServices = monthlyFixedServices / billingMonthDays;

  const totalIncurred = proratedRent + proratedFixedServices + electricAmount + waterAmount + Number(form.unpaidRent) +
                        Number(form.cleaningFee) + Number(form.damageFee) + Number(form.otherFee);
  
  const deposit = Number(contract.deposit || 0);
  const isOffsetDeposit = form.settlementMode === 'offset_deposit';
  const isPaySeparately = form.settlementMode === 'pay_separately';
  const isForfeitDeposit = form.settlementMode === 'forfeit_deposit';
  const depositUsed = isOffsetDeposit ? Math.min(deposit, totalIncurred) : 0;
  const depositForfeited = isForfeitDeposit ? deposit : 0;
  const mustCollect = (isPaySeparately || isForfeitDeposit) ? totalIncurred : Math.max(0, totalIncurred - deposit);
  const mustRefund = isPaySeparately ? deposit : isOffsetDeposit ? Math.max(0, deposit - totalIncurred) : 0;
  const isRefund = mustRefund > 0;
  const isDebt = mustCollect > 0;
  const settlementTransferContent = `P${room.id} TRA PHONG ${formatDisplayDate(form.actualEndDate, '').replace(/\//g, '')}`;
  const settlementQrUrl = buildPaymentQrUrl(bankInfo, mustCollect, settlementTransferContent);
  const copySettlementTransfer = () => {
    const text = `${bankInfo?.bankName || ''}\nSTK: ${bankInfo?.accountNo || ''}\nChủ TK: ${bankInfo?.accountName || ''}\nSố tiền: ${formatMoney(mustCollect)}\nNội dung: ${settlementTransferContent}`;
    navigator.clipboard.writeText(text);
    alert('Đã copy thông tin chuyển khoản tất toán.');
  };

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
                  <label>Số ngày sử dụng <input value={`${roomChargeDays} ngày`} readOnly style={{ background: '#f1f5f9', fontWeight: 'bold' }} /></label>
                  <label>Tiền phòng phát sinh <input value={formatMoney(proratedRent)} readOnly style={{ background: '#f1f5f9', fontWeight: 'bold' }} /></label>
                  <label>Dịch vụ phát sinh <input value={formatMoney(proratedFixedServices)} readOnly style={{ background: '#f1f5f9', fontWeight: 'bold' }} /></label>
                  <p className="small muted" style={{ gridColumn: 'span 2', margin: 0 }}>
                    Tính từ ngày 01 đến ngày {formatDisplayDate(form.actualEndDate)}, gồm cả ngày trả phòng: {formatMoney(monthlyRent)} / {billingMonthDays} x {roomChargeDays} ngày = {formatMoney(proratedRent)}
                  </p>
                  <p className="small muted" style={{ gridColumn: 'span 2', margin: 0 }}>
                    Dịch vụ cố định: {formatMoney(monthlyFixedServices)}/tháng ({occupantCount} người) / {billingMonthDays} x {roomChargeDays} ngày = {formatMoney(proratedFixedServices)}
                  </p>
                  <label style={{ gridColumn: 'span 2' }}>Tiền cọc đang giữ <input value={formatMoney(deposit)} readOnly style={{ background: '#f1f5f9', fontWeight: 'bold' }} /></label>
                </div>
              </section>

              <section>
                <h3 className="form-section-title">💳 Cách xử lý công nợ & tiền cọc</h3>
                <div className="stack" style={{ gap: '10px' }}>
                  <label className="option-row">
                    <input type="radio" name="settlementMode" checked={form.settlementMode === 'offset_deposit'} onChange={() => setForm({...form, settlementMode: 'offset_deposit'})} />
                    <span><b>Đối trừ vào cọc</b><small>Điện nước, phí phát sinh được trừ vào tiền cọc; chỉ hoàn phần cọc còn lại.</small></span>
                  </label>
                  <label className="option-row">
                    <input type="radio" name="settlementMode" checked={form.settlementMode === 'pay_separately'} onChange={() => setForm({...form, settlementMode: 'pay_separately'})} />
                    <span><b>Khách thanh toán điện nước/phí riêng, hoàn nguyên cọc</b><small>Tạo khoản khách cần trả cho phát sinh và hoàn lại toàn bộ tiền cọc.</small></span>
                  </label>
                  <label className="option-row danger">
                    <input type="radio" name="settlementMode" checked={form.settlementMode === 'forfeit_deposit'} onChange={() => setForm({...form, settlementMode: 'forfeit_deposit'})} />
                    <span><b>Trả sớm hợp đồng, không hoàn cọc</b><small>Cọc được bù công nợ trước; phần cọc còn lại ghi nhận giữ lại do khách trả phòng trước hạn.</small></span>
                  </label>
                </div>
              </section>

              {mustRefund > 0 && !isForfeitDeposit && (
                <section>
                  <h3 className="form-section-title">🏦 Thông tin hoàn cọc cho khách</h3>
                  <div className="form-grid-v2">
                    <label>Ngân hàng người thuê
                      <input value={form.refundBankName} onChange={e => setForm({...form, refundBankName: e.target.value})} placeholder="VD: Vietcombank, BIDV..." />
                    </label>
                    <label>Số tài khoản
                      <input value={form.refundBankAccount} onChange={e => setForm({...form, refundBankAccount: e.target.value})} placeholder="Nhập STK nhận hoàn cọc" />
                    </label>
                    <label style={{ gridColumn: 'span 2' }}>Chủ tài khoản
                      <input value={form.refundBankOwner || tenant.name || ''} onChange={e => setForm({...form, refundBankOwner: e.target.value})} placeholder={tenant.name || 'Tên chủ tài khoản'} />
                    </label>
                    <label style={{ gridColumn: 'span 2' }}>Link ảnh QR của người thuê
                      <input value={form.refundQrImageUrl} onChange={e => setForm({...form, refundQrImageUrl: e.target.value})} placeholder="Dán link ảnh QR để lưu vào phiếu chi hoàn cọc" />
                    </label>
                    <p className="small muted" style={{ gridColumn: 'span 2', margin: 0 }}>
                      Khi hoàn tất trả phòng, hệ thống tự tạo phiếu chi hoàn cọc {formatMoney(mustRefund)} để xác nhận chủ nhà đã chuyển lại tiền cho khách.
                    </p>
                  </div>
                </section>
              )}

              {/* Nhóm 2: Điện nước */}
              <section>
                <h3 className="form-section-title">⚡ Chỉ số điện nước</h3>
                <div className="form-grid-v2">
                  <p className="small muted" style={{ gridColumn: 'span 2', margin: 0 }}>{meterSourceText}</p>
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
                  <label>Tiền phòng/phí còn nợ khác <input type="number" value={form.unpaidRent} onChange={e => setForm({...form, unpaidRent: e.target.value})} /></label>
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
                
                <div className="summary-row"><span>Tiền phòng phát sinh ({roomChargeDays} ngày)</span><b>{formatMoney(proratedRent)}</b></div>
                <div className="summary-row"><span>Dịch vụ phát sinh ({roomChargeDays} ngày)</span><b>{formatMoney(proratedFixedServices)}</b></div>
                <div className="summary-row"><span>Tiền điện</span><b>{formatMoney(electricAmount)}</b></div>
                <div className="summary-row"><span>Tiền nước</span><b>{formatMoney(waterAmount)}</b></div>
                <div className="summary-row"><span>Phí khác</span><b>{formatMoney(Number(form.unpaidRent || 0) + Number(form.cleaningFee || 0) + Number(form.damageFee || 0) + Number(form.otherFee || 0))}</b></div>
                <div className="summary-row"><span>Tổng phát sinh</span><b>{formatMoney(totalIncurred)}</b></div>
                {isOffsetDeposit && <div className="summary-row"><span>Tiền cọc đối trừ</span><b style={{ color: 'var(--text-muted)' }}>- {formatMoney(depositUsed)}</b></div>}
                {isForfeitDeposit && <div className="summary-row"><span>Cọc giữ lại do trả sớm</span><b style={{ color: 'var(--danger)' }}>{formatMoney(depositForfeited)}</b></div>}
                {isPaySeparately && <div className="summary-row"><span>Hoàn nguyên cọc</span><b style={{ color: 'var(--success)' }}>{formatMoney(mustRefund)}</b></div>}
                {mustCollect > 0 && <div className="summary-row"><span>Khách thanh toán phát sinh</span><b style={{ color: 'var(--danger)' }}>{formatMoney(mustCollect)}</b></div>}
                
                <div className="summary-total">
                  <span className="summary-label-main">{isForfeitDeposit ? 'Khách cần thanh toán' : isOffsetDeposit ? (isRefund ? 'Số tiền hoàn khách' : 'Khách cần trả thêm') : 'Hoàn cọc & thu phát sinh'}</span>
                  <p className="summary-amount" style={{ color: isRefund ? 'var(--success)' : isDebt ? 'var(--danger)' : 'var(--text-main)' }}>
                    {isForfeitDeposit ? formatMoney(mustCollect) : isOffsetDeposit ? formatMoney(isRefund ? mustRefund : mustCollect) : `${formatMoney(mustRefund)} / ${formatMoney(mustCollect)}`}
                  </p>
                  <p style={{ fontSize: '13px', fontWeight: '500' }}>
                    {isForfeitDeposit
                      ? '⚠️ Khách trả sớm hợp đồng: không hoàn cọc, phát sinh vẫn thu riêng'
                      : isOffsetDeposit
                      ? isRefund ? '✨ Cần hoàn phần cọc còn lại cho khách' : isDebt ? '⚠️ Khách thuê cần đóng thêm tiền sau khi trừ cọc' : '✅ Công nợ đã được tất toán đủ'
                      : '✨ Hoàn toàn bộ cọc, đồng thời thu riêng điện nước/phí phát sinh'}
                  </p>
                </div>

                {mustCollect > 0 && (
                  <div className="settlement-qr-card">
                    <h4>QR khách chuyển khoản</h4>
                    {settlementQrUrl ? <img src={settlementQrUrl} alt="QR tất toán" /> : <p className="small muted">Chưa cấu hình mã VietQR.</p>}
                    <div className="payment-details-v4">
                      <div className="row"><span className="label">Số tiền</span><span className="val">{formatMoney(mustCollect)}</span></div>
                      <div className="row"><span className="label">Nội dung CK</span><span className="val" style={{ color: '#1e40af' }}>{settlementTransferContent}</span></div>
                    </div>
                    <button type="button" className="secondary-btn sm wide" onClick={copySettlementTransfer}>Copy nội dung CK</button>
                  </div>
                )}

                {mustRefund > 0 && !isForfeitDeposit && (
                  <div className="settlement-qr-card">
                    <h4>Phiếu chi hoàn cọc sẽ tạo</h4>
                    {form.refundQrImageUrl ? <img src={form.refundQrImageUrl} alt="QR người thuê nhận hoàn cọc" /> : <p className="small muted">Có thể dán link QR người thuê ở form bên trái để lưu kèm phiếu chi.</p>}
                    <div className="payment-details-v4">
                      <div className="row"><span className="label">Người nhận</span><span className="val">{form.refundBankOwner || tenant.name || 'Khách thuê'}</span></div>
                      <div className="row"><span className="label">Số tiền hoàn</span><span className="val">{formatMoney(mustRefund)}</span></div>
                      {form.refundBankAccount && <div className="row"><span className="label">STK</span><span className="val">{form.refundBankAccount}</span></div>}
                    </div>
                  </div>
                )}

                <div className="warning-box">
                  <b>Lưu ý sau khi hoàn tất:</b>
                  <ul>
                    <li>Hợp đồng sẽ chuyển sang "Đã kết thúc"</li>
                    <li>Phòng {room.id} sẽ trở về trạng thái trống</li>
                    <li>Lịch sử tất toán sẽ được lưu lại</li>
                    {fingerprintRemovalList.length > 0 && <li>Admin cần xóa vân tay: {fingerprintRemovalList.map(m => `${m.tenant.name} (${m.tenant.fingerprintCode})`).join(', ')}</li>}
                    {isForfeitDeposit && <li>Tiền cọc còn lại được ghi nhận giữ lại do khách trả sớm hợp đồng</li>}
                  </ul>
                </div>

                <div className="stack" style={{ gap: '12px', marginTop: 'auto' }}>
                  <button className="primary-btn wide" onClick={() => {
                    if (window.confirm('Xác nhận hoàn tất mọi thủ tục trả phòng và chốt công nợ?')) {
                      onSave({
                        ...form,
                        id: uid('moveout'),
                        unpaidRent: Number(form.unpaidRent || 0),
                        monthlyRent,
                        monthlyFixedServices,
                        roomChargeDays,
                        billingMonthDays,
                        dailyRent,
                        dailyFixedServices,
                        proratedRent,
                        proratedFixedServices,
                        cleaningFee: Number(form.cleaningFee || 0),
                        damageFee: Number(form.damageFee || 0),
                        otherFee: Number(form.otherFee || 0),
                        electricOld,
                        electricNew: Number(form.electricNew || 0),
                        electricUsed,
                        electricAmount,
                        waterOld,
                        waterNew: Number(form.waterNew || 0),
                        waterUsed,
                        waterAmount,
                        totalIncurred,
                        depositUsed,
                        depositForfeited,
                        mustCollect,
                        mustRefund,
                        settlementMode: form.settlementMode,
                        tenantId: tenant.id,
                        tenantName: tenant.name,
                        tenantPhone: tenant.phone,
                        refundBankName: form.refundBankName,
                        refundBankAccount: form.refundBankAccount,
                        refundBankOwner: form.refundBankOwner || tenant.name,
                        refundQrImageUrl: form.refundQrImageUrl,
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
  const hasValidCurrentEndDate = isValidBusinessDate(contract.endDate);
  const currentEndDateText = formatBusinessDate(contract.endDate);
  const currentStartDateText = formatBusinessDate(contract.startDate);
  const currentSignedDateText = formatContractDate(contract.signedDate);
  const defaultStartDate = hasValidCurrentEndDate ? addDaysToDate(contract.endDate, 1) : new Date().toISOString().slice(0, 10);
  
  const [form, setForm] = useState({
    signedDate: new Date().toISOString().slice(0, 10),
    newStartDate: defaultStartDate,
    newEndDate: hasValidCurrentEndDate ? addMonthsToDate(contract.endDate, 12) : '',
    keepPricing: true,
    newRent: contract.rent,
    newDeposit: contract.deposit,
    note: ''
  });

  const isValid = form.newEndDate && (!hasValidCurrentEndDate || form.newEndDate > contract.endDate);
  const rentChanged = !form.keepPricing && Number(form.newRent) !== Number(contract.rent);
  const depositChanged = !form.keepPricing && Number(form.newDeposit) !== Number(contract.deposit);
  const rentDelta = form.keepPricing ? 0 : Number(form.newRent || 0) - Number(contract.rent || 0);
  const depositDelta = form.keepPricing ? 0 : Number(form.newDeposit || 0) - Number(contract.deposit || 0);
  const draftAppendixPrintId = `draft-renewal-appendix-${contract.id}`;

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
                <li>Ngày hết hạn cũ: <b>${currentEndDateText}</b></li>
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
            <p className="muted">{tenant.name} • Hết hạn hiện tại: {currentEndDateText}</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>

        <div className="detail-body-v2 stack" style={{ gap: '24px' }}>
          {/* Thông tin hiện tại */}
          <section className="op-card" style={{ background: '#f8fafc', borderStyle: 'dashed' }}>
            <h3 className="op-card-title">📌 Thông tin hiện tại</h3>
            <div className="op-grid">
              <div className="op-item"><span className="op-label">Ngày bắt đầu</span><span className="op-value">{currentStartDateText}</span></div>
              <div className="op-item"><span className="op-label">Hết hạn hiện tại</span><span className="op-value">{currentEndDateText}</span></div>
              <div className="op-item"><span className="op-label">Ngày ký hợp đồng gốc</span><span className="op-value">{currentSignedDateText}</span></div>
              <div className="op-item"><span className="op-label">Giá thuê</span><span className="op-value">{formatMoney(contract.rent)}</span></div>
              <div className="op-item"><span className="op-label">Tiền cọc</span><span className="op-value">{formatMoney(contract.deposit)}</span></div>
            </div>
            {(isInvalidContractDate(contract.signedDate) || !isValidBusinessDate(contract.startDate) || !hasValidCurrentEndDate) && (
              <div className="warning-box" style={{ marginTop: '12px' }}>Hợp đồng thiếu ngày ký / ngày bắt đầu / ngày hết hạn hợp lệ. Vui lòng cập nhật dữ liệu hợp đồng trước khi gia hạn.</div>
            )}
          </section>

          {/* Form gia hạn */}
          <section className="stack" style={{ gap: '16px' }}>
            <h3 className="form-section-title">✍️ Chi tiết gia hạn</h3>
            <div className="form-grid-v2">
              <label>Ngày ký phụ lục <input type="date" value={form.signedDate} onChange={e => setForm({...form, signedDate: e.target.value})} /><span className="small muted">Chỉ dùng cho phụ lục gia hạn, không ghi đè ngày ký hợp đồng gốc.</span></label>
              <label>Bắt đầu gia hạn <input type="date" value={form.newStartDate} onChange={e => setForm({...form, newStartDate: e.target.value})} /></label>
              <label style={{ gridColumn: 'span 2' }}>
                Ngày hết hạn mới
                <input type="date" value={form.newEndDate} onChange={e => setForm({...form, newEndDate: e.target.value})} />
                {!hasValidCurrentEndDate && <p className="danger small" style={{ marginTop: '4px' }}>⚠ Chưa có ngày hết hạn cũ hợp lệ. Vui lòng chọn ngày mới.</p>}
                {hasValidCurrentEndDate && form.newEndDate <= contract.endDate && <p className="danger small" style={{ marginTop: '4px' }}>⚠ Ngày hết hạn mới phải sau ngày {currentEndDateText}</p>}
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
              {(rentDelta !== 0 || depositDelta !== 0) && (
                <div className="warning-box" style={{ marginTop: '12px' }}>
                  <b>Chứng từ sẽ tự động tạo sau khi gia hạn:</b>
                  <div className="stack" style={{ gap: '4px', marginTop: '8px' }}>
                    {(Math.max(0, rentDelta) + Math.max(0, depositDelta)) > 0 && <span>• Phiếu thu: {formatMoney(Math.max(0, rentDelta) + Math.max(0, depositDelta))}</span>}
                    {(Math.max(0, -rentDelta) + Math.max(0, -depositDelta)) > 0 && <span>• Phiếu chi hoàn khách: {formatMoney(Math.max(0, -rentDelta) + Math.max(0, -depositDelta))}</span>}
                  </div>
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
            <button className="secondary-btn" onClick={() => setShowAppendix(!showAppendix)} disabled={!isValid}>
              {!isValid ? 'Chưa có phụ lục' : showAppendix ? 'Ẩn phụ lục' : 'Xem phụ lục'}
            </button>
            <button className="primary-btn" onClick={() => onSave(form)} disabled={!isValid} style={{ flex: 2 }}>
              🚀 Gia hạn hợp đồng
            </button>
          </div>

          {showAppendix && (
            <div className="appendix-preview-scroll" style={{ marginTop: '20px', borderTop: '1px solid rgba(255,255,255,0.1)', paddingTop: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
                <h3 className="form-section-title" style={{ margin: 0 }}>📄 Xem trước phụ lục (A4)</h3>
                <button className="secondary-btn sm" onClick={() => printDocumentElement(draftAppendixPrintId, `Phụ lục gia hạn P${contract.roomId}`)}>🖨️ In phụ lục này</button>
              </div>
              <div className="appendix-container-scroll" style={{ maxHeight: '500px', overflowY: 'auto', background: 'rgba(0,0,0,0.2)', padding: '20px', borderRadius: '12px' }}>
                <div className="appendix-paper-a4" style={{ transform: 'scale(0.8)', transformOrigin: 'top center', margin: '0 auto', marginBottom: '-150px' }}>
                   <AppendixContent contract={contract} tenant={tenant} form={form} room={room} printId={draftAppendixPrintId} />
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function printAppendixFromElement(elementId, title = 'Phụ lục gia hạn') {
  printDocumentElement(elementId, title);
}

function RenewalAppendixPreviewModal({ contract, form, data, onClose }) {
  const room = data.rooms.find(r => r.id === contract.roomId) || {};
  const tenant = getPrimaryTenantByContract(data, contract.id) || { name: 'N/A', phone: 'N/A' };
  const printId = 'saved-renewal-appendix';

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '900px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2 style={{ fontSize: '24px' }}>Phụ lục gia hạn • Phòng {contract.roomId}</h2>
            <p className="muted">Ngày ký phụ lục: {formatDisplayDate(form.signedDate)} • Hết hạn mới: {formatDisplayDate(form.newEndDate)}</p>
          </div>
          <div className="btn-group">
            <button className="primary-btn" onClick={() => printAppendixFromElement(printId, `Phụ lục gia hạn P${contract.roomId}`)}>🖨️ In phụ lục</button>
            <button className="secondary-btn" onClick={onClose}>Đóng</button>
          </div>
        </div>
        <div className="detail-body-v2">
          <div className="appendix-container-scroll" style={{ maxHeight: '70vh', overflowY: 'auto', background: 'rgba(0,0,0,0.06)', padding: '20px', borderRadius: '12px' }}>
            <div className="appendix-paper-a4" style={{ transform: 'scale(0.86)', transformOrigin: 'top center', margin: '0 auto', marginBottom: '-100px' }}>
              <AppendixContent contract={contract} tenant={tenant} form={form} room={room} printId={printId} />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function AppendixContent({ contract, tenant, form, room, printId }) {
  return (
    <div id={printId} className="appendix-content-v1" style={{
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
          <li>Ngày hết hạn cũ: <b>{formatBusinessDate(contract.endDate)}</b></li>
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

function TenantDetailModal({ tenant, data, onClose }) {
  const memberships = (data.memberships || []).filter(m => m.tenantId === tenant.id).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
  const activeMembership = memberships.find(m => m.status === 'active') || memberships[0];
  const contract = activeMembership ? (data.contracts || []).find(c => c.id === activeMembership.contractId) : null;
  const room = activeMembership ? (data.rooms || []).find(r => r.id === activeMembership.roomId) : null;
  const receipts = contract ? (data.receipts || []).filter(r => r.contractId === contract.id).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0)) : [];
  const currentReceipt = receipts.find(r => r.month === getCurrentMonthLabel() && r.type === 'monthly');
  const debt = currentReceipt ? getReceiptPaymentState(currentReceipt).debt : 0;
  const timeline = [
    { date: tenant.createdAt, title: 'Tạo khách', detail: tenant.name },
    ...memberships.map(m => ({ date: m.joinedDate || m.createdAt, title: 'Vào ở', detail: `P${m.roomId} - ${m.role === 'primary' ? 'Đại diện' : 'Ở cùng'}` })),
    ...(contract ? [{ date: contract.signedDate || contract.createdAt, title: 'Ký hợp đồng', detail: `${formatBusinessDate(contract.startDate)} → ${formatBusinessDate(contract.endDate)}` }] : []),
    ...receipts.filter(r => Number(r.paidAmount || 0) > 0).map(r => ({ date: r.paidDate || r.createdAt, title: `Thanh toán ${r.month}`, detail: formatMoney(r.paidAmount || 0) })),
    ...(memberships.filter(m => m.leftDate).map(m => ({ date: m.leftDate, title: 'Rời phòng', detail: `P${m.roomId}` })))
  ].filter(i => i.date).sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0)).slice(0, 8);

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '820px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2 style={{ fontSize: '26px' }}>{tenant.name}</h2>
            <p className="muted">{tenant.phone || 'Chưa có SĐT'} • CCCD {tenant.cccd || 'Chưa cập nhật'}</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ gap: '16px' }}>
          <div className="tenant-detail-grid">
            <div className="op-card">
              <h3 className="op-card-title">Cư trú</h3>
              <p className="op-value">P{activeMembership?.roomId || '—'}</p>
              <p className="op-label">{activeMembership?.status === 'active' ? 'Đang ở' : 'Đã rời'} • {activeMembership?.role === 'primary' ? 'Đại diện' : 'Ở cùng'}</p>
            </div>
            <div className="op-card">
              <h3 className="op-card-title">Hợp đồng</h3>
              <p className="op-value">{contract ? `${formatBusinessDate(contract.startDate)} → ${formatBusinessDate(contract.endDate)}` : isOwnerOccupiedRoom(room) ? 'Chủ nhà ở' : 'Chưa có'}</p>
              <p className="op-label">{contract ? calculateRentalDuration(contract.startDate, contract.endDate) : 'Không áp dụng'}</p>
            </div>
            <div className="op-card">
              <h3 className="op-card-title">Thanh toán</h3>
              <p className="op-value" style={{ color: debt > 0 ? 'var(--danger)' : 'var(--success)' }}>{debt > 0 ? formatMoney(debt) : 'Không nợ'}</p>
              <p className="op-label">{currentReceipt ? `Phiếu ${currentReceipt.month}` : 'Chưa có phiếu tháng này'}</p>
            </div>
            <div className="op-card">
              <h3 className="op-card-title">Ra vào</h3>
              <p className="op-value">{tenant.licensePlate || 'Chưa có biển số'}</p>
              <p className="op-label">Vân tay: {tenant.fingerprintCode || 'Chưa đăng ký'} • {tenant.fingerprintStatus || 'Chưa đăng ký'}</p>
            </div>
          </div>
          <div className="op-card">
            <h3 className="op-card-title">Timeline</h3>
            {timeline.map((item, idx) => (
              <div key={`${item.title}-${idx}`} className="tenant-timeline-row">
                <span>{formatDisplayDate(String(item.date || '').slice(0, 10))}</span>
                <b>{item.title}</b>
                <em>{item.detail}</em>
              </div>
            ))}
            {!timeline.length && <p className="muted small">Chưa có lịch sử hoạt động.</p>}
          </div>
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
            <label>Biển số xe <input value={form.licensePlate || form.vehicle || ''} onChange={e => setForm({...form, licensePlate: e.target.value})} /></label>
            <label>Mã vân tay <input value={form.fingerprintCode || ''} onChange={e => setForm({...form, fingerprintCode: e.target.value})} placeholder="VD: F502-01" /></label>
            <label>Trạng thái vân tay
              <select value={form.fingerprintStatus || 'Chưa đăng ký'} onChange={e => setForm({...form, fingerprintStatus: e.target.value})}>
                <option>Chưa đăng ký</option>
                <option>Đã đăng ký</option>
                <option>Cần xóa</option>
                <option>Đã xóa</option>
              </select>
            </label>
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

function RoommateModal({ room, contract, onClose, onSave }) {
  const [isSaving, setIsSaving] = useState(false);
  const [form, setForm] = useState({
    name: '',
    phone: '',
    cccd: '',
    cccdDate: '',
    cccdPlace: '',
    birthday: '',
    address: '',
    licensePlate: '',
    fingerprintCode: '',
    fingerprintStatus: 'Chưa đăng ký',
    joinedDate: new Date().toISOString().slice(0, 10),
    note: '',
    idPaste: ''
  });

  const applyPastedIdInfo = (text) => {
    const parsed = parseVietnameseIdCard(text);

    setForm(prev => ({
      ...prev,
      idPaste: text,
      cccd: parsed.cccd || prev.cccd,
      name: parsed.name || prev.name,
      birthday: parsed.birthday || prev.birthday,
      address: parsed.address || prev.address,
      cccdDate: parsed.issueDate || prev.cccdDate,
      cccdPlace: parsed.issuePlace || prev.cccdPlace,
      note: parsed.note || prev.note
    }));
  };

  const handleSave = () => {
    if (isSaving) return;
    if (!form.name.trim()) {
      alert('Vui lòng nhập họ tên người ở cùng.');
      return;
    }
    if (!form.joinedDate) {
      alert('Vui lòng nhập ngày vào ở.');
      return;
    }
    setIsSaving(true);
    const tenantId = uid('tenant');
    const membershipId = uid('membership');
    const tenant = {
      id: tenantId,
      name: form.name.trim(),
      phone: form.phone,
      cccd: form.cccd,
      cccdDate: form.cccdDate,
      cccdPlace: form.cccdPlace,
      birthday: form.birthday,
      address: form.address,
      licensePlate: form.licensePlate,
      fingerprintCode: form.fingerprintCode,
      fingerprintStatus: form.fingerprintStatus,
      role: 'member',
      status: 'active',
      note: form.note,
      createdAt: new Date().toISOString()
    };
    const membership = {
      id: membershipId,
      contractId: contract.id,
      tenantId,
      roomId: room.id,
      role: 'member',
      status: 'active',
      joinedDate: form.joinedDate,
      leftDate: '',
      createdAt: new Date().toISOString()
    };
    try {
      const saved = onSave({ tenant, membership });
      if (saved === false) {
        setIsSaving(false);
      }
    } catch (error) {
      console.error(error);
      setIsSaving(false);
      alert('Không lưu được người ở cùng. Vui lòng thử lại.');
    }
  };

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '620px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>Thêm người ở cùng • Phòng {room.id}</h2>
            <p className="muted">Gắn người ở cùng vào hợp đồng {contract.contractNo || contract.id}</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack">
          <div className="paste-id-card">
            <label>Dán nhanh thông tin CCCD
              <textarea
                value={form.idPaste}
                onChange={e => applyPastedIdInfo(e.target.value)}
                placeholder="Dán nguyên nội dung CCCD ở đây, hệ thống tự điền: họ tên, số CCCD, ngày sinh, thường trú, ngày cấp..."
              />
            </label>
          </div>
          <div className="rental-subsection-title">CCCD mặt trước</div>
          <div className="form-grid-v2">
            <label style={{ gridColumn: 'span 2' }}>Họ và tên <input value={form.name} onChange={e => setForm({...form, name: e.target.value})} placeholder="Nguyễn Văn A" /></label>
            <label>Số CCCD <input value={form.cccd} onChange={e => setForm({...form, cccd: e.target.value})} /></label>
            <label>Ngày sinh <input type="date" value={form.birthday} onChange={e => setForm({...form, birthday: e.target.value})} /></label>
            <label style={{ gridColumn: 'span 2' }}>Địa chỉ thường trú <input value={form.address} onChange={e => setForm({...form, address: e.target.value})} /></label>
          </div>
          <div className="rental-subsection-title">CCCD mặt sau</div>
          <div className="form-grid-v2">
            <label>Ngày cấp CCCD <input type="date" value={form.cccdDate} onChange={e => setForm({...form, cccdDate: e.target.value})} /></label>
            <label>Nơi cấp <input value={form.cccdPlace} onChange={e => setForm({...form, cccdPlace: e.target.value})} /></label>
            <label style={{ gridColumn: 'span 2' }}>Ghi chú CCCD / MRZ <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} placeholder="Giới tính, quốc tịch, quê quán, hạn CCCD, đặc điểm nhận dạng, MRZ..." /></label>
          </div>
          <div className="rental-subsection-title">Liên hệ & quản lý ra vào</div>
          <div className="form-grid-v2">
            <label>Số điện thoại <input value={form.phone} onChange={e => setForm({...form, phone: e.target.value})} placeholder="09xx..." /></label>
            <label>Ngày vào ở <input type="date" value={form.joinedDate} onChange={e => setForm({...form, joinedDate: e.target.value})} /></label>
            <label>Biển số xe <input value={form.licensePlate} onChange={e => setForm({...form, licensePlate: e.target.value})} /></label>
            <label>Mã vân tay <input value={form.fingerprintCode} onChange={e => setForm({...form, fingerprintCode: e.target.value})} placeholder="VD: F502-02" /></label>
            <label>Trạng thái vân tay
              <select value={form.fingerprintStatus} onChange={e => setForm({...form, fingerprintStatus: e.target.value})}>
                <option>Chưa đăng ký</option>
                <option>Đã đăng ký</option>
                <option>Cần xóa</option>
                <option>Đã xóa</option>
              </select>
            </label>
          </div>
          <div className="op-action-footer">
            <button className="secondary-btn" onClick={onClose}>Hủy</button>
            <button className="primary-btn" onClick={handleSave} disabled={isSaving}>{isSaving ? 'Đang lưu...' : '+ Thêm người ở cùng'}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ReceiptModal({ receipt, room, data, bankInfo, onClose, onPrev, onNext }) {
  if (!receipt || !room) return null;
  const enrichedReceipt = enrichReceiptWithTransferUtility(receipt, data);
  const paymentState = getReceiptPaymentState(enrichedReceipt);
  const displayReceipt = {
    ...enrichedReceipt,
    debt: paymentState.debt,
    status: paymentState.status
  };
  const isMonthly = receipt.type === 'monthly';
  const tenant = getTenantForReceipt(data, displayReceipt) || { name: '—', phone: '—' };
  const settlementReport = !isMonthly
    ? (data.moveOutReports || []).find(report => report.contractId === displayReceipt.contractId && report.roomId === displayReceipt.roomId)
    : null;

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
        .paid-stamp-v4 { top: 20mm !important; right: 8mm !important; padding: 6px 12px !important; border-width: 3px !important; border-radius: 10px !important; }
        .paid-stamp-v4 span { font-size: 14px !important; }
        .paid-stamp-v4 small { font-size: 9px !important; }
        
        .info-bar-v4 { gap: 4px !important; padding: 4px 10px !important; margin-bottom: 8px !important; border-radius: 8px !important; }
        .info-bar-v4 .value { font-size: 13px !important; }
        
        .receipt-body-v4 { gap: 10px !important; grid-template-columns: 1fr 260px !important; }
        .receipt-items-section h3 { margin-bottom: 8px !important; padding-bottom: 4px !important; }
        .charge-row-v4 { padding: 4px 0 !important; }
        .charge-row-v4 .name { font-size: 13px !important; }
        .charge-row-v4 .amount { font-size: 14px !important; }
        
        .payment-card-v4 { padding: 10px !important; gap: 8px !important; border-radius: 12px !important; }
        .qr-box-v4 { padding: 8px !important; gap: 4px !important; }
        .qr-box-v4 img { width: 170px !important; height: 170px !important; padding: 6px !important; image-rendering: crisp-edges !important; }
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
    navigator.clipboard.writeText(transferContent(displayReceipt));
    alert('Đã copy nội dung chuyển khoản!');
  };

  return (
    <div className="modal no-print-backdrop" onClick={onClose}>
      <div className="receipt-modal-v3 liquid-glass" onClick={e => e.stopPropagation()}>
        <div className="modal-header no-print" style={{ padding: '10px 24px', borderBottom: '1px solid rgba(255,255,255,0.1)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
            <div style={{ display: 'flex', gap: '4px' }}>
              <button className="secondary-btn" onClick={onPrev} disabled={!onPrev} style={{ borderRadius: '50%', width: '32px', height: '32px', padding: 0 }}>←</button>
              <button className="secondary-btn" onClick={onNext} disabled={!onNext} style={{ borderRadius: '50%', width: '32px', height: '32px', padding: 0 }}>→</button>
            </div>
            <div>
              <h2 style={{ fontSize: '18px', margin: 0 }}>Phiếu thu {getTransferRoomLabel(displayReceipt)}</h2>
              <p className="muted small">Tháng {displayReceipt.month}</p>
            </div>
          </div>
          <button className="secondary-btn" onClick={onClose} style={{ borderRadius: '50%', width: '36px', height: '36px', padding: 0 }}>✕</button>
        </div>

        <div className="modal-body-v3 scrollable">
          <PrintableReceipt 
            receipt={displayReceipt}
            room={room} 
            tenant={tenant} 
            bankInfo={bankInfo} 
            settlementReport={settlementReport}
          />
        </div>

        <div className="modal-footer-v3 no-print" style={{ padding: '10px 24px', borderTop: '1px solid rgba(255,255,255,0.1)', display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
          <button className="secondary-btn" onClick={handleCopyTransfer}>📋 Copy nội dung CK</button>
          <button className="primary-btn" onClick={handlePrintReceipt}>🖨️ In phiếu thu</button>
        </div>
      </div>
    </div>
  );
}

function SettlementReceiptBreakdown({ report, receipt, room, table = false }) {
  const roomChargeDays = Number(report?.roomChargeDays || 0);
  const reportEndDate = report?.actualEndDate ? new Date(report.actualEndDate) : null;
  const inferredMonthDays = reportEndDate && !Number.isNaN(reportEndDate.getTime())
    ? new Date(reportEndDate.getFullYear(), reportEndDate.getMonth() + 1, 0).getDate()
    : 30;
  const billingMonthDays = Number(report?.billingMonthDays || inferredMonthDays);
  const monthlyRent = Number(report?.monthlyRent ?? room?.rent ?? 0);
  const monthlyServices = Number(report?.monthlyFixedServices ?? 0);
  const proratedRent = Number(report?.proratedRent ?? receipt?.rent ?? 0);
  const proratedServices = Number(report?.proratedFixedServices ?? receipt?.fixedServices ?? 0);
  const rentDetails = roomChargeDays
    ? `${formatMoney(monthlyRent)} / ${billingMonthDays} ngày × ${roomChargeDays} ngày`
    : (monthlyRent ? `Tiền phòng tháng: ${formatMoney(monthlyRent)}` : '');
  const serviceDetails = roomChargeDays && monthlyServices
    ? `${formatMoney(monthlyServices)} / ${billingMonthDays} ngày × ${roomChargeDays} ngày`
    : '';
  const rowData = report ? [
    ['Tiền phòng phát sinh', proratedRent, rentDetails],
    ['Dịch vụ phát sinh', proratedServices, serviceDetails],
    ['Tiền điện', report.electricAmount, `${report.electricOld ?? 0} → ${report.electricNew ?? 0} (${report.electricUsed ?? 0} kWh)`],
    ['Tiền nước', report.waterAmount, `${report.waterOld ?? 0} → ${report.waterNew ?? 0} (${report.waterUsed ?? 0} m³)`],
    ['Tiền phòng / công nợ khác', report.unpaidRent, ''],
    ['Phí vệ sinh', report.cleaningFee, ''],
    ['Phí hư hỏng', report.damageFee, ''],
    ['Phí khác', report.otherFee, ''],
  ].filter(([, amount], index) => index < 4 || Number(amount || 0) !== 0) : [];

  if (!report) {
    return table
      ? <tr><td style={{ border: '1px solid black', padding: '8px' }}>Phí chốt tất toán trả phòng</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(receipt.total)}</td></tr>
      : <div className="charge-row-v4"><div className="row-main"><span className="name">📝 Phí tất toán trả phòng</span><span className="amount">{formatMoney(receipt.total)}</span></div></div>;
  }

  if (table) return <>
    {rowData.map(([label, amount, details]) => <tr key={label}><td style={{ border: '1px solid black', padding: '8px' }}>{label}</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>{details || '-'}</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(amount || 0)}</td></tr>)}
    <tr style={{ fontWeight: 'bold', background: '#f8fafc' }}><td colSpan="2" style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>TỔNG PHÁT SINH</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(report.totalIncurred || 0)}</td></tr>
    {Number(report.depositUsed || 0) > 0 && <tr><td colSpan="2" style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>Tiền cọc đối trừ</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>- {formatMoney(report.depositUsed)}</td></tr>}
    {Number(report.depositForfeited || 0) > 0 && <tr><td colSpan="2" style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>Cọc giữ lại do trả sớm</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(report.depositForfeited)}</td></tr>}
    {Number(report.mustRefund || 0) > 0 && <tr><td colSpan="2" style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>Chủ nhà hoàn khách</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(report.mustRefund)}</td></tr>}
  </>;

  return <div className="receipt-charge-group settlement-breakdown">
    <div className="group-title">CHI TIẾT TẤT TOÁN / TRẢ PHÒNG</div>
    {rowData.map(([label, amount, details]) => <div className="charge-row-v4" key={label}>
      <div className="row-main"><span className="name">{label}</span><span className="amount">{formatMoney(amount || 0)}</span></div>
      {details && <p className="details">{details}</p>}
    </div>)}
    <div className="charge-row-v4"><div className="row-main"><span className="name"><b>Tổng phát sinh</b></span><span className="amount">{formatMoney(report.totalIncurred || 0)}</span></div></div>
    {Number(report.depositUsed || 0) > 0 && <div className="charge-row-v4"><div className="row-main"><span className="name">Tiền cọc đối trừ</span><span className="amount">- {formatMoney(report.depositUsed)}</span></div></div>}
    {Number(report.depositForfeited || 0) > 0 && <div className="charge-row-v4"><div className="row-main"><span className="name">Cọc giữ lại do trả sớm</span><span className="amount">{formatMoney(report.depositForfeited)}</span></div></div>}
    {Number(report.mustRefund || 0) > 0 && <div className="charge-row-v4"><div className="row-main"><span className="name">Chủ nhà hoàn khách</span><span className="amount">{formatMoney(report.mustRefund)}</span></div></div>}
    <div className="charge-row-v4"><div className="row-main"><span className="name"><b>Khách cần thanh toán</b></span><span className="amount">{formatMoney(report.mustCollect || 0)}</span></div></div>
  </div>;
}

function PrintableReceipt({ receipt, room, tenant, bankInfo, settlementReport }) {
  const isMonthly = receipt.type === 'monthly';
  const isRenewalAdjustment = receipt.type === 'renewal_adjustment';
  const paymentState = getReceiptPaymentState(receipt);
  const normalizedReceipt = { ...receipt, status: paymentState.status, debt: paymentState.debt };
  const eOld = getElectricOld(receipt);
  const eNew = getElectricNew(receipt);
  const wOld = getWaterOld(receipt);
  const wNew = getWaterNew(receipt);
  const transferOldUtility = receipt.transferOldRoomUtility;
  const extraOther = Math.max(0, Number(receipt.other || 0) - Number(transferOldUtility?.total || 0));
  const hasTransferBreakdown = isMonthly && transferOldUtility;
  const showCurrentElectric = !hasTransferBreakdown || Number(receipt.electricAmount || 0) > 0 || Number(receipt.electricUsed || 0) > 0;
  const showCurrentWater = !hasTransferBreakdown || Number(receipt.waterAmount || 0) > 0 || Number(receipt.waterUsed || 0) > 0;
  const currentMonthLabel = receipt.month ? `${String(Number(receipt.month.split('/')[0]))}/${receipt.month.split('/')[1]}` : '';
  
  const statusColor = normalizedReceipt.status === 'Đã thanh toán' ? '#166534' : normalizedReceipt.status === 'Nợ một phần' ? '#92400e' : '#991b1b';
  const statusBg = normalizedReceipt.status === 'Đã thanh toán' ? '#dcfce7' : normalizedReceipt.status === 'Nợ một phần' ? '#fef3c7' : '#fee2e2';

  return (
    <div id="printable-receipt" className="printable-receipt">
      {paymentState.isPaid && (
        <div className="paid-stamp-v4">
          <span>ĐÃ THANH TOÁN</span>
          <small>{formatMoney(paymentState.paidAmount)}</small>
        </div>
      )}
      <header className="receipt-header-v4">
        <div className="brand-box">
          <h3 className="brand">ROOM MANAGER</h3>
          <p className="sub">Hệ thống quản lý phòng trọ chuyên nghiệp</p>
        </div>
        <div className="status-section">
          <h1 className="title">{isMonthly ? 'PHIẾU THU TIỀN PHÒNG' : isRenewalAdjustment ? 'PHIẾU THU ĐIỀU CHỈNH GIA HẠN' : 'PHIẾU TẤT TOÁN'}</h1>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <span className="status-badge-liquid" style={{ background: statusBg, color: statusColor, fontSize: '11px' }}>
              {normalizedReceipt.status.toUpperCase()}
            </span>
            <span className="code" style={{ fontSize: '12px', color: '#64748b' }}>{receiptCode(receipt)}</span>
          </div>
        </div>
      </header>

      <div className="info-bar-v4">
        <div className="item"><span className="label">Phòng</span><span className="value">{getTransferRoomLabel(receipt)}</span></div>
        <div className="item"><span className="label">Tháng thu</span><span className="value">{receipt.month}</span></div>
        <div className="item"><span className="label">Ngày lập</span><span className="value">{new Date(receipt.createdAt).toLocaleDateString('vi-VN')}</span></div>
        <div className="item"><span className="label">Hạn trả</span><span className="value">Trong 5 ngày</span></div>
      </div>

      <div className="receipt-tenant-note-v4">
        <p><b>Khách thuê:</b> {tenant.name} • <b>SĐT:</b> {tenant.phone}</p>
        <p><b>Ghi chú:</b> {receipt.note || 'Vui lòng thanh toán trong vòng 5 ngày kể từ ngày nhận phiếu. Xin cảm ơn!'}</p>
      </div>

      <div className="receipt-body-v4">
        <div className="receipt-items-section">
          <h3>CHI TIẾT KHOẢN THU</h3>
          
          {isMonthly ? (
            <>
              {hasTransferBreakdown && (
                <div className="receipt-charge-group old-room">
                  <div className="group-title">I. Chi phí phòng cũ chưa thanh toán</div>
                  <div className="group-meta">
                    <span>Phòng cũ: <b>P{transferOldUtility.oldRoomId}</b></span>
                    <span>Thời gian ở: <b>{formatDisplayDate(transferOldUtility.oldStayFrom)}</b> - <b>{formatDisplayDate(transferOldUtility.oldStayTo)}</b></span>
                    <span>Số ngày tính tiền: <b>{transferOldUtility.oldRoomDays || 0} ngày</b></span>
                  </div>

                  <div className="charge-row-v4">
                    <div className="row-main"><span className="name">🏠 Tiền thuê phòng cũ tháng {transferOldUtility.oldBillingMonth}</span><span className="amount">{formatMoney(transferOldUtility.oldRentAmount || 0)}</span></div>
                    <p className="details">{formatMoney(transferOldUtility.oldMonthlyRent || 0)} / {transferOldUtility.oldMonthDays || 30} x {transferOldUtility.oldRoomDays || 0} ngày</p>
                  </div>

                  {(transferOldUtility.oldServiceAmount || 0) > 0 && (
                    <div className="charge-row-v4">
                      <div className="row-main"><span className="name">🛠️ Dịch vụ phòng cũ</span><span className="amount">{formatMoney(transferOldUtility.oldServiceAmount || 0)}</span></div>
                      <p className="details">{transferOldUtility.oldServiceMode === 'monthly' ? 'Thu cố định theo tháng' : `${formatMoney(transferOldUtility.oldMonthlyServiceFee || 0)} / ${transferOldUtility.oldMonthDays || 30} x ${transferOldUtility.oldRoomDays || 0} ngày`}</p>
                    </div>
                  )}

                  <div className="charge-row-v4">
                    <div className="row-main"><span className="name">⚡ Điện phòng cũ</span><span className="amount">{formatMoney(transferOldUtility.electricAmount || 0)}</span></div>
                    <div className="details">
                      <span>CS cũ: <b>{transferOldUtility.electricOld}</b></span>
                      <span>Mới: <b>{transferOldUtility.electricNew}</b></span>
                      <span>Sử dụng: <b>{transferOldUtility.electricUsed}</b> kWh</span>
                    </div>
                  </div>

                  <div className="charge-row-v4">
                    <div className="row-main"><span className="name">💧 Nước phòng cũ</span><span className="amount">{formatMoney(transferOldUtility.waterAmount || 0)}</span></div>
                    <div className="details">
                      <span>CS cũ: <b>{transferOldUtility.waterOld}</b></span>
                      <span>Mới: <b>{transferOldUtility.waterNew}</b></span>
                      <span>Sử dụng: <b>{transferOldUtility.waterUsed}</b> m³</span>
                    </div>
                  </div>
                </div>
              )}

              <div className={hasTransferBreakdown ? 'receipt-charge-group current-room' : ''}>
                {hasTransferBreakdown && (
                  <>
                    <div className="group-title">II. Chi phí phòng hiện tại</div>
                    <div className="group-meta">
                      <span>Phòng hiện tại: <b>P{receipt.roomId}</b></span>
                      <span>Tháng thu: <b>{receipt.month}</b></span>
                      {receipt.currentRoomChargeDays && receipt.currentRoomMonthDays && receipt.currentRoomChargeDays !== receipt.currentRoomMonthDays && <span>Số ngày tính tiền: <b>{receipt.currentRoomChargeDays} ngày</b></span>}
                    </div>
                  </>
                )}

                <div className="charge-row-v4">
                  <div className="row-main"><span className="name">🏠 Tiền thuê phòng{currentMonthLabel ? ` tháng ${currentMonthLabel}` : ''}</span><span className="amount">{formatMoney(receipt.rent)}</span></div>
                  {receipt.currentRoomChargeDays && receipt.currentRoomMonthDays && receipt.currentRoomChargeDays !== receipt.currentRoomMonthDays && (
                    <p className="details">
                      {receipt.currentRoomChargeFrom && receipt.currentRoomChargeTo && <>Từ {formatDisplayDate(receipt.currentRoomChargeFrom)} đến {formatDisplayDate(receipt.currentRoomChargeTo)}. </>}
                      {formatMoney(receipt.currentRoomMonthlyRent || room.rent || 0)} / {receipt.currentRoomMonthDays} x {receipt.currentRoomChargeDays} ngày
                    </p>
                  )}
                </div>

                <div className="charge-row-v4">
                  <div className="row-main"><span className="name">🛠️ {hasTransferBreakdown ? 'Dịch vụ' : 'Dịch vụ cố định'}{currentMonthLabel ? ` tháng ${currentMonthLabel}` : ''}</span><span className="amount">{formatMoney(receipt.fixedServices)}</span></div>
                  <p className="details">
                    {receipt.currentRoomChargeDays && receipt.currentRoomMonthDays && receipt.currentRoomChargeDays !== receipt.currentRoomMonthDays
                      ? `${formatMoney(receipt.currentRoomMonthlyServiceFee || 0)} / ${receipt.currentRoomMonthDays} x ${receipt.currentRoomChargeDays} ngày. `
                      : ''}
                    (1 người: 200.000đ/tháng; từ 2 người trở lên: 400.000đ/tháng)
                  </p>
                </div>

                {showCurrentElectric && (
                  <div className="charge-row-v4">
                    <div className="row-main"><span className="name">⚡ Tiền điện</span><span className="amount">{formatMoney(receipt.electricAmount)}</span></div>
                    <div className="details">
                      <span>CS cũ: <b>{eOld}</b></span>
                      <span>Mới: <b>{eNew}</b></span>
                      <span>Sử dụng: <b>{receipt.electricUsed}</b> kWh</span>
                      <span>Đơn giá: <b>{formatMoney(room.electricPrice)}</b></span>
                    </div>
                  </div>
                )}

                {showCurrentWater && (
                  <div className="charge-row-v4">
                    <div className="row-main"><span className="name">💧 Tiền nước</span><span className="amount">{formatMoney(receipt.waterAmount)}</span></div>
                    <div className="details">
                      <span>CS cũ: <b>{wOld}</b></span>
                      <span>Mới: <b>{wNew}</b></span>
                      <span>Sử dụng: <b>{receipt.waterUsed}</b> m³</span>
                      <span>Đơn giá: <b>{formatMoney(room.waterPrice)}</b></span>
                    </div>
                  </div>
                )}
              </div>

              {extraOther > 0 && (
                <div className="receipt-charge-group">
                  <div className="group-title">{hasTransferBreakdown ? 'III. Khoản phát sinh khác' : 'Khoản phát sinh khác'}</div>
                  <div className="charge-row-v4">
                    <div className="row-main"><span className="name">➕ Phụ phí khác</span><span className="amount">{formatMoney(extraOther)}</span></div>
                  </div>
                </div>
              )}
            </>
          ) : isRenewalAdjustment ? (
            <div className="receipt-charge-group renewal-adjustment-breakdown">
              <div className="group-title">CHI TIẾT ĐIỀU CHỈNH GIA HẠN</div>
              {Number(receipt.rent || 0) > 0 && <div className="charge-row-v4"><div className="row-main"><span className="name">Chênh lệch tăng giá thuê</span><span className="amount">{formatMoney(receipt.rent)}</span></div></div>}
              {Number(receipt.other || 0) > 0 && <div className="charge-row-v4"><div className="row-main"><span className="name">Chênh lệch tăng tiền cọc</span><span className="amount">{formatMoney(receipt.other)}</span></div></div>}
            </div>
          ) : (
            <SettlementReceiptBreakdown report={settlementReport} receipt={receipt} room={room} />
          )}
        </div>

        <div className="payment-column-v4">
          <div className="payment-card-v4">
            {paymentState.isPaid ? (
              <div className="paid-confirm-card-v4">
                <div className="paid-checkmark-v4">✓</div>
                <p>Phiếu đã thanh toán đủ</p>
                <span>Không cần quét QR thanh toán lại</span>
              </div>
            ) : (
              <div className="qr-box-v4">
                <img src={buildVietQrUrl(bankInfo, receipt)} alt="QR VietQR" />
                <div style={{ textAlign: 'center' }}>
                  <p style={{ fontSize: '11px', fontWeight: '700', margin: 0 }}>QUÉT MÃ THANH TOÁN</p>
                  <p style={{ fontSize: '10px', color: '#64748b' }}>Tự động điền số tiền & nội dung</p>
                </div>
              </div>
            )}

            <div className="payment-details-v4">
              <div className="row"><span className="label">Ngân hàng</span><span className="val">{bankInfo.bankName}</span></div>
              <div className="row"><span className="label">Số tài khoản</span><span className="val">{bankInfo.accountNo}</span></div>
              <div className="row"><span className="label">Chủ tài khoản</span><span className="val">{bankInfo.accountName}</span></div>
              <div className="row"><span className="label">Nội dung CK</span><span className="val" style={{ color: '#1e40af' }}>{transferContent(receipt)}</span></div>
            </div>

            <div className="total-summary-v4">
              <p className="label">{paymentState.isPaid ? 'Tổng cộng đã thu' : paymentState.isAdjustment ? 'Cần thu thêm' : 'Tổng cộng cần trả'}</p>
              <p className="total-amount">{formatMoney(paymentState.isPaid ? receipt.total : (paymentState.debt || receipt.total))}</p>
              {paymentState.isAdjustment && paymentState.basePaidAmount > 0 && <p className="remaining-amount-v4" style={{ color: '#475569' }}>Đã thu trước đó: {formatMoney(paymentState.basePaidAmount)}</p>}
              {paymentState.isPartial && <p className="remaining-amount-v4">Còn nợ: {formatMoney(paymentState.debt)}</p>}
            </div>
          </div>
        </div>
      </div>

      <footer className="receipt-footer-v4">
        <div className="signatures-v4 screen-optional-signatures">
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
  const displayReceipt = enrichReceiptWithTransferUtility(receipt, data);
  const tenant = getTenantForReceipt(data, displayReceipt) || { name: 'N/A' };
  const isMonthly = displayReceipt.type === 'monthly';
  const isRenewalAdjustment = displayReceipt.type === 'renewal_adjustment';
  const transferOldUtility = displayReceipt.transferOldRoomUtility;
  const extraOther = Math.max(0, Number(displayReceipt.other || 0) - Number(transferOldUtility?.total || 0));
  const showCurrentElectric = !transferOldUtility || Number(displayReceipt.electricAmount || 0) > 0 || Number(displayReceipt.electricUsed || 0) > 0;
  const showCurrentWater = !transferOldUtility || Number(displayReceipt.waterAmount || 0) > 0 || Number(displayReceipt.waterUsed || 0) > 0;
  const settlementReport = !isMonthly
    ? (data.moveOutReports || []).find(report => report.contractId === displayReceipt.contractId && report.roomId === displayReceipt.roomId)
    : null;
  return (
    <div className="receipt-page" style={{ padding: '40px', background: 'white', color: 'black', fontFamily: 'serif', position: 'relative', borderBottom: '1px dashed #eee' }}>
      <div style={{ textAlign: 'center', marginBottom: '20px' }}><h1 style={{ fontSize: '20px', fontWeight: 'bold', margin: 0 }}>{isMonthly ? 'PHIẾU THU TIỀN PHÒNG' : isRenewalAdjustment ? 'PHIẾU THU ĐIỀU CHỈNH GIA HẠN' : 'PHIẾU CHỐT TẤT TOÁN'}</h1><p style={{ fontSize: '14px' }}>{isMonthly ? `Tháng ${displayReceipt.month}` : isRenewalAdjustment ? `Áp dụng tháng ${displayReceipt.month}` : 'Quyết toán trả phòng'}</p></div>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}><div><p>Phòng: <b>{getTransferRoomLabel(displayReceipt)}</b></p><p>Khách thuê: <b>{tenant.name}</b></p><p>Ngày lập: {new Date(displayReceipt.createdAt).toLocaleDateString('vi-VN')}</p></div><div style={{ textAlign: 'right' }}><p>Trạng thái: <b>{displayReceipt.status}</b></p></div></div>
      <table className="contract-table" style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }}>
        <thead><tr style={{ background: '#f8fafc' }}><th style={{ border: '1px solid black', padding: '8px' }}>Nội dung</th><th style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>Chỉ số</th><th style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>Thành tiền</th></tr></thead>
        <tbody>
          {isMonthly ? (
            <>
              <tr><td style={{ border: '1px solid black', padding: '8px' }}>Tiền phòng</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.rent)}</td></tr>
              <tr><td style={{ border: '1px solid black', padding: '8px' }}>{transferOldUtility ? 'Dịch vụ tháng' : 'Dịch vụ cố định'}</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>1 người: 200.000đ/tháng; từ 2 người: 400.000đ/tháng</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.fixedServices)}</td></tr>
              {showCurrentElectric && (
                <tr>
                  <td style={{ border: '1px solid black', padding: '8px' }}>Tiền điện</td>
                  <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>
                    CS cũ: {displayReceipt.electricOld} → CS mới: {displayReceipt.electricNew}<br/>
                    (Sử dụng: {displayReceipt.electricUsed} kWh)
                  </td>
                  <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.electricAmount)}</td>
                </tr>
              )}
              {showCurrentWater && (
                <tr>
                  <td style={{ border: '1px solid black', padding: '8px' }}>Tiền nước</td>
                  <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>
                    CS cũ: {displayReceipt.waterOld} → CS mới: {displayReceipt.waterNew}<br/>
                    (Sử dụng: {displayReceipt.waterUsed} m³)
                  </td>
                  <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.waterAmount)}</td>
                </tr>
              )}
              {transferOldUtility && (
                <>
                  <tr style={{ background: '#f8fafc', fontWeight: 'bold' }}><td colSpan="3" style={{ border: '1px solid black', padding: '8px' }}>Chi phí phòng cũ chưa thanh toán - P{transferOldUtility.oldRoomId} ({formatDisplayDate(transferOldUtility.oldStayFrom)} - {formatDisplayDate(transferOldUtility.oldStayTo)}, {transferOldUtility.oldRoomDays || 0} ngày)</td></tr>
                  <tr>
                    <td style={{ border: '1px solid black', padding: '8px' }}>Tiền thuê phòng cũ tháng {transferOldUtility.oldBillingMonth}</td>
                    <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>{formatMoney(transferOldUtility.oldMonthlyRent || 0)} / {transferOldUtility.oldMonthDays || 30} x {transferOldUtility.oldRoomDays || 0} ngày</td>
                    <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(transferOldUtility.oldRentAmount || 0)}</td>
                  </tr>
                  {(transferOldUtility.oldServiceAmount || 0) > 0 && (
                    <tr>
                      <td style={{ border: '1px solid black', padding: '8px' }}>Dịch vụ phòng cũ</td>
                      <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>{transferOldUtility.oldServiceMode === 'monthly' ? 'Thu cố định theo tháng' : `${formatMoney(transferOldUtility.oldMonthlyServiceFee || 0)} / ${transferOldUtility.oldMonthDays || 30} x ${transferOldUtility.oldRoomDays || 0} ngày`}</td>
                      <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(transferOldUtility.oldServiceAmount || 0)}</td>
                    </tr>
                  )}
                  <tr>
                    <td style={{ border: '1px solid black', padding: '8px' }}>Điện phòng cũ P{transferOldUtility.oldRoomId}</td>
                    <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>Điện: {transferOldUtility.electricOld} → {transferOldUtility.electricNew} ({transferOldUtility.electricUsed} kWh)</td>
                    <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(transferOldUtility.electricAmount || 0)}</td>
                  </tr>
                  <tr>
                    <td style={{ border: '1px solid black', padding: '8px' }}>Nước phòng cũ P{transferOldUtility.oldRoomId}</td>
                    <td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>Nước: {transferOldUtility.waterOld} → {transferOldUtility.waterNew} ({transferOldUtility.waterUsed} m³)</td>
                    <td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(transferOldUtility.waterAmount || 0)}</td>
                  </tr>
                </>
              )}
              {extraOther > 0 && <tr><td style={{ border: '1px solid black', padding: '8px' }}>Phụ phí khác</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>-</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(extraOther)}</td></tr>}
            </>
          ) : isRenewalAdjustment ? (
            <>
              {Number(displayReceipt.rent || 0) > 0 && <tr><td style={{ border: '1px solid black', padding: '8px' }}>Chênh lệch tăng giá thuê</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>Theo phụ lục gia hạn</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.rent)}</td></tr>}
              {Number(displayReceipt.other || 0) > 0 && <tr><td style={{ border: '1px solid black', padding: '8px' }}>Chênh lệch tăng tiền cọc</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>Theo phụ lục gia hạn</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.other)}</td></tr>}
            </>
          ) : (
            <SettlementReceiptBreakdown report={settlementReport} receipt={displayReceipt} room={room} table />
          )}
          <tr style={{ fontWeight: 'bold' }}><td colSpan="2" style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>TỔNG CỘNG</td><td style={{ border: '1px solid black', padding: '8px', textAlign: 'right' }}>{formatMoney(displayReceipt.total)}</td></tr>
        </tbody>
      </table>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}><div style={{ textAlign: 'center', width: '200px' }}><p><b>QUÉT MÃ THANH TOÁN</b></p><img src={buildVietQrUrl(bankInfo, displayReceipt)} alt="QR" style={{ width: '120px', border: '1px solid #eee', padding: '5px' }} /><p style={{ fontSize: '10px' }}>{bankInfo.bankName} - {bankInfo.accountNo}</p></div><div style={{ textAlign: 'center', width: '200px' }}><p><b>CHỦ NHÀ KÝ TÊN</b></p><div style={{ height: '80px' }}></div><p><b>DIỆM THỊ BÌNH</b></p></div></div>
      <p style={{ fontStyle: 'italic', fontSize: '12px', marginTop: '20px', textAlign: 'center' }}>Quý khách vui lòng thanh toán trong vòng 5 ngày kể từ ngày nhận phiếu. Trân trọng!</p>
    </div>
  );
}

function RoomOpsModal({ mode, room, onClose, onSave }) {
  const today = new Date().toISOString().slice(0, 10);
  const oldElectric = room.electricNew ?? room.electricOld ?? room.electricStart ?? room.initialElectric ?? 0;
  const oldWater = room.waterNew ?? room.waterOld ?? room.waterStart ?? room.initialWater ?? 0;
  const configs = {
    asset: {
      title: `Thêm tài sản • Phòng ${room.id}`,
      submit: '+ Thêm tài sản',
      initial: {
        name: '',
        category: 'Nội thất',
        quantity: 1,
        handoverStatus: 'Tốt',
        currentStatus: 'Tốt',
        handoverDate: today,
        value: 0,
        deductDeposit: false,
        note: ''
      }
    },
    maintenance: {
      title: `Tạo yêu cầu bảo trì • Phòng ${room.id}`,
      submit: '+ Tạo bảo trì',
      initial: {
        title: '',
        category: 'Thiết bị',
        priority: 'Trung bình',
        reporter: 'Khách thuê',
        assignee: '',
        cost: 0,
        costOwner: 'Chủ trọ',
        status: 'Mới tạo',
        createdDate: today,
        note: ''
      }
    },
    file: {
      title: `Tải tệp lên • Phòng ${room.id}`,
      submit: '+ Lưu tệp',
      initial: {
        name: '',
        group: 'Hợp đồng',
        size: '',
        uploader: 'Admin',
        uploadedAt: today,
        note: ''
      }
    },
    meter: {
      title: `Ghi chỉ số điện nước • Phòng ${room.id}`,
      submit: 'Lưu chỉ số',
      initial: {
        electricOld: oldElectric,
        electricNew: oldElectric,
        waterOld: oldWater,
        waterNew: oldWater,
        readingDate: today,
        reader: 'Admin',
        electricPhoto: '',
        waterPhoto: '',
        note: ''
      }
    }
  };
  const config = configs[mode];
  const [form, setForm] = useState(config.initial);
  const update = (key, value) => setForm(prev => ({ ...prev, [key]: value }));

  const handleSubmit = () => {
    if (mode === 'asset' && !form.name.trim()) return alert('Vui lòng nhập tên tài sản.');
    if (mode === 'maintenance' && !form.title.trim()) return alert('Vui lòng nhập tiêu đề sự cố.');
    if (mode === 'file' && !form.name.trim()) return alert('Vui lòng nhập tên tệp.');
    if (mode === 'meter' && (Number(form.electricNew) < Number(form.electricOld) || Number(form.waterNew) < Number(form.waterOld))) {
      return alert('Chỉ số mới không được nhỏ hơn chỉ số cũ.');
    }
    onSave({
      ...form,
      id: uid(mode),
      roomId: room.id,
      quantity: Number(form.quantity || 0),
      value: Number(form.value || 0),
      cost: Number(form.cost || 0),
      electricOld: Number(form.electricOld || 0),
      electricNew: Number(form.electricNew || 0),
      waterOld: Number(form.waterOld || 0),
      waterNew: Number(form.waterNew || 0),
      createdAt: new Date().toISOString()
    });
  };

  return (
    <div className="modal" onClick={onClose}>
      <div className="detail-modal-v2 liquid-glass" style={{ maxWidth: '680px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>{config.title}</h2>
            <p className="muted small">Dữ liệu được lưu trực tiếp vào hồ sơ phòng và cập nhật badge tương ứng.</p>
          </div>
          <button className="secondary-btn" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ gap: '18px' }}>
          {mode === 'asset' && (
            <div className="form-grid-v2">
              <label style={{ gridColumn: 'span 2' }}>Tên tài sản <input value={form.name} onChange={e => update('name', e.target.value)} placeholder="Giường, tủ, máy lạnh..." /></label>
              <label>Loại <input value={form.category} onChange={e => update('category', e.target.value)} /></label>
              <label>Số lượng <input type="number" value={form.quantity} onChange={e => update('quantity', e.target.value)} /></label>
              <label>Tình trạng bàn giao <input value={form.handoverStatus} onChange={e => update('handoverStatus', e.target.value)} /></label>
              <label>Tình trạng hiện tại <input value={form.currentStatus} onChange={e => update('currentStatus', e.target.value)} /></label>
              <label>Ngày bàn giao <input type="date" value={form.handoverDate} onChange={e => update('handoverDate', e.target.value)} /></label>
              <label>Giá trị ước tính <input type="number" value={form.value} onChange={e => update('value', e.target.value)} /></label>
              <label style={{ gridColumn: 'span 2', flexDirection: 'row', alignItems: 'center' }}><input type="checkbox" checked={form.deductDeposit} onChange={e => update('deductDeposit', e.target.checked)} /> Có trừ cọc nếu hỏng</label>
              <label style={{ gridColumn: 'span 2' }}>Ghi chú <textarea value={form.note} onChange={e => update('note', e.target.value)} /></label>
            </div>
          )}
          {mode === 'maintenance' && (
            <div className="form-grid-v2">
              <label style={{ gridColumn: 'span 2' }}>Tiêu đề sự cố <input value={form.title} onChange={e => update('title', e.target.value)} placeholder="Máy lạnh không lạnh, rò nước..." /></label>
              <label>Loại sự cố <select value={form.category} onChange={e => update('category', e.target.value)}><option>Điện</option><option>Nước</option><option>Thiết bị</option><option>Vệ sinh</option><option>Khác</option></select></label>
              <label>Mức độ <select value={form.priority} onChange={e => update('priority', e.target.value)}><option>Thấp</option><option>Trung bình</option><option>Khẩn cấp</option></select></label>
              <label>Người báo <input value={form.reporter} onChange={e => update('reporter', e.target.value)} /></label>
              <label>Người xử lý <input value={form.assignee} onChange={e => update('assignee', e.target.value)} /></label>
              <label>Chi phí <input type="number" value={form.cost} onChange={e => update('cost', e.target.value)} /></label>
              <label>Bên chịu phí <select value={form.costOwner} onChange={e => update('costOwner', e.target.value)}><option>Chủ trọ</option><option>Khách thuê</option><option>Chia sẻ</option></select></label>
              <label>Trạng thái <select value={form.status} onChange={e => update('status', e.target.value)}><option>Mới tạo</option><option>Đang xử lý</option><option>Chờ linh kiện</option><option>Hoàn tất</option><option>Đã hủy</option></select></label>
              <label>Ngày tạo <input type="date" value={form.createdDate} onChange={e => update('createdDate', e.target.value)} /></label>
              <label style={{ gridColumn: 'span 2' }}>Ghi chú <textarea value={form.note} onChange={e => update('note', e.target.value)} /></label>
            </div>
          )}
          {mode === 'file' && (
            <div className="form-grid-v2">
              <label style={{ gridColumn: 'span 2' }}>Tên tệp <input value={form.name} onChange={e => update('name', e.target.value)} placeholder="Hop-dong-P202.pdf" /></label>
              <label>Nhóm tệp <select value={form.group} onChange={e => update('group', e.target.value)}><option>Hợp đồng</option><option>CCCD</option><option>Thanh toán</option><option>Điện nước</option><option>Tài sản</option><option>Bảo trì</option><option>Khác</option></select></label>
              <label>Dung lượng <input value={form.size} onChange={e => update('size', e.target.value)} placeholder="2.4 MB" /></label>
              <label>Người tải lên <input value={form.uploader} onChange={e => update('uploader', e.target.value)} /></label>
              <label>Ngày tải lên <input type="date" value={form.uploadedAt} onChange={e => update('uploadedAt', e.target.value)} /></label>
              <label style={{ gridColumn: 'span 2' }}>Ghi chú/đường dẫn <textarea value={form.note} onChange={e => update('note', e.target.value)} /></label>
            </div>
          )}
          {mode === 'meter' && (
            <div className="form-grid-v2">
              <label>Điện cũ <input type="number" value={form.electricOld} onChange={e => update('electricOld', e.target.value)} /></label>
              <label>Điện mới <input type="number" value={form.electricNew} onChange={e => update('electricNew', e.target.value)} /></label>
              <label>Nước cũ <input type="number" value={form.waterOld} onChange={e => update('waterOld', e.target.value)} /></label>
              <label>Nước mới <input type="number" value={form.waterNew} onChange={e => update('waterNew', e.target.value)} /></label>
              <label>Ngày ghi <input type="date" value={form.readingDate} onChange={e => update('readingDate', e.target.value)} /></label>
              <label>Người ghi <input value={form.reader} onChange={e => update('reader', e.target.value)} /></label>
              <label>Ảnh đồng hồ điện <input value={form.electricPhoto} onChange={e => update('electricPhoto', e.target.value)} placeholder="Tên file hoặc đường dẫn" /></label>
              <label>Ảnh đồng hồ nước <input value={form.waterPhoto} onChange={e => update('waterPhoto', e.target.value)} placeholder="Tên file hoặc đường dẫn" /></label>
              <label style={{ gridColumn: 'span 2' }}>Ghi chú <textarea value={form.note} onChange={e => update('note', e.target.value)} /></label>
            </div>
          )}
          <div className="btn-group">
            <button className="secondary-btn" onClick={onClose}>Hủy</button>
            <button className="primary-btn" onClick={handleSubmit}>{config.submit}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PaymentModal({ receipt, onClose, onSave }) {
  const receiptPaymentState = getReceiptPaymentState(receipt || {});
  const [paidAmount, setPaidAmount] = useState(receiptPaymentState.debt || receipt?.total || 0);
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
            <p>Phòng: <b>{getTransferRoomLabel(receipt)}</b></p>
            <p>{receiptPaymentState.isAdjustment ? 'Khoản cần thu thêm' : 'Tổng tiền'}: <b>{formatMoney(receiptPaymentState.debt || receipt.total)}</b></p>
            {receiptPaymentState.isAdjustment && <p className="muted small">Đã thu trước đó: {formatMoney(receiptPaymentState.basePaidAmount)}</p>}
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
            {Number(paidAmount) >= (receiptPaymentState.debt || receipt.total) ? 
              <span className="status-badge-liquid active">Thanh toán đủ</span> : 
              Number(paidAmount) > 0 ? 
                <span className="status-badge-liquid notice">Thanh toán một phần</span> : 
                <span className="status-badge-liquid debt">Chưa trả tiền</span>
            }
          </div>
          <button className="primary-btn wide" onClick={() => { 
            if (receiptPaymentState.isAdjustment) {
              const adjustmentDue = Number(receipt.adjustmentDueAmount || receiptPaymentState.adjustmentDue || receipt.total || 0);
              const nextAdjustmentPaid = Number(receipt.adjustmentPaidAmount || 0) + Number(paidAmount || 0);
              let status = 'Chưa thanh toán';
              if (nextAdjustmentPaid >= adjustmentDue) status = 'Đã thanh toán';
              else if (nextAdjustmentPaid > 0) status = 'Nợ một phần';
              onSave({
                ...receipt,
                adjustmentDueAmount: adjustmentDue,
                adjustmentPaidAmount: nextAdjustmentPaid,
                adjustmentPaidDate: paidDate,
                status
              });
              return;
            }
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

function ExpensesTab({ data, onAction }) {
  const [filter, setFilter] = useState(() => ({ supplierId: '', categoryId: '', status: 'all', month: getCurrentMonthLabel() }));

  const filteredExpenses = useMemo(() => {
    return (data.expensePayments || []).filter(e => {
      if (filter.month && e.month !== filter.month) return false;
      if (filter.supplierId && e.supplierId !== filter.supplierId) return false;
      if (filter.categoryId && e.categoryId !== filter.categoryId) return false;
      if (filter.status !== 'all' && e.status !== filter.status) return false;
      return true;
    }).sort((a, b) => new Date(b.paymentDate) - new Date(a.paymentDate));
  }, [data.expensePayments, filter]);

  return (
    <div className="stack" style={{ gap: '16px' }}>
      <div className="widget liquid-glass no-print" style={{ padding: '16px' }}>
        <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap', alignItems: 'flex-end' }}>
          <label style={{ margin: 0 }}>Tháng <input type="month" value={filter.month ? `${filter.month.split('/')[1]}-${filter.month.split('/')[0]}` : ''} onChange={e => {
            if (!e.target.value) return setFilter({...filter, month: ''});
            const [y, m] = e.target.value.split('-'); 
            setFilter({...filter, month: `${m}/${y}`}); 
          }} /></label>
          <label style={{ margin: 0 }}>Nhà CC / Nhà cung cấp <select value={filter.supplierId} onChange={e => setFilter({...filter, supplierId: e.target.value})}><option value="">Tất cả</option>{(data.suppliers || []).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></label>
          <label style={{ margin: 0 }}>Loại chi phí <select value={filter.categoryId} onChange={e => setFilter({...filter, categoryId: e.target.value})}><option value="">Tất cả</option>{(data.expenseCategories || []).map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select></label>
          <label style={{ margin: 0 }}>Trạng thái <select value={filter.status} onChange={e => setFilter({...filter, status: e.target.value})}><option value="all">Tất cả</option><option value="paid">Đã thanh toán</option><option value="partial">Thanh toán một phần</option><option value="unpaid">Chưa thanh toán</option></select></label>
          <div style={{ display: 'flex', gap: '8px' }}>
            <button className="primary-btn" onClick={() => onAction('create_expense')}>+ Tạo phiếu chi</button>
            <button className="secondary-btn" onClick={() => onAction('manage_suppliers')}>🏢 Quản lý Nhà CC</button>
          </div>
        </div>
      </div>
      <div className="widget liquid-glass" style={{ padding: 0 }}>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Mã phiếu</th><th>Ngày</th><th>Nhà CC</th><th>Loại</th><th>Nội dung</th><th>Tổng tiền</th><th>Đã trả</th><th>Trạng thái</th><th>Thao tác</th></tr></thead>
            <tbody>
              {filteredExpenses.map(e => {
                const supplier = (data.suppliers || []).find(s => s.id === e.supplierId);
                const category = (data.expenseCategories || []).find(c => c.id === e.categoryId);
                return (
                  <tr key={e.id}>
                    <td><span className="small muted mono">{e.expenseCode}</span></td>
                    <td>{e.paymentDate ? e.paymentDate.split('-').reverse().join('/') : 'N/A'}</td>
                    <td><b>{supplier?.name || e.recipientName || 'Vãng lai'}</b></td>
                    <td><span className="small muted">{category?.name}</span></td>
                    <td>{e.title}</td>
                    <td style={{ fontWeight: '700' }}>{formatMoney(e.totalAmount)}</td>
                    <td>{formatMoney(e.paidAmount)}</td>
                    <td>
                      <span className={`status-badge-liquid ${e.status === 'paid' ? 'active' : e.status === 'partial' ? 'notice' : 'debt'}`}>
                        {e.status === 'paid' ? 'ĐÃ THANH TOÁN' : e.status === 'partial' ? 'THANH TOÁN MỘT PHẦN' : 'CHƯA THANH TOÁN'}
                      </span>
                    </td>
                    <td>
                      <div style={{ display: 'flex', gap: '4px' }}>
                        <button className="secondary-btn sm" onClick={() => onAction('edit_expense', e)}>✏️</button>
                        <button className="secondary-btn sm" onClick={() => onAction('view_qr', e)}>📱</button>
                        <button className="secondary-btn sm" style={{ color: '#ef4444' }} onClick={() => onAction('delete_expense', e)}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filteredExpenses.length === 0 && <tr><td colSpan="9" style={{ textAlign: 'center', padding: '40px' }}>Không có dữ liệu chi phí.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SuppliersTab({ data, onAction }) {
  return (
    <div className="stack" style={{ gap: '16px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <button className="secondary-btn" onClick={() => onAction('back')}>⬅ Quay lại Chi phí</button>
        <button className="primary-btn" onClick={() => onAction('add_supplier')}>+ Thêm Nhà CC</button>
      </div>
      <div className="widget liquid-glass" style={{ padding: 0 }}>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Tên Nhà CC</th><th>SĐT</th><th>Ngân hàng</th><th>Số tài khoản</th><th>Ghi chú</th><th>Thao tác</th></tr></thead>
            <tbody>
              {data.suppliers.map(s => (
                <tr key={s.id}>
                  <td><b>{s.name}</b></td>
                  <td>{s.phone}</td>
                  <td>{s.bankName}</td>
                  <td>{s.bankAccount}</td>
                  <td><span className="small muted">{s.note}</span></td>
                  <td>
                    <div style={{ display: 'flex', gap: '4px' }}>
                      <button className="secondary-btn sm" onClick={() => onAction('edit_supplier', s)}>✏️ Sửa</button>
                      <button className="secondary-btn sm" style={{ color: '#ef4444' }} onClick={() => onAction('delete_supplier', s)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
              {data.suppliers.length === 0 && <tr><td colSpan="6" style={{ textAlign: 'center', padding: '40px' }}>Chưa có nhà cung cấp nào.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function FinancialReportTab({ data, onAction }) {
  const [selectedMonth, setSelectedMonth] = useState(() => getCurrentMonthLabel());

  const stats = useMemo(() => {
    const monthIncome = (data.receipts || []).filter(r => r.month === selectedMonth);
    const totalIncome = monthIncome.reduce((sum, r) => sum + (r.paidAmount || 0), 0);
    const incomeDebt = monthIncome.reduce((sum, r) => sum + (r.total - (r.paidAmount || 0)), 0);

    const monthExpense = (data.expensePayments || []).filter(e => e.month === selectedMonth);
    const totalExpense = monthExpense.reduce((sum, e) => sum + (e.paidAmount || 0), 0);
    const expenseDebt = monthExpense.reduce((sum, e) => sum + (e.totalAmount - (e.paidAmount || 0)), 0);

    return {
      totalIncome,
      totalExpense,
      profit: totalIncome - totalExpense,
      incomeDebt,
      expenseDebt,
      incomeVouchers: monthIncome,
      expenseVouchers: monthExpense
    };
  }, [data, selectedMonth]);

  return (
    <div className="stack" style={{ gap: '24px' }}>
      <div className="widget liquid-glass no-print" style={{ padding: '16px' }}>
        <div style={{ display: 'flex', gap: '16px', alignItems: 'center' }}>
          <h2 style={{ fontSize: '18px', margin: 0 }}>Báo cáo tháng</h2>
          <input type="month" value={selectedMonth ? `${selectedMonth.split('/')[1]}-${selectedMonth.split('/')[0]}` : ''} onChange={e => {
            if (!e.target.value) return;
            const [y, m] = e.target.value.split('-'); 
            setSelectedMonth(`${m}/${y}`); 
          }} />
        </div>
      </div>

      <div className="stats-grid">
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--success)' }}>
          <p className="stat-label">Tổng thu</p>
          <p className="stat-value">{formatMoney(stats.totalIncome)}</p>
          <p className="stat-note">Tiền thực nhận từ khách</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--danger)' }}>
          <p className="stat-label">Tổng chi</p>
          <p className="stat-value">{formatMoney(stats.totalExpense)}</p>
          <p className="stat-note">Tiền thực chi cho nhà CC</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid var(--primary)' }}>
          <p className="stat-label">Lợi nhuận</p>
          <p className="stat-value" style={{ color: stats.profit >= 0 ? 'var(--success)' : 'var(--danger)' }}>{formatMoney(stats.profit)}</p>
          <p className="stat-note">Thu - Chi</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid #f43f5e' }}>
          <p className="stat-label">Nợ khách thuê</p>
          <p className="stat-value">{formatMoney(stats.incomeDebt)}</p>
          <p className="stat-note">Chưa thu từ khách</p>
        </div>
        <div className="stat-card-liquid" style={{ borderLeft: '4px solid #8b5cf6' }}>
          <p className="stat-label">Nợ nhà cung cấp</p>
          <p className="stat-value">{formatMoney(stats.expenseDebt)}</p>
          <p className="stat-note">Chưa trả cho NCC</p>
        </div>
      </div>

      <div className="dashboard-grid-main" style={{ gridTemplateColumns: '1fr 1fr' }}>
        <div className="widget liquid-glass" style={{ padding: 0 }}>
          <div style={{ padding: '16px', borderBottom: '1px solid rgba(0,0,0,0.05)' }}>
            <h3 style={{ fontSize: '16px' }}>📥 Chi tiết khoản thu</h3>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Phòng</th><th>Loại</th><th>Tổng</th><th>Đã thu</th></tr></thead>
              <tbody>
                {stats.incomeVouchers.map(r => (
                  <tr key={r.id}>
                    <td><b>P{r.roomId}</b></td>
                    <td><span className="small muted">{receiptTypeLabel(r.type)}</span></td>
                    <td>{formatMoney(r.total)}</td>
                    <td style={{ color: 'var(--success)', fontWeight: '600' }}>{formatMoney(r.paidAmount)}</td>
                  </tr>
                ))}
                {stats.incomeVouchers.length === 0 && <tr><td colSpan="4" style={{ textAlign: 'center', padding: '20px' }}>Không có khoản thu nào.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>

        <div className="widget liquid-glass" style={{ padding: 0 }}>
          <div style={{ padding: '16px', borderBottom: '1px solid rgba(0,0,0,0.05)' }}>
            <h3 style={{ fontSize: '16px' }}>📤 Chi tiết khoản chi</h3>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Nhà CC</th><th>Loại</th><th>Tổng</th><th>Đã trả</th></tr></thead>
              <tbody>
                {stats.expenseVouchers.map(e => {
                  const supplier = (data.suppliers || []).find(s => s.id === e.supplierId);
                  const category = (data.expenseCategories || []).find(c => c.id === e.categoryId);
                  return (
                    <tr key={e.id}>
                      <td><b>{supplier?.name || 'Vãng lai'}</b></td>
                      <td><span className="small muted">{category?.name}</span></td>
                      <td>{formatMoney(e.totalAmount || 0)}</td>
                      <td style={{ color: 'var(--danger)', fontWeight: '600' }}>{formatMoney(e.paidAmount || 0)}</td>
                    </tr>
                  );
                })}
                {stats.expenseVouchers.length === 0 && <tr><td colSpan="4" style={{ textAlign: 'center', padding: '20px' }}>Không có khoản chi nào.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

function ExpenseModal({ expense, data, onClose, onSave }) {
  const [form, setForm] = useState(expense || {
    id: uid('exp'),
    supplierId: '',
    categoryId: 'cat_other',
    month: getCurrentMonthLabel(),
    paymentDate: new Date().toISOString().slice(0, 10),
    title: '',
    description: '',
    totalAmount: 0,
    paidAmount: 0,
    status: 'unpaid',
    paymentMethod: 'transfer',
    note: ''
  });

  // Tự sinh mã phiếu chi
  useEffect(() => {
    if (!expense && !form.expenseCode) {
      const [m, y] = (form.month || getCurrentMonthLabel()).split('/');
      const prefix = `PC-${y}${m}`;
      const count = (data.expensePayments || []).filter(e => e.expenseCode && e.expenseCode.startsWith(prefix)).length + 1;
      const code = `${prefix}-${String(count).padStart(3, '0')}`;
      setForm(prev => ({ ...prev, expenseCode: code }));
    }
  }, [form.month, expense, data.expensePayments]);

  useEffect(() => {
    let status = 'unpaid';
    if (form.paidAmount >= form.totalAmount && form.totalAmount > 0) status = 'paid';
    else if (form.paidAmount > 0) status = 'partial';
    setForm(prev => ({ ...prev, status }));
  }, [form.totalAmount, form.paidAmount]);

  // Tự chọn loại chi phí theo nhà cung cấp
  const handleSupplierChange = (val) => {
    const sup = (data.suppliers || []).find(s => s.id === val);
    const updates = { supplierId: val };
    if (sup && sup.defaultCategory) {
      updates.categoryId = sup.defaultCategory;
    }
    // Nếu chọn vãng lai, xóa tên người nhận cũ
    if (!val) updates.recipientName = '';
    setForm(prev => ({ ...prev, ...updates }));
  };

  return (
    <div className="modal">
      <div className="detail-modal-v2" style={{ maxWidth: '700px' }}>
        <div className="modal-header">
          <div><h2 style={{ fontSize: '20px' }}>{expense ? 'Sửa phiếu chi' : 'Tạo phiếu chi mới'}</h2><p className="muted">Quản lý khoản thanh toán cho nhà cung cấp</p></div>
          <button className="secondary-btn sm" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ gap: '20px' }}>
          <div className="form-grid-v2">
            <label>Mã phiếu chi (Tự sinh)
              <input type="text" value={form.expenseCode || ''} readOnly style={{ background: '#f8fafc', fontWeight: 'bold' }} />
            </label>
            <label>Nhà CC / Nhà cung cấp
              <select value={form.supplierId} onChange={e => handleSupplierChange(e.target.value)}>
                <option value="">-- Vãng lai / Không có trong danh sách --</option>
                {(data.suppliers || []).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
          </div>
          
          {!form.supplierId && (
            <label>Tên người nhận / Đơn vị nhận tiền (Bắt buộc khi vãng lai)
              <input type="text" value={form.recipientName || ''} onChange={e => setForm({...form, recipientName: e.target.value})} placeholder="Nhập tên người nhận tiền..." />
            </label>
          )}

          <div className="form-grid-v2">
            <label>Loại chi phí
              <select value={form.categoryId} onChange={e => setForm({...form, categoryId: e.target.value})}>
                {data.expenseCategories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
            <label>Tháng áp dụng
              <input type="month" value={form.month ? `${form.month.split('/')[1]}-${form.month.split('/')[0]}` : ''} onChange={e => {
                if (!e.target.value) return;
                const [y, m] = e.target.value.split('-'); 
                setForm({...form, month: `${m}/${y}`}); 
              }} />
            </label>
            <label>Ngày thanh toán
              <input type="date" value={form.paymentDate} onChange={e => setForm({...form, paymentDate: e.target.value})} />
            </label>
          </div>
          <label>Nội dung chi
            <input type="text" placeholder="VD: Tiền điện tháng 5, Sửa vòi nước..." value={form.title} onChange={e => setForm({...form, title: e.target.value})} />
          </label>
          <div className="form-grid-v2">
            <label>Tổng tiền (VND)
              <input type="number" value={form.totalAmount} onChange={e => setForm({...form, totalAmount: Number(e.target.value)})} />
            </label>
            <label>Đã trả (VND)
              <input type="number" value={form.paidAmount} onChange={e => setForm({...form, paidAmount: Number(e.target.value)})} />
            </label>
            <label>Phương thức
              <select value={form.paymentMethod} onChange={e => setForm({...form, paymentMethod: e.target.value})}>
                <option value="transfer">Chuyển khoản</option>
                <option value="cash">Tiền mặt</option>
                <option value="qr">Quét mã QR</option>
              </select>
            </label>
            <label>Trạng thái tự tính
              <span className={`status-badge-liquid ${form.status === 'paid' ? 'active' : form.status === 'partial' ? 'notice' : 'debt'}`} style={{ height: '48px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                {form.status === 'paid' ? 'Đã thanh toán' : form.status === 'partial' ? 'Thanh toán một phần' : 'Chưa thanh toán'}
              </span>
            </label>
          </div>
          <label>Ghi chú
            <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})} placeholder="Thêm ghi chú nếu cần..."></textarea>
          </label>
          <div className="op-action-footer">
            <button className="secondary-btn" onClick={onClose}>Hủy</button>
            <button className="primary-btn" onClick={() => {
              if (!form.title) return alert('Vui lòng nhập nội dung chi.');
              if (!form.supplierId && !form.recipientName) return alert('Vui lòng nhập tên người nhận khi chọn Vãng lai.');
              onSave(form);
            }}>Lưu phiếu chi</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function SupplierModal({ supplier, onClose, onSave }) {
  const [form, setForm] = useState(supplier || {
    id: uid('sup'),
    name: '',
    phone: '',
    email: '',
    address: '',
    bankName: '',
    bankAccount: '',
    bankOwner: '',
    note: ''
  });

  return (
    <div className="modal">
      <div className="detail-modal-v2">
        <div className="modal-header">
          <div><h2 style={{ fontSize: '20px' }}>{supplier ? 'Sửa Nhà cung cấp' : 'Thêm Nhà cung cấp'}</h2><p className="muted">Thông tin đối tác / người nhận thanh toán</p></div>
          <button className="secondary-btn sm" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ gap: '20px' }}>
          <div className="form-grid-v2">
            <label>Tên Nhà cung cấp / Tên riêng
              <input type="text" value={form.name} onChange={e => setForm({...form, name: e.target.value})} placeholder="VD: Điện lực, Internet VNPT..." />
            </label>
            <label>Số điện thoại
              <input type="text" value={form.phone} onChange={e => setForm({...form, phone: e.target.value})} />
            </label>
          </div>
          <div className="form-grid-v2">
            <label>Tên Ngân hàng
              <input type="text" value={form.bankName} onChange={e => setForm({...form, bankName: e.target.value})} placeholder="VD: MB, VCB..." />
            </label>
            <label>Số tài khoản
              <input type="text" value={form.bankAccount} onChange={e => setForm({...form, bankAccount: e.target.value})} />
            </label>
            <label>Tên chủ tài khoản
              <input type="text" value={form.bankOwner} onChange={e => setForm({...form, bankOwner: e.target.value})} />
            </label>
          </div>
          <label>Ghi chú
            <textarea value={form.note} onChange={e => setForm({...form, note: e.target.value})}></textarea>
          </label>
          <div className="op-action-footer">
            <button className="secondary-btn" onClick={onClose}>Hủy</button>
            <button className="primary-btn" onClick={() => {
              if (!form.name) return alert('Vui lòng nhập tên nhà cung cấp.');
              onSave(form);
            }}>Lưu thông tin</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ExpenseQRModal({ expense, supplier, onClose }) {
  const recipientName = supplier?.name || expense.recipientName || 'Vãng lai';
  const bankName = supplier?.bankName || expense.recipientBankName || '';
  const bankAccount = supplier?.bankAccount || expense.recipientBankAccount || '';
  const bankOwner = supplier?.bankOwner || expense.recipientBankOwner || recipientName;
  const qrImageUrl = expense.recipientQrImageUrl || expense.qrImageUrl || '';
  const remainingAmount = Math.max(0, Number(expense.totalAmount || 0) - Number(expense.paidAmount || 0));
  const displayAmount = remainingAmount > 0 ? remainingAmount : Number(expense.totalAmount || 0);
  const transferText = `${expense.expenseCode || ''} ${expense.title || ''}`.trim();
  const copyExpenseTransfer = () => {
    const text = `${bankName}\nSTK: ${bankAccount}\nChủ TK: ${bankOwner}\nSố tiền: ${formatMoney(displayAmount)}\nNội dung: ${transferText}`;
    navigator.clipboard.writeText(text);
    alert('Đã copy thông tin chuyển khoản phiếu chi.');
  };

  return (
    <div className="modal no-print" onClick={onClose}>
      <div className="detail-modal-v2" style={{ maxWidth: '400px', textAlign: 'center' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2 style={{ fontSize: '18px' }}>Mã QR Thanh toán</h2>
          <button className="secondary-btn sm" onClick={onClose}>✕</button>
        </div>
        <div className="detail-body-v2 stack" style={{ alignItems: 'center' }}>
          <p>Thanh toán cho: <b>{recipientName}</b></p>
          <p style={{ fontSize: '24px', fontWeight: '800', color: 'var(--primary)' }}>{formatMoney(displayAmount)}</p>
          
          {qrImageUrl && (
            <div className="settlement-qr-card" style={{ width: '100%' }}>
              <h4>QR người nhận</h4>
              <img src={qrImageUrl} alt="QR người nhận tiền" />
            </div>
          )}

          {bankAccount ? (
            <div className="widget" style={{ background: '#f8fafc', padding: '20px', borderRadius: '16px', border: '1px dashed #cbd5e1' }}>
              <p className="small muted">THÔNG TIN CHUYỂN KHOẢN</p>
              <p style={{ marginTop: '10px' }}>Ngân hàng: <b>{bankName || 'Chưa cập nhật'}</b></p>
              <p>Số TK: <b>{bankAccount}</b></p>
              <p>Chủ TK: <b>{bankOwner || 'Chưa cập nhật'}</b></p>
              <p style={{ marginTop: '10px' }} className="small">Nội dung: <b>{transferText}</b></p>
            </div>
          ) : (
            <div className="warning-box">Phiếu chi này chưa có thông tin ngân hàng. Có thể lưu link ảnh QR hoặc cập nhật STK người nhận.</div>
          )}
          
          {(bankAccount || qrImageUrl) && <button className="secondary-btn wide" onClick={copyExpenseTransfer}>📋 Copy nội dung CK</button>}
          <button className="primary-btn wide" style={{ marginTop: '20px' }} onClick={() => window.print()}>🖨️ In chứng từ</button>
        </div>
      </div>
    </div>
  );
}

export default App;
