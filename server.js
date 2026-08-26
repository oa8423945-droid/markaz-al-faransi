const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const ExcelJS = require('exceljs');

const PORT = Number(process.env.PORT || 3210);
const ROOT = __dirname;
const PUBLIC = path.join(ROOT, 'public');
const DATA_DIR = path.join(ROOT, 'data');
const PID_FILE = path.join(ROOT, 'garage-server.pid');
const PRIMARY_DATA_FILE = 'F:\\codx\\المركز الفرنسي\\main data 2.xlsx';
const DATA_FILE = process.env.GARAGE_DATA_FILE
  ? path.resolve(process.env.GARAGE_DATA_FILE)
  : (fs.existsSync(PRIMARY_DATA_FILE) ? PRIMARY_DATA_FILE : path.join(DATA_DIR, 'main data 2.xlsx'));

const CUSTOMER_SHEET = 'بيانات العميل ';
const VISIT_SHEET = 'قسم الزياره';
const INVENTORY_SHEET = 'قسم المخرن';
const SUPPLIER_SHEET = 'قسم الموردين';
const EXPENSE_SHEET = 'المصروفات ';
const MOVEMENT_SHEET = 'حركة المخزن';
const ACCOUNT_SHEET = 'حركة الحسابات';
const EMPLOYEE_SHEET = 'الموظفين';
const CUSTOMER_HEADERS = ['تاريخ التسجيل', 'الاسم', 'رقم التلفون', 'رقم العربيه', 'نوع العربيه', 'كود العميل', 'المبلغ المستحق على العميل', 'المبلغ المستحق على المركز'];
const VISIT_HEADERS = ['تاريخ الزياره', 'قراءة العداد', 'نوع الصيانه', 'اسم الفني', 'كود القطع الغيار المستخدمه', 'اجمالي سعر قطع الغيار', 'المصنعية', 'اجمالي', 'كود العميل', 'رقم العربيه', 'كود الزيارة', 'ملاحظات', 'اجمالي تكلفة قطع الغيار', 'كود حركة المخزن', 'طريقة الدفع', 'تفاصيل طرق الدفع', 'المبلغ المدفوع', 'المبلغ المتبقي'];
const INVENTORY_HEADERS = ['اسم المنتج', 'تفاصيل المنتج', 'بلد المنشاء', 'كود المنتج', 'الكميه', 'سعر الشراء', 'سعر البيع', 'هامش الربح', 'اسم المورد', 'المبلغ المدفوع', 'المبلغ المتبقي'];
const SUPPLIER_HEADERS = ['كود المورد', 'اسم المورد', 'رقم التلفون', 'تاريخ التعاقد', 'المبلغ المطلوب', 'ملاحظات', 'المبلغ المدفوع', 'تاريخ الدفع'];
const EXPENSE_HEADERS = ['البيان', 'المبلغ', 'التاريخ'];
const MOVEMENT_HEADERS = ['كود الحركة', 'التاريخ', 'نوع الحركة', 'كود المنتج', 'اسم المنتج', 'تفاصيل المنتج', 'بلد المنشأ', 'الكمية', 'سعر الوحدة', 'الإجمالي', 'كود الزيارة', 'كود العميل', 'اسم المورد', 'ملاحظات'];
const ACCOUNT_HEADERS = ['كود الحركة', 'التاريخ', 'الوقت', 'تاريخ التنفيذ', 'اتجاه الحركة', 'نوع الحركة', 'البيان', 'كود العميل', 'اسم العميل', 'كود الزيارة', 'كود المورد', 'اسم المورد', 'كود الموظف', 'اسم الموظف', 'كود المنتج', 'اسم المنتج', 'الكمية', 'إجمالي المبلغ', 'المبلغ المدفوع', 'المبلغ المستحق', 'طريقة الدفع', 'تفاصيل طرق الدفع', 'رصيد المركز', 'ملاحظات'];
const EMPLOYEE_HEADERS = ['كود الموظف', 'الاسم', 'رقم التلفون', 'تاريخ التوظيف', 'التخصص', 'عدد المشاركات في الأعمال', 'المرتب الأسبوعي', 'مبالغ على الموظف', 'مبالغ للموظف', 'الحالة', 'سبب الإيقاف', 'تاريخ الإيقاف', 'ملاحظات'];

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.writeFileSync(PID_FILE, String(process.pid));
const clearPidFile = () => {
  try {
    if (fs.existsSync(PID_FILE) && clean(fs.readFileSync(PID_FILE, 'utf8')) === String(process.pid)) fs.unlinkSync(PID_FILE);
  } catch (_) {}
};
process.on('exit', clearPidFile);
process.on('SIGINT', () => { clearPidFile(); process.exit(0); });
process.on('SIGTERM', () => { clearPidFile(); process.exit(0); });

function clean(value) {
  return String(value ?? '').trim();
}

function currentTime() {
  return new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });
}

// PDFKit يعكس أقواس الإنجليزي وسط النص العربي؛ نحول محتواها إلى جزء واضح بين شرطات.
function pdfRtlText(doc, text, x, y, width, options = {}) {
  const value = String(text ?? '').replace(/\(([^()]*)\)/g, ' — $1 — ').replace(/\s*—\s*—\s*/g, ' — ').replace(/\s+/g, '\u00A0');
  const style = () => doc.font(options.bold ? 'ArabicBold' : 'Arabic').fontSize(options.size || 9).fillColor(options.color || '#111827');
  style().text(value, x, y, { width, align: 'right', lineGap: options.lineGap || 0 });
}

function normalizedPlate(value) {
  return clean(value).replace(/[\s-]/g, '').toLocaleLowerCase('ar-EG');
}

function normalizedPhone(value) {
  const arabicDigits = '٠١٢٣٤٥٦٧٨٩';
  let phone = clean(value)
    .replace(/[٠-٩]/g, (digit) => String(arabicDigits.indexOf(digit)))
    .replace(/\D/g, '');
  if (phone.startsWith('0020')) phone = `0${phone.slice(4)}`;
  else if (phone.startsWith('20') && phone.length === 12) phone = `0${phone.slice(2)}`;
  return phone;
}

function normalizedAutoCode(value, prefix) {
  const text = clean(value);
  const match = text.match(new RegExp(`^${prefix}-?0*(\\d+)$`, 'i'));
  return match ? `${prefix}${Number(match[1])}` : text;
}

function findSheetName(workbook, wanted) {
  const exact = workbook.SheetNames.find((name) => clean(name) === clean(wanted));
  return exact || wanted;
}

function loadWorkbook() {
  if (!fs.existsSync(DATA_FILE)) throw new Error('ملف البيانات غير موجود. انسخ Book1.xlsx إلى مجلد data.');
  return XLSX.readFile(DATA_FILE, { cellDates: false });
}

function rowsFromSheet(workbook, wanted, headers) {
  const name = findSheetName(workbook, wanted);
  const sheet = workbook.Sheets[name];
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  if (!rows.length) return [];
  const actualHeaders = rows[0].map((header) => clean(header));
  return rows.slice(1).filter((row) => row.some((cell) => clean(cell))).map((row) =>
    Object.fromEntries(headers.map((header) => {
      const index = actualHeaders.findIndex((actual) => actual === clean(header));
      return [header, index >= 0 ? row[index] ?? '' : ''];
    }))
  );
}

function writeSheet(workbook, wanted, headers, objects) {
  const name = findSheetName(workbook, wanted);
  workbook.Sheets[name] = XLSX.utils.aoa_to_sheet([
    headers,
    ...objects.map((object) => headers.map((header) => object[header] ?? '')),
  ]);
  if (!workbook.SheetNames.includes(name)) workbook.SheetNames.push(name);
}

function nextCode(existing, prefix) {
  const max = existing.reduce((result, value) => {
    const match = normalizedAutoCode(value, prefix).match(new RegExp(`^${prefix}(\\d+)$`, 'i'));
    return Math.max(result, match ? Number(match[1]) : 0);
  }, 0);
  return `${prefix}${max + 1}`;
}

function appendAccount(accounts, account) {
  if (!account.time) account.time = currentTime();
  const currentBalance = accounts.reduce((balance, entry) => {
    const cash = Number(entry.paid) || 0;
    return balance + (entry.direction === 'وارد' ? cash : entry.direction === 'صادر' ? -cash : 0);
  }, 0);
  const cash = Number(account.paid) || 0;
  account.balance = currentBalance + (account.direction === 'وارد' ? cash : account.direction === 'صادر' ? -cash : 0);
  accounts.push(account);
  return account;
}

function syncCustomerDebts(customers, accounts) {
  for (const customer of customers) {
    const linked = accounts.filter((account) => account.customerCode === customer.code);
    customer.dueFromCustomer = Math.max(0, linked.reduce((sum, account) => sum + (account.direction === 'مديونية على الغير' ? account.due : account.direction === 'وارد' && !['سداد مديونية', 'سداد مديونية عميل'].includes(account.type) ? account.due : ['سداد مديونية', 'سداد مديونية عميل'].includes(account.type) ? -account.paid : 0), 0));
    customer.dueFromCenter = linked.reduce((sum, account) => sum + (account.direction === 'مديونية على المركز' ? account.due : 0), 0);
  }
}

function paymentInfo(input, fallbackAmount = 0) {
  const allowed = new Set(['نقدي', 'إنستاباي', 'فودافون كاش', 'تحويل بنكي', 'آجل']);
  let payments = Array.isArray(input.payments) ? input.payments.map((entry) => ({
    method: clean(entry.method), amount: Math.max(0, Number(entry.amount) || 0),
  })).filter((entry) => allowed.has(entry.method) && entry.amount > 0) : [];
  if (!payments.length) {
    const method = allowed.has(clean(input.paymentMethod)) ? clean(input.paymentMethod) : 'نقدي';
    const amount = clean(input.paid) ? Math.max(0, Number(input.paid) || 0) : Math.max(0, Number(fallbackAmount) || 0);
    if (amount) payments = [{ method, amount }];
  }
  const paid = payments.filter((entry) => entry.method !== 'آجل').reduce((sum, entry) => sum + entry.amount, 0);
  const allocated = payments.reduce((sum, entry) => sum + entry.amount, 0);
  const methods = [...new Set(payments.map((entry) => entry.method))];
  return { payments, paid, allocated, paymentMethod: methods.join(' + ') || 'آجل', paymentDetails: payments.map((entry) => `${entry.method}: ${entry.amount}`).join(' | ') };
}

function migrateStoredAutoCodes() {
  const workbook = loadWorkbook();
  let changed = false;
  const migrateSheet = (wanted, columns) => {
    const name = findSheetName(workbook, wanted);
    const sheet = workbook.Sheets[name];
    if (!sheet?.['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const headerIndexes = {};
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const header = clean(sheet[XLSX.utils.encode_cell({ r: range.s.r, c: column })]?.v);
      if (columns[header]) headerIndexes[column] = columns[header];
    }
    for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
      for (const [column, prefix] of Object.entries(headerIndexes)) {
        const address = XLSX.utils.encode_cell({ r: row, c: Number(column) });
        const cell = sheet[address];
        if (!cell) continue;
        const normalized = normalizedAutoCode(cell.v, prefix);
        if (normalized !== clean(cell.v)) {
          cell.v = normalized;
          cell.t = 's';
          delete cell.w;
          changed = true;
        }
      }
    }
  };
  migrateSheet(CUSTOMER_SHEET, { 'كود العميل': 'C' });
  migrateSheet(VISIT_SHEET, { 'كود العميل': 'C', 'كود الزيارة': 'V', 'كود حركة المخزن': 'M' });
  migrateSheet(MOVEMENT_SHEET, { 'كود الحركة': 'M', 'كود الزيارة': 'V', 'كود العميل': 'C' });
  migrateSheet(SUPPLIER_SHEET, { 'كود المورد': 'S' });
  migrateSheet(ACCOUNT_SHEET, { 'كود الحركة': 'A', 'كود المورد': 'S', 'كود الموظف': 'E' });
  migrateSheet(EMPLOYEE_SHEET, { 'كود الموظف': 'E' });
  if (changed) XLSX.writeFile(workbook, DATA_FILE, { compression: true });
}

function readData() {
  const workbook = loadWorkbook();
  const customerRows = rowsFromSheet(workbook, CUSTOMER_SHEET, CUSTOMER_HEADERS);
  const visitRows = rowsFromSheet(workbook, VISIT_SHEET, VISIT_HEADERS);
  const inventoryRows = rowsFromSheet(workbook, INVENTORY_SHEET, INVENTORY_HEADERS);
  const movementRows = rowsFromSheet(workbook, MOVEMENT_SHEET, MOVEMENT_HEADERS);
  const accountRows = rowsFromSheet(workbook, ACCOUNT_SHEET, ACCOUNT_HEADERS);
  const employeeRows = rowsFromSheet(workbook, EMPLOYEE_SHEET, EMPLOYEE_HEADERS);
  const customers = customerRows.map((row) => ({
    name: clean(row['الاسم']), phone: clean(row['رقم التلفون']), plate: clean(row['رقم العربيه']),
    carType: clean(row['نوع العربيه']), code: normalizedAutoCode(row['كود العميل'], 'C'), registeredDate: clean(row['تاريخ التسجيل']),
    dueFromCustomer: Number(row['المبلغ المستحق على العميل']) || 0, dueFromCenter: Number(row['المبلغ المستحق على المركز']) || 0,
  })).filter((customer) => customer.code || customer.name || customer.plate);
  const visits = visitRows.map((row) => ({
    date: clean(row['تاريخ الزياره']), mileage: Number(row['قراءة العداد']) || 0,
    serviceType: clean(row['نوع الصيانه']), technician: clean(row['اسم الفني']),
    partsCodes: clean(row['كود القطع الغيار المستخدمه']), partsTotal: Number(row['اجمالي سعر قطع الغيار']) || 0,
    labor: Number(row['المصنعية']) || 0, total: Number(row['اجمالي']) || 0,
    customerCode: normalizedAutoCode(row['كود العميل'], 'C'), plate: clean(row['رقم العربيه']),
    code: normalizedAutoCode(row['كود الزيارة'], 'V'), notes: clean(row['ملاحظات']), partsCost: Number(row['اجمالي تكلفة قطع الغيار']) || 0,
    stockMovementCode: normalizedAutoCode(row['كود حركة المخزن'], 'M'),
    paymentMethod: clean(row['طريقة الدفع']), paymentDetails: clean(row['تفاصيل طرق الدفع']), paid: Number(row['المبلغ المدفوع']) || 0, due: Number(row['المبلغ المتبقي']) || 0,
  })).filter((visit) => visit.date || visit.code || visit.customerCode || visit.plate);
  const inventory = inventoryRows.map((row) => ({
    name: clean(row['اسم المنتج']), details: clean(row['تفاصيل المنتج']), country: clean(row['بلد المنشاء']),
    code: clean(row['كود المنتج']), qty: Number(row['الكميه']) || 0,
    buy: Number(row['سعر الشراء']) || 0, sell: Number(row['سعر البيع']) || 0,
    margin: Number(row['هامش الربح']) || 0, supplier: clean(row['اسم المورد']),
    paid: Number(row['المبلغ المدفوع']) || 0, due: Number(row['المبلغ المتبقي']) || 0,
  })).filter((item) => item.code || item.name);
  const supplierRows = rowsFromSheet(workbook, SUPPLIER_SHEET, SUPPLIER_HEADERS);
  const expenseRows = rowsFromSheet(workbook, EXPENSE_SHEET, EXPENSE_HEADERS);
  const suppliers = supplierRows.map((row) => ({
    code: normalizedAutoCode(row['كود المورد'], 'S'), name: clean(row['اسم المورد']), phone: clean(row['رقم التلفون']), contractDate: clean(row['تاريخ التعاقد']), due: Number(row['المبلغ المطلوب']) || 0,
    notes: clean(row['ملاحظات']), paid: Number(row['المبلغ المدفوع']) || 0, paymentDate: clean(row['تاريخ الدفع']),
  })).filter((item) => item.name || item.due || item.paid);
  for (const supplier of suppliers) {
    if (!supplier.code) supplier.code = nextCode(suppliers.map((item) => item.code), 'S');
  }
  const expenses = expenseRows.map((row) => ({
    description: clean(row['البيان']), amount: Number(row['المبلغ']) || 0, date: clean(row['التاريخ']),
  })).filter((item) => item.description || item.amount);
  const movements = movementRows.map((row) => ({
    code: normalizedAutoCode(row['كود الحركة'], 'M'), date: clean(row['التاريخ']), type: clean(row['نوع الحركة']),
    productCode: clean(row['كود المنتج']), productName: clean(row['اسم المنتج']), details: clean(row['تفاصيل المنتج']),
    country: clean(row['بلد المنشأ']), qty: Number(row['الكمية']) || 0, unitPrice: Number(row['سعر الوحدة']) || 0,
    total: Number(row['الإجمالي']) || 0, visitCode: normalizedAutoCode(row['كود الزيارة'], 'V'), customerCode: normalizedAutoCode(row['كود العميل'], 'C'),
    supplier: clean(row['اسم المورد']), notes: clean(row['ملاحظات']),
  })).filter((item) => item.code || item.productCode || item.type);
  for (const visit of visits) {
    if (customers.some((customer) => customer.code === visit.customerCode)) continue;
    const customerByPlate = customers.find((customer) => normalizedPlate(customer.plate) === normalizedPlate(visit.plate));
    if (customerByPlate) visit.customerCode = customerByPlate.code;
  }
  for (const movement of movements) {
    if (!movement.visitCode) continue;
    const relatedVisit = visits.find((visit) => visit.code === movement.visitCode);
    if (relatedVisit) movement.customerCode = relatedVisit.customerCode;
  }
  const accounts = accountRows.map((row) => ({
    code: clean(row['كود الحركة']).toUpperCase().startsWith('T') ? normalizedAutoCode(row['كود الحركة'], 'T') : normalizedAutoCode(row['كود الحركة'], 'A'),
    date: clean(row['التاريخ']), time: clean(row['الوقت']), executionDate: clean(row['تاريخ التنفيذ']) || clean(row['التاريخ']), direction: clean(row['اتجاه الحركة']), type: clean(row['نوع الحركة']),
    description: clean(row['البيان']), customerCode: normalizedAutoCode(row['كود العميل'], 'C'), customerName: clean(row['اسم العميل']),
    visitCode: normalizedAutoCode(row['كود الزيارة'], 'V'), supplierCode: normalizedAutoCode(row['كود المورد'], 'S'), supplierName: clean(row['اسم المورد']),
    employeeCode: normalizedAutoCode(row['كود الموظف'], 'E'), employeeName: clean(row['اسم الموظف']),
    productCode: clean(row['كود المنتج']), productName: clean(row['اسم المنتج']), qty: Number(row['الكمية']) || 0,
    total: Number(row['إجمالي المبلغ']) || 0, paid: Number(row['المبلغ المدفوع']) || 0,
    due: Number(row['المبلغ المستحق']) || 0, paymentMethod: clean(row['طريقة الدفع']), paymentDetails: clean(row['تفاصيل طرق الدفع']), balance: Number(row['رصيد المركز']) || 0, notes: clean(row['ملاحظات']),
  })).filter((item) => item.code || item.type || item.supplierName);
  let runningBalance = 0;
  for (const account of accounts) {
    if (!account.direction) {
      if (account.type.includes('إيراد') || account.type.includes('إيداع')) account.direction = 'وارد';
      else if (account.type.includes('سداد') || account.type.includes('توريد') || account.type.includes('مصروف') || account.type.includes('سحب')) account.direction = 'صادر';
    }
    const cash = Number(account.paid) || 0;
    runningBalance += account.direction === 'وارد' ? cash : account.direction === 'صادر' ? -cash : 0;
    account.balance = runningBalance;
  }
  syncCustomerDebts(customers, accounts);
  const employees = employeeRows.map((row) => ({
    code: normalizedAutoCode(row['كود الموظف'], 'E'), name: clean(row['الاسم']), phone: clean(row['رقم التلفون']), hireDate: clean(row['تاريخ التوظيف']),
    specialty: clean(row['التخصص']), contributions: Number(row['عدد المشاركات في الأعمال']) || 0, weeklySalary: Number(row['المرتب الأسبوعي']) || 0,
    debtOnEmployee: Number(row['مبالغ على الموظف']) || 0, dueToEmployee: Number(row['مبالغ للموظف']) || 0,
    status: clean(row['الحالة']) || 'يعمل', stopReason: clean(row['سبب الإيقاف']), stopDate: clean(row['تاريخ الإيقاف']), notes: clean(row['ملاحظات']),
  })).filter((item) => item.code || item.name);
  for (const employee of employees) {
    employee.contributions = visits.filter((visit) => clean(visit.technician).toLocaleLowerCase('ar-EG') === employee.name.toLocaleLowerCase('ar-EG')).length;
  }
  return { customers, visits, inventory, suppliers, expenses, movements, accounts, employees };
}

function writeInventorySheet(workbook, inventory) {
  const name = findSheetName(workbook, INVENTORY_SHEET);
  const rows = [INVENTORY_HEADERS, ...inventory.map((item) => [
    item.name, item.details, item.country, item.code, item.qty, item.buy, item.sell,
    item.sell - item.buy, item.supplier, item.paid, item.due,
  ])];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  inventory.forEach((_, index) => {
    const row = index + 2;
    sheet[`H${row}`] = { t: 'n', f: `G${row}-F${row}`, v: Number(inventory[index].sell) - Number(inventory[index].buy) };
  });
  workbook.Sheets[name] = sheet;
  if (!workbook.SheetNames.includes(name)) workbook.SheetNames.push(name);
}

function saveData(customers, visits, inventory = null, suppliers = null, movements = null, accounts = null, employees = null) {
  const workbook = loadWorkbook();
  if (accounts) syncCustomerDebts(customers, accounts);
  writeSheet(workbook, CUSTOMER_SHEET, CUSTOMER_HEADERS, customers.map((customer) => ({
    'الاسم': customer.name, 'رقم التلفون': customer.phone, 'رقم العربيه': customer.plate,
    'نوع العربيه': customer.carType, 'كود العميل': customer.code, 'تاريخ التسجيل': customer.registeredDate,
    'المبلغ المستحق على العميل': customer.dueFromCustomer || 0, 'المبلغ المستحق على المركز': customer.dueFromCenter || 0,
  })));
  writeSheet(workbook, VISIT_SHEET, VISIT_HEADERS, visits.map((visit) => ({
    'تاريخ الزياره': visit.date, 'قراءة العداد': visit.mileage, 'نوع الصيانه': visit.serviceType,
    'اسم الفني': visit.technician, 'كود القطع الغيار المستخدمه': visit.partsCodes,
    'اجمالي سعر قطع الغيار': visit.partsTotal, 'المصنعية': visit.labor, 'اجمالي': visit.total,
    'كود العميل': visit.customerCode, 'رقم العربيه': visit.plate, 'كود الزيارة': visit.code,
    'ملاحظات': visit.notes, 'اجمالي تكلفة قطع الغيار': visit.partsCost, 'كود حركة المخزن': visit.stockMovementCode,
    'طريقة الدفع': visit.paymentMethod, 'تفاصيل طرق الدفع': visit.paymentDetails, 'المبلغ المدفوع': visit.paid, 'المبلغ المتبقي': visit.due,
  })));
  if (inventory) {
    writeInventorySheet(workbook, inventory);
  }
  if (suppliers) {
    writeSheet(workbook, SUPPLIER_SHEET, SUPPLIER_HEADERS, suppliers.map((supplier) => ({
      'كود المورد': supplier.code, 'اسم المورد': supplier.name, 'رقم التلفون': supplier.phone, 'تاريخ التعاقد': supplier.contractDate, 'المبلغ المطلوب': supplier.due,
      'ملاحظات': supplier.notes, 'المبلغ المدفوع': supplier.paid, 'تاريخ الدفع': supplier.paymentDate,
    })));
  }
  if (movements) {
    writeSheet(workbook, MOVEMENT_SHEET, MOVEMENT_HEADERS, movements.map((movement) => ({
      'كود الحركة': movement.code, 'التاريخ': movement.date, 'نوع الحركة': movement.type,
      'كود المنتج': movement.productCode, 'اسم المنتج': movement.productName, 'تفاصيل المنتج': movement.details,
      'بلد المنشأ': movement.country, 'الكمية': movement.qty, 'سعر الوحدة': movement.unitPrice,
      'الإجمالي': movement.total, 'كود الزيارة': movement.visitCode, 'كود العميل': movement.customerCode,
      'اسم المورد': movement.supplier, 'ملاحظات': movement.notes,
    })));
  }
  if (accounts) {
    writeSheet(workbook, ACCOUNT_SHEET, ACCOUNT_HEADERS, accounts.map((account) => ({
      'كود الحركة': account.code, 'التاريخ': account.date, 'الوقت': account.time, 'نوع الحركة': account.type, 'البيان': account.description,
      'تاريخ التنفيذ': account.executionDate || account.date, 'اتجاه الحركة': account.direction,
      'كود العميل': account.customerCode, 'اسم العميل': account.customerName, 'كود الزيارة': account.visitCode,
      'كود المورد': account.supplierCode, 'اسم المورد': account.supplierName,
      'كود الموظف': account.employeeCode, 'اسم الموظف': account.employeeName,
      'كود المنتج': account.productCode, 'اسم المنتج': account.productName, 'الكمية': account.qty,
      'إجمالي المبلغ': account.total, 'المبلغ المدفوع': account.paid,
      'المبلغ المستحق': account.due, 'طريقة الدفع': account.paymentMethod, 'تفاصيل طرق الدفع': account.paymentDetails, 'رصيد المركز': account.balance, 'ملاحظات': account.notes,
    })));
  }
  if (employees) {
    writeSheet(workbook, EMPLOYEE_SHEET, EMPLOYEE_HEADERS, employees.map((employee) => ({
      'كود الموظف': employee.code, 'الاسم': employee.name, 'رقم التلفون': employee.phone, 'تاريخ التوظيف': employee.hireDate,
      'التخصص': employee.specialty, 'عدد المشاركات في الأعمال': employee.contributions, 'المرتب الأسبوعي': employee.weeklySalary,
      'مبالغ على الموظف': employee.debtOnEmployee, 'مبالغ للموظف': employee.dueToEmployee, 'الحالة': employee.status || 'يعمل',
      'سبب الإيقاف': employee.stopReason, 'تاريخ الإيقاف': employee.stopDate, 'ملاحظات': employee.notes,
    })));
  }
  XLSX.writeFile(workbook, DATA_FILE, { compression: true });
}

function repairStoredCustomerLinks() {
  const workbook = loadWorkbook();
  const rawVisits = rowsFromSheet(workbook, VISIT_SHEET, VISIT_HEADERS);
  const rawMovements = rowsFromSheet(workbook, MOVEMENT_SHEET, MOVEMENT_HEADERS);
  const data = readData();
  const needsVisitRepair = data.visits.some((visit) => {
    const raw = rawVisits.find((row) => normalizedAutoCode(row['كود الزيارة'], 'V') === visit.code);
    return raw && normalizedAutoCode(raw['كود العميل'], 'C') !== visit.customerCode;
  });
  const needsMovementRepair = data.movements.some((movement) => {
    const raw = rawMovements.find((row) => normalizedAutoCode(row['كود الحركة'], 'M') === movement.code && clean(row['كود المنتج']) === movement.productCode);
    return raw && normalizedAutoCode(raw['كود العميل'], 'C') !== movement.customerCode;
  });
  if (needsVisitRepair || needsMovementRepair) saveData(data.customers, data.visits, null, null, data.movements);
}

function ensureStoredSupplierCodes() {
  const workbook = loadWorkbook();
  const rawSuppliers = rowsFromSheet(workbook, SUPPLIER_SHEET, SUPPLIER_HEADERS);
  const supplierSheetName = findSheetName(workbook, SUPPLIER_SHEET);
  const actualHeaders = XLSX.utils.sheet_to_json(workbook.Sheets[supplierSheetName], { header: 1, defval: '', raw: false })[0] || [];
  const missingContractHeader = !actualHeaders.some((header) => clean(header) === 'تاريخ التعاقد');
  if (!missingContractHeader && !rawSuppliers.some((row) => clean(row['اسم المورد']) && !clean(row['كود المورد']))) return;
  const data = readData();
  saveData(data.customers, data.visits, null, data.suppliers, null);
}

function ensureStoredAccountSchema() {
  const workbook = loadWorkbook();
  const name = findSheetName(workbook, ACCOUNT_SHEET);
  const sheet = workbook.Sheets[name];
  if (!sheet) return;
  const actualHeaders = (XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false })[0] || []).map(clean);
  if (ACCOUNT_HEADERS.every((header) => actualHeaders.includes(header))) return;
  const data = readData();
  saveData(data.customers, data.visits, null, null, null, data.accounts);
}

function ensureStoredCustomerSchema() {
  const workbook = loadWorkbook();
  const name = findSheetName(workbook, CUSTOMER_SHEET);
  const sheet = workbook.Sheets[name];
  if (!sheet) return;
  const actualHeaders = (XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false })[0] || []).map(clean);
  if (CUSTOMER_HEADERS.every((header) => actualHeaders.includes(header))) return;
  const data = readData();
  saveData(data.customers, data.visits, null, null, null, data.accounts);
}

function ensureStoredEmployeeSchema() {
  const workbook = loadWorkbook();
  const name = findSheetName(workbook, EMPLOYEE_SHEET);
  const sheet = workbook.Sheets[name];
  const actualHeaders = sheet ? (XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false })[0] || []).map(clean) : [];
  if (sheet && EMPLOYEE_HEADERS.every((header) => actualHeaders.includes(header))) return;
  const data = readData();
  saveData(data.customers, data.visits, null, null, null, null, data.employees);
}

function partsAndQuantities(partsCodes, inventory) {
  const text = clean(partsCodes);
  if (!text) return '—';
  const byCode = new Map((inventory || []).map((item) => [normalizedAutoCode(item.code, 'P').toLocaleLowerCase('en-US'), item]));
  return text.split(/[،,]+/).map((entry) => {
    const itemText = clean(entry).replace(/^\(+|\)+$/g, '').trim();
    const match = itemText.match(/^(.*?)\s*(?:x|×)\s*(\d+(?:\.\d+)?)\s*$/i);
    const code = clean(match ? match[1] : itemText);
    const quantity = match ? match[2] : '';
    const product = byCode.get(normalizedAutoCode(code, 'P').toLocaleLowerCase('en-US'));
    if (!product) return quantity ? `${code} × ${quantity}` : code;
    const statement = [clean(product.name), clean(product.details)].filter(Boolean).join(' — ');
    return `${statement || code}${quantity ? ` × ${quantity}` : ''}`;
  }).filter(Boolean).join(' | ') || text;
}

function dailyCloseWorkbook(data, fromDate, toDate = fromDate, reportTitle = 'إقفال اليومية') {
  const accounts = data.accounts.filter((account) => {
    const date = account.executionDate || account.date;
    return date >= fromDate && date <= toDate;
  });
  const incoming = accounts.filter((account) => account.direction === 'وارد').reduce((sum, account) => sum + account.paid, 0);
  const outgoing = accounts.filter((account) => account.direction === 'صادر').reduce((sum, account) => sum + account.paid, 0);
  const currentBalance = data.accounts.length ? data.accounts[data.accounts.length - 1].balance : 0;
  const beforeBalance = data.accounts.filter((account) => (account.executionDate || account.date) < fromDate).reduce((balance, account) => account.direction === 'وارد' ? balance + account.paid : account.direction === 'صادر' ? balance - account.paid : balance, 0);
  const rows = [
    [reportTitle, `${fromDate} — ${toDate}`], [],
    ['الملخص السريع'],
    ['رصيد المركز قبل الفترة', beforeBalance], ['إجمالي الإضافات (وارد)', incoming], ['إجمالي الخصومات (صادر)', outgoing], ['صافي الحركة', incoming - outgoing], ['الرصيد الحالي للمركز', currentBalance], [],
  ];
  const addSection = (title, headers, values) => {
    rows.push([title], headers, ...(values.length ? values : [['لا توجد حركات']]), []);
  };
  const withdrawals = accounts.filter((account) => account.direction === 'صادر' && account.code.startsWith('T'));
  addSection('الوارد', ['التاريخ', 'الوقت', 'البيان', 'النوع', 'كود الحركة', 'المبلغ', 'طريقة الدفع'], accounts.filter((account) => account.direction === 'وارد').map((account) => [account.executionDate || account.date, account.time || '—', account.description, account.type, account.code, account.paid, account.paymentMethod]));
  addSection('الصادر والخصومات', ['التاريخ', 'الوقت', 'البيان', 'النوع', 'كود الحركة', 'المبلغ', 'طريقة الدفع'], accounts.filter((account) => account.direction === 'صادر').map((account) => [account.executionDate || account.date, account.time || '—', account.description, account.type, account.code, account.paid, account.paymentMethod]));
  addSection('المسحوبات والمصروفات التشغيلية', ['البيان', 'النوع', 'المبلغ', 'طريقة الدفع', 'الملاحظات'], withdrawals.map((account) => [account.description, account.type, account.paid, account.paymentMethod, account.notes]));
  const additionalExpenses = accounts.filter((account) => account.direction === 'صادر' && account.type === 'مصروفات إضافية');
  addSection('المصروفات الإضافية', ['البيان', 'المبلغ', 'طريقة الدفع', 'الملاحظات'], additionalExpenses.map((account) => [account.description, account.paid, account.paymentMethod, account.notes]));
  const visitsWithParts = data.visits.filter((visit) => visit.date >= fromDate && visit.date <= toDate && visit.partsCodes);
  addSection('قطع الغيار المستخدمة في الزيارات', ['العميل', 'كود العميل', 'كود الزيارة', 'القطع والكميات', 'تكلفة الشراء', 'سعر البيع', 'هامش الربح'], visitsWithParts.map((visit) => {
    const customer = data.customers.find((item) => item.code === visit.customerCode) || {};
    return [customer.name, visit.customerCode, visit.code, partsAndQuantities(visit.partsCodes, data.inventory), visit.partsCost || 0, visit.partsTotal || 0, (visit.partsTotal || 0) - (visit.partsCost || 0)];
  }));
  const supplierPayments = accounts.filter((account) => account.type === 'سداد مستحقات');
  addSection('دفعات الموردين', ['المورد', 'كود المورد', 'المبلغ', 'طريقة الدفع', 'البيان'], supplierPayments.map((account) => [account.supplierName, account.supplierCode, account.paid, account.paymentMethod, account.description]));
  const sales = accounts.filter((account) => account.type === 'إيراد زيارة صيانة');
  addSection('المبيعات', ['العميل', 'كود العميل', 'كود الزيارة', 'نوع العربية', 'المصنعية', 'إجمالي الفاتورة', 'المبلغ المدفوع', 'المتبقي', 'طريقة الدفع'], sales.map((account) => {
    const visit = data.visits.find((item) => item.code === account.visitCode) || {};
    const customer = data.customers.find((item) => item.code === account.customerCode) || {};
    return [account.customerName || customer.name, account.customerCode, account.visitCode, customer.carType, visit.labor || 0, visit.total || account.total || 0, account.paid, account.due, account.paymentMethod];
  }));
  addSection('كشف الحساب', ['التاريخ', 'الوقت', 'البيان', 'نوع الحركة', 'كود الحركة', 'ملاحظات', 'مدين', 'دائن', 'الرصيد بعد العملية'], accounts.map((account) => [account.executionDate || account.date, account.time || '—', account.description, account.type, account.code, account.notes, account.direction === 'صادر' ? account.paid : 0, account.direction === 'وارد' ? account.paid : 0, account.balance]));
  addSection('كل الحركات المالية', ACCOUNT_HEADERS, accounts.map((account) => [account.code, account.date, account.time, account.executionDate, account.direction, account.type, account.description, account.customerCode, account.customerName, account.visitCode, account.supplierCode, account.supplierName, account.employeeCode, account.employeeName, account.productCode, account.productName, account.qty, account.total, account.paid, account.due, account.paymentMethod, account.paymentDetails, account.balance, account.notes]));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = Array.from({ length: 20 }, (_, index) => ({ wch: index === 0 ? 28 : 18 }));
  XLSX.utils.book_append_sheet(workbook, sheet, 'إقفال اليومية');
  return workbook;
}

async function dailyCloseBuffer(data, fromDate, toDate = fromDate, reportTitle = 'إقفال اليومية') {
  const raw = XLSX.write(dailyCloseWorkbook(data, fromDate, toDate, reportTitle), { type: 'buffer', bookType: 'xlsx', compression: true });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(raw);
  const sheet = workbook.getWorksheet('إقفال اليومية');
  // نُبقي شكل ملف Excel بسيطاً مثل النسخة المرجعية، مع اتجاه عربي فقط.
  sheet.views = [{ rightToLeft: true, showGridLines: true }];
  const maxColumns = 20;
  sheet.eachRow((row, rowNumber) => {
    for (let column = 1; column <= maxColumns; column += 1) {
      const cell = row.getCell(column);
      if (typeof cell.value === 'string') cell.value = cell.value.replace(/(\d{4})-(\d{2})-(\d{2})/g, '$1/$2/$3');
    }
  });
  workbook.creator = 'المركز الفرنسي';
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function rtlWorkbookBuffer(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  workbook.worksheets.forEach((sheet) => { sheet.views = [{ rightToLeft: true, showGridLines: true }]; });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function reportHtml(data, fromDate, toDate, title) {
  const escape = (value) => clean(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const accounts = data.accounts.filter((account) => { const date = account.executionDate || account.date; return date >= fromDate && date <= toDate; });
  const incoming = accounts.filter((a) => a.direction === 'وارد').reduce((sum, a) => sum + a.paid, 0);
  const outgoing = accounts.filter((a) => a.direction === 'صادر').reduce((sum, a) => sum + a.paid, 0);
  const beforeBalance = data.accounts.filter((account) => (account.executionDate || account.date) < fromDate).reduce((balance, account) => account.direction === 'وارد' ? balance + account.paid : account.direction === 'صادر' ? balance - account.paid : balance, 0);
  const balance = data.accounts.length ? data.accounts[data.accounts.length - 1].balance : 0;
  const parts = data.visits.filter((visit) => visit.date >= fromDate && visit.date <= toDate && visit.partsCodes);
  const section = (heading, headers, bodyRows, empty = 'لا توجد بيانات') => `<section><h2>${heading}</h2><table><thead><tr>${headers.map((header) => `<th>${header}</th>`).join('')}</tr></thead><tbody>${bodyRows || `<tr><td colspan="${headers.length}">${empty}</td></tr>`}</tbody></table></section>`;
  const incomingRows = accounts.filter((a) => a.direction === 'وارد').map((a) => `<tr><td>${escape(a.executionDate || a.date)}</td><td>${escape(a.time || '—')}</td><td>${escape(a.description)}</td><td>${escape(a.type)}</td><td>${escape(a.code)}</td><td>${a.paid}</td><td>${escape(a.paymentMethod || '—')}</td></tr>`).join('');
  const outgoingRows = accounts.filter((a) => a.direction === 'صادر').map((a) => `<tr><td>${escape(a.executionDate || a.date)}</td><td>${escape(a.time || '—')}</td><td>${escape(a.description)}</td><td>${escape(a.type)}</td><td>${escape(a.code)}</td><td>${a.paid}</td><td>${escape(a.paymentMethod || '—')}</td></tr>`).join('');
  const partRows = parts.map((visit) => { const customer = data.customers.find((item) => item.code === visit.customerCode) || {}; const cost = Number(visit.partsCost || 0); const sale = Number(visit.partsTotal || 0); return `<tr><td>${escape(customer.name || '—')}</td><td>${escape(visit.code)}</td><td>${escape(partsAndQuantities(visit.partsCodes, data.inventory))}</td><td>${cost}</td><td>${sale}</td><td>${sale - cost}</td></tr>`; }).join('');
  return `<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><title>${escape(title)}</title><style>body{font-family:Tahoma,Arial;margin:28px;color:#111;background:#fff}h1{text-align:center;color:#17243c;margin:0 0 4px}p{text-align:center;color:#667085}.summary{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin:20px 0}.summary div{border:1px solid #d7dde7;border-radius:8px;padding:11px;text-align:center;background:#f8fafc}.summary b{display:block;font-size:18px;margin-top:6px;color:#111827}section{border:1px solid #d7dde7;border-radius:10px;margin:16px 0;padding:12px;page-break-inside:avoid}h2{font-size:15px;border-right:4px solid #2676ee;padding-right:8px;margin:0 0 10px;color:#17243c}table{border-collapse:collapse;width:100%;font-size:10px}th,td{border:1px solid #d7dde7;padding:7px;text-align:right}th{background:#eef5ff;color:#17243c}footer{border-top:2px solid #f7941d;margin-top:22px;padding-top:10px;text-align:center;color:#667085;font-size:10px}@media print{button{display:none}body{margin:8mm}.summary{grid-template-columns:repeat(5,1fr)}}button{background:#2676ee;color:#fff;border:0;padding:9px 15px;border-radius:7px;font-weight:bold;cursor:pointer}</style></head><body><button onclick="window.print()">طباعة / حفظ PDF</button><h1>${escape(title)}</h1><p>${escape(fromDate)} — ${escape(toDate)}</p><div class="summary"><div>رصيد قبل الفترة<b>${beforeBalance}</b></div><div>الإضافات / الوارد<b>${incoming}</b></div><div>الخصومات / الصادر<b>${outgoing}</b></div><div>صافي الحركة<b>${incoming - outgoing}</b></div><div>الرصيد الحالي<b>${balance}</b></div></div>${section('الوارد', ['التاريخ','الوقت','البيان','النوع','كود الحركة','المبلغ','طريقة الدفع'], incomingRows)}${section('الصادر والخصومات', ['التاريخ','الوقت','البيان','النوع','كود الحركة','المبلغ','طريقة الدفع'], outgoingRows)}${section('قطع الغيار المستخدمة', ['العميل','كود الزيارة','القطع والكميات','تكلفة الشراء','سعر البيع','هامش الربح'], partRows)}<footer>المركز الفرنسي - تقرير يومية مالي</footer><script>window.addEventListener('load',()=>setTimeout(()=>window.print(),300))</script></body></html>`;
}

function supplierDebtRows(data) {
  return data.suppliers.map((supplier) => ({ code: supplier.code, name: supplier.name,
    paid: Number(supplier.paid) || 0, due: Number(supplier.due) || 0,
    total: (Number(supplier.paid) || 0) + (Number(supplier.due) || 0) }))
    .filter((supplier) => supplier.total > 0).sort((first, second) => second.due - first.due);
}

async function supplierDebtsWorkbook(data) {
  const rows = supplierDebtRows(data);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'المركز الفرنسي'; workbook.created = new Date();
  const sheet = workbook.addWorksheet('ديون الموردين', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 3 }] });
  sheet.mergeCells('A1:E1'); sheet.getCell('A1').value = 'تقرير ديون الموردين';
  sheet.getCell('A1').font = { name: 'Tahoma', size: 16, bold: true, color: { argb: 'FF17243C' } };
  sheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }; sheet.getRow(1).height = 30;
  sheet.getCell('A2').value = 'تاريخ التقرير'; sheet.getCell('B2').value = new Date(); sheet.getCell('B2').numFmt = 'yyyy-mm-dd';
  sheet.addRow(['كود المورد', 'اسم المورد', 'إجمالي المبلغ', 'المبلغ المدفوع', 'المبلغ المتبقي']);
  rows.forEach((supplier) => sheet.addRow([supplier.code, supplier.name, supplier.total, supplier.paid, supplier.due]));
  const firstDataRow = 4; const lastDataRow = sheet.rowCount;
  const sums = rows.reduce((result, supplier) => ({ total: result.total + supplier.total, paid: result.paid + supplier.paid, due: result.due + supplier.due }), { total: 0, paid: 0, due: 0 });
  const totalValues = rows.length ? [{ formula: `SUM(C${firstDataRow}:C${lastDataRow})`, result: sums.total }, { formula: `SUM(D${firstDataRow}:D${lastDataRow})`, result: sums.paid }, { formula: `SUM(E${firstDataRow}:E${lastDataRow})`, result: sums.due }] : [0, 0, 0];
  const totalRow = sheet.addRow([null, 'الإجمالي', ...totalValues]);
  sheet.getRow(3).height = 25; sheet.getRow(3).eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD966' } };
    cell.font = { name: 'Tahoma', bold: true, color: { argb: 'FF111827' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { bottom: { style: 'medium', color: { argb: 'FFE2A900' } } };
  });
  sheet.eachRow((row, rowNumber) => { if (rowNumber < 4) return;
    row.font = { name: 'Tahoma', size: 11, bold: rowNumber === totalRow.number }; row.alignment = { vertical: 'middle', horizontal: 'right' }; row.height = 22;
    row.eachCell((cell) => { cell.border = { bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } } }; });
  });
  totalRow.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF4CC' } }; });
  ['C', 'D', 'E'].forEach((column) => { sheet.getColumn(column).numFmt = '#,##0.00 "ج"'; });
  sheet.columns = [{ width: 16 }, { width: 28 }, { width: 20 }, { width: 20 }, { width: 20 }];
  if (rows.length) sheet.autoFilter = `A3:E${sheet.rowCount - 1}`;
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function supplierDebtsPdf(data) {
  return new Promise((resolve, reject) => { try {
    const rows = supplierDebtRows(data); const doc = new PDFDocument({ size: 'A4', margins: { top: 38, right: 42, bottom: 45, left: 42 }, bufferPages: true });
    const chunks = []; doc.on('data', (chunk) => chunks.push(chunk)); doc.on('end', () => resolve(Buffer.concat(chunks))); doc.on('error', reject);
    const fonts = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
    const regular = ['tahoma.ttf', 'arial.ttf'].map((name) => path.join(fonts, name)).find(fs.existsSync);
    const bold = ['tahomabd.ttf', 'arialbd.ttf'].map((name) => path.join(fonts, name)).find(fs.existsSync);
    if (!regular || !bold) throw new Error('تعذر العثور على خط عربي مناسب لإنشاء التقرير.');
    doc.registerFont('Arabic', regular); doc.registerFont('ArabicBold', bold);
    const reverse = (value) => [...String(value)].reverse().join('');
    const money = (value) => Number(value || 0).toFixed(2);
    const rtl = (text, x, y, width, options = {}) => pdfRtlText(doc, text, x, y, width, options);
    // يرسم PDF من اليسار لليمين، لذا نعكس المصفوفة ليبدأ الجدول فعليًا بكود المورد من اليمين.
    const left = 42; const width = doc.page.width - 84; const columns = [90, 90, 95, 155, 85];
    const drawHeader = (continued = false) => { rtl(continued ? 'تقرير ديون الموردين - تابع' : 'تقرير ديون الموردين', left, 45, width, { bold: true, size: 20, color: '#17243C' }); rtl(`تاريخ التقرير: ${reverse(new Date().toISOString().slice(0, 10))}`, left, 75, width, { size: 9, color: '#667085' }); doc.moveTo(left, 96).lineTo(left + width, 96).lineWidth(3).strokeColor('#F7941D').stroke(); };
    const drawTableHeader = (y) => { const labels = ['المبلغ المتبقي', 'المبلغ المدفوع', 'إجمالي المبلغ', 'اسم المورد', 'كود المورد']; let x = left; labels.forEach((label, index) => { doc.rect(x, y, columns[index], 30).fillAndStroke('#FFD966', '#D9A900'); rtl(label, x + 4, y + 9, columns[index] - 8, { bold: true, size: 8 }); x += columns[index]; }); return y + 30; };
    drawHeader(); let y = drawTableHeader(112);
    const drawDataRow = (values, total = false) => { if (y > 750) { doc.addPage(); drawHeader(true); y = drawTableHeader(112); } let x = left; values.forEach((value, index) => { doc.rect(x, y, columns[index], 29).fillAndStroke(total ? '#FFF4CC' : '#FFFFFF', '#E5E7EB'); rtl(value, x + 4, y + 9, columns[index] - 8, { bold: total, size: 8 }); x += columns[index]; }); y += 29; };
    rows.forEach((supplier) => drawDataRow([money(supplier.due), money(supplier.paid), money(supplier.total), supplier.name || '—', supplier.code || '—']));
    drawDataRow([money(rows.reduce((sum, item) => sum + item.due, 0)), money(rows.reduce((sum, item) => sum + item.paid, 0)), money(rows.reduce((sum, item) => sum + item.total, 0)), 'الإجمالي', ''], true);
    rtl('المركز الفرنسي - تقرير مالي داخلي', left, 775, width, { bold: true, size: 8, color: '#667085' }); doc.end();
  } catch (error) { reject(error); } });
}

function inventoryAuditRows(data, from = '', to = '') {
  if (from || to) {
    const totals = new Map();
    (data.movements || []).filter((m) => (!from || m.date >= from) && (!to || m.date <= to)).forEach((m) => { const item = (data.inventory || []).find((i) => i.code === m.productCode); if (!item) return; const current = totals.get(item.code) || { item, quantity: 0 }; current.quantity += m.type === 'صادر' ? -Number(m.qty || 0) : Number(m.qty || 0); totals.set(item.code, current); });
    return [...totals.values()].map(({ item, quantity }) => ({ code: clean(item.code) || '—', statement: [clean(item.name), clean(item.details)].filter(Boolean).join(' — ') || '—', quantity: Math.max(0, quantity), value: Math.max(0, quantity) * (Number(item.buy) || 0) })).filter((r) => r.quantity > 0);
  }
  return (data.inventory || []).map((item) => {
    const quantity = Number(item.qty) || 0;
    const unitCost = Number(item.buy) || 0;
    return {
      code: clean(item.code) || '—',
      statement: [clean(item.name), clean(item.details)].filter(Boolean).join(' — ') || '—',
      quantity,
      value: quantity * unitCost,
    };
  }).sort((first, second) => first.statement.localeCompare(second.statement, 'ar'));
}

function inventoryMovementRows(data, from = '', to = '') {
  return (data.movements || []).filter((m) => (!from || m.date >= from) && (!to || m.date <= to)).map((m) => ({ code: m.code || '—', date: m.date || '—', time: m.time || '—', direction: m.type || '—', productCode: m.productCode || '—', productName: [m.productName, m.details].filter(Boolean).join(' — ') || '—', qty: Number(m.qty) || 0, amount: Number(m.total) || 0, customerCode: m.customerCode || '—', supplierCode: m.supplierCode || '—' }));
}

async function inventoryMovementsWorkbook(data, from, to) {
  const rows = inventoryMovementRows(data, from, to); const workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet('جرد حركة المخزن', { views: [{ rightToLeft: true, showGridLines: true }] });
  sheet.addRow(['جرد حركة المخزن', `من ${from} إلى ${to}`]); sheet.addRow([]); sheet.addRow(['كود الحركة','اليوم','الوقت','نوع الحركة','كود المنتج','بيان المنتج','الكمية','المبلغ','كود العميل','كود المورد']);
  rows.forEach((r) => { const row = sheet.addRow([r.code,r.date,r.time,r.direction,r.productCode,r.productName,r.qty,r.amount,r.customerCode,r.supplierCode]); const color = r.direction === 'صادر' ? 'FFFDE2E2' : 'FFE2F7E9'; row.eachCell((cell) => { cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:color}}; }); });
  sheet.getRow(1).font={name:'Tahoma',size:16,bold:true}; sheet.getRow(1).alignment={horizontal:'right'}; sheet.getRow(3).eachCell((c) => { c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFD966'}}; c.font={name:'Tahoma',size:16,bold:true,color:{argb:'FF111827'}}; c.alignment={horizontal:'center'}; });
  [18,14,12,14,16,42,12,16,16,16].forEach((w,i)=>sheet.getColumn(i+1).width=w); sheet.getColumn(7).numFmt='#,##0'; sheet.getColumn(8).numFmt='#,##0.00'; return Buffer.from(await workbook.xlsx.writeBuffer());
}

function inventoryMovementsPdf(data, from, to) { return new Promise((resolve, reject) => { try { const rows=inventoryMovementRows(data,from,to); const doc=new PDFDocument({size:'A4',margins:{top:38,right:32,bottom:40,left:32},bufferPages:true}); const chunks=[]; doc.on('data',c=>chunks.push(c)); doc.on('end',()=>resolve(Buffer.concat(chunks))); doc.on('error',reject); const fonts=path.join(process.env.WINDIR||'C:\\Windows','Fonts'); const regular=['tahoma.ttf','arial.ttf'].map(n=>path.join(fonts,n)).find(fs.existsSync); const bold=['tahomabd.ttf','arialbd.ttf'].map(n=>path.join(fonts,n)).find(fs.existsSync); if(!regular||!bold) throw new Error('تعذر العثور على خط عربي مناسب.'); doc.registerFont('Arabic',regular);doc.registerFont('ArabicBold',bold); const rtl=(t,x,y,w,o={})=>pdfRtlText(doc,t,x,y,w,o); const cols=[55,55,55,65,70,130,45,60,55,55]; const labels=['كود المورد','كود العميل','المبلغ','الكمية','بيان المنتج','كود المنتج','نوع الحركة','الوقت','اليوم','كود الحركة']; const left=25; let y=48; rtl('تقرير جرد حركة المخزن',left,y,545,{bold:true,size:19}); rtl(`الفترة: ${from} إلى ${to}`,left,y+28,545,{size:9,color:'#667085'}); y+=58; const head=()=>{let x=left; labels.forEach((l,i)=>{doc.rect(x,y,cols[i],28).fillAndStroke('#FFD966','#D9A900');rtl(l,x+2,y+8,cols[i]-4,{bold:true,size:6});x+=cols[i]});y+=28}; const row=(vals)=>{if(y>760){doc.addPage();y=48;head()}let x=left;vals.forEach((v,i)=>{doc.rect(x,y,cols[i],25).fillAndStroke('#fff','#e5e7eb');rtl(v,x+2,y+7,cols[i]-4,{size:6});x+=cols[i]});y+=25}; head(); rows.forEach(r=>row([r.supplierCode,r.customerCode,r.amount,r.qty,r.productName,r.productCode,r.direction,r.time,r.date,r.code])); doc.end(); } catch(e){reject(e);} }); }

async function inventoryAuditWorkbook(data, from = '', to = '') {
  const rows = inventoryAuditRows(data, from, to);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'المركز الفرنسي'; workbook.created = new Date();
  const sheet = workbook.addWorksheet('جرد المخزن', { views: [{ rightToLeft: true, showGridLines: true }] });
  sheet.addRow(['جرد المخزن', `تاريخ التقرير: ${new Date().toISOString().slice(0, 10)}`]);
  sheet.addRow([]);
  sheet.addRow(['كود المنتج', 'البيان', 'الكمية الحالية', 'قيمة المخزون']);
  rows.forEach((item) => sheet.addRow([item.code, item.statement, item.quantity, item.value]));
  sheet.addRow([]);
  sheet.addRow(['إجمالي الأصناف الموجودة', rows.length]);
  const totalRow = sheet.addRow(['إجمالي قيمة البضاعة', { formula: `SUM(D4:D${Math.max(4, rows.length + 3)})`, result: rows.reduce((sum, item) => sum + item.value, 0) }]);
  sheet.getColumn(1).width = 18; sheet.getColumn(2).width = 44; sheet.getColumn(3).width = 18; sheet.getColumn(4).width = 22;
  sheet.getColumn(3).numFmt = '#,##0'; sheet.getColumn(4).numFmt = '#,##0.00';
  totalRow.getCell(2).numFmt = '#,##0.00';
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function inventoryAuditPdf(data, from = '', to = '') {
  return new Promise((resolve, reject) => { try {
    const rows = inventoryAuditRows(data, from, to);
    const doc = new PDFDocument({ size: 'A4', margins: { top: 38, right: 40, bottom: 46, left: 40 }, bufferPages: true });
    const chunks = []; doc.on('data', (chunk) => chunks.push(chunk)); doc.on('end', () => resolve(Buffer.concat(chunks))); doc.on('error', reject);
    const fonts = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
    const regular = ['tahoma.ttf', 'arial.ttf'].map((name) => path.join(fonts, name)).find(fs.existsSync);
    const bold = ['tahomabd.ttf', 'arialbd.ttf'].map((name) => path.join(fonts, name)).find(fs.existsSync);
    if (!regular || !bold) throw new Error('تعذر العثور على خط عربي مناسب لإنشاء التقرير.');
    doc.registerFont('Arabic', regular); doc.registerFont('ArabicBold', bold);
    const rtl = (text, x, y, width, options = {}) => pdfRtlText(doc, text, x, y, width, options);
    const reverse = (value) => [...String(value)].reverse().join('');
    const money = (value) => Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    // كود المنتج هو أول عمود عربيًا؛ لذلك يوضع في أقصى اليمين عند الرسم.
    const left = 40; const width = doc.page.width - 80; const columns = [105, 80, 245, 85];
    const header = (continued = false) => { rtl(continued ? 'جرد المخزن - تابع' : 'تقرير جرد المخزن', left, 44, width, { bold: true, size: 20, color: '#17243C' }); rtl(`تاريخ التقرير: ${reverse(new Date().toISOString().slice(0, 10))}`, left, 74, width, { size: 9, color: '#667085' }); doc.moveTo(left, 96).lineTo(left + width, 96).lineWidth(3).strokeColor('#F7941D').stroke(); };
    const tableHeader = (y) => { const labels = ['قيمة المخزون', 'الكمية الحالية', 'البيان', 'كود المنتج']; let x = left; labels.forEach((label, index) => { doc.rect(x, y, columns[index], 30).fillAndStroke('#FFD966', '#D9A900'); rtl(label, x + 4, y + 9, columns[index] - 8, { bold: true, size: 8 }); x += columns[index]; }); return y + 30; };
    header(); let y = tableHeader(112);
    const drawRow = (values, total = false) => { const height = Math.max(31, doc.heightOfString(String(values[2]), { width: columns[2] - 10, align: 'right' }) + 16); if (y + height > 746) { doc.addPage(); header(true); y = tableHeader(112); } let x = left; values.forEach((value, index) => { doc.rect(x, y, columns[index], height).fillAndStroke(total ? '#FFF4CC' : '#FFFFFF', '#E5E7EB'); rtl(value, x + 5, y + 8, columns[index] - 10, { bold: total, size: 8 }); x += columns[index]; }); y += height; };
    if (rows.length) rows.forEach((item) => drawRow([money(item.value), item.quantity, item.statement, item.code]));
    else drawRow(['0.00', '0', 'لا توجد أصناف في المخزن', '—']);
    y += 12;
    const totalValue = rows.reduce((sum, item) => sum + item.value, 0);
    doc.roundedRect(left, y, width, 50, 8).fillAndStroke('#F8FAFC', '#DDE3EC');
    rtl(`إجمالي الأصناف المختلفة: ${reverse(rows.length)}`, left + width / 2, y + 10, width / 2 - 12, { bold: true, size: 10 });
    rtl(`إجمالي قيمة البضاعة: ${reverse(money(totalValue))} ج`, left + 12, y + 28, width - 24, { bold: true, size: 11, color: '#17243C' });
    rtl('المركز الفرنسي - تقرير جرد داخلي', left, Math.max(y + 68, 780), width, { bold: true, size: 8, color: '#667085' });
    doc.end();
  } catch (error) { reject(error); } });
}

function externalDebtRows(data) {
  const customers = (data.customers || []).filter((customer) => Number(customer.dueFromCustomer) > 0).map((customer) => ({
    code: customer.code || '—', statement: `عميل — ${customer.name || 'غير معروف'}`, amount: Number(customer.dueFromCustomer) || 0,
  }));
  const employees = (data.employees || []).filter((employee) => Number(employee.debtOnEmployee) > 0).map((employee) => ({
    code: employee.code || '—', statement: `موظف — ${employee.name || 'غير معروف'}`, amount: Number(employee.debtOnEmployee) || 0,
  }));
  const knownCodes = new Set([...customers, ...employees].map((item) => item.code));
  const otherPeople = (data.accounts || []).filter((account) => account.direction === 'مديونية على الغير' && !knownCodes.has(account.customerCode) && !knownCodes.has(account.employeeCode)).map((account) => ({
    code: account.code || '—', statement: clean(account.description) || 'جهة خارجية', amount: Number(account.due) || 0,
  })).filter((item) => item.amount > 0);
  return [...customers, ...employees, ...otherPeople].sort((first, second) => second.amount - first.amount);
}

async function externalDebtsWorkbook(data) {
  const rows = externalDebtRows(data);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'المركز الفرنسي'; workbook.created = new Date();
  const sheet = workbook.addWorksheet('الديون الخارجية للمركز', { views: [{ rightToLeft: true, showGridLines: true }] });
  sheet.addRow(['تقرير الديون الخارجية للمركز', `تاريخ التقرير: ${new Date().toISOString().slice(0, 10)}`]);
  sheet.addRow([]); sheet.addRow(['الكود', 'البيان', 'المبلغ المطلوب']);
  rows.forEach((item) => sheet.addRow([item.code, item.statement, item.amount]));
  sheet.addRow([]); sheet.addRow(['إجمالي الأشخاص', rows.length]);
  const totalRow = sheet.addRow(['إجمالي المبلغ المطلوب', { formula: `SUM(C4:C${Math.max(4, rows.length + 3)})`, result: rows.reduce((sum, item) => sum + item.amount, 0) }]);
  sheet.getColumn(1).width = 18; sheet.getColumn(2).width = 42; sheet.getColumn(3).width = 22;
  sheet.getColumn(3).numFmt = '#,##0.00'; totalRow.getCell(2).numFmt = '#,##0.00';
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function externalDebtsPdf(data) {
  return new Promise((resolve, reject) => { try {
    const rows = externalDebtRows(data); const doc = new PDFDocument({ size: 'A4', margins: { top: 38, right: 40, bottom: 46, left: 40 }, bufferPages: true });
    const chunks = []; doc.on('data', (chunk) => chunks.push(chunk)); doc.on('end', () => resolve(Buffer.concat(chunks))); doc.on('error', reject);
    const fonts = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
    const regular = ['tahoma.ttf', 'arial.ttf'].map((name) => path.join(fonts, name)).find(fs.existsSync);
    const bold = ['tahomabd.ttf', 'arialbd.ttf'].map((name) => path.join(fonts, name)).find(fs.existsSync);
    if (!regular || !bold) throw new Error('تعذر العثور على خط عربي مناسب لإنشاء التقرير.');
    doc.registerFont('Arabic', regular); doc.registerFont('ArabicBold', bold);
    const rtl = (text, x, y, width, options = {}) => pdfRtlText(doc, text, x, y, width, options);
    const reverse = (value) => [...String(value)].reverse().join(''); const money = (value) => Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const pdfStatement = (value) => String(value).replace(/[A-Za-z0-9][A-Za-z0-9 .-]*/g, (latinText) => reverse(latinText));
    const left = 40; const width = doc.page.width - 80; const columns = [145, 280, 90];
    const header = (continued = false) => { rtl(continued ? 'الديون الخارجية للمركز - تابع' : 'تقرير الديون الخارجية للمركز', left, 44, width, { bold: true, size: 19, color: '#17243C' }); rtl(`تاريخ التقرير: ${reverse(new Date().toISOString().slice(0, 10))}`, left, 74, width, { size: 9, color: '#667085' }); doc.moveTo(left, 96).lineTo(left + width, 96).lineWidth(3).strokeColor('#F7941D').stroke(); };
    const tableHeader = (y) => { const labels = ['المبلغ المطلوب', 'البيان', 'الكود']; let x = left; labels.forEach((label, index) => { doc.rect(x, y, columns[index], 30).fillAndStroke('#FFD966', '#D9A900'); rtl(label, x + 4, y + 9, columns[index] - 8, { bold: true, size: 8 }); x += columns[index]; }); return y + 30; };
    header(); let y = tableHeader(112);
    const drawRow = (values) => { const height = Math.max(31, doc.heightOfString(String(values[1]), { width: columns[1] - 10, align: 'right' }) + 16); if (y + height > 746) { doc.addPage(); header(true); y = tableHeader(112); } let x = left; values.forEach((value, index) => { doc.rect(x, y, columns[index], height).fillAndStroke('#FFFFFF', '#E5E7EB'); rtl(value, x + 5, y + 8, columns[index] - 10, { size: 8 }); x += columns[index]; }); y += height; };
    if (rows.length) rows.forEach((item) => drawRow([money(item.amount), pdfStatement(item.statement), item.code]));
    else drawRow(['0.00', 'لا توجد ديون خارجية مسجلة', '—']);
    y += 12; const total = rows.reduce((sum, item) => sum + item.amount, 0);
    doc.roundedRect(left, y, width, 50, 8).fillAndStroke('#F8FAFC', '#DDE3EC');
    rtl(`إجمالي الأشخاص: ${reverse(rows.length)}`, left + width / 2, y + 10, width / 2 - 12, { bold: true, size: 10 });
    rtl(`إجمالي المبلغ المطلوب: ${reverse(money(total))} ج`, left + 12, y + 28, width - 24, { bold: true, size: 11, color: '#17243C' });
    rtl('المركز الفرنسي - تقرير مالي داخلي', left, Math.max(y + 68, 780), width, { bold: true, size: 8, color: '#667085' }); doc.end();
  } catch (error) { reject(error); } });
}

function json(response, status, body) {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  response.end(JSON.stringify(body));
}

async function bodyOf(request) {
  let body = '';
  for await (const chunk of request) {
    body += chunk;
    if (body.length > 1_000_000) throw new Error('الطلب أكبر من المسموح');
  }
  return body ? JSON.parse(body) : {};
}

function visitInvoicePdf(visit, customer) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'A4', margins: { top: 34, right: 40, bottom: 34, left: 40 }, bufferPages: true });
      const chunks = [];
      doc.on('data', (chunk) => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      const windowsFonts = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
      const regularFont = ['tahoma.ttf', 'arial.ttf'].map((name) => path.join(windowsFonts, name)).find(fs.existsSync);
      const boldFont = ['tahomabd.ttf', 'arialbd.ttf'].map((name) => path.join(windowsFonts, name)).find(fs.existsSync);
      if (!regularFont || !boldFont) throw new Error('تعذر العثور على خط عربي مناسب لإنشاء الفاتورة.');
      doc.registerFont('Arabic', regularFont);
      doc.registerFont('ArabicBold', boldFont);

      const pageWidth = doc.page.width;
      const right = pageWidth - 40;
      const left = 40;
      const contentWidth = right - left;
      // PDFKit يعكس الأرقام داخل السطر العربي؛ نعكس النص الرقمي مسبقًا ليظهر صحيحًا في الفاتورة.
      const rtlNumber = (value) => [...String(value)].reverse().join('');
      const money = (value) => `${rtlNumber(Number(value || 0).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 }))}\u00A0جنيه`;
      const valueText = (value) => clean(value) || '—';
      const rtl = (text, x, y, width, options = {}) => pdfRtlText(doc, text, x, y, width, { size: 10, ...options });
      const logoPath = path.join(PUBLIC, 'invoice-logo.b64');
      if (fs.existsSync(logoPath)) {
        const logo = Buffer.from(fs.readFileSync(logoPath, 'utf8').trim(), 'base64');
        doc.image(logo, right - 82, 34, { fit: [82, 82], align: 'center', valign: 'center' });
      }
      rtl('المركز الفرنسي', left, 47, 265, { bold: true, size: 24 });
      doc.font('Arabic').fontSize(9).fillColor('#667085').text('French Center For Car Maintenance', left, 78, { width: 265, align: 'left' });
      doc.moveTo(left, 124).lineTo(right, 124).lineWidth(3).strokeColor('#F7941D').stroke();
      rtl('فاتورة زيارة صيانة', left, 134, contentWidth / 2, { bold: true, size: 12, color: '#2676EE' });
      rtl(`التاريخ: ${rtlNumber(valueText(visit.date))}`, right - contentWidth / 2, 134, contentWidth / 2, { bold: true, size: 10 });

      const gap = 10;
      const boxWidth = (contentWidth - gap) / 2;
      const boxHeight = 52;
      const drawBox = (x, y, label, value) => {
        doc.roundedRect(x, y, boxWidth, boxHeight, 7).fillAndStroke('#F8FAFC', '#DDE3EC');
        rtl(label, x + 9, y + 7, boxWidth - 18, { size: 8, color: '#667085' });
        rtl(valueText(value), x + 9, y + 25, boxWidth - 18, { bold: true, size: 10 });
      };
      let y = 164;
      const details = [
        ['اسم العميل', customer.name], ['كود العميل', visit.customerCode],
        ['كود الزيارة', visit.code], ['كود الحركة', visit.stockMovementCode],
        ['نوع العربية', customer.carType], ['لوحة العربية', visit.plate || customer.plate],
        ['قراءة العداد', `${rtlNumber(Number(visit.mileage || 0).toLocaleString('en-US'))} كم`], ['نوع الصيانة', visit.serviceType],
      ];
      for (let index = 0; index < details.length; index += 2) {
        drawBox(right - boxWidth, y, details[index][0], details[index][1]);
        drawBox(left, y, details[index + 1][0], details[index + 1][1]);
        y += boxHeight + gap;
      }

      if (clean(visit.notes)) {
        doc.roundedRect(left, y, contentWidth, 56, 7).fillAndStroke('#F8FAFC', '#DDE3EC');
        rtl('الملاحظات', left + 10, y + 7, contentWidth - 20, { size: 8, color: '#667085' });
        rtl(visit.notes, left + 10, y + 24, contentWidth - 20, { bold: true, size: 9, lineGap: 2 });
        y += 66;
      }

      doc.roundedRect(left, y, contentWidth, 78, 7).strokeColor('#DDE3EC').stroke();
      doc.rect(left, y, contentWidth, 26).fill('#EEF5FF');
      rtl('قطع الغيار المستخدمة', left + 10, y + 7, contentWidth - 20, { bold: true, size: 9 });
      rtl(visit.partsCodes || 'لا توجد قطع غيار', left + 12, y + 37, contentWidth - 24, { bold: true, size: 10, lineGap: 3 });
      y += 92;

      const totalsWidth = 300;
      const totalsX = left;
      const totalRow = (label, value, isTotal = false) => {
        doc.rect(totalsX, y, totalsWidth, 34).fillAndStroke(isTotal ? '#EDF8F3' : '#FFFFFF', '#D5DCE7');
        rtl(label, totalsX + totalsWidth / 2, y + 10, totalsWidth / 2 - 12, { bold: isTotal, size: isTotal ? 11 : 9 });
        rtl(value, totalsX + 10, y + 10, totalsWidth / 2 - 18, { bold: true, size: isTotal ? 12 : 10 });
        y += 34;
      };
      totalRow('تكلفة قطع الغيار', money(visit.partsTotal));
      totalRow('تكلفة المصنعية', money(visit.labor));
      totalRow('الإجمالي', money(visit.total), true);

      const footerY = Math.max(y + 24, 742);
      doc.moveTo(left, footerY).lineTo(right, footerY).lineWidth(2).strokeColor('#F7941D').stroke();
      rtl(`رقم التواصل: ${rtlNumber('01212891063')}`, left, footerY + 10, contentWidth, { bold: true, size: 9 });
      rtl('العنوان: المركز الكائن كوبري قباء أسفل جسر السويس، القاهرة، مصر', left, footerY + 27, contentWidth, { size: 8 });
      rtl('شكرًا لثقتكم في المركز الفرنسي', left, footerY + 44, contentWidth, { bold: true, size: 8, color: '#667085' });
      doc.end();
    } catch (error) {
      reject(error);
    }
  });
}

async function api(request, response, pathname, searchParams) {
  if (request.method === 'GET' && pathname === '/api/data') {
    return json(response, 200, { ...readData(), databaseFile: path.basename(DATA_FILE) });
  }

  if (request.method === 'GET' && pathname === '/api/export') {
    const buffer = await rtlWorkbookBuffer(DATA_FILE);
    response.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="Book1.xlsx"',
      'Content-Length': buffer.length,
    });
    return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/visit-invoice') {
    const code = clean(searchParams?.get('code'));
    const data = readData();
    const visit = data.visits.find((item) => item.code === code);
    if (!visit) return json(response, 404, { error: 'الزيارة غير موجودة.' });
    const customer = data.customers.find((item) => item.code === visit.customerCode) || {};
    const buffer = await visitInvoicePdf(visit, customer);
    const safeCode = code.replace(/[^a-z0-9_-]/gi, '') || 'visit';
    response.writeHead(200, {
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="invoice-${safeCode}.pdf"`,
      'Content-Length': buffer.length,
      'Cache-Control': 'no-store',
    });
    return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/daily-close') {
    const date = clean(searchParams?.get('date')) || new Date().toISOString().slice(0, 10);
    const buffer = await dailyCloseBuffer(readData(), date);
    response.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="daily-close-${date}.xlsx"`,
      'Content-Length': buffer.length,
    });
    return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/financial-report') {
    const from = clean(searchParams?.get('from')) || new Date().toISOString().slice(0, 10);
    const to = clean(searchParams?.get('to')) || from;
    const title = clean(searchParams?.get('title')) || 'الملخص المالي';
    const format = clean(searchParams?.get('format')) || 'xlsx';
    const data = readData();
    if (format === 'pdf') {
      const html = reportHtml(data, from, to, title);
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
      return response.end(html);
    }
    const buffer = await dailyCloseBuffer(data, from, to, title);
    response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="financial-report-${from}-${to}.xlsx"`, 'Content-Length': buffer.length });
    return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/supplier-debts') {
    const format = clean(searchParams?.get('format')) || 'xlsx';
    const data = readData();
    if (format === 'pdf') {
      const buffer = await supplierDebtsPdf(data);
      response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': 'attachment; filename="supplier-debts.pdf"', 'Content-Length': buffer.length, 'Cache-Control': 'no-store' });
      return response.end(buffer);
    }
    const buffer = await supplierDebtsWorkbook(data);
    response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="supplier-debts.xlsx"', 'Content-Length': buffer.length, 'Cache-Control': 'no-store' });
    return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/inventory-audit') {
    const format = clean(searchParams?.get('format')) || 'xlsx';
    const from = clean(searchParams?.get('from')) || ''; const to = clean(searchParams?.get('to')) || '';
    const data = readData();
    if (format === 'pdf') {
      const buffer = await inventoryAuditPdf(data, from, to);
      response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': 'attachment; filename="inventory-audit.pdf"', 'Content-Length': buffer.length, 'Cache-Control': 'no-store' });
      return response.end(buffer);
    }
    const buffer = await inventoryAuditWorkbook(data, from, to);
    response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="inventory-audit.xlsx"', 'Content-Length': buffer.length, 'Cache-Control': 'no-store' });
    return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/inventory-movements-report') {
    const format = clean(searchParams?.get('format')) || 'xlsx'; const from = clean(searchParams?.get('from')) || '0000-01-01'; const to = clean(searchParams?.get('to')) || '9999-12-31'; const data = readData();
    const buffer = format === 'pdf' ? await inventoryMovementsPdf(data, from, to) : await inventoryMovementsWorkbook(data, from, to);
    response.writeHead(200, { 'Content-Type': format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="inventory-movements.${format === 'pdf' ? 'pdf' : 'xlsx'}"`, 'Content-Length': buffer.length, 'Cache-Control': 'no-store' }); return response.end(buffer);
  }

  if (request.method === 'GET' && pathname === '/api/external-debts') {
    const format = clean(searchParams?.get('format')) || 'xlsx';
    const data = readData();
    if (format === 'pdf') {
      const buffer = await externalDebtsPdf(data);
      response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': 'attachment; filename="external-debts.pdf"', 'Content-Length': buffer.length, 'Cache-Control': 'no-store' });
      return response.end(buffer);
    }
    const buffer = await externalDebtsWorkbook(data);
    response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="external-debts.xlsx"', 'Content-Length': buffer.length, 'Cache-Control': 'no-store' });
    return response.end(buffer);
  }

  if (request.method === 'POST' && pathname === '/api/customers') {
    const input = await bodyOf(request);
    const data = readData();
    if (!clean(input.name) || !clean(input.phone) || !clean(input.plate) || !clean(input.carType))
      return json(response, 400, { error: 'أكمل بيانات العميل الأساسية.' });
    const phoneOwner = data.customers.find((customer) => normalizedPhone(customer.phone) === normalizedPhone(input.phone));
    if (phoneOwner)
      return json(response, 409, { error: `رقم التليفون مسجل بالفعل للعميل ${phoneOwner.name} — الكود ${phoneOwner.code}.` });
    if (data.customers.some((customer) => normalizedPlate(customer.plate) === normalizedPlate(input.plate)))
      return json(response, 409, { error: 'لوحة العربية مسجلة بالفعل لعميل موجود.' });
    const customer = {
      code: nextCode(data.customers.map((item) => item.code), 'C'), name: clean(input.name),
      phone: clean(input.phone), plate: clean(input.plate), carType: clean(input.carType), registeredDate: new Date().toISOString().slice(0, 10), dueFromCustomer: 0, dueFromCenter: 0,
    };
    data.customers.push(customer);
    saveData(data.customers, data.visits);
    return json(response, 201, customer);
  }

  if (request.method === 'POST' && pathname === '/api/visits') {
    const input = await bodyOf(request);
    const data = readData();
    const customer = data.customers.find((item) => item.code === clean(input.customerCode));
    if (!customer) return json(response, 404, { error: 'العميل غير موجود.' });
    if (!clean(input.date) || !clean(input.serviceType)) return json(response, 400, { error: 'اكتب تاريخ ونوع الصيانة.' });
    const requestedParts = Array.isArray(input.parts) ? input.parts : [];
    const groupedParts = new Map();
    for (const part of requestedParts) {
      const code = clean(part.code);
      const qty = Math.max(0, Math.floor(Number(part.qty) || 0));
      if (code && qty) groupedParts.set(code, (groupedParts.get(code) || 0) + qty);
    }
    let partsTotal = 0;
    let partsCost = 0;
    const partsDetails = [];
    const visitCode = nextCode(data.visits.map((item) => item.code), 'V');
    const movementCode = groupedParts.size ? nextCode(data.movements.map((item) => item.code), 'M') : '';
    const outgoingMovements = [];
    for (const [code, qty] of groupedParts) {
      const stockItem = data.inventory.find((item) => item.code === code);
      if (!stockItem) return json(response, 400, { error: `قطعة الغيار ${code} غير موجودة في المخزن.` });
      if (qty > stockItem.qty) return json(response, 400, { error: `الكمية المطلوبة من ${stockItem.name} أكبر من المتاح (${stockItem.qty}).` });
      stockItem.qty -= qty;
      partsTotal += stockItem.sell * qty;
      partsCost += stockItem.buy * qty;
      partsDetails.push(`(${stockItem.code} x ${qty})`);
      outgoingMovements.push({
        code: movementCode, date: clean(input.date), type: 'صادر', productCode: stockItem.code,
        productName: stockItem.name, details: stockItem.details, country: stockItem.country, qty,
        unitPrice: stockItem.sell, total: stockItem.sell * qty, visitCode, customerCode: customer.code,
        supplier: stockItem.supplier, notes: clean(input.serviceType),
      });
    }
    const labor = Number(input.labor) || 0;
    if (labor <= 0) return json(response, 400, { error: 'يجب كتابة قيمة المصنعية قبل حفظ الزيارة.' });
    const visitTotal = partsTotal + labor;
    const payment = paymentInfo(input, visitTotal);
    const { paymentMethod, paymentDetails, paid } = payment;
    if (Math.abs(payment.allocated - visitTotal) > 0.001) return json(response, 400, { error: `وزّع إجمالي الفاتورة كاملًا على طرق الدفع (${visitTotal}).` });
    if (payment.allocated > visitTotal) return json(response, 400, { error: 'مجموع طرق الدفع أكبر من إجمالي الزيارة.' });
    if (paid > visitTotal) return json(response, 400, { error: 'المبلغ المدفوع أكبر من إجمالي الزيارة.' });
    const visit = {
      code: visitCode, customerCode: customer.code,
      plate: customer.plate, date: clean(input.date), mileage: Number(input.mileage) || 0,
      serviceType: clean(input.serviceType), technician: clean(input.technician),
      partsCodes: partsDetails.join(', '), partsTotal, partsCost, labor, total: visitTotal,
      notes: clean(input.notes), stockMovementCode: movementCode, paymentMethod, paymentDetails, paid, due: visitTotal - paid,
    };
    data.visits.push(visit);
    const participatingEmployee = data.employees.find((employee) => employee.name.toLocaleLowerCase('ar-EG') === visit.technician.toLocaleLowerCase('ar-EG'));
    if (participatingEmployee) participatingEmployee.contributions += 1;
    data.movements.push(...outgoingMovements);
    const account = { code: nextCode(data.accounts.map((item) => item.code), 'A'), date: visit.date, executionDate: visit.date,
      direction: 'وارد', type: 'إيراد زيارة صيانة', description: `زيارة ${visit.code}`, customerCode: customer.code, customerName: customer.name,
      visitCode: visit.code, supplierCode: '', supplierName: '', productCode: '', productName: '', qty: 0,
      total: visit.total, paid: visit.paid, due: visit.due, paymentMethod, paymentDetails, notes: `${visit.code} · ${customer.name}` };
    appendAccount(data.accounts, account);
    saveData(data.customers, data.visits, data.inventory, null, data.movements, data.accounts, data.employees);
    return json(response, 201, visit);
  }

  if (request.method === 'POST' && pathname === '/api/employees') {
    const input = await bodyOf(request);
    const data = readData();
    const name = clean(input.name); const phone = clean(input.phone);
    if (!name || !phone || !clean(input.hireDate) || !clean(input.specialty)) return json(response, 400, { error: 'أكمل اسم الموظف ورقم التليفون وتاريخ التوظيف والتخصص.' });
    if (data.employees.some((employee) => normalizedPhone(employee.phone) === normalizedPhone(phone))) return json(response, 409, { error: 'رقم تليفون الموظف مسجل بالفعل.' });
    const employee = { code: nextCode(data.employees.map((item) => item.code), 'E'), name, phone, hireDate: clean(input.hireDate), specialty: clean(input.specialty),
      contributions: 0, weeklySalary: Math.max(0, Number(input.weeklySalary) || 0), debtOnEmployee: 0, dueToEmployee: 0,
      status: 'يعمل', stopReason: '', stopDate: '', notes: clean(input.notes) };
    data.employees.push(employee);
    saveData(data.customers, data.visits, null, null, null, null, data.employees);
    return json(response, 201, employee);
  }

  if (request.method === 'POST' && pathname === '/api/employees/status') {
    const input = await bodyOf(request); const data = readData();
    const employee = data.employees.find((item) => item.code === clean(input.employeeCode));
    if (!employee) return json(response, 404, { error: 'الموظف غير موجود.' });
    const status = clean(input.status); const reason = clean(input.reason);
    if (!['موقوف مؤقتًا', 'منتهي الخدمة', 'يعمل'].includes(status)) return json(response, 400, { error: 'اختر حالة صحيحة للموظف.' });
    if (status !== 'يعمل' && !reason) return json(response, 400, { error: 'اكتب سبب الإيقاف أو إنهاء الخدمة.' });
    employee.status = status; employee.stopReason = status === 'يعمل' ? '' : reason; employee.stopDate = status === 'يعمل' ? '' : (clean(input.stopDate) || new Date().toISOString().slice(0, 10));
    saveData(data.customers, data.visits, null, null, null, null, data.employees);
    return json(response, 200, employee);
  }

  if (request.method === 'POST' && pathname === '/api/suppliers') {
    const input = await bodyOf(request);
    const data = readData();
    const name = clean(input.name);
    const phone = clean(input.phone);
    if (!name || !phone) return json(response, 400, { error: 'اكتب اسم المورد ورقم التليفون.' });
    if (data.suppliers.some((supplier) => supplier.name.toLocaleLowerCase('ar-EG') === name.toLocaleLowerCase('ar-EG')))
      return json(response, 409, { error: 'اسم المورد مسجل بالفعل.' });
    if (data.suppliers.some((supplier) => normalizedPhone(supplier.phone) === normalizedPhone(phone)))
      return json(response, 409, { error: 'رقم تليفون المورد مسجل بالفعل.' });
    const supplier = { code: nextCode(data.suppliers.map((item) => item.code), 'S'), name, phone, contractDate: clean(input.contractDate) || new Date().toISOString().slice(0, 10), due: 0, notes: clean(input.notes), paid: 0, paymentDate: '' };
    data.suppliers.push(supplier);
    saveData(data.customers, data.visits, null, data.suppliers);
    return json(response, 201, supplier);
  }

  if (request.method === 'POST' && pathname === '/api/supplier-transactions') {
    const input = await bodyOf(request);
    const data = readData();
    const type = clean(input.type);
    const supplierKey = clean(input.supplierCode);
    const supplier = data.suppliers.find((item) => item.code === supplierKey || item.name === supplierKey) || null;
    const today = new Date().toISOString().slice(0, 10);
    if (type === 'سداد مستحقات') {
      const amount = Math.max(0, Number(input.amount) || 0);
      const payment = paymentInfo(input, amount);
      if (payment.payments.some((entry) => entry.method === 'آجل')) return json(response, 400, { error: 'لا يمكن تسجيل السداد بطريقة آجل.' });
      if (!supplier) return json(response, 400, { error: 'اختر المورد قبل تسجيل السداد.' });
      if (!amount) return json(response, 400, { error: 'اكتب مبلغ السداد.' });
      if (amount > supplier.due) return json(response, 400, { error: `مبلغ السداد أكبر من المستحق على المورد (${supplier.due}).` });
      if (payment.paid !== amount) return json(response, 400, { error: 'مجموع مبالغ طرق الدفع يجب أن يساوي مبلغ السداد.' });
      supplier.due -= amount;
      supplier.paid += amount;
      supplier.paymentDate = today;
      const account = { code: nextCode(data.accounts.map((item) => item.code), 'A'), date: today, executionDate: today, direction: 'صادر', type,
        description: `سداد مستحقات ${supplier.name}`, customerCode: '', customerName: '', visitCode: '',
        supplierCode: supplier.code, supplierName: supplier.name, productCode: '', productName: '', qty: 0,
        total: amount, paid: amount, due: supplier.due, paymentMethod: payment.paymentMethod, paymentDetails: payment.paymentDetails, notes: clean(input.notes) };
      appendAccount(data.accounts, account);
      saveData(data.customers, data.visits, null, data.suppliers, null, data.accounts);
      return json(response, 201, { account, supplier });
    }
    if (type === 'توريد بضاعة') {
      const productCode = clean(input.productCode);
      const item = data.inventory.find((entry) => entry.code.toLocaleLowerCase() === productCode.toLocaleLowerCase());
      const qty = Math.max(0, Math.floor(Number(input.qty) || 0));
      const buy = Math.max(0, Number(input.buy) || 0);
      const sell = Math.max(0, Number(input.sell) || 0);
      const payment = paymentInfo(input, 0);
      if (payment.payments.some((entry) => entry.method === 'آجل')) return json(response, 400, { error: 'التوريد لا يقبل طريقة دفع آجل.' });
      const paid = payment.paid;
      if (!item) return json(response, 400, { error: 'اختر منتجًا موجودًا من المخزن.' });
      if (!qty) return json(response, 400, { error: 'اكتب كمية صحيحة.' });
      const total = qty * buy;
      if (payment.allocated > total) return json(response, 400, { error: 'مجموع طرق الدفع أكبر من إجمالي التوريد.' });
      if (paid > total) return json(response, 400, { error: 'المبلغ المدفوع أكبر من إجمالي التوريد.' });
      const due = total - paid;
      item.qty += qty; item.buy = buy; item.sell = sell; item.margin = sell - buy;
      if (supplier) item.supplier = supplier.name;
      item.paid = Number(item.paid || 0) + paid; item.due = Number(item.due || 0) + due;
      if (supplier) { supplier.paid += paid; supplier.due += due; supplier.paymentDate = paid ? today : supplier.paymentDate; }
      const movement = { code: nextCode(data.movements.map((entry) => entry.code), 'M'), date: today, time: currentTime(), type: 'وارد',
        productCode: item.code, productName: item.name, details: item.details, country: item.country, qty,
        unitPrice: buy, total, visitCode: '', customerCode: '', supplier: supplier?.name || '', notes: 'توريد بضاعة من قسم الموردين' };
      const account = { code: nextCode(data.accounts.map((entry) => entry.code), 'A'), date: today, executionDate: today, direction: 'صادر', type,
        description: `توريد ${item.name}`, customerCode: '', customerName: '', visitCode: '',
        supplierCode: supplier?.code || '', supplierName: supplier?.name || '', productCode: item.code,
        productName: item.name, qty, total, paid, due, paymentMethod: payment.paymentMethod, paymentDetails: payment.paymentDetails, notes: clean(input.notes) };
      data.movements.push(movement); appendAccount(data.accounts, account);
      saveData(data.customers, data.visits, data.inventory, data.suppliers, data.movements, data.accounts);
      return json(response, 201, { account, movement, item, supplier });
    }
    return json(response, 400, { error: 'اختر نوع الحركة.' });
  }

  if (request.method === 'POST' && pathname === '/api/accounts') {
    const input = await bodyOf(request);
    const data = readData();
    const operation = clean(input.operation);
    const type = clean(input.type);
    const amount = Math.max(0, Number(input.amount) || 0);
    const payment = paymentInfo(input, amount);
    const executionDate = clean(input.executionDate) || new Date().toISOString().slice(0, 10);
    if (!['سحب', 'إيداع', 'دين', 'دفع مستحقات'].includes(operation)) return json(response, 400, { error: 'اختر سحب أو إيداع أو دين أو دفع مستحقات.' });
    if (!type) return json(response, 400, { error: 'اكتب نوع الحركة.' });
    if (!amount) return json(response, 400, { error: 'اكتب مبلغ الحركة.' });
    if (payment.allocated > amount) return json(response, 400, { error: 'مجموع طرق الدفع أكبر من مبلغ الحركة.' });
    if (payment.paid > amount) return json(response, 400, { error: 'مجموع مبالغ طرق الدفع أكبر من مبلغ الحركة.' });
    const customerKey = clean(input.customerCode);
    const supplierKey = clean(input.supplierCode);
    const employeeKey = clean(input.employeeCode);
    const customer = data.customers.find((item) => item.code === customerKey || item.name === customerKey) || null;
    const supplier = data.suppliers.find((item) => item.code === supplierKey || item.name === supplierKey) || null;
    const employee = data.employees.find((item) => item.code === employeeKey || item.name === employeeKey) || null;
    const employeeMovement = ['مرتبات', 'سلفة موظف', 'سداد سلفة موظف', 'مستحق لموظف', 'دفع مستحق موظف'].includes(type);
    if (['سداد مديونية', 'سداد مديونية عميل', 'سداد مستحقات', 'سداد مستحقات للمورد', 'سداد سلفة موظف', 'دفع مستحق موظف', 'توريد بضاعة'].includes(type) && payment.payments.some((entry) => entry.method === 'آجل')) return json(response, 400, { error: 'هذه الحركة لا تقبل طريقة دفع آجل.' });
    if (type === 'سداد مديونية' && !customer) return json(response, 400, { error: 'اختر العميل قبل سداد مديونيته.' });
    if (type === 'سداد مديونية' && amount > customer.dueFromCustomer) return json(response, 400, { error: `المبلغ أكبر من مديونية العميل (${customer.dueFromCustomer}).` });
    if (type === 'سداد مديونية عميل' && !customer) return json(response, 400, { error: 'اختر العميل قبل سداد مديونيته.' });
    if (type === 'سداد مديونية عميل' && amount > customer.dueFromCustomer) return json(response, 400, { error: `المبلغ أكبر من مديونية العميل (${customer.dueFromCustomer}).` });
    if (type === 'سداد مستحقات للمورد' && !supplier) return json(response, 400, { error: 'اختر المورد قبل سداد مستحقاته.' });
    if (type === 'سداد مستحقات للمورد' && amount > supplier.due) return json(response, 400, { error: `المبلغ أكبر من مستحقات المورد (${supplier.due}).` });
    if (employeeMovement && !employee) return json(response, 400, { error: 'اختر الموظف المرتبط بالحركة.' });
    if (type === 'مرتبات' && employee.status && employee.status !== 'يعمل') return json(response, 400, { error: 'لا يمكن صرف مرتب لموظف موقوف عن العمل.' });
    if (type === 'سداد سلفة موظف' && amount > employee.debtOnEmployee) return json(response, 400, { error: `المبلغ أكبر من السلفة المتبقية على الموظف (${employee.debtOnEmployee}).` });
    if (type === 'دفع مستحق موظف' && amount > employee.dueToEmployee) return json(response, 400, { error: `المبلغ أكبر من المستحق للموظف (${employee.dueToEmployee}).` });
    const debtSide = clean(input.debtSide);
    if (operation === 'دين' && !customer) return json(response, 400, { error: 'اختر العميل المرتبط بالدين.' });
    if (operation === 'دين' && !['على العميل', 'على المركز'].includes(debtSide)) return json(response, 400, { error: 'حدد هل الدين على العميل أم على المركز.' });
    let direction = operation === 'سحب' ? 'صادر' : operation === 'إيداع' ? 'وارد' : operation === 'دفع مستحقات' ? (type === 'سداد مستحقات للمورد' ? 'صادر' : 'وارد') : debtSide === 'على العميل' ? 'مديونية على الغير' : 'مديونية على المركز';
    if (type === 'سلفة موظف' || type === 'دفع مستحق موظف') direction = 'صادر';
    if (type === 'مرتبات') direction = 'صادر';
    if (type === 'سداد سلفة موظف') direction = 'وارد';
    if (type === 'مستحق لموظف') direction = 'مديونية على المركز';
    if (type === 'سلفة موظف') employee.debtOnEmployee += amount;
    if (type === 'سداد سلفة موظف') employee.debtOnEmployee = Math.max(0, employee.debtOnEmployee - amount);
    if (type === 'مستحق لموظف') employee.dueToEmployee += amount;
    if (type === 'دفع مستحق موظف') employee.dueToEmployee = Math.max(0, employee.dueToEmployee - amount);
    if (type === 'سداد مستحقات للمورد') { supplier.due = Math.max(0, supplier.due - amount); supplier.paid += amount; supplier.paymentDate = executionDate; }
    const account = { code: nextCode(data.accounts.map((item) => item.code), 'T'), date: new Date().toISOString().slice(0, 10),
      executionDate, direction, type: operation === 'دين' ? `دين ${debtSide}` : type, description: clean(input.description) || type,
      customerCode: employeeMovement ? '' : (customer?.code || ''), customerName: employeeMovement ? '' : (customer?.name || ''), visitCode: '',
      supplierCode: employeeMovement ? '' : (supplier?.code || ''), supplierName: employeeMovement ? '' : (supplier?.name || ''),
      employeeCode: employee?.code || '', employeeName: employee?.name || '',
      productCode: '', productName: '', qty: 0, total: amount, paid: operation === 'دين' ? 0 : payment.paid, due: operation === 'دين' ? amount : amount - payment.paid,
      paymentMethod: operation === 'دين' ? 'آجل' : payment.paymentMethod, paymentDetails: operation === 'دين' ? `آجل: ${amount}` : payment.paymentDetails, notes: clean(input.notes) };
    appendAccount(data.accounts, account);
    saveData(data.customers, data.visits, null, supplier ? data.suppliers : null, null, data.accounts, data.employees);
    return json(response, 201, account);
  }

  if (request.method === 'POST' && pathname === '/api/inventory') {
    const input = await bodyOf(request);
    const data = readData();
    const mode = clean(input.mode);
    const qty = Math.max(0, Math.floor(Number(input.qty) || 0));
    const buy = Number(input.buy) || 0;
    const sell = Number(input.sell) || 0;
    const payment = paymentInfo(input, 0);
    const paid = payment.paid;
    if (!qty || buy < 0 || sell < 0) return json(response, 400, { error: 'اكتب كمية وأسعار صحيحة.' });
    const purchaseTotal = qty * buy;
    if (payment.allocated > purchaseTotal) return json(response, 400, { error: 'مجموع طرق الدفع أكبر من إجمالي سعر الشراء.' });
    if (paid > purchaseTotal) return json(response, 400, { error: 'المبلغ المدفوع أكبر من إجمالي سعر الشراء.' });
    const due = purchaseTotal - paid;
    const supplierName = clean(input.supplier);
    let item;
    if (mode === 'new') {
      const code = clean(input.code);
      if (!clean(input.name) || !clean(input.details) || !clean(input.country) || !code)
        return json(response, 400, { error: 'أكمل اسم المنتج وتفاصيله وبلد المنشأ والكود.' });
      if (data.inventory.some((existing) => existing.code.toLocaleLowerCase() === code.toLocaleLowerCase()))
        return json(response, 409, { error: 'كود المنتج موجود بالفعل في المخزن.' });
      item = { name: clean(input.name), details: clean(input.details), country: clean(input.country), code,
        qty, buy, sell, margin: sell - buy, supplier: supplierName, paid, due };
      data.inventory.push(item);
    } else if (mode === 'old') {
      item = data.inventory.find((existing) => existing.code.toLocaleLowerCase() === clean(input.code).toLocaleLowerCase());
      if (!item) return json(response, 404, { error: 'المنتج القديم غير موجود في المخزن.' });
      item.qty += qty;
      item.buy = buy;
      item.sell = sell;
      item.margin = sell - buy;
      item.supplier = supplierName || item.supplier;
      item.paid = Number(item.paid || 0) + paid;
      item.due = Number(item.due || 0) + due;
    } else return json(response, 400, { error: 'اختر منتجًا جديدًا أو قديمًا.' });
    const movement = {
      code: nextCode(data.movements.map((entry) => entry.code), 'M'), date: new Date().toISOString().slice(0, 10), time: currentTime(),
      type: 'وارد', productCode: item.code, productName: item.name, details: item.details, country: item.country,
      qty, unitPrice: buy, total: purchaseTotal, visitCode: '', customerCode: '', supplier: supplierName || item.supplier,
      notes: mode === 'new' ? 'إضافة منتج جديد' : 'إضافة كمية لمنتج قديم',
    };
    data.movements.push(movement);
    const supplierForAccount = data.suppliers.find((entry) => entry.name === (supplierName || item.supplier));
    const account = { code: nextCode(data.accounts.map((entry) => entry.code), 'A'), date: movement.date, executionDate: movement.date,
      direction: 'صادر', type: 'توريد بضاعة', description: `توريد ${item.name}`, customerCode: '', customerName: '', visitCode: '',
      supplierCode: supplierForAccount?.code || '', supplierName: supplierForAccount?.name || supplierName,
      productCode: item.code, productName: item.name, qty, total: purchaseTotal, paid, due,
      paymentMethod: payment.paymentMethod, paymentDetails: payment.paymentDetails, notes: mode === 'new' ? 'إضافة منتج جديد' : 'إضافة كمية لمنتج قديم' };
    appendAccount(data.accounts, account);
    if (supplierName) {
      let supplier = data.suppliers.find((entry) => entry.name === supplierName);
      if (!supplier) {
        supplier = { code: nextCode(data.suppliers.map((entry) => entry.code), 'S'), name: supplierName, phone: '', contractDate: new Date().toISOString().slice(0, 10), due: 0, notes: 'تمت إضافته من شاشة المخزن', paid: 0, paymentDate: '' };
        data.suppliers.push(supplier);
      }
      supplier.paid += paid;
      supplier.due += due;
      supplier.paymentDate = new Date().toISOString().slice(0, 10);
    }
    saveData(data.customers, data.visits, data.inventory, data.suppliers, data.movements, data.accounts);
    return json(response, 201, { item, purchaseTotal, due, movement, account });
  }

  json(response, 404, { error: 'المسار غير موجود.' });
}

const MIME = { '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.svg': 'image/svg+xml' };
const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url, `http://${request.headers.host}`);
    if (url.pathname.startsWith('/api/')) return await api(request, response, url.pathname, url.searchParams);
    const relative = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname.slice(1));
    const file = path.resolve(PUBLIC, relative);
    if (!file.startsWith(PUBLIC) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      response.writeHead(404); return response.end('Not found');
    }
    response.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    fs.createReadStream(file).pipe(response);
  } catch (error) {
    console.error(error);
    json(response, 500, { error: error.message || 'حدث خطأ غير متوقع.' });
  }
});

server.on('error', (error) => {
  clearPidFile();
  if (error.code === 'EADDRINUSE') console.error(`المنفذ ${PORT} مستخدم بواسطة نسخة أخرى من النظام.`);
  else console.error(error);
  process.exit(1);
});
try {
  migrateStoredAutoCodes();
  repairStoredCustomerLinks();
  ensureStoredSupplierCodes();
  ensureStoredAccountSchema();
  ensureStoredCustomerSchema();
  ensureStoredEmployeeSchema();
} catch (error) {
  console.error('تعذر تحديث تنسيق الأكواد داخل ملف Excel:', error.message);
}
server.listen(PORT, () => console.log(`نظام المركز يعمل على http://localhost:${PORT} — قاعدة البيانات: ${DATA_FILE}`));
