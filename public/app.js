const state = { customers: [], visits: [], inventory: [], suppliers: [], employees: [], expenses: [], movements: [], accounts: [], selectedCustomer: null, movementMode: '', selectedProduct: null, productMovementMode: '', selectedSupplierName: '' };
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const fmt = (value) => Number(value || 0).toLocaleString('ar-EG');
const esc = (value) => String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', "'":'&#39;', '"':'&quot;' }[char]));
const normalized = (value) => String(value || '').trim().replace(/[\s-]/g, '').toLocaleLowerCase('ar-EG');
const codeKey = (value) => { const text = String(value || '').trim(); const match = text.match(/^([a-zA-Z])[-\s]*0*(\d+)$/); return match ? `${match[1].toUpperCase()}${Number(match[2])}` : normalized(text); };
const paymentMethodOptions = () => ['نقدي', 'إنستاباي', 'فودافون كاش', 'تحويل بنكي', 'آجل'].map((value) => `<option value="${value}">${value}</option>`).join('');
function paymentRow() { return `<div class="split-payment-row"><select class="split-payment-method">${paymentMethodOptions()}</select><input class="split-payment-amount" type="number" min="0" step="0.01" placeholder="المبلغ"><button type="button" class="remove-split-payment" aria-label="حذف">×</button></div>`; }
function installSplitPayment(form, paidFieldName = '') {
  if (!form || form.querySelector('.split-payment-editor')) return;
  const method = form.elements.paymentMethod;
  const methodLabel = method?.closest('label');
  if (!methodLabel) return;
  methodLabel.classList.add('legacy-payment-field');
  if (paidFieldName && form.elements[paidFieldName]) form.elements[paidFieldName].closest('label')?.classList.add('legacy-payment-field');
  const editor = document.createElement('div');
  editor.className = 'split-payment-editor wide';
  editor.innerHTML = `<div class="split-payment-head"><div><b>طرق الدفع</b><small>اكتب مبلغ كل طريقة، ويمكن إضافة أكثر من طريقة.</small></div><button type="button" class="secondary add-split-payment">＋ إضافة طريقة دفع</button></div><div class="split-payment-rows">${paymentRow()}</div>`;
  methodLabel.before(editor);
  const addRow = () => { editor.querySelector('.split-payment-rows').insertAdjacentHTML('beforeend', paymentRow()); bindSplitPaymentRows(editor); };
  editor.querySelector('.add-split-payment').addEventListener('click', addRow);
  bindSplitPaymentRows(editor);
}
function bindSplitPaymentRows(editor) {
  const sync = () => {
    const form = editor.closest('form');
    const paid = [...editor.querySelectorAll('.split-payment-row')].reduce((sum, row) => row.querySelector('.split-payment-method').value === 'آجل' ? sum : sum + (Number(row.querySelector('.split-payment-amount').value) || 0), 0);
    if (form.elements.paid) form.elements.paid.value = paid;
    if (form.id === 'inventoryForm') calculateInventoryPurchase();
    if (form.id === 'supplierTransactionForm') calculateSupplierTransaction();
  };
  editor.querySelectorAll('.remove-split-payment').forEach((button) => { button.onclick = () => { if (editor.querySelectorAll('.split-payment-row').length > 1) { button.closest('.split-payment-row').remove(); sync(); } }; });
  editor.querySelectorAll('.split-payment-method,.split-payment-amount').forEach((input) => { input.oninput = sync; input.onchange = sync; });
}
function resetSplitPayment(form) {
  const rows = form.querySelector('.split-payment-rows');
  if (rows) rows.innerHTML = paymentRow();
  if (rows) bindSplitPaymentRows(form.querySelector('.split-payment-editor'));
}
function addPaymentsToInput(input, form) {
  input.payments = [...form.querySelectorAll('.split-payment-row')].map((row) => ({ method: row.querySelector('.split-payment-method').value, amount: Number(row.querySelector('.split-payment-amount').value) || 0 })).filter((entry) => entry.amount > 0);
  if (!input.payments.length) throw new Error('اختر طريقة الدفع واكتب المبلغ قبل الحفظ. استخدم آجل للمبلغ غير المدفوع.');
  input.paymentMethod = [...new Set(input.payments.map((entry) => entry.method))].join(' + ') || 'آجل';
  input.paid = input.payments.filter((entry) => entry.method !== 'آجل').reduce((sum, entry) => sum + entry.amount, 0);
  return input;
}
let partChoiceCounter = 0;
let financialPeriod = 'daily';
let dailyClosingVisible = false;
let statementPeriod = 'daily';
let supplierDebtsVisible = false;

async function request(url, options) {
  const response = await fetch(url, { headers: { 'Content-Type': 'application/json' }, ...options });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || 'تعذر تنفيذ الطلب');
  return data;
}

function toast(message, error = false) {
  const element = $('#toast');
  element.textContent = message;
  element.className = `toast show${error ? ' error' : ''}`;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => element.className = 'toast', 2600);
}

async function load() {
  try {
    const data = await request('/api/data');
    state.customers = data.customers;
    state.visits = data.visits;
    state.inventory = data.inventory || [];
    state.suppliers = data.suppliers || [];
    state.employees = data.employees || [];
    state.expenses = data.expenses || [];
    state.movements = data.movements || [];
    state.accounts = data.accounts || [];
    $('#databaseFileName').textContent = data.databaseFile || 'main data 2.xlsx';
    renderDashboard();
  } catch (error) { toast(error.message, true); }
}

function go(page) {
  $$('.page').forEach((element) => element.classList.toggle('active', element.id === page));
  $$('.nav-btn').forEach((element) => element.classList.toggle('active', element.dataset.page === page));
  $('.sidebar').classList.remove('open');
  if (page === 'customers') resetCustomerView();
  if (page === 'clients') renderCustomerDirectory();
  if (page === 'inventory') { closeInventoryMovements(); renderInventory(); }
  if (page === 'suppliers') renderSuppliers();
  if (page === 'employees') renderEmployees();
  if (page === 'accounts') renderAccounts();
}

function renderDashboard() {
  const month = new Date().toISOString().slice(0, 7);
  const monthVisits = state.visits.filter((visit) => visit.date.startsWith(month));
  const monthExpenses = state.expenses.filter((expense) => expense.date.startsWith(month)).reduce((sum, expense) => sum + expense.amount, 0);
  const suppliersPaid = state.suppliers.filter((supplier) => supplier.paymentDate.startsWith(month)).reduce((sum, supplier) => sum + supplier.paid, 0);
  const monthlyIncome = monthVisits.reduce((sum, visit) => sum + visit.total, 0);
  const goodsProfit = monthVisits.reduce((sum, visit) => sum + Math.max(0, visit.partsTotal - visit.partsCost), 0);
  const laborProfit = monthVisits.reduce((sum, visit) => sum + visit.labor, 0);
  const currentCenterBalance = state.accounts.length ? Number(state.accounts[state.accounts.length - 1].balance || 0) : 0;
  const lowStock = state.inventory.filter((item) => item.qty <= 3);
  $('#newCustomersMonth').textContent = fmt(state.customers.filter((customer) => customer.registeredDate.startsWith(month)).length);
  $('#monthVisitsCount').textContent = fmt(monthVisits.length);
  $('#lowStockCount').textContent = fmt(lowStock.length);
  $('#monthlyIncome').textContent = `${fmt(monthlyIncome)} ج`;
  $('#availableAmount').textContent = `${fmt(currentCenterBalance)} ج`;
  $('#goodsProfit').textContent = `${fmt(goodsProfit)} ج`;
  $('#laborProfit').textContent = `${fmt(laborProfit)} ج`;
  $('#monthlyExpenses').textContent = `${fmt(monthExpenses)} ج`;
  $('#suppliersPaid').textContent = `${fmt(suppliersPaid)} ج`;
  $('#suppliersDue').textContent = `${fmt(state.suppliers.reduce((sum, supplier) => sum + supplier.due, 0))} ج`;
  $('#inventoryValue').textContent = `${fmt(state.inventory.reduce((sum, item) => sum + item.buy * item.qty, 0))} ج`;
  const lowStockPanel = $('.low-stock-panel');
  lowStockPanel.classList.toggle('has-items', Boolean(lowStock.length));
  $('#lowStockItems').innerHTML = lowStock.length
    ? `<div class="low-stock-cards">${lowStock.map((item) => `<div class="low-stock-product-card"><div class="low-stock-product-title"><h3>${esc(item.name)}</h3><strong>${fmt(item.qty)} متاح</strong></div><div class="low-stock-product-meta"><span><small>تفاصيل المنتج</small><b>${esc(item.details || 'بدون تفاصيل')}</b></span><span><small>بلد المنشأ</small><b>${esc(item.country || 'غير مسجل')}</b></span></div></div>`).join('')}</div>`
    : '<div class="empty-state">✓ كل الكميات جيدة حاليًا</div>';
  const recentVisits = [...monthVisits].reverse().slice(0, 5);
  $('#recentVisits').innerHTML = recentVisits.length
    ? `<div class="recent-visit-list">${recentVisits.map((visit) => {
      const customer = state.customers.find((item) => item.code === visit.customerCode)
        || state.customers.find((item) => normalized(item.plate) === normalized(visit.plate)) || {};
      return `<button class="recent-visit-item" data-customer="${esc(customer.code || visit.customerCode)}"><div><b>${esc(customer.name || 'عميل غير معروف')}</b><span>— ${esc(visit.serviceType)}</span></div><time>${esc(visit.date)}</time></button>`;
    }).join('')}</div>`
    : '<div class="empty-state compact-empty">لا توجد زيارات خلال هذا الشهر.</div>';
  $$('#recentVisits .recent-visit-item').forEach((item) => item.addEventListener('click', () => {
    go('clients');
    openCustomerRecord(item.dataset.customer);
  }));
}

function installGlobalCodeSearch() {
  const dashboardTitle = $('#dashboard .page-title');
  const search = document.createElement('div'); search.className = 'global-code-search';
  search.innerHTML = `<div><b>بحث سريع بالكود</b><small>عميل، زيارة، منتج، مورد، موظف أو حركة</small></div><label><input id="globalCodeSearchInput" placeholder="اكتب الكود مثل C1 أو V1 أو E1"><button id="globalCodeSearchButton" class="primary" type="button">بحث</button></label><div id="globalCodeSearchResults" class="global-code-results hidden"></div>`;
  dashboardTitle.after(search);
  $('#globalCodeSearchButton').addEventListener('click', runGlobalCodeSearch);
  $('#globalCodeSearchInput').addEventListener('keydown', (event) => { if (event.key === 'Enter') runGlobalCodeSearch(); });
  $('#globalCodeSearchInput').addEventListener('input', renderGlobalCodeSuggestions);
}

function globalSearchEntries() {
  return [
    ...state.customers.map((item) => ({ kind: 'customer', code: item.code, type: 'عميل', label: `${item.name} · ${item.phone}` })),
    ...state.visits.map((item) => ({ kind: 'visit', code: item.code, type: 'زيارة', label: `${item.serviceType} · ${item.date}` })),
    ...state.inventory.map((item) => ({ kind: 'product', code: item.code, type: 'منتج', label: `${item.name} · ${item.details || ''}` })),
    ...state.suppliers.map((item) => ({ kind: 'supplier', code: item.code, type: 'مورد', label: `${item.name} · ${item.phone || ''}` })),
    ...state.employees.map((item) => ({ kind: 'employee', code: item.code, type: 'موظف', label: `${item.name} · ${item.specialty}` })),
    ...state.accounts.map((item) => ({ kind: 'account', code: item.code, type: 'حركة مالية', label: `${item.type} · ${item.description || ''} · ${fmt(item.paid)} ج` })),
    ...state.movements.map((item) => ({ kind: 'movement', code: item.code, type: 'حركة مخزن', label: `${item.type} · ${item.productName} · ${fmt(item.qty)}` })),
  ];
}

function renderGlobalCodeSuggestions() {
  const raw = $('#globalCodeSearchInput').value.trim(); const results = $('#globalCodeSearchResults');
  if (!raw) { results.classList.add('hidden'); results.innerHTML = ''; return; }
  const key = codeKey(raw); const text = normalized(raw);
  const matches = globalSearchEntries().filter((entry) => codeKey(entry.code).startsWith(key) || normalized(entry.code).includes(text)).slice(0, 8);
  results.classList.remove('hidden');
  results.innerHTML = matches.length ? matches.map((entry) => `<button type="button" class="global-code-result" data-kind="${entry.kind}" data-code="${esc(entry.code)}"><span>${esc(entry.type)}</span><b>${esc(entry.code)}</b><small>${esc(entry.label)}</small><i>فتح ←</i></button>`).join('') : `<div class="global-code-no-result">لا توجد أكواد مطابقة.</div>`;
  $$('#globalCodeSearchResults .global-code-result').forEach((button) => button.addEventListener('click', () => openGlobalSearchEntry(button.dataset.kind, button.dataset.code)));
}

function openGlobalSearchEntry(kind, code) {
  $('#globalCodeSearchResults').classList.add('hidden');
  if (kind === 'customer') { go('clients'); openCustomerRecord(code); }
  if (kind === 'visit') { const visit = state.visits.find((item) => item.code === code); go('clients'); if (visit) openCustomerRecord(visit.customerCode); showVisit(code); }
  if (kind === 'supplier') { const supplier = state.suppliers.find((item) => item.code === code); state.selectedSupplierName = supplier?.name || ''; go('suppliers'); showSupplierDetails(code); }
  if (kind === 'employee') { go('employees'); showEmployeeDetails(code); }
  if (kind === 'account') { go('accounts'); showAccountDetails(code); }
  if (kind === 'movement') openMovementByCode(code);
  if (kind === 'product') { go('inventory'); openProductDetails(code); }
}

function runGlobalCodeSearch() {
  const raw = $('#globalCodeSearchInput').value.trim(); if (!raw) return toast('اكتب الكود الذي تريد البحث عنه.', true);
  const key = codeKey(raw); const exact = (value) => codeKey(value) === key;
  const entry = globalSearchEntries().find((item) => exact(item.code));
  if (entry) { openGlobalSearchEntry(entry.kind, entry.code); return; }
  toast(`لم يتم العثور على كود ${raw}.`, true);
}

function renderInventory() {
  const element = $('#inventoryTable');
  if (!state.inventory.length) {
    element.innerHTML = '<div class="empty-state">لا توجد منتجات مسجلة في شيت المخزن.</div>';
    return;
  }
  const sortedInventory = [...state.inventory].sort((a, b) => {
    const aLow = a.qty <= 3 ? 0 : 1;
    const bLow = b.qty <= 3 ? 0 : 1;
    return aLow - bLow || a.qty - b.qty || a.name.localeCompare(b.name, 'ar');
  });
  const rows = sortedInventory.map((item) => `<tr class="inventory-product-row ${item.qty <= 3 ? 'inventory-low-row' : ''}" data-product-code="${esc(item.code)}" tabindex="0"><td><b>${esc(item.name)}</b></td><td>${esc(item.code)}</td><td>${esc(item.details || '—')}</td><td>${esc(item.country || '—')}</td><td class="${item.qty <= 3 ? 'stock-danger' : 'stock-ok'}">${fmt(item.qty)}</td><td>${fmt(item.buy)} ج</td><td>${fmt(item.sell)} ج</td><td>${fmt(item.buy * item.qty)} ج</td></tr>`).join('');
  element.innerHTML = `<div class="table-wrap"><table><thead><tr><th>المنتج</th><th>الكود</th><th>تفاصيل المنتج</th><th>بلد المنشأ</th><th>الكمية</th><th>سعر الشراء</th><th>سعر البيع</th><th>قيمة المخزون</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  $$('.inventory-product-row').forEach((row) => {
    row.addEventListener('click', () => openProductDetails(row.dataset.productCode));
    row.addEventListener('keydown', (event) => (event.key === 'Enter' || event.key === ' ') && openProductDetails(row.dataset.productCode));
  });
}

function supplierOptions(selected = '') {
  return state.suppliers.map((supplier) => `<option value="${esc(supplier.name)}">`).join('');
}

function upgradeSearchableChoices() {
  $('#searchCode').placeholder = 'مثال: C1';
  $('#directoryCustomerCode').placeholder = 'مثال: C1';
  $('#directoryVisitCode').placeholder = 'مثال: V1';
  const technician = $('#visitForm [name="technician"]');
  if (technician && !$('#employeeTechnicianOptions')) { technician.setAttribute('list', 'employeeTechnicianOptions'); technician.insertAdjacentHTML('afterend', '<datalist id="employeeTechnicianOptions"></datalist>'); }
  const supplierSelect = $('#inventorySupplier');
  if (supplierSelect?.tagName === 'SELECT') {
    const input = document.createElement('input');
    input.id = 'inventorySupplier';
    input.name = 'supplier';
    input.setAttribute('list', 'inventorySupplierOptions');
    input.placeholder = 'اضغط للاختيار أو اكتب اسم المورد';
    const list = document.createElement('datalist');
    list.id = 'inventorySupplierOptions';
    supplierSelect.replaceWith(input);
    input.after(list);
  }
  const productSearch = $('#oldProductSearch');
  if (productSearch && !$('#oldProductOptions')) {
    productSearch.setAttribute('list', 'oldProductOptions');
    productSearch.placeholder = 'اضغط للاختيار أو اكتب كود أو اسم المنتج';
    const list = document.createElement('datalist');
    list.id = 'oldProductOptions';
    productSearch.after(list);
    $('#productSuggestions').classList.add('hidden');
  }
  if (!$('#visitForm [name="paymentMethod"]')) {
    const laborLabel = $('#visitForm [name="labor"]').closest('label');
    laborLabel.insertAdjacentHTML('afterend', `<label>طريقة الدفع<select name="paymentMethod" required>${paymentMethodOptions()}</select></label><label>المبلغ المدفوع<input name="paid" type="number" min="0" step="0.01" placeholder="فارغ = سداد الإجمالي"></label>`);
  }
  if (!$('#inventoryForm [name="paymentMethod"]')) {
    $('.inventory-common-fields').insertAdjacentHTML('beforeend', `<label>طريقة الدفع<select name="paymentMethod" required>${paymentMethodOptions()}</select></label>`);
  }
  if (!$('#supplierTransactionForm [name="paymentMethod"]')) {
    $('#supplierTransactionForm .transaction-notes').insertAdjacentHTML('beforebegin', `<label class="transaction-notes">طريقة الدفع<select name="paymentMethod" required>${paymentMethodOptions()}</select></label>`);
  }
  installSplitPayment($('#visitForm'), 'paid');
  installSplitPayment($('#inventoryForm'), 'paid');
  installSplitPayment($('#supplierTransactionForm'), 'paid');
  installSplitPayment($('#manualAccountForm'));
}

function openInventoryDialog() {
  resetInventoryDialog();
  $('#inventoryDialog').showModal();
}

function resetInventoryDialog() {
  const form = $('#inventoryForm');
  form.reset();
  resetSplitPayment(form);
  form.elements.paymentMethod.value = 'نقدي';
  form.classList.add('hidden');
  $('#inventoryModeChoice').classList.remove('hidden');
  $('#newProductFields').classList.add('hidden');
  $('#oldProductFields').classList.add('hidden');
  $('#selectedOldProduct').classList.add('hidden');
  $('#productSuggestions').innerHTML = '';
  $('#inventorySupplier').value = '';
  $('#inventorySupplierOptions').innerHTML = supplierOptions();
  calculateInventoryPurchase();
}

function chooseInventoryMode(mode) {
  const form = $('#inventoryForm');
  form.reset();
  resetSplitPayment(form);
  form.elements.paymentMethod.value = 'نقدي';
  form.elements.mode.value = mode;
  form.elements.code.value = '';
  $('#inventoryModeChoice').classList.add('hidden');
  form.classList.remove('hidden');
  $('#newProductFields').classList.toggle('hidden', mode !== 'new');
  $('#oldProductFields').classList.toggle('hidden', mode !== 'old');
  $('#selectedOldProduct').classList.add('hidden');
  $('#inventorySupplier').value = '';
  $('#inventorySupplierOptions').innerHTML = supplierOptions();
  calculateInventoryPurchase();
  if (mode === 'old') {
    $('#oldProductSearch').value = '';
    renderProductSuggestions();
    $('#oldProductSearch').focus();
  }
}

function calculateInventoryPurchase() {
  const form = $('#inventoryForm');
  const qty = Math.max(0, Number(form.elements.qty?.value) || 0);
  const buy = Math.max(0, Number(form.elements.buy?.value) || 0);
  const sell = Math.max(0, Number(form.elements.sell?.value) || 0);
  const paid = Math.max(0, Number(form.elements.paid?.value) || 0);
  const total = qty * buy;
  const due = Math.max(0, total - paid);
  $('#inventoryPurchaseTotal').textContent = `${fmt(total)} ج`;
  $('#inventoryMargin').textContent = `${fmt(sell - buy)} ج`;
  $('#inventoryDue').value = due;
}

function renderProductSuggestions() {
  const query = normalized($('#oldProductSearch').value);
  const matches = state.inventory.filter((item) => !query || normalized(`${item.code} ${item.name} ${item.details}`).includes(query));
  $('#oldProductOptions').innerHTML = matches.map((item) => `<option value="${esc(item.code)}" label="${esc(item.name)} — ${esc(item.details || 'بدون تفاصيل')} — متاح ${fmt(item.qty)}">`).join('');
  const exactItem = state.inventory.find((item) => normalized(item.code) === query);
  if (exactItem) selectOldProduct(exactItem.code);
}

function selectOldProduct(code) {
  const item = state.inventory.find((entry) => entry.code === code);
  if (!item) return;
  const form = $('#inventoryForm');
  form.elements.code.value = item.code;
  form.elements.buy.value = item.buy;
  form.elements.sell.value = item.sell;
  form.elements.supplier.value = item.supplier || '';
  $('#oldProductSearch').value = item.code;
  const selected = $('#selectedOldProduct');
  selected.classList.remove('hidden');
  selected.innerHTML = `<b>${esc(item.name)}</b><span>${esc(item.details || 'بدون تفاصيل')} · ${esc(item.country || 'بلد المنشأ غير مسجل')} · الكمية الحالية ${fmt(item.qty)}</span>`;
  calculateInventoryPurchase();
}

function installProductDetailsUI() {
  const dialog = document.createElement('dialog');
  dialog.id = 'productDetailsDialog';
  dialog.innerHTML = '<div class="product-details-dialog"><button type="button" class="dialog-close product-dialog-close" aria-label="إغلاق تفاصيل المنتج">×</button><div id="productDetailsContent"></div></div>';
  document.body.appendChild(dialog);
  dialog.querySelector('.dialog-close').addEventListener('click', () => dialog.close());
  const movementDialog = document.createElement('dialog');
  movementDialog.id = 'singleMovementDialog';
  movementDialog.innerHTML = '<div class="single-movement-dialog"><button type="button" class="dialog-close">×</button><div id="singleMovementContent"></div></div>';
  document.body.appendChild(movementDialog);
  movementDialog.querySelector('.dialog-close').addEventListener('click', () => movementDialog.close());
}

function openProductDetails(code) {
  const item = state.inventory.find((entry) => entry.code === code);
  if (!item) return;
  state.selectedProduct = item;
  state.productMovementMode = '';
  $('#productDetailsContent').innerHTML = `<div class="dialog-title"><span>▣</span><div><h2>${esc(item.name)}</h2><p>${esc(item.details || 'بدون تفاصيل')} · ${esc(item.code)}</p></div></div>
    <div class="product-summary-grid"><div><small>تفاصيل المنتج</small><b>${esc(item.details || '—')}</b></div><div><small>بلد المنشأ</small><b>${esc(item.country || '—')}</b></div><div><small>الكمية الحالية</small><b>${fmt(item.qty)}</b></div><div><small>سعر الشراء</small><b>${fmt(item.buy)} ج</b></div><div><small>سعر البيع</small><b>${fmt(item.sell)} ج</b></div><div><small>المورد</small><b>${esc(item.supplier || 'غير مسجل')}</b></div><div><small>هامش الربح</small><b>${fmt(item.sell - item.buy)} ج</b></div></div>
    <div class="product-detail-actions"><button id="addStockForProduct" class="primary" type="button">＋ إضافة بضاعة</button><button id="goToProductSupplier" class="secondary" type="button">اذهب إلى المورد</button></div>
    <div class="product-movement-head"><div><h3>سجل الحركة</h3><p>الوارد والصادر الخاص بهذا المنتج.</p></div><div class="product-movement-filters"><button type="button" class="inbound-filter" data-product-movement="وارد">↓ الوارد</button><button type="button" class="outbound-filter" data-product-movement="صادر">↑ الصادر</button></div></div>
    <div id="productMovementResults"></div>`;
  $('#addStockForProduct').addEventListener('click', () => {
    $('#productDetailsDialog').close();
    openInventoryDialog();
    chooseInventoryMode('old');
    selectOldProduct(item.code);
  });
  $('#goToProductSupplier').addEventListener('click', () => {
    if (!item.supplier) return toast('لا يوجد مورد مسجل لهذا المنتج.', true);
    state.selectedSupplierName = item.supplier;
    $('#productDetailsDialog').close();
    go('suppliers');
  });
  $$('#productDetailsContent [data-product-movement]').forEach((button) => button.addEventListener('click', () => {
    state.productMovementMode = button.dataset.productMovement;
    renderProductMovements();
  }));
  renderProductMovements();
  $('#productDetailsDialog').showModal();
}

function renderProductMovements() {
  const item = state.selectedProduct;
  if (!item) return;
  $$('#productDetailsContent [data-product-movement]').forEach((button) => button.classList.toggle('active', button.dataset.productMovement === state.productMovementMode));
  const movements = state.movements.filter((movement) => movement.productCode === item.code && (!state.productMovementMode || movement.type === state.productMovementMode));
  if (!movements.length) {
    $('#productMovementResults').innerHTML = '<div class="empty-state">لا توجد حركات مسجلة لهذا المنتج.</div>';
    return;
  }
  const rows = movements.map((movement) => `<tr class="product-movement-row" data-movement-code="${esc(movement.code)}" data-product-code="${esc(movement.productCode)}" tabindex="0"><td><b>${esc(movement.code)}</b></td><td><span class="movement-type ${movement.type === 'صادر' ? 'out' : 'in'}">${esc(movement.type)}</span></td><td>${esc(movement.date || '—')}</td><td>${fmt(movement.qty)}</td><td>${fmt(movement.total)} ج</td><td>${esc(movement.visitCode || '—')}</td></tr>`).join('');
  $('#productMovementResults').innerHTML = `<div class="table-wrap"><table><thead><tr><th>كود الحركة</th><th>النوع</th><th>التاريخ</th><th>الكمية</th><th>الإجمالي</th><th>كود الزيارة</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  $$('#productMovementResults .product-movement-row').forEach((row) => {
    row.addEventListener('click', () => showProductMovementDetails(row.dataset.movementCode, row.dataset.productCode));
    row.addEventListener('keydown', (event) => (event.key === 'Enter' || event.key === ' ') && showProductMovementDetails(row.dataset.movementCode, row.dataset.productCode));
  });
}

function showProductMovementDetails(code, productCode) {
  const movement = state.movements.find((item) => item.code === code && item.productCode === productCode);
  if (!movement) return;
  const detail = (label, value) => `<div><small>${label}</small><b>${esc(value)}</b></div>`;
  let content = '';
  if (movement.type === 'صادر') {
    const visit = state.visits.find((item) => item.code === movement.visitCode) || {};
    const customer = state.customers.find((item) => item.code === (visit.customerCode || movement.customerCode))
      || state.customers.find((item) => normalized(item.plate) === normalized(visit.plate)) || {};
    content = `${customer.code ? `<button type="button" class="movement-related-link" data-related-customer="${esc(customer.code)}">العميل: ${esc(customer.name)} ←</button>` : ''}
      <div class="single-movement-grid">${detail('نوع الحركة', 'صادر')}${detail('لوحة العربية', visit.plate || customer.plate || '—')}${detail('التاريخ', movement.date || visit.date || '—')}${detail('الكمية', fmt(movement.qty))}${detail('نوع الصيانة', visit.serviceType || '—')}${detail('كود الزيارة', visit.code || movement.visitCode || '—')}${detail('اسم الفني', visit.technician || '—')}${detail('قيمة الفاتورة الكلية', `${fmt(visit.total || movement.total)} ج`)}</div>`;
  } else {
    const supplier = state.suppliers.find((item) => item.name === movement.supplier) || {};
    content = `${supplier.name ? `<button type="button" class="movement-related-link" data-related-supplier="${esc(supplier.name)}">المورد: ${esc(supplier.name)} ←</button>` : ''}
      <div class="single-movement-grid">${detail('نوع الحركة', 'وارد')}${detail('التاريخ', movement.date || '—')}${detail('عدد القطع', fmt(movement.qty))}${detail('كود المورد', supplier.code || '—')}${detail('سعر الوحدة', `${fmt(movement.unitPrice)} ج`)}${detail('إجمالي الوارد', `${fmt(movement.total)} ج`)}</div>`;
  }
  $('#singleMovementContent').innerHTML = `<div class="dialog-title"><span>${movement.type === 'صادر' ? '↑' : '↓'}</span><div><h2>تفاصيل الحركة ${esc(movement.code)}</h2><p>${esc(movement.productName)} · ${esc(movement.productCode)}</p></div></div>${content}`;
  $('#singleMovementContent [data-related-customer]')?.addEventListener('click', (event) => {
    $('#singleMovementDialog').close(); $('#productDetailsDialog').close(); go('clients');
    $$('.nav-btn').forEach((button) => button.classList.toggle('active', button.dataset.page === 'clients'));
    openCustomerRecord(event.currentTarget.dataset.relatedCustomer);
  });
  $('#singleMovementContent [data-related-supplier]')?.addEventListener('click', (event) => {
    state.selectedSupplierName = event.currentTarget.dataset.relatedSupplier;
    $('#singleMovementDialog').close(); $('#productDetailsDialog').close(); go('suppliers');
    $$('.nav-btn').forEach((button) => button.classList.toggle('active', button.dataset.page === 'suppliers'));
  });
  $('#singleMovementDialog').showModal();
}

function customerVisits(customer) {
  return state.visits.filter((visit) => visit.customerCode === customer.code || normalized(visit.plate) === normalized(customer.plate));
}

function renderCustomerDirectory() {
  $('#clientDirectoryView').classList.remove('hidden');
  $('#clientRecordView').classList.add('hidden');
  $('#directoryCount').textContent = `${fmt(state.customers.length)} عميل مسجل`;
  const element = $('#customersDirectory');
  if (!state.customers.length) {
    element.innerHTML = '<div class="empty-state">لا يوجد عملاء مسجلون بعد.</div>';
    return;
  }
  const rows = state.customers.map((customer) => {
    const visits = customerVisits(customer);
    const income = visits.reduce((sum, visit) => sum + visit.total, 0);
    return `<tr class="customer-directory-row" data-customer="${esc(customer.code)}"><td><b>${esc(customer.code)}</b></td><td>${esc(customer.name)}</td><td>${esc(customer.phone)}</td><td>${esc(customer.plate)}</td><td>${esc(customer.carType)}</td><td>${fmt(visits.length)}</td><td>${fmt(income)} ج</td></tr>`;
  }).join('');
  element.innerHTML = `<div class="table-wrap"><table><thead><tr><th>كود العميل</th><th>الاسم</th><th>التليفون</th><th>لوحة العربية</th><th>نوع العربية</th><th>الزيارات</th><th>إجمالي الدخل</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  $$('.customer-directory-row').forEach((row) => row.addEventListener('click', () => openCustomerRecord(row.dataset.customer)));
}

function openCustomerRecord(code) {
  const customer = state.customers.find((item) => item.code.toLocaleLowerCase() === String(code).toLocaleLowerCase());
  if (!customer) return toast('كود العميل غير موجود.', true);
  state.selectedCustomer = customer;
  const visits = customerVisits(customer);
  const income = visits.reduce((sum, visit) => sum + visit.total, 0);
  $('#clientDirectoryView').classList.add('hidden');
  $('#clientRecordView').classList.remove('hidden');
  $('#clientRecordContent').innerHTML = `<div class="panel customer-record-head"><div class="avatar">${esc(customer.name.slice(0, 1))}</div><div class="customer-record-identity"><h2>${esc(customer.name)}</h2><div class="customer-meta"><span>الكود: <b>${esc(customer.code)}</b></span><span>التليفون: <b>${esc(customer.phone)}</b></span><span>العربية: <b>${esc(customer.carType)}</b></span><span>اللوحة: <b>${esc(customer.plate)}</b></span></div></div><div class="customer-record-stats"><div class="record-stat"><small>إجمالي الزيارات</small><strong>${fmt(visits.length)}</strong></div><div class="record-stat"><small>إجمالي الدخل</small><strong>${fmt(income)} ج</strong></div><div class="record-stat debt-receivable"><small>المستحق على العميل</small><strong>${fmt(customer.dueFromCustomer)} ج</strong></div><div class="record-stat debt-payable"><small>المستحق على المركز</small><strong>${fmt(customer.dueFromCenter)} ج</strong></div></div><button id="recordNewVisit" class="primary">＋ زيارة جديدة</button></div><div class="panel visit-summary"><div class="panel-head"><div><h2>سجل الزيارات</h2><p>اضغط على أي زيارة لعرض تفاصيلها.</p></div></div>${visits.length ? table([...visits].reverse()) : '<div class="empty-state">لا توجد زيارات لهذا العميل.</div>'}</div>`;
  $('#recordNewVisit').addEventListener('click', openVisitDialog);
  $('#clientRecordContent').querySelectorAll('.visit-row').forEach((row) => row.addEventListener('click', () => showVisit(row.dataset.visit)));
}

function searchClientDirectory() {
  const customerCode = $('#directoryCustomerCode').value.trim();
  const visitCode = $('#directoryVisitCode').value.trim();
  if (!customerCode && !visitCode) return toast('اكتب كود العميل أو كود الزيارة.', true);
  if (customerCode) return openCustomerRecord(customerCode);
  const visit = state.visits.find((item) => item.code.toLocaleLowerCase() === visitCode.toLocaleLowerCase());
  if (!visit) return toast('كود الزيارة غير موجود.', true);
  showVisit(visit.code);
}

function table(visits, clickable = true) {
  const body = visits.map((visit) => {
    const customer = state.customers.find((item) => item.code === visit.customerCode) || {};
    return `<tr class="${clickable ? 'visit-row' : ''}" data-visit="${esc(visit.code)}"><td>${esc(visit.code || '—')}</td><td>${esc(customer.name || '—')}<br><span class="tag">${esc(visit.plate || customer.plate || '')}</span></td><td>${esc(visit.date)}</td><td>${esc(visit.serviceType)}</td><td>${fmt(visit.total)} ج</td></tr>`;
  }).join('');
  return `<div class="table-wrap"><table><thead><tr><th>كود الزيارة</th><th>العميل والعربية</th><th>التاريخ</th><th>نوع الصيانة</th><th>الإجمالي</th></tr></thead><tbody>${body}</tbody></table></div>`;
}

function resetCustomerView() {
  $('#customerChoices').classList.remove('hidden');
  $('#newCustomerView').classList.add('hidden');
  $('#oldCustomerView').classList.add('hidden');
  $('#customerResult').classList.add('hidden');
  $('#searchMessage').classList.remove('hidden');
  state.selectedCustomer = null;
}

function showMode(mode) {
  $('#customerChoices').classList.add('hidden');
  $(`#${mode === 'new' ? 'newCustomerView' : 'oldCustomerView'}`).classList.remove('hidden');
}

function searchCustomer() {
  const code = $('#searchCode').value.trim().toLocaleLowerCase();
  const plate = normalized($('#searchPlate').value);
  if (!code && !plate) return toast('اكتب كود العميل أو لوحة العربية أولًا.', true);
  const customer = state.customers.find((item) => (code && item.code.toLocaleLowerCase() === code) || (plate && normalized(item.plate) === plate));
  if (!customer) {
    $('#customerResult').classList.add('hidden');
    $('#searchMessage').classList.remove('hidden');
    $('#searchMessage').textContent = 'لم يتم العثور على عميل بهذه البيانات.';
    return;
  }
  state.selectedCustomer = customer;
  const visits = state.visits.filter((visit) => visit.customerCode === customer.code || normalized(visit.plate) === normalized(customer.plate));
  $('#searchMessage').classList.add('hidden');
  $('#customerResult').classList.remove('hidden');
  $('#customerResult').innerHTML = `<div class="panel customer-card"><div class="avatar">${esc(customer.name.slice(0,1))}</div><div><h2>${esc(customer.name)}</h2><div class="customer-meta"><span>الكود: <b>${esc(customer.code)}</b></span><span>التليفون: <b>${esc(customer.phone)}</b></span><span>العربية: <b>${esc(customer.carType)}</b></span><span>اللوحة: <b>${esc(customer.plate)}</b></span><span>الزيارات السابقة: <b>${fmt(visits.length)}</b></span><span>المستحق على العميل: <b>${fmt(customer.dueFromCustomer)} ج</b></span><span>المستحق على المركز: <b>${fmt(customer.dueFromCenter)} ج</b></span></div></div><button id="newVisitBtn" class="primary">＋ زيارة جديدة</button></div><div class="panel visit-summary"><h3>الزيارات السابقة (${fmt(visits.length)})</h3>${visits.length ? table([...visits].reverse()) : '<div class="empty-state">لا توجد زيارات سابقة لهذا العميل.</div>'}</div>`;
  $('#newVisitBtn').addEventListener('click', openVisitDialog);
  $$('.visit-row').forEach((row) => row.addEventListener('click', () => showVisit(row.dataset.visit)));
}

function openVisitDialog() {
  const form = $('#visitForm');
  form.reset();
  resetSplitPayment(form);
  form.elements.date.value = new Date().toISOString().slice(0, 10);
  form.elements.labor.value = '';
  $('#employeeTechnicianOptions').innerHTML = state.employees.filter((employee) => !employee.status || employee.status === 'يعمل').map((employee) => `<option value="${esc(employee.name)}" label="${esc(employee.code)} · ${esc(employee.specialty)}">`).join('');
  form.elements.paymentMethod.value = 'نقدي';
  resetVisitParts();
  $('#visitCustomerName').textContent = `${state.selectedCustomer.name} — ${state.selectedCustomer.plate}`;
  $('#visitDialog').showModal();
}

function partOptions() {
  return state.inventory.map((item) =>
    `<option value="${esc(item.code)}" label="${esc(item.name)} — ${esc(item.details || 'بدون تفاصيل')} — ${esc(item.country || 'بلد المنشأ غير مسجل')} — متاح ${fmt(item.qty)}">`
  ).join('');
}

function resetVisitParts() {
  const container = $('#visitPartsRows');
  container.innerHTML = '';
  $('#addVisitPart').disabled = !state.inventory.length;
  if (!state.inventory.length) {
    container.innerHTML = '<div class="inventory-empty">لا توجد منتجات مسجلة في شيت المخزن بعد.</div>';
  } else {
    addVisitPartRow();
  }
  calculateVisitParts();
}

function addVisitPartRow() {
  const container = $('#visitPartsRows');
  container.querySelector('.inventory-empty')?.remove();
  const row = document.createElement('div');
  const choiceListId = `visitPartChoices${++partChoiceCounter}`;
  row.className = 'part-entry';
  row.innerHTML = `<label>المنتج<input class="part-select" list="${choiceListId}" placeholder="اضغط للاختيار أو اكتب كود المنتج"><datalist id="${choiceListId}">${partOptions()}</datalist></label><label>الكمية<input class="part-qty" type="number" min="1" value="1"></label><div class="part-info">اختر المنتج لعرض السعر</div><button class="remove-part" type="button" title="حذف القطعة">×</button>`;
  container.appendChild(row);
  row.querySelector('.part-select').addEventListener('input', calculateVisitParts);
  row.querySelector('.part-qty').addEventListener('input', calculateVisitParts);
  row.querySelector('.remove-part').addEventListener('click', () => { row.remove(); calculateVisitParts(); });
}

function selectedVisitParts() {
  return $$('#visitPartsRows .part-entry').map((row) => ({
    code: row.querySelector('.part-select').value,
    qty: Math.max(1, Math.floor(Number(row.querySelector('.part-qty').value) || 1)),
  })).filter((part) => part.code);
}

function calculateVisitParts() {
  let total = 0;
  $$('#visitPartsRows .part-entry').forEach((row) => {
    const item = state.inventory.find((entry) => entry.code === row.querySelector('.part-select').value);
    const qtyInput = row.querySelector('.part-qty');
    const qty = Math.max(1, Math.floor(Number(qtyInput.value) || 1));
    if (item) {
      qtyInput.max = item.qty;
      row.querySelector('.part-info').innerHTML = `<span>سعر الوحدة</span><b>${fmt(item.sell)} ج</b>`;
      total += item.sell * qty;
    } else {
      qtyInput.removeAttribute('max');
      row.querySelector('.part-info').textContent = 'اختر المنتج لعرض السعر';
    }
  });
  $('#visitPartsTotal').textContent = `${fmt(total)} ج`;
  return total;
}

function installInventoryMovementUI() {
  const addButton = $('#openInventoryAdd');
  const inventoryActions = document.createElement('div');
  inventoryActions.className = 'inventory-page-actions';
  addButton.before(inventoryActions);
  inventoryActions.appendChild(addButton);
  const movementButton = document.createElement('button');
  movementButton.id = 'openInventoryMovements';
  movementButton.type = 'button';
  movementButton.className = 'primary inventory-movement-button';
  movementButton.textContent = '↔ حركة المخزن';
  inventoryActions.appendChild(movementButton);

  const movementPage = document.createElement('div');
  movementPage.id = 'inventoryMovementPage';
  movementPage.className = 'inventory-movement-page hidden';
  movementPage.innerHTML = `<div class="section-bar movement-page-head"><button id="backToInventory" type="button" class="back-btn">→ رجوع للمخزن</button><div><h1>حركة المخزن</h1><p>ابحث بكود الحركة أو اختر الصادر والوارد.</p></div></div>
    <label class="movement-search">كود الحركة<input id="movementCodeSearch" placeholder="مثال: M1"></label>
    <div class="movement-mode-choice"><button type="button" data-movement-mode="وارد">↓ الوارد</button><button type="button" data-movement-mode="صادر">↑ الصادر</button></div>
    <div id="movementResults" class="movement-results"><div class="empty-state">اختر الوارد أو الصادر لعرض الحركات.</div></div>
  `;
  $('#inventory').appendChild(movementPage);
  movementButton.addEventListener('click', openInventoryMovements);
  $('#backToInventory').addEventListener('click', closeInventoryMovements);
  movementPage.querySelectorAll('[data-movement-mode]').forEach((button) => button.addEventListener('click', () => {
    state.movementMode = button.dataset.movementMode;
    renderInventoryMovements();
  }));
  $('#movementCodeSearch').addEventListener('input', renderInventoryMovements);
}

function installSuppliersUI() {
  const inventoryNav = $('.nav-btn[data-page="inventory"]');
  const supplierNav = document.createElement('button');
  supplierNav.type = 'button';
  supplierNav.className = 'nav-btn';
  supplierNav.dataset.page = 'suppliers';
  supplierNav.innerHTML = '<span>♧</span> الموردين';
  inventoryNav.after(supplierNav);

  const page = document.createElement('section');
  page.id = 'suppliers';
  page.className = 'page';
  page.innerHTML = `<div class="page-title"><div><span class="eyebrow">بيانات الموردين</span><h1>الموردين</h1><p>عرض البيانات المسجلة في شيت قسم الموردين.</p></div><div class="supplier-page-actions"><button id="addSupplierButton" class="primary" type="button">＋ إضافة مورد</button><button id="addSupplierTransactionButton" class="primary" type="button">＋ تسجيل حركة</button></div></div>
    <div class="panel"><div class="panel-head"><div><h2>كل الموردين</h2><p id="suppliersCount"></p></div></div><div id="suppliersTable"></div></div>`;
  $('main').appendChild(page);

  const supplierDialog = document.createElement('dialog');
  supplierDialog.id = 'supplierAddDialog';
  supplierDialog.innerHTML = `<form id="supplierAddForm"><button type="button" class="dialog-close">×</button><div class="dialog-title"><span>＋</span><div><h2>إضافة مورد جديد</h2><p>سيتم إنشاء كود المورد تلقائيًا.</p></div></div><div class="form-grid"><label>اسم المورد<input name="name" required></label><label>رقم التليفون<input name="phone" required></label><label>تاريخ التعاقد<input name="contractDate" type="date" required></label><label class="wide">ملاحظات<textarea name="notes" rows="2"></textarea></label></div><div class="form-actions"><button class="primary" type="submit">حفظ المورد</button></div></form>`;
  document.body.appendChild(supplierDialog);
  supplierDialog.querySelector('.dialog-close').addEventListener('click', () => supplierDialog.close());

  const transactionDialog = document.createElement('dialog');
  transactionDialog.id = 'supplierTransactionDialog';
  transactionDialog.innerHTML = `<form id="supplierTransactionForm"><button type="button" class="dialog-close">×</button><div class="dialog-title"><span>↔</span><div><h2>تسجيل حركة مورد</h2><p>سداد مستحقات أو توريد بضاعة.</p></div></div>
    <div class="form-grid"><label>اختر المورد (اختياري في التوريد)<input name="supplierCode" id="transactionSupplier" list="transactionSupplierOptions" placeholder="اكتب كود أو اسم المورد"><datalist id="transactionSupplierOptions"></datalist></label><label>نوع الحركة<select name="type" id="supplierTransactionType"><option value="">اختر نوع الحركة</option><option value="سداد مستحقات">سداد مستحقات</option><option value="توريد بضاعة">توريد بضاعة</option></select></label></div>
    <div id="paymentTransactionFields" class="form-grid hidden"><div id="supplierDueHint" class="supplier-due-hint wide">اختر المورد لعرض المبلغ المستحق.</div><label>مبلغ السداد<input name="amount" type="number" min="0.01" step="0.01"></label></div>
    <div id="supplyTransactionFields" class="form-grid hidden"><label class="wide">المنتج<input name="productCode" id="transactionProduct" list="transactionProductOptions" placeholder="اكتب كود أو اسم المنتج"><datalist id="transactionProductOptions"></datalist></label><label>الكمية<input name="qty" type="number" min="1"></label><label>سعر الشراء<input name="buy" type="number" min="0" step="0.01"></label><label>سعر البيع<input name="sell" type="number" min="0" step="0.01"></label><label>المبلغ المدفوع<input name="paid" type="number" min="0" step="0.01" value="0"></label><label>المبلغ المستحق<input id="transactionDue" readonly value="0"></label></div>
    <label class="transaction-notes">ملاحظات<textarea name="notes" rows="2"></textarea></label><div class="form-actions"><button class="primary" type="submit">حفظ الحركة</button></div></form>`;
  document.body.appendChild(transactionDialog);
  transactionDialog.querySelector('.dialog-close').addEventListener('click', () => transactionDialog.close());
  $('#addSupplierButton').addEventListener('click', openSupplierAdd);
  $('#addSupplierTransactionButton').addEventListener('click', openSupplierTransaction);
  $('#supplierTransactionType').addEventListener('change', updateSupplierTransactionFields);
  ['qty','buy','paid'].forEach((name) => $('#supplierTransactionForm').elements[name].addEventListener('input', calculateSupplierTransaction));
  $('#transactionSupplier').addEventListener('input', selectSupplierTransactionValue);
  $('#transactionProduct').addEventListener('input', selectProductTransactionValue);
  $('#supplierAddForm').addEventListener('submit', submitSupplierAdd);
  $('#supplierTransactionForm').addEventListener('submit', submitSupplierTransaction);
  const detailsDialog = document.createElement('dialog');
  detailsDialog.id = 'supplierDetailsDialog';
  detailsDialog.innerHTML = '<div class="supplier-details-dialog"><button type="button" class="dialog-close">×</button><div id="supplierDetailsContent"></div></div>';
  document.body.appendChild(detailsDialog);
  detailsDialog.querySelector('.dialog-close').addEventListener('click', () => detailsDialog.close());
}

function renderSuppliers() {
  $('#suppliersCount').textContent = `${fmt(state.suppliers.length)} مورد مسجل`;
  if (!state.suppliers.length) {
    $('#suppliersTable').innerHTML = '<div class="empty-state">لا توجد بيانات موردين مسجلة في شيت الموردين.</div>';
    return;
  }
  const rows = state.suppliers.map((supplier) => `<tr class="supplier-row ${supplier.name === state.selectedSupplierName ? 'supplier-highlight' : ''}" data-supplier-code="${esc(supplier.code)}" tabindex="0"><td><b>${esc(supplier.code || '—')}</b></td><td><b>${esc(supplier.name)}</b></td><td>${esc(supplier.phone || '—')}</td><td>${esc(supplier.contractDate || '—')}</td><td>${fmt(supplier.due)} ج</td><td>${fmt(supplier.paid)} ج</td><td>${esc(supplier.paymentDate || '—')}</td><td>${esc(supplier.notes || '—')}</td></tr>`).join('');
  $('#suppliersTable').innerHTML = `<div class="table-wrap"><table><thead><tr><th>كود المورد</th><th>اسم المورد</th><th>رقم التليفون</th><th>تاريخ التعاقد</th><th>المبلغ المطلوب</th><th>المبلغ المدفوع</th><th>تاريخ الدفع</th><th>ملاحظات</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  $$('#suppliersTable .supplier-row').forEach((row) => {
    row.addEventListener('click', () => showSupplierDetails(row.dataset.supplierCode));
    row.addEventListener('keydown', (event) => (event.key === 'Enter' || event.key === ' ') && showSupplierDetails(row.dataset.supplierCode));
  });
  $('.supplier-highlight')?.scrollIntoView({ block: 'center', behavior: 'smooth' });
}

function showSupplierDetails(code) {
  const supplier = state.suppliers.find((item) => item.code === code);
  if (!supplier) return;
  const accounts = state.accounts.filter((account) => account.supplierCode === supplier.code || account.supplierName === supplier.name);
  const detail = (label, value) => `<div><small>${label}</small><b>${esc(value)}</b></div>`;
  const rows = [...accounts].reverse().map((account) => `<tr><td><b>${esc(account.code)}</b></td><td>${esc(account.date || '—')}</td><td>${esc(account.type)}</td><td>${esc(account.productName || '—')}</td><td>${fmt(account.total)} ج</td><td>${fmt(account.paid)} ج</td><td>${fmt(account.due)} ج</td></tr>`).join('');
  $('#supplierDetailsContent').innerHTML = `<div class="dialog-title"><span>♧</span><div><h2>${esc(supplier.name)}</h2><p>${esc(supplier.code)}</p></div></div>
    <div class="supplier-summary-grid">${detail('كود المورد', supplier.code)}${detail('رقم التليفون', supplier.phone || '—')}${detail('تاريخ التعاقد', supplier.contractDate || 'غير مسجل')}${detail('إجمالي المدفوع', `${fmt(supplier.paid)} ج`)}${detail('إجمالي المستحق', `${fmt(supplier.due)} ج`)}${detail('آخر تاريخ دفع', supplier.paymentDate || '—')}</div>
    <div class="supplier-movements-title"><h3>الحركات المالية</h3><p>${fmt(accounts.length)} حركة مسجلة</p></div>
    ${rows ? `<div class="table-wrap"><table><thead><tr><th>الكود</th><th>التاريخ</th><th>النوع</th><th>المنتج</th><th>الإجمالي</th><th>المدفوع</th><th>المستحق</th></tr></thead><tbody>${rows}</tbody></table></div>` : '<div class="empty-state">لا توجد حركات مالية مسجلة لهذا المورد بعد.</div>'}`;
  $('#supplierDetailsDialog').showModal();
}

function supplierTransactionOptions() {
  $('#transactionSupplierOptions').innerHTML = state.suppliers.map((supplier) => `<option value="${esc(supplier.code)}" label="${esc(supplier.name)} · ${esc(supplier.phone || '')}">`).join('');
  $('#transactionProductOptions').innerHTML = state.inventory.map((item) => `<option value="${esc(item.code)}" label="${esc(item.name)} · ${esc(item.details || '')}">`).join('');
}

function openSupplierAdd() {
  $('#supplierAddForm').reset();
  $('#supplierAddForm').elements.contractDate.value = new Date().toISOString().slice(0, 10);
  $('#supplierAddDialog').showModal();
}

function openSupplierTransaction() {
  const form = $('#supplierTransactionForm');
  form.reset();
  resetSplitPayment(form);
  form.elements.paymentMethod.value = 'نقدي';
  supplierTransactionOptions();
  updateSupplierTransactionFields();
  calculateSupplierTransaction();
  $('#supplierTransactionDialog').showModal();
}

function updateSupplierTransactionFields() {
  const type = $('#supplierTransactionType').value;
  $('#paymentTransactionFields').classList.toggle('hidden', type !== 'سداد مستحقات');
  $('#supplyTransactionFields').classList.toggle('hidden', type !== 'توريد بضاعة');
  updateSupplierDueHint();
}

function selectSupplierTransactionValue() {
  const input = $('#transactionSupplier');
  const query = normalized(input.value);
  const supplier = state.suppliers.find((item) => normalized(item.code) === query || normalized(item.name) === query);
  if (supplier) input.value = supplier.code;
  updateSupplierDueHint();
}

function updateSupplierDueHint() {
  const input = $('#transactionSupplier');
  const query = normalized(input.value);
  const supplier = state.suppliers.find((item) => normalized(item.code) === query || normalized(item.name) === query);
  $('#supplierDueHint').innerHTML = supplier
    ? `<span>المبلغ المستحق حاليًا على ${esc(supplier.name)}</span><strong>${fmt(supplier.due)} ج</strong>`
    : '<span>اختر المورد لعرض المبلغ المستحق.</span>';
}

function selectProductTransactionValue() {
  const input = $('#transactionProduct');
  const query = normalized(input.value);
  const item = state.inventory.find((entry) => normalized(entry.code) === query || normalized(entry.name) === query);
  if (!item) return;
  input.value = item.code;
  const form = $('#supplierTransactionForm');
  form.elements.buy.value = item.buy;
  form.elements.sell.value = item.sell;
  if (!form.elements.supplierCode.value && item.supplier) {
    const supplier = state.suppliers.find((entry) => entry.name === item.supplier);
    if (supplier) form.elements.supplierCode.value = supplier.code;
  }
  calculateSupplierTransaction();
}

function calculateSupplierTransaction() {
  const form = $('#supplierTransactionForm');
  const total = (Number(form.elements.qty.value) || 0) * (Number(form.elements.buy.value) || 0);
  const due = Math.max(0, total - (Number(form.elements.paid.value) || 0));
  $('#transactionDue').value = due;
}

async function submitSupplierAdd(event) {
  event.preventDefault();
  try {
    const supplier = await request('/api/suppliers', { method: 'POST', body: JSON.stringify(Object.fromEntries(new FormData(event.currentTarget))) });
    $('#supplierAddDialog').close();
    await load(); renderSuppliers(); toast(`تم إضافة المورد ${supplier.name} — الكود ${supplier.code}`);
  } catch (error) { toast(error.message, true); }
}

async function submitSupplierTransaction(event) {
  event.preventDefault();
  try {
    const input = addPaymentsToInput(Object.fromEntries(new FormData(event.currentTarget)), event.currentTarget);
    const result = await request('/api/supplier-transactions', { method: 'POST', body: JSON.stringify(input) });
    $('#supplierTransactionDialog').close();
    await load(); renderSuppliers(); toast(`تم تسجيل الحركة ${result.account.code} بنجاح.`);
  } catch (error) { toast(error.message, true); }
}

function installEmployeesUI() {
  const supplierNav = $('.nav-btn[data-page="suppliers"]');
  const nav = document.createElement('button'); nav.type = 'button'; nav.className = 'nav-btn'; nav.dataset.page = 'employees'; nav.innerHTML = '<span>♟</span> الموظفين';
  supplierNav.after(nav);
  const page = document.createElement('section'); page.id = 'employees'; page.className = 'page';
  page.innerHTML = `<div class="page-title"><div><span class="eyebrow">فريق العمل</span><h1>الموظفين</h1><p>بيانات الموظفين ورواتبهم وأرصدتهم المالية.</p></div><div class="employee-page-actions"><button id="addEmployeeButton" class="primary" type="button">＋ إضافة موظف</button><button id="stoppedEmployeesButton" class="secondary" type="button">الموقوفين عن العمل</button></div></div><div id="activeEmployeesView" class="panel"><div class="panel-head"><div><h2>الموظفين العاملين</h2><p id="employeesCount"></p></div></div><div id="employeesTable"></div></div><div id="stoppedEmployeesView" class="panel hidden"><div class="panel-head"><div><h2>الموقوفين عن العمل</h2><p id="stoppedEmployeesCount"></p></div><button id="activeEmployeesBack" class="back-btn" type="button">→ رجوع للعاملين</button></div><div id="stoppedEmployeesTable"></div></div>`;
  $('main').appendChild(page);
  const addDialog = document.createElement('dialog'); addDialog.id = 'employeeAddDialog';
  addDialog.innerHTML = `<form id="employeeAddForm"><button type="button" class="dialog-close">×</button><div class="dialog-title"><span>♟</span><div><h2>إضافة موظف</h2><p>سيتم إنشاء كود E تلقائيًا.</p></div></div><div class="form-grid"><label>الاسم<input name="name" required></label><label>رقم التليفون<input name="phone" required></label><label>تاريخ التوظيف<input name="hireDate" type="date" required></label><label>التخصص<input name="specialty" required placeholder="مثال: ميكانيكا أو كهرباء"></label><label>المرتب الأسبوعي<input name="weeklySalary" type="number" min="0" step="0.01" required></label><label class="wide">ملاحظات<textarea name="notes" rows="2"></textarea></label></div><div class="form-actions"><button class="primary" type="submit">حفظ الموظف</button></div></form>`;
  document.body.appendChild(addDialog); addDialog.querySelector('.dialog-close').addEventListener('click', () => addDialog.close());
  const details = document.createElement('dialog'); details.id = 'employeeDetailsDialog';
  details.innerHTML = '<div class="employee-details-dialog"><button type="button" class="dialog-close">×</button><div id="employeeDetailsContent"></div></div>';
  document.body.appendChild(details); details.querySelector('.dialog-close').addEventListener('click', () => details.close());
  const statusDialog = document.createElement('dialog'); statusDialog.id = 'employeeStatusDialog';
  statusDialog.innerHTML = `<form id="employeeStatusForm"><button type="button" class="dialog-close">×</button><input type="hidden" name="employeeCode"><div class="dialog-title"><span>!</span><div><h2>تغيير حالة الموظف</h2><p id="employeeStatusName"></p></div></div><div class="form-grid"><label>الحالة<select name="status" required><option value="موقوف مؤقتًا">موقوف مؤقتًا</option><option value="منتهي الخدمة">منتهي الخدمة</option></select></label><label>تاريخ الإيقاف<input name="stopDate" type="date" required></label><label class="wide">سبب الإيقاف أو إنهاء الخدمة<textarea name="reason" rows="3" required></textarea></label></div><div class="form-actions"><button class="danger-button" type="submit">حفظ ونقل للموقوفين</button></div></form>`;
  document.body.appendChild(statusDialog); statusDialog.querySelector('.dialog-close').addEventListener('click', () => statusDialog.close());
  $('#addEmployeeButton').addEventListener('click', () => { $('#employeeAddForm').reset(); $('#employeeAddForm').elements.hireDate.value = new Date().toISOString().slice(0,10); addDialog.showModal(); });
  $('#employeeAddForm').addEventListener('submit', submitEmployee);
  $('#employeeStatusForm').addEventListener('submit', submitEmployeeStatus);
  $('#stoppedEmployeesButton').addEventListener('click', showStoppedEmployees);
  $('#activeEmployeesBack').addEventListener('click', showActiveEmployees);
}

function renderEmployees() {
  const active = state.employees.filter((employee) => !employee.status || employee.status === 'يعمل');
  $('#employeesCount').textContent = `${fmt(active.length)} موظف يعمل`;
  if (!active.length) { $('#employeesTable').innerHTML = '<div class="empty-state">لا يوجد موظفون عاملون حاليًا.</div>'; return; }
  const rows = active.map(employeeTableRow).join('');
  $('#employeesTable').innerHTML = `<div class="table-wrap"><table><thead><tr><th>الكود</th><th>الاسم</th><th>التليفون</th><th>تاريخ التوظيف</th><th>التخصص</th><th>المشاركات</th><th>المرتب الأسبوعي</th><th>عليه</th><th>له</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  $$('#employeesTable .employee-row').forEach((row) => row.addEventListener('click', () => showEmployeeDetails(row.dataset.employeeCode)));
}

function employeeTableRow(employee, stopped = false) {
  return `<tr class="employee-row ${stopped ? 'stopped-employee-row' : ''}" data-employee-code="${esc(employee.code)}"><td><b>${esc(employee.code)}</b></td><td><b>${esc(employee.name)}</b></td><td>${esc(employee.phone)}</td><td>${esc(employee.hireDate)}</td><td>${esc(employee.specialty)}</td><td>${fmt(employee.contributions)}</td><td>${fmt(employee.weeklySalary)} ج</td><td>${fmt(employee.debtOnEmployee)} ج</td><td>${fmt(employee.dueToEmployee)} ج</td>${stopped ? `<td>${esc(employee.status)}</td><td>${esc(employee.stopDate || '—')}</td><td>${esc(employee.stopReason || '—')}</td>` : ''}</tr>`;
}

function showStoppedEmployees() {
  $('#activeEmployeesView').classList.add('hidden'); $('#stoppedEmployeesView').classList.remove('hidden');
  const stopped = state.employees.filter((employee) => employee.status && employee.status !== 'يعمل');
  $('#stoppedEmployeesCount').textContent = `${fmt(stopped.length)} موظف`;
  $('#stoppedEmployeesTable').innerHTML = stopped.length ? `<div class="table-wrap"><table><thead><tr><th>الكود</th><th>الاسم</th><th>التليفون</th><th>تاريخ التوظيف</th><th>التخصص</th><th>المشاركات</th><th>المرتب</th><th>عليه</th><th>له</th><th>الحالة</th><th>تاريخ الإيقاف</th><th>السبب</th></tr></thead><tbody>${stopped.map((employee) => employeeTableRow(employee, true)).join('')}</tbody></table></div>` : '<div class="empty-state">لا يوجد موظفون موقوفون عن العمل.</div>';
  $$('#stoppedEmployeesTable .employee-row').forEach((row) => row.addEventListener('click', () => showEmployeeDetails(row.dataset.employeeCode)));
}

function showActiveEmployees() { $('#stoppedEmployeesView').classList.add('hidden'); $('#activeEmployeesView').classList.remove('hidden'); renderEmployees(); }

function showEmployeeDetails(code) {
  const employee = state.employees.find((item) => item.code === code); if (!employee) return;
  const accounts = state.accounts.filter((account) => account.employeeCode === employee.code);
  const rows = [...accounts].reverse().map((account) => `<tr class="account-row" data-account-code="${esc(account.code)}"><td>${esc(account.executionDate || account.date)}</td><td>${esc(account.type)}</td><td>${esc(account.description)}</td><td>${fmt(account.total)} ج</td><td>${esc(account.direction)}</td></tr>`).join('');
  const isActive = !employee.status || employee.status === 'يعمل';
  $('#employeeDetailsContent').innerHTML = `<div class="dialog-title"><span>♟</span><div><h2>${esc(employee.name)}</h2><p>${esc(employee.code)} · ${esc(employee.specialty)}</p></div></div><div class="employee-status-actions">${isActive ? '<button id="stopEmployeeButton" class="danger-button" type="button">إيقاف أو إنهاء الخدمة</button>' : '<button id="reactivateEmployeeButton" class="primary" type="button">إعادة الموظف للعمل</button>'}</div><div class="employee-summary-grid"><div><small>رقم التليفون</small><b>${esc(employee.phone)}</b></div><div><small>تاريخ التوظيف</small><b>${esc(employee.hireDate)}</b></div><div><small>الحالة</small><b>${esc(employee.status || 'يعمل')}</b></div><div><small>عدد المشاركات</small><b>${fmt(employee.contributions)}</b></div><div><small>المرتب الأسبوعي</small><b>${fmt(employee.weeklySalary)} ج</b></div><div class="employee-debt"><small>مبالغ عليه</small><b>${fmt(employee.debtOnEmployee)} ج</b></div><div class="employee-credit"><small>مبالغ له</small><b>${fmt(employee.dueToEmployee)} ج</b></div>${!isActive ? `<div><small>تاريخ الإيقاف</small><b>${esc(employee.stopDate || '—')}</b></div><div><small>سبب الإيقاف</small><b>${esc(employee.stopReason || '—')}</b></div>` : ''}</div><div class="supplier-movements-title"><h3>الحركات المالية</h3><p>${fmt(accounts.length)} حركة</p></div>${rows ? `<div class="table-wrap"><table><thead><tr><th>التاريخ</th><th>النوع</th><th>البيان</th><th>المبلغ</th><th>الاتجاه</th></tr></thead><tbody>${rows}</tbody></table></div>` : '<div class="empty-state">لا توجد حركات مالية للموظف.</div>'}`;
  $('#stopEmployeeButton')?.addEventListener('click', () => openEmployeeStatusDialog(employee));
  $('#reactivateEmployeeButton')?.addEventListener('click', () => reactivateEmployee(employee.code));
  $$('#employeeDetailsContent .account-row').forEach((row) => row.addEventListener('click', () => { $('#employeeDetailsDialog').close(); showAccountDetails(row.dataset.accountCode); }));
  $('#employeeDetailsDialog').showModal();
}

function openEmployeeStatusDialog(employee) {
  $('#employeeDetailsDialog').close(); const form = $('#employeeStatusForm'); form.reset(); form.elements.employeeCode.value = employee.code; form.elements.stopDate.value = new Date().toISOString().slice(0,10); $('#employeeStatusName').textContent = `${employee.name} — ${employee.code}`; $('#employeeStatusDialog').showModal();
}

async function submitEmployeeStatus(event) {
  event.preventDefault();
  try { const input = Object.fromEntries(new FormData(event.currentTarget)); await request('/api/employees/status', { method: 'POST', body: JSON.stringify(input) }); $('#employeeStatusDialog').close(); await load(); showStoppedEmployees(); toast('تم نقل الموظف إلى الموقوفين عن العمل.'); }
  catch (error) { toast(error.message, true); }
}

async function reactivateEmployee(code) {
  try { await request('/api/employees/status', { method: 'POST', body: JSON.stringify({ employeeCode: code, status: 'يعمل', reason: '' }) }); $('#employeeDetailsDialog').close(); await load(); showActiveEmployees(); toast('تمت إعادة الموظف للعمل.'); }
  catch (error) { toast(error.message, true); }
}

async function submitEmployee(event) {
  event.preventDefault();
  try { const employee = await request('/api/employees', { method: 'POST', body: JSON.stringify(Object.fromEntries(new FormData(event.currentTarget))) }); $('#employeeAddDialog').close(); await load(); renderEmployees(); toast(`تم إضافة الموظف ${employee.name} — ${employee.code}`); }
  catch (error) { toast(error.message, true); }
}

function installAccountsUI() {
  const supplierNav = $('.nav-btn[data-page="suppliers"]');
  const nav = document.createElement('button');
  nav.type = 'button'; nav.className = 'nav-btn'; nav.dataset.page = 'accounts'; nav.innerHTML = '<span>ج</span> الحسابات';
  supplierNav.after(nav);
  const page = document.createElement('section');
  page.id = 'accounts'; page.className = 'page';
  page.innerHTML = `<div class="page-title"><div><span class="eyebrow">الحركة المالية</span><h1>الحسابات</h1><p>سجل كل الحركات المالية الواردة والصادرة.</p></div><button id="addManualAccountButton" class="primary" type="button">＋ تسجيل حركة</button></div><div class="panel"><div class="panel-head"><div><h2>الحركات المالية</h2><p id="accountsCount"></p></div></div><div id="accountsTable"></div></div>`;
  $('main').appendChild(page);
  const accountActions = document.createElement('div');
  accountActions.className = 'account-page-actions';
  $('#addManualAccountButton').replaceWith(accountActions);
  accountActions.innerHTML = '<button id="addManualAccountButton" class="primary" type="button">＋ تسجيل حركة</button><button id="financialSummariesButton" class="primary" type="button">ملخصات مالية</button>';
  const summaries = document.createElement('div');
  summaries.id = 'financialSummariesView'; summaries.className = 'hidden';
  summaries.innerHTML = `<div class="financial-summary-head"><button id="financialSummaryBack" class="back-btn" type="button">→ رجوع للحسابات</button><div><h1>الملخصات المالية</h1><p id="financialRangeLabel">ملخص الوضع المالي الحالي.</p></div><label>التاريخ المرجعي<input id="financialSummaryDate" type="date"></label></div><div class="financial-period-tabs main-period-tabs"><button class="active" data-financial-period="daily">يومي</button><button data-financial-period="weekly">أسبوعي</button><button data-financial-period="monthly">شهري</button><button data-financial-period="quarterly">ربع سنوي</button><button data-financial-period="halfyear">نصف سنوي</button><button data-financial-period="yearly">سنوي</button></div><div class="financial-export-actions"><button id="dailyClosingButton" class="secondary" type="button">قفل اليومية</button><button id="accountStatementButton" class="secondary" type="button">كشف حساب</button><button id="supplierDebtsButton" class="secondary" type="button">ديون الموردين</button></div><div id="financialSummaryContent"></div><div id="dailyClosingView" class="hidden"></div><div id="accountStatementView" class="hidden"></div><div id="supplierDebtsView" class="hidden"></div>`;
  page.appendChild(summaries);
  const dialog = document.createElement('dialog');
  dialog.id = 'manualAccountDialog';
  dialog.innerHTML = `<form id="manualAccountForm"><button type="button" class="dialog-close">×</button><div class="dialog-title"><span>ج</span><div><h2>تسجيل حركة مالية</h2><p>سيتم إنشاء كود T تلقائيًا.</p></div></div><div class="form-grid"><label>نوع التسجيل<select name="operation" required><option value="">اختر</option><option value="سحب">سحب</option><option value="إيداع">إيداع</option><option value="دين">دين</option></select></label><label id="debtSideField" class="hidden">الدين على مَن؟<select name="debtSide"><option value="على العميل">مستحق على العميل</option><option value="على المركز">مستحق على المركز</option></select></label><label>تاريخ التنفيذ<input name="executionDate" type="date" required></label><label>نوع الحركة<input name="type" list="accountTypeOptions" placeholder="اختر أو اكتب نوعًا آخر" required><datalist id="accountTypeOptions"><option value="إذن صرف"><option value="مرتبات"><option value="مصروفات تشغيل"><option value="مشتريات نقدية"><option value="مصروفات إضافية"><option value="إيراد نقدي"><option value="سلفة"><option value="أجل"></datalist></label><label>المبلغ<input name="amount" type="number" min="0.01" step="0.01" required></label><label class="wide">ملاحظات<textarea name="notes" rows="3" placeholder="اكتب أخد إيه ولمين وأي تفاصيل أخرى"></textarea></label></div><div class="form-actions"><button class="primary" type="submit">حفظ الحركة</button></div></form>`;
  document.body.appendChild(dialog);
  const manualGrid = $('#manualAccountForm .form-grid');
  manualGrid.querySelector('.wide').insertAdjacentHTML('beforebegin', `<label class="wide">البيان<input name="description" required placeholder="اكتب المشتريات أو سبب المصروف بالتفصيل"></label><label id="manualCustomerField">العميل (اختياري)<input name="customerCode" list="manualCustomerOptions" placeholder="كود أو اسم العميل"><datalist id="manualCustomerOptions"></datalist></label><label id="manualSupplierField">المورد (اختياري)<input name="supplierCode" list="manualSupplierOptions" placeholder="كود أو اسم المورد"><datalist id="manualSupplierOptions"></datalist></label><label id="manualEmployeeField" class="hidden">الموظف<input name="employeeCode" list="manualEmployeeOptions" placeholder="كود أو اسم الموظف"><datalist id="manualEmployeeOptions"></datalist></label><label>طريقة الدفع<select name="paymentMethod" required>${paymentMethodOptions()}</select></label>`);
  $('#accountTypeOptions').insertAdjacentHTML('beforeend', '<option value="سلفة موظف"><option value="سداد سلفة موظف"><option value="مستحق لموظف"><option value="دفع مستحق موظف">');
  dialog.querySelector('.dialog-close').addEventListener('click', () => dialog.close());
  $('#addManualAccountButton').addEventListener('click', openManualAccount);
  $('#financialSummariesButton').addEventListener('click', openFinancialSummaries);
  $('#financialSummaryBack').addEventListener('click', closeFinancialSummaries);
  $('#financialSummaryDate').addEventListener('change', renderFinancialSummaries);
  $$('#financialSummariesView [data-financial-period]').forEach((button) => button.addEventListener('click', () => { financialPeriod = button.dataset.financialPeriod; dailyClosingVisible = false; $$('#financialSummariesView [data-financial-period]').forEach((item) => item.classList.toggle('active', item === button)); renderFinancialSummaries(); }));
  $('#accountStatementButton').addEventListener('click', () => { $('#accountStatementView').classList.toggle('hidden'); statementPeriod = 'daily'; renderFinancialSummaries(); });
  $('#supplierDebtsButton').addEventListener('click', () => { supplierDebtsVisible = !supplierDebtsVisible; $('#accountStatementView').classList.add('hidden'); renderFinancialSummaries(); });
  $('#dailyClosingButton').addEventListener('click', () => {
    dailyClosingVisible = !dailyClosingVisible;
    if (dailyClosingVisible) {
      financialPeriod = 'daily';
      $$('#financialSummariesView [data-financial-period]').forEach((item) => item.classList.toggle('active', item.dataset.financialPeriod === 'daily'));
    }
    renderFinancialSummaries();
  });
  $('#manualAccountForm').addEventListener('submit', submitManualAccount);
  $('#manualAccountForm').elements.operation.addEventListener('change', updateManualDebtFields);
  $('#manualAccountForm').elements.type.addEventListener('input', updateManualEmployeeFields);
  $('#manualAccountForm').elements.employeeCode.addEventListener('input', selectManualEmployee);
  const details = document.createElement('dialog');
  details.id = 'accountDetailsDialog';
  details.innerHTML = '<div class="account-details-dialog"><button type="button" class="dialog-close">×</button><div id="accountDetailsContent"></div></div>';
  document.body.appendChild(details);
  details.querySelector('.dialog-close').addEventListener('click', () => details.close());
  const debtDialog = document.createElement('dialog');
  debtDialog.id = 'debtDetailsDialog';
  debtDialog.innerHTML = '<div class="debt-details-dialog"><button type="button" class="dialog-close">×</button><div id="debtDetailsContent"></div></div>';
  document.body.appendChild(debtDialog);
  debtDialog.querySelector('.dialog-close').addEventListener('click', () => debtDialog.close());
}

function toggleDownloadChoices(show) {
  $('#downloadChoicesView').classList.toggle('hidden', !show);
  ['.financial-summary-head','.financial-export-actions','#financialSummaryContent','#dailyClosingView','#accountStatementView','#supplierDebtsView'].forEach((selector) => $('#financialSummariesView').querySelector(selector)?.classList.toggle('download-page-hidden', show));
}

function openFinancialSummaries() {
  $('#accounts > .page-title').classList.add('hidden');
  $('#accounts > .panel').classList.add('hidden');
  $('#financialSummariesView').classList.remove('hidden');
  $('#financialSummaryDate').value = new Date().toISOString().slice(0, 10);
  renderFinancialSummaries();
}

function closeFinancialSummaries() {
  $('#financialSummariesView').classList.add('hidden');
  $('#accounts > .page-title').classList.remove('hidden');
  $('#accounts > .panel').classList.remove('hidden');
}

function financialDateRange(period, referenceValue) {
  const reference = new Date(`${referenceValue}T12:00:00`);
  const start = new Date(reference); const end = new Date(reference);
  if (period === 'weekly') { const day = (reference.getDay() + 6) % 7; start.setDate(reference.getDate() - day); end.setDate(start.getDate() + 6); }
  if (period === 'monthly') { start.setDate(1); end.setMonth(start.getMonth() + 1, 0); }
  if (period === 'quarterly') { start.setMonth(Math.floor(reference.getMonth() / 3) * 3, 1); end.setMonth(start.getMonth() + 3, 0); }
  if (period === 'halfyear') { start.setMonth(reference.getMonth() < 6 ? 0 : 6, 1); end.setMonth(start.getMonth() + 6, 0); }
  if (period === 'yearly') { start.setMonth(0, 1); end.setMonth(12, 0); }
  const iso = (date) => { const adjusted = new Date(date.getTime() - date.getTimezoneOffset() * 60000); return adjusted.toISOString().slice(0, 10); };
  const titles = { daily: 'الملخص اليومي', weekly: 'الملخص الأسبوعي', monthly: 'الملخص الشهري', quarterly: 'الملخص ربع السنوي', halfyear: 'الملخص نصف السنوي', yearly: 'الملخص السنوي' };
  return { from: iso(start), to: iso(end), title: titles[period] };
}

function renderFinancialSummaries() {
  const date = $('#financialSummaryDate').value || new Date().toISOString().slice(0, 10);
  const range = financialDateRange(financialPeriod, date);
  const accounts = state.accounts.filter((account) => { const movementDate = account.executionDate || account.date; return movementDate >= range.from && movementDate <= range.to; });
  const incoming = accounts.filter((account) => account.direction === 'وارد').reduce((sum, account) => sum + Number(account.paid || 0), 0);
  const outgoing = accounts.filter((account) => account.direction === 'صادر').reduce((sum, account) => sum + Number(account.paid || 0), 0);
  const currentBalance = state.accounts.length ? Number(state.accounts[state.accounts.length - 1].balance || 0) : 0;
  const section = (title, headers, rows) => `<div class="financial-report-section"><h2>${title}</h2>${rows.length ? `<div class="table-wrap"><table><thead><tr>${headers.map((header) => `<th>${header}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '<div class="empty-state compact-empty">لا توجد بيانات في هذا اليوم.</div>'}</div>`;
  const withdrawals = accounts.filter((account) => account.direction === 'صادر' && account.code.startsWith('T'));
  const additionalExpenses = accounts.filter((account) => account.direction === 'صادر' && account.type === 'مصروفات إضافية');
  const visitsWithParts = state.visits.filter((visit) => visit.date >= range.from && visit.date <= range.to && visit.partsCodes);
  const supplierPayments = accounts.filter((account) => account.type === 'سداد مستحقات');
  const sales = accounts.filter((account) => account.type === 'إيراد زيارة صيانة');
  const query = `from=${encodeURIComponent(range.from)}&to=${encodeURIComponent(range.to)}&title=${encodeURIComponent(range.title)}`;
  const financialExcelUrl = `/api/financial-report?${query}&format=xlsx`;
  const financialPdfUrl = `/api/financial-report?${query}&format=pdf`;
  $('#financialRangeLabel').textContent = `${range.title}: من ${range.from} إلى ${range.to}`;
  const centerDebt = state.suppliers.reduce((sum, supplier) => sum + Number(supplier.due || 0), 0) + state.customers.reduce((sum, customer) => sum + Number(customer.dueFromCenter || 0), 0) + state.employees.reduce((sum, employee) => sum + Number(employee.dueToEmployee || 0), 0);
  const othersDebt = state.customers.reduce((sum, customer) => sum + Number(customer.dueFromCustomer || 0), 0) + state.employees.reduce((sum, employee) => sum + Number(employee.debtOnEmployee || 0), 0);
  $('#financialSummaryContent').innerHTML = `<div class="financial-summary-cards"><button type="button" class="treasury-card"><small>رصيد الخزنة الحالي</small><strong>${fmt(currentBalance)} ج</strong><span>وارد الفترة ${fmt(incoming)} ج · صادر الفترة ${fmt(outgoing)} ج</span></button><div><small>إجمالي الوارد</small><strong>${fmt(incoming)} ج</strong></div><div><small>إجمالي المصروف</small><strong>${fmt(outgoing)} ج</strong></div><div><small>صافي الدخل</small><strong>${fmt(incoming - outgoing)} ج</strong></div><button id="debtSummaryCard" type="button" class="treasury-card debt-summary-card"><small>المديونية</small><strong>${fmt(centerDebt + othersDebt)} ج</strong><span>على المركز ${fmt(centerDebt)} ج · على الغير ${fmt(othersDebt)} ج</span></button></div>`;
  $('#debtSummaryCard').addEventListener('click', () => openDebtDialog('center'));
  $('#dailyClosingButton').textContent = dailyClosingVisible ? 'إغلاق قفل اليومية' : 'قفل اليومية';
  $('#dailyClosingView').classList.toggle('hidden', !dailyClosingVisible);
  $('#dailyClosingView').innerHTML = `<div class="daily-closing-title"><div><h2>قفل اليومية — ${range.from}</h2><p>تفاصيل حركة اليوم كاملة.</p></div><details class="download-menu daily-closing-download"><summary>تنزيل</summary><div><a href="${financialPdfUrl}" target="_blank">تنزيل PDF</a><a href="${financialExcelUrl}" download>تنزيل Excel</a></div></details></div>
    ${section('حركات اليوم', ['التاريخ','الوقت','البيان','نوع الحركة','كود الحركة','وارد/صادر','المبلغ'], accounts.map((a) => `<tr><td>${esc(a.executionDate || a.date)}</td><td>${esc(a.time || '—')}</td><td>${esc(a.description || '—')}</td><td>${esc(a.type)}</td><td><b>${esc(a.code)}</b></td><td>${esc(a.direction || '—')}</td><td>${fmt(a.paid)} ج</td></tr>`))}
    ${section('المسحوبات والمصروفات التشغيلية', ['البيان','النوع','المبلغ','الدفع','الملاحظات'], withdrawals.map((a) => `<tr><td>${esc(a.description || '—')}</td><td>${esc(a.type)}</td><td>${fmt(a.paid)} ج</td><td>${esc(a.paymentMethod || '—')}</td><td>${esc(a.notes || '—')}</td></tr>`))}
    ${section('المصروفات الإضافية', ['البيان','المبلغ','طريقة الدفع','الملاحظات'], additionalExpenses.map((a) => `<tr><td>${esc(a.description || '—')}</td><td>${fmt(a.paid)} ج</td><td>${esc(a.paymentMethod || '—')}</td><td>${esc(a.notes || '—')}</td></tr>`))}
    ${section('قطع الغيار المستخدمة في الزيارات', ['العميل','كود الزيارة','القطع والكميات','قيمة القطع'], visitsWithParts.map((visit) => { const customer = state.customers.find((item) => item.code === visit.customerCode) || {}; return `<tr><td>${esc(customer.name || '—')}<br><small>${esc(visit.customerCode)}</small></td><td>${esc(visit.code)}</td><td>${esc(visit.partsCodes)}</td><td>${fmt(visit.partsTotal)} ج</td></tr>`; }))}
    ${section('دفعات الموردين', ['المورد','المبلغ','طريقة الدفع'], supplierPayments.map((a) => `<tr><td>${esc(a.supplierName)}<br><small>${esc(a.supplierCode)}</small></td><td>${fmt(a.paid)} ج</td><td>${esc(a.paymentMethod || '—')}</td></tr>`))}
    ${section('المبيعات', ['العميل','كود الزيارة','نوع العربية','المصنعية','المدفوع','طريقة الدفع'], sales.map((a) => { const visit = state.visits.find((v) => v.code === a.visitCode) || {}; const customer = state.customers.find((c) => c.code === a.customerCode) || {}; return `<tr><td>${esc(a.customerName || customer.name || '—')}<br><small>${esc(a.customerCode)}</small></td><td>${esc(a.visitCode)}</td><td>${esc(customer.carType || '—')}</td><td>${fmt(visit.labor)} ج</td><td>${fmt(a.paid)} ج</td><td>${esc(a.paymentMethod || '—')}</td></tr>`; }))}`;
  const statementRange = financialDateRange(statementPeriod, date);
  const statementAccounts = state.accounts.filter((account) => { const movementDate = account.executionDate || account.date; return movementDate >= statementRange.from && movementDate <= statementRange.to; });
  const statementRows = statementAccounts.map((account) => `<tr class="account-row" data-account-code="${esc(account.code)}"><td>${esc(account.executionDate || account.date)}</td><td>${esc(account.time || '—')}</td><td>${esc(account.description || '—')}</td><td>${esc(account.type)}</td><td><b>${esc(account.code)}</b></td><td>${esc(account.notes || '—')}</td><td>${account.direction === 'صادر' ? `${fmt(account.paid)} ج` : '—'}</td><td>${account.direction === 'وارد' ? `${fmt(account.paid)} ج` : '—'}</td><td><b>${fmt(account.balance)} ج</b></td></tr>`).join('');
  const statementQuery = `from=${encodeURIComponent(statementRange.from)}&to=${encodeURIComponent(statementRange.to)}&title=${encodeURIComponent(`كشف الحساب — ${statementRange.title}`)}`;
  $('#accountStatementView').innerHTML = `<div class="statement-head"><div><h2>كشف الحساب</h2><p>من ${statementRange.from} إلى ${statementRange.to}</p></div><details class="download-menu"><summary>تنزيل كشف الحساب</summary><div><a href="/api/financial-report?${statementQuery}&format=pdf" target="_blank">تنزيل PDF</a><a href="/api/financial-report?${statementQuery}&format=xlsx" download>تنزيل Excel</a></div></details></div><div class="financial-period-tabs statement-period-tabs"><button class="${statementPeriod === 'daily' ? 'active' : ''}" data-statement-period="daily">يومي</button><button class="${statementPeriod === 'weekly' ? 'active' : ''}" data-statement-period="weekly">أسبوعي</button><button class="${statementPeriod === 'monthly' ? 'active' : ''}" data-statement-period="monthly">شهري</button><button class="${statementPeriod === 'quarterly' ? 'active' : ''}" data-statement-period="quarterly">ربع سنوي</button><button class="${statementPeriod === 'halfyear' ? 'active' : ''}" data-statement-period="halfyear">نصف سنوي</button><button class="${statementPeriod === 'yearly' ? 'active' : ''}" data-statement-period="yearly">سنوي</button></div><div class="financial-report-section"><h2>${statementRange.title}</h2><div class="table-wrap"><table><thead><tr><th>التاريخ</th><th>الوقت</th><th>البيان</th><th>نوع الحركة</th><th>كود الحركة</th><th>الملاحظات</th><th>مدين</th><th>دائن</th><th>الرصيد بعد العملية</th></tr></thead><tbody>${statementRows || '<tr><td colspan="9">لا توجد حركات في هذه الفترة.</td></tr>'}</tbody></table></div></div>`;
  const supplierDebtRows = state.suppliers.filter((supplier) => Number(supplier.paid || 0) + Number(supplier.due || 0) > 0).sort((first, second) => Number(second.due || 0) - Number(first.due || 0));
  const supplierDebtTotals = supplierDebtRows.reduce((totals, supplier) => ({ paid: totals.paid + Number(supplier.paid || 0), due: totals.due + Number(supplier.due || 0) }), { paid: 0, due: 0 });
  $('#supplierDebtsButton').textContent = supplierDebtsVisible ? 'إغلاق ديون الموردين' : 'ديون الموردين';
  $('#supplierDebtsView').classList.toggle('hidden', !supplierDebtsVisible);
  $('#supplierDebtsView').innerHTML = `<div class="statement-head supplier-debts-head"><div><h2>ديون الموردين</h2><p>إجمالي ما تم دفعه والمتبقي لكل مورد.</p></div><details class="download-menu"><summary>تنزيل</summary><div><a href="/api/supplier-debts?format=pdf" download>تنزيل PDF</a><a href="/api/supplier-debts?format=xlsx" download>تنزيل Excel</a></div></details></div><div class="financial-report-section"><div class="table-wrap"><table><thead><tr><th>كود المورد</th><th>اسم المورد</th><th>إجمالي المبلغ</th><th>المبلغ المدفوع</th><th>المبلغ المتبقي</th></tr></thead><tbody>${supplierDebtRows.map((supplier) => `<tr class="supplier-debt-row" data-supplier-code="${esc(supplier.code)}"><td><b>${esc(supplier.code || '—')}</b></td><td>${esc(supplier.name || '—')}</td><td>${fmt(Number(supplier.paid || 0) + Number(supplier.due || 0))} ج</td><td>${fmt(supplier.paid)} ج</td><td><b>${fmt(supplier.due)} ج</b></td></tr>`).join('') || '<tr><td colspan="5">لا توجد ديون مسجلة على الموردين.</td></tr>'}<tr class="supplier-debt-total"><td></td><td><b>الإجمالي</b></td><td><b>${fmt(supplierDebtTotals.paid + supplierDebtTotals.due)} ج</b></td><td><b>${fmt(supplierDebtTotals.paid)} ج</b></td><td><b>${fmt(supplierDebtTotals.due)} ج</b></td></tr></tbody></table></div></div>`;
  $$('#accountStatementView [data-statement-period]').forEach((button) => button.addEventListener('click', () => { statementPeriod = button.dataset.statementPeriod; renderFinancialSummaries(); }));
  $$('#accountStatementView .account-row').forEach((row) => row.addEventListener('click', () => showAccountDetails(row.dataset.accountCode)));
  $$('#supplierDebtsView .supplier-debt-row').forEach((row) => row.addEventListener('click', () => { const supplier = state.suppliers.find((item) => item.code === row.dataset.supplierCode); state.selectedSupplierName = supplier?.name || ''; go('suppliers'); showSupplierDetails(row.dataset.supplierCode); }));
}

function openDebtDialog(mode = 'center') {
  const centerRows = [
    ...state.suppliers.filter((supplier) => Number(supplier.due) > 0).map((supplier) => ({ kind: 'supplier', code: supplier.code, statement: supplier.name, amount: supplier.due, notes: supplier.notes || 'مستحقات مورد' })),
    ...state.customers.filter((customer) => Number(customer.dueFromCenter) > 0).map((customer) => ({ kind: 'customer', code: customer.code, statement: customer.name, amount: customer.dueFromCenter, notes: 'مستحق على المركز للعميل' })),
    ...state.employees.filter((employee) => Number(employee.dueToEmployee) > 0).map((employee) => ({ kind: 'employee', code: employee.code, statement: employee.name, amount: employee.dueToEmployee, notes: 'مستحق للموظف' })),
  ];
  const othersRows = [...state.customers.filter((customer) => Number(customer.dueFromCustomer) > 0).map((customer) => ({ kind: 'customer', code: customer.code, statement: customer.name, amount: customer.dueFromCustomer, notes: 'مستحق على العميل' })), ...state.employees.filter((employee) => Number(employee.debtOnEmployee) > 0).map((employee) => ({ kind: 'employee', code: employee.code, statement: employee.name, amount: employee.debtOnEmployee, notes: 'سلفة على الموظف' }))];
  const rows = mode === 'center' ? centerRows : othersRows;
  $('#debtDetailsContent').innerHTML = `<div class="dialog-title"><span>ج</span><div><h2>المديونية</h2><p>عرض المبالغ المستحقة وربطها بصاحبها.</p></div></div><div class="debt-mode-tabs"><button class="${mode === 'center' ? 'active' : ''}" data-debt-mode="center">مديونية على المركز</button><button class="${mode === 'others' ? 'active' : ''}" data-debt-mode="others">مديونية على الغير</button></div><div class="debt-total">الإجمالي <strong>${fmt(rows.reduce((sum, row) => sum + Number(row.amount), 0))} ج</strong></div>${rows.length ? `<div class="table-wrap"><table><thead><tr><th>البيان</th><th>الكود</th><th>المبلغ</th><th>ملاحظات</th></tr></thead><tbody>${rows.map((row) => `<tr class="debt-person-row" data-kind="${row.kind}" data-code="${esc(row.code)}"><td><b>${esc(row.statement)}</b></td><td>${esc(row.code)}</td><td>${fmt(row.amount)} ج</td><td>${esc(row.notes)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="empty-state">لا توجد مديونيات مسجلة في هذا القسم.</div>'}`;
  $$('#debtDetailsContent [data-debt-mode]').forEach((button) => button.addEventListener('click', () => openDebtDialog(button.dataset.debtMode)));
  $$('#debtDetailsContent .debt-person-row').forEach((row) => row.addEventListener('click', () => {
    $('#debtDetailsDialog').close();
    if (row.dataset.kind === 'supplier') { const supplier = state.suppliers.find((item) => item.code === row.dataset.code); state.selectedSupplierName = supplier?.name || ''; go('suppliers'); }
    else if (row.dataset.kind === 'employee') { go('employees'); showEmployeeDetails(row.dataset.code); }
    else { go('clients'); openCustomerRecord(row.dataset.code); }
  }));
  if (!$('#debtDetailsDialog').open) $('#debtDetailsDialog').showModal();
}

function renderAccounts() {
  $('#accountsCount').textContent = `${fmt(state.accounts.length)} حركة مسجلة`;
  if (!state.accounts.length) { $('#accountsTable').innerHTML = '<div class="empty-state">لا توجد حركات مالية مسجلة بعد.</div>'; return; }
  const rows = [...state.accounts].reverse().map((account) => `<tr class="account-row" data-account-code="${esc(account.code)}" tabindex="0"><td><b>${esc(account.code)}</b></td><td>${esc(account.description || '—')}</td><td>${esc(account.executionDate || account.date)}</td><td>${esc(account.time || '—')}</td><td><span class="account-direction ${account.direction === 'وارد' ? 'in' : 'out'}">${esc(account.direction || '—')}</span></td><td>${esc(account.type)}</td><td>${esc(account.customerCode || '—')}<br><small>${esc(account.customerName || '')}</small></td><td>${esc(account.supplierCode || '—')}<br><small>${esc(account.supplierName || '')}</small></td><td class="money-value">${fmt(account.paid)} ج</td><td>${esc(account.paymentMethod || '—')}</td><td class="money-value">${fmt(account.balance)} ج</td></tr>`).join('');
  $('#accountsTable').innerHTML = `<div class="table-wrap"><table><thead><tr><th>كود الحركة</th><th>البيان</th><th>تاريخ التنفيذ</th><th>الوقت</th><th>وارد/صادر</th><th>النوع</th><th>العميل</th><th>المورد</th><th>المبلغ الفعلي</th><th>طريقة الدفع</th><th>رصيد المركز</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  $$('#accountsTable .account-row').forEach((row) => {
    row.addEventListener('click', () => showAccountDetails(row.dataset.accountCode));
    row.addEventListener('keydown', (event) => (event.key === 'Enter' || event.key === ' ') && showAccountDetails(row.dataset.accountCode));
  });
}

function openManualAccount() {
  $('#manualAccountForm').reset();
  resetSplitPayment($('#manualAccountForm'));
  $('#manualAccountForm').elements.executionDate.value = new Date().toISOString().slice(0, 10);
  $('#manualAccountForm').elements.paymentMethod.value = 'نقدي';
  $('#manualCustomerOptions').innerHTML = state.customers.map((customer) => `<option value="${esc(customer.code)}" label="${esc(customer.name)}">`).join('');
  $('#manualSupplierOptions').innerHTML = state.suppliers.map((supplier) => `<option value="${esc(supplier.code)}" label="${esc(supplier.name)}">`).join('');
  $('#manualEmployeeOptions').innerHTML = state.employees.map((employee) => `<option value="${esc(employee.code)}" label="${esc(employee.name)} · ${esc(employee.specialty)}">`).join('');
  updateManualDebtFields();
  updateManualEmployeeFields();
  $('#manualAccountDialog').showModal();
}

function updateManualDebtFields() {
  const form = $('#manualAccountForm');
  const isDebt = form.elements.operation.value === 'دين';
  $('#debtSideField').classList.toggle('hidden', !isDebt);
  form.querySelector('.split-payment-editor').classList.toggle('hidden', isDebt);
  form.elements.customerCode.required = isDebt;
  if (isDebt && !form.elements.type.value) form.elements.type.value = 'سلفة';
}

function updateManualEmployeeFields() {
  const form = $('#manualAccountForm');
  const type = form.elements.type.value;
  const employeeMovement = ['مرتبات', 'سلفة', 'سلفة موظف', 'سداد سلفة موظف', 'مستحق لموظف', 'دفع مستحق موظف'].includes(type);
  $('#manualEmployeeField').classList.toggle('hidden', !employeeMovement);
  $('#manualCustomerField').classList.toggle('hidden', employeeMovement);
  $('#manualSupplierField').classList.toggle('hidden', employeeMovement);
  form.elements.employeeCode.required = employeeMovement;
  if (!employeeMovement) { form.elements.employeeCode.value = ''; return; }
  form.elements.customerCode.value = '';
  form.elements.supplierCode.value = '';
  if (type === 'مرتبات' || type === 'سلفة' || type === 'سلفة موظف' || type === 'دفع مستحق موظف') form.elements.operation.value = 'سحب';
  if (type === 'سداد سلفة موظف' || type === 'مستحق لموظف') form.elements.operation.value = 'إيداع';
  $('#debtSideField').classList.add('hidden');
  form.querySelector('.split-payment-editor').classList.toggle('hidden', type === 'مستحق لموظف');
}

function selectManualEmployee() {
  const form = $('#manualAccountForm'); const query = normalized(form.elements.employeeCode.value);
  const employee = state.employees.find((item) => normalized(item.code) === query || normalized(item.name) === query);
  if (!employee) return;
  form.elements.employeeCode.value = employee.code;
  if (form.elements.type.value === 'مرتبات') {
    form.elements.amount.value = employee.weeklySalary || '';
    const paymentAmount = form.querySelector('.split-payment-amount');
    if (paymentAmount && !paymentAmount.value) paymentAmount.value = employee.weeklySalary || '';
    if (!form.elements.description.value) form.elements.description.value = `مرتب أسبوعي — ${employee.name}`;
  }
}

async function submitManualAccount(event) {
  event.preventDefault();
  try {
    let input = Object.fromEntries(new FormData(event.currentTarget));
    if (input.type === 'سلفة' && input.employeeCode) input.type = 'سلفة موظف';
    if (input.operation === 'دين' || input.type === 'مستحق لموظف') {
      input.payments = [{ method: 'آجل', amount: Number(input.amount) || 0 }];
      input.paymentMethod = 'آجل'; input.paid = 0;
    } else input = addPaymentsToInput(input, event.currentTarget);
    const account = await request('/api/accounts', { method: 'POST', body: JSON.stringify(input) });
    $('#manualAccountDialog').close(); await load(); renderAccounts(); toast(`تم تسجيل الحركة ${account.code}.`);
  } catch (error) { toast(error.message, true); }
}

function showAccountDetails(code) {
  const account = state.accounts.find((item) => item.code === code);
  if (!account) return;
  const detail = (label, value) => `<div><small>${label}</small><b>${esc(value)}</b></div>`;
  const customerLink = account.customerCode ? `<button class="financial-related-link" data-customer-code="${esc(account.customerCode)}">الذهاب للعميل: ${esc(account.customerName || account.customerCode)}</button>` : '';
  const supplierLink = account.supplierCode ? `<button class="financial-related-link" data-supplier-code="${esc(account.supplierCode)}">الذهاب للمورد: ${esc(account.supplierName || account.supplierCode)}</button>` : '';
  const visitLink = account.visitCode ? `<button class="financial-related-link" data-visit-code="${esc(account.visitCode)}">فتح الزيارة ${esc(account.visitCode)}</button>` : '';
  const productLink = account.productCode ? `<button class="financial-related-link" data-product-code="${esc(account.productCode)}">فتح المنتج: ${esc(account.productName || account.productCode)}</button>` : '';
  const employeeLink = account.employeeCode ? `<button class="financial-related-link" data-employee-code="${esc(account.employeeCode)}">فتح الموظف: ${esc(account.employeeName || account.employeeCode)}</button>` : '';
  $('#accountDetailsContent').innerHTML = `<div class="dialog-title"><span>ج</span><div><h2>تفاصيل الحركة ${esc(account.code)}</h2><p>${esc(account.type)}</p></div></div><div class="financial-related-actions">${customerLink}${supplierLink}${visitLink}${productLink}${employeeLink}</div><div class="account-detail-grid">${detail('تاريخ التسجيل', account.date || '—')}${detail('تاريخ التنفيذ', account.executionDate || account.date || '—')}${detail('وقت التسجيل', account.time || '—')}${detail('اتجاه الحركة', account.direction || '—')}${detail('نوع الحركة', account.type || '—')}${detail('البيان', account.description || '—')}${detail('المبلغ', `${fmt(account.total)} ج`)}${detail('المبلغ المدفوع', `${fmt(account.paid)} ج`)}${detail('المبلغ المستحق', `${fmt(account.due)} ج`)}${detail('طريقة الدفع', account.paymentMethod || '—')}${detail('تفاصيل طرق الدفع', account.paymentDetails || account.paymentMethod || '—')}${detail('رصيد المركز بعد الحركة', `${fmt(account.balance)} ج`)}${detail('المورد', account.supplierName || '—')}${detail('كود المورد', account.supplierCode || '—')}${detail('العميل', account.customerName || '—')}${detail('كود العميل', account.customerCode || '—')}${detail('الموظف', account.employeeName || '—')}${detail('كود الموظف', account.employeeCode || '—')}${detail('كود الزيارة', account.visitCode || '—')}${detail('المنتج', account.productName || '—')}${detail('كود المنتج', account.productCode || '—')}${detail('الكمية', fmt(account.qty))}${detail('الملاحظات', account.notes || '—')}</div>`;
  $('#accountDetailsContent [data-customer-code]')?.addEventListener('click', (event) => { $('#accountDetailsDialog').close(); go('clients'); openCustomerRecord(event.currentTarget.dataset.customerCode); });
  $('#accountDetailsContent [data-supplier-code]')?.addEventListener('click', (event) => { const supplier = state.suppliers.find((item) => item.code === event.currentTarget.dataset.supplierCode); $('#accountDetailsDialog').close(); state.selectedSupplierName = supplier?.name || ''; go('suppliers'); });
  $('#accountDetailsContent [data-visit-code]')?.addEventListener('click', (event) => { $('#accountDetailsDialog').close(); showVisit(event.currentTarget.dataset.visitCode); });
  $('#accountDetailsContent [data-product-code]')?.addEventListener('click', (event) => { $('#accountDetailsDialog').close(); go('inventory'); openProductDetails(event.currentTarget.dataset.productCode); });
  $('#accountDetailsContent [data-employee-code]')?.addEventListener('click', (event) => { $('#accountDetailsDialog').close(); go('employees'); showEmployeeDetails(event.currentTarget.dataset.employeeCode); });
  $('#accountDetailsDialog').showModal();
}

function openInventoryMovements() {
  state.movementMode = '';
  $('#movementCodeSearch').value = '';
  $('#movementResults').innerHTML = '<div class="empty-state">اختر الوارد أو الصادر لعرض الحركات.</div>';
  $('#inventory').querySelectorAll(':scope > .page-title, :scope > .panel').forEach((element) => element.classList.add('hidden'));
  $('#inventoryMovementPage').classList.remove('hidden');
}

function closeInventoryMovements() {
  $('#inventoryMovementPage').classList.add('hidden');
  $('#inventory').querySelectorAll(':scope > .page-title, :scope > .panel').forEach((element) => element.classList.remove('hidden'));
}

function renderInventoryMovements() {
  const query = normalized($('#movementCodeSearch').value);
  const movements = state.movements.filter((movement) => {
    const modeMatches = query ? true : movement.type === state.movementMode;
    return modeMatches && (!query || normalized(movement.code).includes(query));
  });
  const results = $('#movementResults');
  if (!movements.length) {
    results.innerHTML = '<div class="empty-state">لا توجد حركات مطابقة.</div>';
    return;
  }
  const rows = movements.map((movement) => `<tr><td><b>${esc(movement.code)}</b></td><td><span class="movement-type ${movement.type === 'صادر' ? 'out' : 'in'}">${esc(movement.type)}</span></td><td>${esc(movement.date || '—')}</td><td><b>${esc(movement.productName)}</b><br><small>${esc(movement.productCode)}</small></td><td>${fmt(movement.qty)}</td><td>${fmt(movement.unitPrice)} ج</td><td>${fmt(movement.total)} ج</td><td>${esc(movement.visitCode || '—')}</td></tr>`).join('');
  results.innerHTML = `<div class="table-wrap"><table><thead><tr><th>كود الحركة</th><th>النوع</th><th>التاريخ</th><th>المنتج</th><th>الكمية</th><th>سعر الوحدة</th><th>الإجمالي</th><th>كود الزيارة</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

function showVisit(code) {
  const visit = state.visits.find((item) => item.code === code);
  if (!visit) return;
  const customer = state.customers.find((item) => item.code === visit.customerCode) || {};
  const box = (label, value, className = '') => `<div class="visit-detail-box ${className}"><small>${label}</small><b>${esc(value)}</b></div>`;
  const movementBox = visit.stockMovementCode
    ? `<div class="visit-detail-box"><small>كود حركة المخزن</small><button type="button" class="movement-code-link" data-movement-code="${esc(visit.stockMovementCode)}">${esc(visit.stockMovementCode)} ←</button></div>`
    : box('كود حركة المخزن', 'لا توجد حركة مخزن');
  $('#visitDetails').innerHTML = `<div class="dialog-title visit-dialog-title"><span>✓</span><div><h2>تفاصيل زيارة الصيانة</h2><p>${esc(customer.name || '—')} — ${esc(visit.code)}</p></div><button id="printVisitInvoice" type="button" class="primary visit-print-button">تنزيل الفاتورة PDF</button></div><div class="visit-detail-layout">
    ${box('اسم العميل', customer.name || '—')}${box('التاريخ', visit.date || '—')}
    ${box('كود العميل', visit.customerCode || '—')}${box('كود الزيارة', visit.code || '—')}
    ${movementBox}${box('نوع الحركة', visit.stockMovementCode ? 'صادر' : '—')}
    ${box('نوع العربية', customer.carType || '—')}${box('لوحة العربية', visit.plate || customer.plate || '—')}
    ${box('قراءة العداد', `${fmt(visit.mileage)} كم`)}${box('نوع الصيانة', visit.serviceType || '—')}
    ${box('اسم الفني', visit.technician || '—')}${box('المصنعية', `${fmt(visit.labor)} ج`)}
    <div class="visit-special-box parts-box"><small>قطع الغيار المستخدمة</small><b>${esc(visit.partsCodes || 'لا توجد قطع غيار')}</b></div>
    <div class="visit-special-box money-box"><small>المبالغ</small><div><span>قطع الغيار <b>${fmt(visit.partsTotal)} ج</b></span><span>المصنعية <b>${fmt(visit.labor)} ج</b></span><span>الإجمالي <b>${fmt(visit.total)} ج</b></span><span>المدفوع <b>${fmt(visit.paid)} ج</b></span><span>المتبقي <b>${fmt(visit.due)} ج</b></span><span>طرق الدفع <b>${esc(visit.paymentDetails || visit.paymentMethod || '—')}</b></span></div></div>
    ${box('الملاحظات', visit.notes || '—', 'notes-box')}
  </div>`;
  $('#visitDetails .movement-code-link')?.addEventListener('click', (event) => openMovementByCode(event.currentTarget.dataset.movementCode));
  $('#printVisitInvoice').addEventListener('click', () => {
    window.location.href = `/api/visit-invoice?code=${encodeURIComponent(visit.code)}`;
  });
  $('#detailsDialog').showModal();
}

function openMovementByCode(code) {
  $('#detailsDialog').close();
  go('inventory');
  openInventoryMovements();
  $$('.nav-btn').forEach((button) => button.classList.toggle('active', button.dataset.page === 'inventory'));
  $('#movementCodeSearch').value = code;
  renderInventoryMovements();
}

installInventoryMovementUI();
installSuppliersUI();
installEmployeesUI();
installAccountsUI();
installProductDetailsUI();
upgradeSearchableChoices();
installGlobalCodeSearch();

$$('.nav-btn').forEach((button) => button.addEventListener('click', () => go(button.dataset.page)));
$$('[data-go]').forEach((button) => button.addEventListener('click', () => go(button.dataset.go)));
$$('.choice-card').forEach((button) => button.addEventListener('click', () => showMode(button.dataset.mode)));
$$('.back-btn').forEach((button) => button.addEventListener('click', resetCustomerView));
$$('[data-close]').forEach((button) => button.addEventListener('click', () => button.closest('dialog').close()));
$('#menuBtn').addEventListener('click', () => $('.sidebar').classList.toggle('open'));
$('#searchBtn').addEventListener('click', searchCustomer);
$('#addVisitPart').addEventListener('click', addVisitPartRow);
$('#directorySearchBtn').addEventListener('click', searchClientDirectory);
$('#directoryCustomerCode').addEventListener('keydown', (event) => event.key === 'Enter' && searchClientDirectory());
$('#directoryVisitCode').addEventListener('keydown', (event) => event.key === 'Enter' && searchClientDirectory());
$('#clientRecordBack').addEventListener('click', renderCustomerDirectory);
$('#openInventoryAdd').addEventListener('click', openInventoryDialog);
$$('[data-inventory-mode]').forEach((button) => button.addEventListener('click', () => chooseInventoryMode(button.dataset.inventoryMode)));
$('#inventoryModeBack').addEventListener('click', resetInventoryDialog);
$('#oldProductSearch').addEventListener('input', renderProductSuggestions);
$('#oldProductSearch').addEventListener('focus', renderProductSuggestions);
$('#oldProductSearch').addEventListener('click', renderProductSuggestions);
['qty','buy','sell','paid'].forEach((name) => $('#inventoryForm').elements[name].addEventListener('input', calculateInventoryPurchase));
const scrollToLowStock = () => $('#lowStockPanel').scrollIntoView({ behavior: 'smooth', block: 'center' });
$('#lowStockSummary').addEventListener('click', scrollToLowStock);
$('#lowStockSummary').addEventListener('keydown', (event) => {
  if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); scrollToLowStock(); }
});
$('#searchCode').addEventListener('keydown', (event) => event.key === 'Enter' && searchCustomer());
$('#searchPlate').addEventListener('keydown', (event) => event.key === 'Enter' && searchCustomer());

$('#customerForm').addEventListener('submit', async (event) => {
  event.preventDefault();
  const form = event.currentTarget;
  try {
    const input = Object.fromEntries(new FormData(form));
    const customer = await request('/api/customers', { method: 'POST', body: JSON.stringify(input) });
    state.customers.push(customer);
    state.selectedCustomer = customer;
    form.reset();
    renderDashboard();
    toast(`تم حفظ العميل في main data 2.xlsx — الكود ${customer.code}`);
    $('#newCustomerView').classList.add('hidden');
    $('#oldCustomerView').classList.remove('hidden');
    $('#searchCode').value = customer.code;
    $('#searchPlate').value = '';
    searchCustomer();
    openVisitDialog();
  } catch (error) { toast(error.message, true); }
});

$('#visitForm').addEventListener('submit', async (event) => {
  event.preventDefault();
  try {
    const input = addPaymentsToInput(Object.fromEntries(new FormData(event.currentTarget)), event.currentTarget);
    input.customerCode = state.selectedCustomer.code;
    if (Number(input.labor) <= 0) throw new Error('يجب كتابة قيمة المصنعية قبل حفظ الزيارة.');
    input.parts = selectedVisitParts();
    const uniqueCodes = new Set(input.parts.map((part) => part.code));
    if (uniqueCodes.size !== input.parts.length) throw new Error('لا تضف نفس قطعة الغيار أكثر من مرة؛ عدّل الكمية في نفس الصف.');
    const visit = await request('/api/visits', { method: 'POST', body: JSON.stringify(input) });
    state.visits.push(visit);
    $('#visitDialog').close();
    renderDashboard();
    if ($('#clients').classList.contains('active')) openCustomerRecord(state.selectedCustomer.code);
    else searchCustomer();
    toast(`تم حفظ الزيارة في main data 2.xlsx — الكود ${visit.code}`);
  } catch (error) { toast(error.message, true); }
});

$('#inventoryForm').addEventListener('submit', async (event) => {
  event.preventDefault();
  try {
    const form = event.currentTarget;
    const input = addPaymentsToInput(Object.fromEntries(new FormData(form)), form);
    if (input.mode === 'old' && !input.code) throw new Error('اختر منتجًا قديمًا من نتائج البحث.');
    const result = await request('/api/inventory', { method: 'POST', body: JSON.stringify(input) });
    $('#inventoryDialog').close();
    await load();
    renderInventory();
    toast(`تم حفظ ${result.item.name} في المخزن بنجاح.`);
  } catch (error) { toast(error.message, true); }
});

$('#today').textContent = new Intl.DateTimeFormat('ar-EG', { dateStyle: 'full' }).format(new Date());
load();

