const LEGACY_BANK_GENERAL_TEMPLATE_PATH = 'FORM connect/bank-general.docx';

const LEGACY_BANK_GENERAL_PLACEHOLDERS = Object.freeze([
  'doc_number',
  'issue_date_th',
  'bank_name',
  'bank_address',
  'case_headline',
  'accounts_no',
  'accounts_name',
  'statement_date_from',
  'statement_date_to',
]);

function asText(value) {
  return String(value == null ? '' : value).trim();
}

function firstFilled(...values) {
  for (const value of values) {
    const text = asText(value);
    if (text) return text;
  }
  return '';
}

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function getBankAccounts(payload) {
  const canonical = asArray(payload.bank_accounts);
  if (canonical.length) return canonical;
  const legacyNos = asArray(payload.bankAccounts);
  if (!legacyNos.length) return [];
  const legacyNames = asArray(payload.bankAccountNames);
  return legacyNos.map((accountNo, index) => ({
    account_no: accountNo,
    account_name: legacyNames[index] || '',
  }));
}

function getFirstAccount(payload) {
  const accounts = getBankAccounts(payload);
  return accounts.length ? accounts[0] || {} : {};
}

function getRequestMode(payload) {
  return firstFilled(
    payload.bank_request_mode,
    payload.bankRequestMode,
    payload.bankSubType
  ).toLowerCase();
}

function getCompatibilityReport(payload = {}) {
  const blockers = [];
  const warnings = [];
  const accounts = getBankAccounts(payload);
  const mode = getRequestMode(payload) || 'fullaccount';

  if (mode && mode !== 'fullaccount') {
    blockers.push(`template นี้รองรับเฉพาะ fullaccount แต่ได้รับ ${mode || 'unknown'}`);
  }
  if (!firstFilled(payload.doc_number, payload.docNumber)) {
    blockers.push('missing doc_number');
  }
  if (!firstFilled(payload.issue_date_th_full, payload.issue_date_th, payload.issueDateTh)) {
    blockers.push('missing issue_date_th');
  }
  if (!firstFilled(payload.bank_name_th, payload.bank_name, payload.provider_name_th, payload.bankCode)) {
    blockers.push('missing bank_name');
  }
  if (!firstFilled(payload.recipient_address, payload.bank_address)) {
    blockers.push('missing bank_address');
  }
  if (!firstFilled(payload.case_headline, payload.case_detail, payload.caseDetail)) {
    blockers.push('missing case_headline');
  }
  if (!firstFilled(payload.period_start_th, payload.statement_date_from, payload.bankDateStart)) {
    blockers.push('missing statement_date_from');
  }
  if (!firstFilled(payload.period_end_th, payload.statement_date_to, payload.bankDateEnd)) {
    blockers.push('missing statement_date_to');
  }
  if (!accounts.length) {
    blockers.push('missing bank account for accounts_no/accounts_name');
  }
  if (accounts.length > 1) {
    warnings.push('template นี้มีช่องบัญชีเดียว ระบบจะใช้เฉพาะบัญชีแรก');
  }
  if (firstFilled(payload.promptpay_id, payload.bankPromptPay)) {
    warnings.push('template นี้ไม่มี placeholder สำหรับ promptpay โดยตรง');
  }
  if (firstFilled(payload.slip_amount, payload.xxxAmount)) {
    warnings.push('template นี้ยังไม่มี placeholder สำหรับโหมดข้อมูลติด XXX');
  }
  warnings.push('ผู้ลงนามและท้ายเอกสารยังเป็นข้อความ hardcoded ใน DOCX ฉบับนี้');

  return {
    compatible: blockers.length === 0,
    blockers,
    warnings,
    mode: mode || 'fullaccount',
    placeholder_count: LEGACY_BANK_GENERAL_PLACEHOLDERS.length,
  };
}

function buildLegacyPlaceholderMap(payload = {}) {
  const firstAccount = getFirstAccount(payload);

  return {
    doc_number: firstFilled(payload.doc_number, payload.docNumber),
    issue_date_th: firstFilled(
      payload.issue_date_th_full,
      payload.issue_date_th,
      payload.issueDateTh
    ),
    bank_name: firstFilled(
      payload.bank_name_th,
      payload.bank_name,
      payload.provider_name_th,
      payload.bankCode
    ),
    bank_address: firstFilled(payload.recipient_address, payload.bank_address),
    case_headline: firstFilled(
      payload.case_headline,
      payload.case_detail,
      payload.caseDetail
    ),
    accounts_no: firstFilled(
      firstAccount.account_no,
      firstAccount.accountNo,
      payload.account_no,
      payload.bankAccountNo
    ),
    accounts_name: firstFilled(
      firstAccount.account_name,
      firstAccount.accountName,
      payload.account_name,
      payload.bankAccountName
    ),
    statement_date_from: firstFilled(
      payload.period_start_th,
      payload.statement_date_from,
      payload.bankDateStart
    ),
    statement_date_to: firstFilled(
      payload.period_end_th,
      payload.statement_date_to,
      payload.bankDateEnd
    ),
  };
}

function resolveLegacyBankGeneralTemplate(payload = {}) {
  const report = getCompatibilityReport(payload);
  return {
    template_key: 'bank-general-legacy',
    template_path: LEGACY_BANK_GENERAL_TEMPLATE_PATH,
    placeholders: LEGACY_BANK_GENERAL_PLACEHOLDERS.slice(),
    placeholder_map: buildLegacyPlaceholderMap(payload),
    report,
  };
}

export {
  LEGACY_BANK_GENERAL_PLACEHOLDERS,
  LEGACY_BANK_GENERAL_TEMPLATE_PATH,
  buildLegacyPlaceholderMap,
  getCompatibilityReport,
  resolveLegacyBankGeneralTemplate,
};
