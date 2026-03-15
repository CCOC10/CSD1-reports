# DOCX Template Placeholders

สเปกนี้อิงจากฟอร์มปัจจุบันใน [CSD1 connect.html](/Users/mrchp/Library/Mobile%20Documents/com~apple~CloudDocs/WEB%20dev/CSD1%20report/CSD1%20connect.html) และตั้งชื่อ placeholder ให้ใช้ชุดเดียวกันทุก template

## Syntax

- ค่าเดี่ยว: `{{doc_number}}`
- block/loop: `{{#phone_numbers}}...{{/phone_numbers}}`
- ค่า optional: ให้เว้น placeholder ได้ ถ้าไม่มีค่า generator ใส่ `''`

## Template Files

1. `bank-general.docx`
2. `bank-atm.docx`
3. `cdr-imei.docx`
4. `imei-lookup.docx`
5. `base-ais.docx`
6. `ip-ais.docx`
7. `truemoney.docx`
8. `tax.docx`

## Common Placeholders

ใช้ในทุกไฟล์

| Placeholder | ความหมาย | source/current logic |
| --- | --- | --- |
| `{{template_key}}` | key ของ template | derived |
| `{{file_name}}` | ชื่อไฟล์เต็ม เช่น `ชป.15-SCB-ตช 0026.21-542.pdf` | derived |
| `{{file_category}}` | เช่น `SCB`, `AIS`, `BASE AIS`, `TAX` | `getRequestFileCategory()` |
| `{{request_id}}` | รหัสระบบ `REQ-...` | `S.requestId` |
| `{{unit}}` | ชป. เช่น `ชป.15` | `S.unit` |
| `{{doc_number}}` | เลขท้ายหนังสือ เช่น `542` | `S.docNumber` |
| `{{book_no}}` | `ตช 0026.21/{{doc_number}}` | derived |
| `{{case_detail}}` | รายละเอียดเคสเต็ม | `S.caseDetail` |
| `{{case_headline}}` | headline ที่เอาไปแทนในย่อหน้าเหตุ | current `getCaseHeadline()` |
| `{{case_type}}` | ปกติ `ความอาญา` | current print renderer |
| `{{complainant}}` | ผู้กล่าวหา | per template |
| `{{defendant}}` | ผู้ต้องหา | per template |
| `{{legal_basis}}` | ฐานอำนาจตามกฎหมาย | per template |
| `{{urgent_text}}` | ปกติ `ด่วนที่สุด` | current print renderer |
| `{{warning_text}}` | ข้อความคำเตือนเต็ม | `PRINT_WARNING_TEXT` |
| `{{agency_name}}` | `สำนักงานตำรวจแห่งชาติ` | `PRINT_AGENCY` |
| `{{issue_place}}` | สถานที่ออกหมาย | `PRINT_ISSUE_PLACE` |
| `{{issue_date_iso}}` | วันที่ออกหมาย `YYYY-MM-DD` | today/derived |
| `{{issue_day}}` | วันแบบไทย | derived |
| `{{issue_month_th}}` | เดือนภาษาไทย | derived |
| `{{issue_year_be}}` | พ.ศ. | derived |
| `{{issue_date_th_full}}` | เช่น `13 เดือน มีนาคม พุทธศักราช 2569` | derived |
| `{{recipient_label}}` | `หมายถึง` หรือ `หมายมายัง` | per template |
| `{{recipient_name}}` | ชื่อผู้รับเอกสาร | per template/meta |
| `{{recipient_address}}` | ที่อยู่ผู้รับ | per template/meta |
| `{{return_address}}` | ที่อยู่ส่งกลับ | `PRINT_RETURN_ADDRESS` |
| `{{signer_rank}}` | บรรทัดก่อนลายเซ็น เช่น `ว่าที่ พันตำรวจเอก` | current print renderer |
| `{{signer_name}}` | ชื่อผู้ลงนาม | current `getPrintSigner()` |
| `{{signer_position}}` | ตำแหน่งผู้ลงนาม | current `getPrintSigner()` |
| `{{contact_name}}` | ชื่อผู้ประสานงานท้ายเอกสาร | current `getPrintContact()` |
| `{{contact_position}}` | ตำแหน่งผู้ประสานงานท้ายเอกสาร | current `getPrintContact()` |
| `{{contact_phone}}` | เบอร์ผู้ประสานงานท้ายเอกสาร | current `getPrintContact()` |
| `{{contact_email}}` | email ท้ายเอกสาร | current `getPrintContact()` |
| `{{commander_name}}` | หัวหน้า ชป. | `S.commander` หรือ leader sheet |
| `{{commander_position}}` | ตำแหน่งหัวหน้า ชป. | current logic/leader sheet |

## Shared Derived Placeholders

ใช้ซ้ำหลาย template

| Placeholder | ความหมาย |
| --- | --- |
| `{{period_start_iso}}` | วันเริ่มต้น |
| `{{period_end_iso}}` | วันสิ้นสุด |
| `{{period_start_th}}` | วันเริ่มต้นภาษาไทย |
| `{{period_end_th}}` | วันสิ้นสุดภาษาไทย |
| `{{period_text}}` | ข้อความห้วงเวลาแบบพร้อมใช้ |
| `{{provider_code}}` | เช่น `AIS`, `SCB`, `KBANK` |
| `{{provider_name_th}}` | ชื่อเต็มภาษาไทย |
| `{{provider_short_name}}` | short name สำหรับย่อหน้าเหตุ |

## `bank-general.docx`

ใช้กับ statement ธนาคาร และรองรับ 3 โหมด: `fullaccount`, `promptpay`, `xxx`

หมายเหตุ: ไฟล์ [bank-general.docx](/Users/mrchp/Library/Mobile%20Documents/com~apple~CloudDocs/WEB%20dev/CSD1%20report/FORM%20connect/bank-general.docx) ที่มีอยู่ตอนนี้ยังใช้ legacy placeholder set ไม่ตรงสเปกกลางทั้งหมด ให้ใช้ alias mapper ที่ [bank-general.alias-mapper.js](/Users/mrchp/Library/Mobile%20Documents/com~apple~CloudDocs/WEB%20dev/CSD1%20report/FORM%20connect/bank-general.alias-mapper.js) ในช่วงทดสอบ/เปลี่ยนผ่าน

### placeholders

- `{{bank_code}}`
- `{{bank_name_th}}`
- `{{bank_short_name}}`
- `{{bank_request_mode}}` = `fullaccount` \| `promptpay` \| `xxx`
- `{{bank_request_mode_label}}` = `เลขบัญชีเต็ม` \| `พร้อมเพย์` \| `ข้อมูลติด XXX`
- `{{period_start_iso}}`
- `{{period_end_iso}}`
- `{{period_start_th}}`
- `{{period_end_th}}`
- `{{period_text}}`
- `{{promptpay_id}}`
- `{{promptpay_name}}`
- `{{slip_amount}}`
- `{{slip_date_iso}}`
- `{{slip_date_th}}`
- `{{slip_time}}`
- `{{slip_datetime_th}}`
- `{{slip_account_name}}`

### loop blocks

```
{{#bank_accounts}}
{{account_index}}
{{account_no}}
{{account_name}}
{{/bank_accounts}}
```

### notes

- `{{recipient_name}}` และ `{{recipient_address}}` ดึงจาก bank metadata
- ถ้าเป็น `promptpay` ให้ `bank_accounts` เป็น array ว่าง
- ถ้าเป็น `xxx` ให้ใช้ `slip_*` เป็นหลัก

## `bank-atm.docx`

### placeholders

- `{{bank_code}}`
- `{{bank_name_th}}`
- `{{bank_short_name}}`
- `{{atm_account_no}}`
- `{{atm_account_name}}`
- `{{atm_date_iso}}`
- `{{atm_date_th}}`
- `{{atm_time}}`
- `{{atm_datetime_th}}`
- `{{atm_location}}`
- `{{atm_terminal_id}}`

## `cdr-imei.docx`

ใช้กับ `AIS / TRUE / DTAC / OTHER` แบบ CDR ปกติ

### placeholders

- `{{network_code}}`
- `{{network_name_th}}`
- `{{network_short_name}}`
- `{{period_start_iso}}`
- `{{period_end_iso}}`
- `{{period_start_th}}`
- `{{period_end_th}}`
- `{{period_text}}`

### loop blocks

```
{{#phone_numbers}}
{{phone_index}}
{{phone_no}}
{{/phone_numbers}}
```

## `imei-lookup.docx`

ใช้กับ `TRUE / DTAC` กรณีลอย IMEI

### placeholders

- `{{network_code}}`
- `{{network_name_th}}`
- `{{network_short_name}}`
- `{{imei_no}}`
- `{{request_scope}}` = `imei_lookup`

## `base-ais.docx`

ใช้กับ `AIS > BASE`

### placeholders

- `{{network_code}}` = `AIS`
- `{{network_name_th}}`
- `{{request_scope}}` = `current_base_location`

### loop blocks

```
{{#phone_numbers}}
{{phone_index}}
{{phone_no}}
{{/phone_numbers}}
```

## `ip-ais.docx`

ใช้กับ `AIS > IP Address`

### placeholders

- `{{network_code}}` = `AIS`
- `{{network_name_th}}`
- `{{request_scope}}` = `ip_usage_lookup`

### loop blocks

```
{{#ip_entries}}
{{ip_index}}
{{ip}}
{{date_iso}}
{{date_th}}
{{time}}
{{datetime_th}}
{{port}}
{{/ip_entries}}
```

## `truemoney.docx`

### placeholders

- `{{wallet_id}}`
- `{{wallet_id_type}}` = `phone` \| `citizen_id`
- `{{wallet_name}}`
- `{{period_start_iso}}`
- `{{period_end_iso}}`
- `{{period_start_th}}`
- `{{period_end_th}}`
- `{{period_text}}`

## `tax.docx`

### placeholders

- `{{taxpayer_type}}` = `individual` \| `corporate`
- `{{taxpayer_type_label}}` = `บุคคลธรรมดา` \| `นิติบุคคล`
- `{{taxpayer_id_label}}` = `เลขประจำตัวประชาชน/ผู้เสียภาษี` \| `เลขประจำตัวนิติบุคคล`
- `{{taxpayer_id}}`
- `{{taxpayer_name}}`
- `{{tax_year_start}}`
- `{{tax_year_end}}`
- `{{tax_year_text}}`

### loop blocks

```
{{#tax_years}}
{{year}}
{{/tax_years}}
```

## Recommended Generator Output Shape

โครง JSON กลางที่แนะนำให้ป้อนเข้า template:

```json
{
  "template_key": "bank-general",
  "file_name": "ชป.15-SCB-ตช 0026.21-542.pdf",
  "unit": "ชป.15",
  "doc_number": "542",
  "book_no": "ตช 0026.21/542",
  "case_detail": "รายละเอียดเคส",
  "issue_date_iso": "2026-03-13",
  "issue_day": "13",
  "issue_month_th": "มีนาคม",
  "issue_year_be": "2569",
  "recipient_name": "กรรมการผู้จัดการใหญ่ ธนาคารไทยพาณิชย์ จำกัด (มหาชน) สำนักงานใหญ่",
  "recipient_address": "9 ถนนรัชดาภิเษก แขวงจตุจักร เขตจตุจักร กรุงเทพมหานคร 10900",
  "signer_name": "พ.ต.อ. ...",
  "contact_name": "พ.ต.ต. ...",
  "bank_accounts": [
    { "account_index": 1, "account_no": "1234567890", "account_name": "นายตัวอย่าง" }
  ]
}
```

## Fields Missing Or Should Be Added

ฟอร์มปัจจุบันยังควรเพิ่ม field เหล่านี้ ถ้าจะให้ DOCX ใช้งานจริงได้ลื่น:

- `{{slip_partial_account_name}}`
  ใช้กรณี `ข้อมูลติด XXX` ถ้าต้องใส่ชื่อบัญชีบางส่วนแยกจากชื่อบัญชีเต็ม
- `{{recipient_title_override}}`
  กรณีบาง template ต้องเปลี่ยนผู้รับเป็น wording พิเศษ
- `{{custom_body_intro}}`
  ใช้แก้ wording รายคดีโดยไม่ต้องแตก template ย่อยเพิ่ม
- `{{custom_legal_basis}}`
  ใช้ override ฐานกฎหมายรายเคส

## Mapping Rule

- ชื่อ placeholder ใน DOCX ใช้ lowercase + underscore ทั้งหมด
- ค่าที่เป็นวันที่ให้มีทั้ง `*_iso`, `*_th`, และ `*_text` ถ้าต้องนำไปใช้หลายตำแหน่ง
- รายการซ้ำทั้งหมดใช้ loop block ไม่ใช้ placeholder แบบ `account_1`, `account_2`
- metadata ผู้รับ/ที่อยู่ของธนาคารและค่ายมือถือไม่ควร hardcode ใน DOCX ให้ generator เติมให้
