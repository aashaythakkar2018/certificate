import './setupEnv';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { certificateToSheetRow } from '../src/services/sheets/googleSheet';
import { Certificate } from '../src/db/certificates';

const base: Certificate = {
  id: 7,
  order_number: '1042',
  customer_first_name: 'Asha',
  customer_last_name: 'Mehta',
  customer_email: 'asha@example.com',
  design_name: 'Echoes of Earth',
  design_code: 'EOE',
  certificate_number: 37,
  edition_total: 250,
  sku: 'RHY-EOE-GRS-037/250',
  certificate_url: null,
  storage_key: null,
  status: 'generated',
  attempt_count: 1,
  error_message: null,
  created_at: new Date('2026-10-01T10:00:00Z'),
  updated_at: new Date('2026-10-02T11:30:00Z'),
  sent_at: null,
};

test('maps a certificate to a Sheet row with dashboard wording', () => {
  assert.deepEqual(certificateToSheetRow(base), {
    certificateId: 7,
    issuedAt: '2026-10-01T10:00:00.000Z',
    orderNumber: '1042',
    customerName: 'Asha Mehta',
    customerEmail: 'asha@example.com',
    design: 'Echoes of Earth',
    designCode: 'EOE',
    sku: 'RHY-EOE-GRS-037/250',
    edition: '037/250',
    status: 'Ready to download',
    updatedAt: '2026-10-02T11:30:00.000Z',
  });
});

test('handles a missing last name and SKU', () => {
  const row = certificateToSheetRow({ ...base, customer_last_name: null, sku: null, status: 'failed' });
  assert.equal(row.customerName, 'Asha');
  assert.equal(row.sku, '');
  assert.equal(row.status, 'Failed');
});
