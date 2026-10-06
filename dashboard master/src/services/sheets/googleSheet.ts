import { env } from '../../config/env';
import { Certificate } from '../../db/certificates';
import { formatEditionNumber } from '../certificate/templateData';
import { logger } from '../../utils/logger';

/**
 * Mirrors certificates into a Google Sheet via an Apps Script web app (see
 * google-sheets/Code.gs and README section 13). The Sheet is a read-only
 * copy for the team - the local database stays the source of truth, so a
 * failed sync is logged and never blocks issuing a certificate. "Sync to
 * Google Sheet" on the dashboard re-sends everything to catch up.
 */

const STATUS_LABELS: Record<string, string> = {
  generated: 'Ready to download',
  processing: 'Processing',
  pending: 'Pending',
  failed: 'Failed',
  emailed: 'Emailed',
};

/** One Sheet row, keyed by certificateId (the Apps Script updates in place by it). */
export interface SheetRow {
  certificateId: number;
  issuedAt: string;
  orderNumber: string;
  customerName: string;
  customerEmail: string;
  design: string;
  designCode: string;
  sku: string;
  edition: string;
  status: string;
  updatedAt: string;
}

export function isSheetSyncConfigured(): boolean {
  return Boolean(env.GOOGLE_SHEETS_WEBHOOK_URL);
}

export function certificateToSheetRow(cert: Certificate): SheetRow {
  return {
    certificateId: cert.id,
    issuedAt: new Date(cert.created_at).toISOString(),
    orderNumber: cert.order_number,
    customerName: [cert.customer_first_name, cert.customer_last_name].filter(Boolean).join(' '),
    customerEmail: cert.customer_email,
    design: cert.design_name,
    designCode: cert.design_code,
    sku: cert.sku ?? '',
    edition: formatEditionNumber(cert.certificate_number, cert.edition_total),
    status: STATUS_LABELS[cert.status] ?? cert.status,
    updatedAt: new Date(cert.updated_at).toISOString(),
  };
}

/** Sends rows to the Apps Script web app. Throws if it isn't configured or rejects them. */
export async function pushRowsToSheet(rows: SheetRow[]): Promise<{ written: number }> {
  if (!env.GOOGLE_SHEETS_WEBHOOK_URL) {
    throw new Error('GOOGLE_SHEETS_WEBHOOK_URL is not set in .env');
  }
  // Apps Script answers a POST with a 302 to the result; fetch follows it.
  const res = await fetch(env.GOOGLE_SHEETS_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ secret: env.GOOGLE_SHEETS_SECRET ?? '', rows }),
    signal: AbortSignal.timeout(30_000),
  });
  const text = await res.text();
  let json: { ok?: boolean; written?: number; error?: string };
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error(
      `Google Sheet returned HTTP ${res.status} and not JSON - check the web app URL and that it is deployed with access "Anyone".`
    );
  }
  if (!res.ok || !json.ok) {
    throw new Error(`Google Sheet rejected the update: ${json.error ?? `HTTP ${res.status}`}`);
  }
  return { written: json.written ?? rows.length };
}

/**
 * Fire-and-forget sync of one certificate after it changes. Never throws:
 * the certificate is already saved locally, so a Sheet outage only logs.
 */
export function syncCertificateToSheet(cert: Certificate): void {
  if (!isSheetSyncConfigured()) return;
  pushRowsToSheet([certificateToSheetRow(cert)]).catch((err) => {
    logger.warn('Google Sheet sync failed', {
      certificateId: cert.id,
      error: err instanceof Error ? err.message : String(err),
    });
  });
}
