console.log('Starte Skript:', __filename, 'CWD:', process.cwd());

import fs from 'node:fs';
import path from 'node:path';
import xlsx from 'xlsx';

interface RawRow {
  sl_store: string;
  sl_address: string;
  sl_city: string;
  sl_zip: string|number;
  sl_latitude: number;
  sl_longitude: number;
  sl_phone?: string;
  Polizeipräsidium?: string;
  sl_tags?: string;
}

const wb = xlsx.readFile(
  path.resolve(__dirname, '..', 'polizeiReviere.xlsx')
);
const raw = xlsx.utils.sheet_to_json<RawRow>(wb.Sheets['inet_store_locator']);

const clean = raw
  .filter(r =>
    !/polizeiposten/i.test(r.sl_store) &&
    !/kriminalinspektion/i.test(r.sl_store) &&
    !/wasserschutz/i.test(r.sl_store) &&
    !(r.sl_tags ?? '').toLowerCase().includes('polizeipräsidium') &&
    !(r.Polizeipräsidium ?? '').toLowerCase().includes('hfpol')
  )
  .sort((a, b) =>
    (a.Polizeipräsidium ?? '').localeCompare(b.Polizeipräsidium ?? '') ||
    a.sl_store.localeCompare(b.sl_store)
  );

const records = clean.map((r, i) => ({
  id: `revier-${i.toString().padStart(4, '0')}`,
  name: r.sl_store,
  address: `${r.sl_address}, ${r.sl_zip} ${r.sl_city}`,
  coordinates: { lat: r.sl_latitude, lng: r.sl_longitude },
  tel: (r.sl_phone ?? '').replace(/\s+/g, ' ').trim()
}));

const out = `import type { PolizeiRevier } from '../stores/appStore';\n\n` +
            `export const polizeiReviere: PolizeiRevier[] = ${JSON.stringify(records, null, 2)};\n`;

fs.mkdirSync(path.resolve('src/data'), { recursive: true });
fs.writeFileSync('src/data/polizeiReviere.ts', out, 'utf8');

console.log(`✅  ${records.length} Dienststellen exportiert → src/data/polizeiReviere.ts`);
