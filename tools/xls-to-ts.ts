console.log('Starte Skript:', __filename, 'CWD:', process.cwd());

import fs from 'node:fs';
import path from 'node:path';
import xlsx from 'xlsx';

function slugify(value: string): string {
  return value
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

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

const records = clean.map(r => {
  const id = `pr-${slugify(r.sl_city)}-${slugify(r.sl_store)}`;
  return {
    id,
    name: r.sl_store,
    address: `${r.sl_address}, ${r.sl_zip} ${r.sl_city}`,
    coordinates: { lat: Number(r.sl_latitude), lng: Number(r.sl_longitude) },
    tel: (r.sl_phone ?? '').replace(/\s+/g, ' ').trim(),
    __city: r.sl_city
  };
});

// remove helper property before writing files
const plainRecords = records.map(({ __city, ...rest }) => rest);

const grouped = records.reduce((acc: Record<string, any[]>, r) => {
  const city = r.__city;
  const { __city, ...item } = r;
  if (!acc[city]) acc[city] = [];
  acc[city].push(item);
  return acc;
}, {} as Record<string, any[]>);

const outFlat = `import type { PolizeiRevier } from '../stores/appStore';\n\n` +
  `export const polizeiReviere: PolizeiRevier[] = ${JSON.stringify(plainRecords, null, 2)};\n`;

const outGroupedTs = `import type { PolizeiRevier } from '../stores/appStore';\n\n` +
  `export const groupedPolizeiReviere: Record<string, PolizeiRevier[]> = ${JSON.stringify(grouped, null, 2)};\n`;

fs.mkdirSync(path.resolve('src/data'), { recursive: true });
fs.writeFileSync('src/data/polizeiReviere.ts', outFlat, 'utf8');
fs.writeFileSync('src/data/groupedPolizeiReviere.ts', outGroupedTs, 'utf8');
fs.writeFileSync('src/data/groupedPolizeiReviere.json', JSON.stringify(grouped, null, 2), 'utf8');

console.log(`✅  ${plainRecords.length} Dienststellen exportiert → src/data`);

