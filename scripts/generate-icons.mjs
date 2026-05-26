// Generates solid-colour placeholder PNG icons for the Office Add-in manifest.
// Replace public/icons/*.png with your own branded artwork when ready.

import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(__dirname, "..", "public", "icons");
const sizes = [16, 32, 64, 80, 128];
const FORCE = process.argv.includes("--force");

// Brand colour from the React PM tool (--primary: #4f46e5).
const R = 0x4f;
const G = 0x46;
const B = 0xe5;

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let k = 0; k < 8; k++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, "ascii");
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function makePng(size) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const rowSize = 1 + size * 3;
  const raw = Buffer.alloc(rowSize * size);
  for (let y = 0; y < size; y++) {
    const off = y * rowSize;
    raw[off] = 0; // filter type None
    for (let x = 0; x < size; x++) {
      const p = off + 1 + x * 3;
      raw[p] = R;
      raw[p + 1] = G;
      raw[p + 2] = B;
    }
  }
  const idat = deflateSync(raw);

  return Buffer.concat([
    signature,
    chunk("IHDR", ihdr),
    chunk("IDAT", idat),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

mkdirSync(outDir, { recursive: true });
for (const s of sizes) {
  const path = resolve(outDir, `icon-${s}.png`);
  if (!FORCE && existsSync(path)) continue;
  writeFileSync(path, makePng(s));
  console.log(`Wrote ${path}`);
}
