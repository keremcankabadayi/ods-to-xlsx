/*
 * Parola korumalı ODS dosyalarını tarayıcıda çözer ve şifresiz bir ODS (zip) döndürür.
 *
 * Desteklenen şifrelemeler:
 *  - LibreOffice 24.2+ "wholesome": Argon2id + AES-256-GCM, tüm paket tek "encrypted-package" girdisinde
 *  - ODF 1.2/1.3: PBKDF2 (HMAC-SHA1) + AES-256-CBC, her dosya ayrı şifreli
 * Desteklenmeyen: OpenOffice 2.x / ODF 1.1 dönemi Blowfish-CFB.
 *
 * Bağımlılıklar: XLSX (SheetJS, XLSX.CFB zip okuyup yazmak için), hashwasm.argon2id (yalnızca Argon2 için).
 */
(function (root) {
  'use strict';

  const subtle = root.crypto.subtle;

  class WrongPasswordError extends Error {
    constructor() { super('Parola yanlış'); this.name = 'WrongPasswordError'; }
  }

  function b64(s) {
    const bin = atob(s);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function attrs(s) {
    const out = {};
    for (const m of s.matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)) out[m[1]] = m[2];
    return out;
  }

  // Manifest basit ve LibreOffice tarafından üretildiği için regex ile ayrıştırmak yeterli
  // (DOMParser'a bağımlı olmamak Node'da test etmeyi de mümkün kılıyor).
  function parseManifest(xml) {
    const entries = [];
    const re = /<manifest:file-entry\b([^>]*?)(?:\/>|>([\s\S]*?)<\/manifest:file-entry>)/g;
    for (const m of xml.matchAll(re)) {
      const entry = { attrs: attrs(m[1]), raw: m[0], enc: null };
      const body = m[2] || '';
      if (body.includes('encryption-data')) {
        const tag = (name) => {
          const t = body.match(new RegExp('<manifest:' + name + '\\b([^>]*)>'));
          return t ? attrs(t[1]) : {};
        };
        entry.enc = {
          data: tag('encryption-data'),
          algorithm: tag('algorithm'),
          startKey: tag('start-key-generation'),
          derivation: tag('key-derivation'),
        };
      }
      entries.push(entry);
    }
    return entries;
  }

  // Zip'i central directory üzerinden okur. SheetJS'in zip okuyucusu, LibreOffice'in şifreli
  // girdilerde kullandığı imzasız "data descriptor"ları yanlış yorumladığı için kendi okuyucumuz var.
  async function readZip(input) {
    const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Zip dosyası okunamadı');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const utf8 = new TextDecoder();
    const files = new Map();
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Zip dizini bozuk');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const local = dv.getUint32(p + 42, true);
      const name = utf8.decode(bytes.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (name.endsWith('/')) continue;
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      const raw = bytes.subarray(start, start + csize);
      if (method === 0) files.set(name, raw);
      else if (method === 8) files.set(name, await inflateRaw(raw));
      else throw new Error('Desteklenmeyen zip sıkıştırması: ' + method);
    }
    return files;
  }

  function writeZip(files) {
    const zip = XLSX.CFB.utils.cfb_new();
    // ODF: "mimetype" ilk girdi olmalı
    const order = [...files.keys()].sort((a, b) => (a === 'mimetype' ? -1 : b === 'mimetype' ? 1 : 0));
    for (const p of order) XLSX.CFB.utils.cfb_add(zip, p, files.get(p));
    XLSX.CFB.utils.cfb_del(zip, '\u0001Sh33tJ5'); // SheetJS'in eklediği yer tutucu
    return new Uint8Array(XLSX.CFB.write(zip, { fileType: 'zip', type: 'array', compression: true }));
  }

  async function inflateRaw(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function tryInflate(bytes) {
    try { return await inflateRaw(bytes); } catch (e) { return bytes; }
  }

  async function digest(name, bytes) {
    return new Uint8Array(await subtle.digest(name, bytes));
  }

  function eq(a, b) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  async function startKey(enc, password) {
    const name = (enc.startKey['manifest:start-key-generation-name'] || 'sha1').toLowerCase();
    const pw = new TextEncoder().encode(password);
    if (name.endsWith('sha256')) return digest('SHA-256', pw);
    if (name.endsWith('sha1')) return digest('SHA-1', pw);
    throw new Error('Desteklenmeyen anahtar üretimi: ' + name);
  }

  async function deriveKey(enc, password) {
    const d = enc.derivation;
    const kdf = d['manifest:key-derivation-name'] || '';
    const size = parseInt(d['manifest:key-size'] || '16', 10);
    const salt = b64(d['manifest:salt']);
    const start = await startKey(enc, password);

    if (kdf === 'PBKDF2') {
      const base = await subtle.importKey('raw', start, 'PBKDF2', false, ['deriveBits']);
      const bits = await subtle.deriveBits(
        { name: 'PBKDF2', hash: 'SHA-1', salt, iterations: parseInt(d['manifest:iteration-count'], 10) },
        base, size * 8);
      return new Uint8Array(bits);
    }
    if (/argon2id$/.test(kdf)) {
      if (!root.hashwasm) throw new Error('Argon2 kütüphanesi yüklenemedi');
      return root.hashwasm.argon2id({
        password: start,
        salt,
        iterations: parseInt(d['loext:argon2-iterations'], 10),
        memorySize: parseInt(d['loext:argon2-memory'], 10),
        parallelism: parseInt(d['loext:argon2-lanes'], 10),
        hashLength: size,
        outputType: 'binary',
      });
    }
    throw new Error('Desteklenmeyen anahtar türetme: ' + kdf);
  }

  // AES-CBC: ODF, W3C xmlenc dolgusunu kullanır (son bayt = dolgu uzunluğu, diğerleri rastgele olabilir).
  // WebCrypto yalnızca PKCS#7 kabul ettiği için şifreli metne geçerli bir PKCS#7 bloğu ekleyip çözüyoruz,
  // ardından asıl dolguyu elle kaldırıyoruz.
  async function aesCbcDecrypt(keyBytes, iv, data) {
    if (data.length === 0 || data.length % 16) throw new WrongPasswordError();
    const key = await subtle.importKey('raw', keyBytes, 'AES-CBC', false, ['encrypt', 'decrypt']);
    const lastBlock = data.subarray(data.length - 16);
    const pad = new Uint8Array(16).fill(16);
    const extra = new Uint8Array(await subtle.encrypt({ name: 'AES-CBC', iv: lastBlock }, key, pad)).subarray(0, 16);
    const joined = new Uint8Array(data.length + 16);
    joined.set(data);
    joined.set(extra, data.length);
    const plain = new Uint8Array(await subtle.decrypt({ name: 'AES-CBC', iv }, key, joined));
    const n = plain[plain.length - 1];
    return n >= 1 && n <= 16 ? plain.subarray(0, plain.length - n) : plain;
  }

  async function aesGcmDecrypt(keyBytes, iv, data) {
    const key = await subtle.importKey('raw', keyBytes, 'AES-GCM', false, ['decrypt']);
    try {
      return new Uint8Array(await subtle.decrypt({ name: 'AES-GCM', iv, tagLength: 128 }, key, data));
    } catch (e) {
      throw new WrongPasswordError(); // GCM doğrulaması başarısız = yanlış parola
    }
  }

  async function decryptEntry(enc, password, data) {
    const alg = enc.algorithm['manifest:algorithm-name'] || '';
    if (/blowfish/i.test(alg)) {
      throw new Error('Eski Blowfish şifrelemesi desteklenmiyor. Dosyayı LibreOffice ile açıp yeniden kaydedin.');
    }
    const iv = b64(enc.algorithm['manifest:initialisation-vector']);
    const key = await deriveKey(enc, password);

    if (/aes256-gcm$/.test(alg)) {
      // LibreOffice IV'yi şifreli verinin başına da yazıyor; varsa atla.
      const body = eq(data.subarray(0, iv.length), iv) ? data.subarray(iv.length) : data;
      return aesGcmDecrypt(key, iv, body);
    }
    if (/aes(128|192|256)-cbc$/.test(alg)) {
      const plain = await aesCbcDecrypt(key, iv, data);
      const type = (enc.data['manifest:checksum-type'] || '').toLowerCase();
      const expected = enc.data['manifest:checksum'];
      if (expected) {
        const head = plain.subarray(0, 1024);
        const hash = type.endsWith('sha256-1k') ? await digest('SHA-256', head)
          : /sha1\/1k$/.test(type) ? await digest('SHA-1', head) : null;
        if (hash && !eq(hash, b64(expected))) throw new WrongPasswordError();
      }
      return plain;
    }
    throw new Error('Desteklenmeyen şifreleme: ' + alg);
  }

  function getManifest(files) {
    const m = files.get('META-INF/manifest.xml');
    return m ? new TextDecoder().decode(m) : '';
  }

  async function isEncrypted(bytes) {
    try {
      return getManifest(await readZip(bytes)).includes('encryption-data');
    } catch (e) {
      return false;
    }
  }

  async function decrypt(bytes, password) {
    const files = await readZip(bytes);
    const manifest = getManifest(files);
    const encrypted = parseManifest(manifest).filter((e) => e.enc);
    if (!encrypted.length) return bytes;

    // LibreOffice 24.2+: tüm paket tek girdi
    const pkg = encrypted.find((e) => e.attrs['manifest:full-path'] === 'encrypted-package');
    if (pkg) {
      const data = files.get('encrypted-package');
      if (!data) throw new Error('encrypted-package bulunamadı');
      const plain = await decryptEntry(pkg.enc, password, data);
      const inner = plain[0] === 0x50 && plain[1] === 0x4b ? plain : await inflateRaw(plain);
      if (inner[0] !== 0x50 || inner[1] !== 0x4b) throw new WrongPasswordError();
      return inner;
    }

    // ODF 1.2/1.3: dosya başına şifreleme
    let newManifest = manifest;
    for (const e of encrypted) {
      const path = e.attrs['manifest:full-path'];
      const data = files.get(path);
      if (!data) continue;
      const plain = await decryptEntry(e.enc, password, data);
      files.set(path, await tryInflate(plain));
      const stripped = e.raw
        .replace(/\s*<manifest:encryption-data[\s\S]*?<\/manifest:encryption-data>\s*/, '')
        .replace(/>\s*<\/manifest:file-entry>$/, '/>');
      newManifest = newManifest.replace(e.raw, stripped);
    }
    files.set('META-INF/manifest.xml', new TextEncoder().encode(newManifest));
    return writeZip(files);
  }

  root.OdsCrypt = { isEncrypted, decrypt, WrongPasswordError };
})(typeof globalThis !== 'undefined' ? globalThis : this);
