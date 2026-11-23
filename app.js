(() => {
  'use strict';

  const drop = document.getElementById('drop');
  const input = document.getElementById('file');
  const results = document.getElementById('results');
  const list = document.getElementById('list');
  const downloadAllBtn = document.getElementById('download-all');
  const clearBtn = document.getElementById('clear');

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const converted = []; // { name, url }

  function outName(name) {
    return name.replace(/\.f?ods$/i, '') + '.xlsx';
  }

  // SheetJS bozuk dosyaları düz metin olarak okuyabildiği için imzayı önceden kontrol et.
  // .ods = ZIP paketi ("PK\x03\x04"), .fods = tek parça XML.
  function looksLikeOds(name, buf) {
    const head = new Uint8Array(buf, 0, Math.min(buf.byteLength, 4096));
    if (/\.fods$/i.test(name)) {
      return new TextDecoder().decode(head).includes('office:document');
    }
    return head[0] === 0x50 && head[1] === 0x4b && head[2] === 0x03 && head[3] === 0x04;
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(1) + ' MB';
  }

  function triggerDownload(url, name) {
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function addRow(file) {
    const li = document.createElement('li');
    li.className = 'item pending';
    li.innerHTML = '<div class="meta"><span class="name"></span><span class="info">Dönüştürülüyor…</span></div>';
    li.querySelector('.name').textContent = file.name;
    list.appendChild(li);
    results.hidden = false;
    return li;
  }

  async function convert(file) {
    const li = addRow(file);
    const info = li.querySelector('.info');
    try {
      if (!/\.f?ods$/i.test(file.name)) throw new Error('Bu bir .ods dosyası değil');
      const data = await file.arrayBuffer();
      if (!looksLikeOds(file.name, data)) throw new Error('Geçerli bir ODS dosyası değil');
      const wb = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
      if (!wb.SheetNames.length) throw new Error('Dosyada sayfa bulunamadı');

      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true, cellStyles: true });
      const blob = new Blob([out], { type: XLSX_MIME });
      const url = URL.createObjectURL(blob);
      const name = outName(file.name);
      converted.push({ name, url });

      li.className = 'item done';
      info.textContent = `${wb.SheetNames.length} sayfa · ${formatSize(blob.size)}`;
      const btn = document.createElement('a');
      btn.className = 'btn';
      btn.href = url;
      btn.download = name;
      btn.textContent = 'İndir';
      li.appendChild(btn);
    } catch (err) {
      console.error(err);
      li.className = 'item error';
      info.textContent = 'Hata: ' + (err && err.message ? err.message : 'dosya okunamadı');
    }
    downloadAllBtn.hidden = converted.length < 2;
  }

  async function handleFiles(files) {
    for (const f of files) await convert(f);
  }

  input.addEventListener('change', () => {
    handleFiles(Array.from(input.files));
    input.value = '';
  });

  drop.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });

  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => {
    e.preventDefault();
    drop.classList.add('over');
  }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => {
    e.preventDefault();
    drop.classList.remove('over');
  }));
  drop.addEventListener('drop', (e) => handleFiles(Array.from(e.dataTransfer.files)));

  // Sayfanın başka bir yerine bırakılan dosyanın tarayıcıda açılmasını engelle
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  downloadAllBtn.addEventListener('click', async () => {
    for (const { url, name } of converted) {
      triggerDownload(url, name);
      await new Promise((r) => setTimeout(r, 300));
    }
  });

  clearBtn.addEventListener('click', () => {
    converted.forEach(({ url }) => URL.revokeObjectURL(url));
    converted.length = 0;
    list.innerHTML = '';
    results.hidden = true;
  });
})();
