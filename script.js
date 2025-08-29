const aylar = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık'];
const aySel = document.getElementById('ay');
aylar.forEach(m=>{ const o=document.createElement('option'); o.value=m; o.textContent=m; aySel.appendChild(o); });
aySel.value = aylar[new Date().getMonth()];

const daireSel = document.getElementById('daire');
for(let i=1;i<=44;i++){ const o=document.createElement('option'); o.value=i; o.textContent=i; daireSel.appendChild(o); }

const turSel = document.getElementById('tur');
const rowDaire = document.getElementById('rowDaire');
const rowKategori = document.getElementById('rowKategori');
turSel.addEventListener('change', ()=>{
  const isAidat = turSel.value === 'Aidat';
  rowDaire.style.display = isAidat ? '' : 'none';
  rowKategori.style.display = isAidat ? 'none' : '';
});

const TBL = document.querySelector('#tbl tbody');
const load = () => JSON.parse(localStorage.getItem('records')||'[]');
const save = (arr) => localStorage.setItem('records', JSON.stringify(arr));
const render = () => {
  const arr = load();
  TBL.innerHTML = '';
  arr.forEach(r=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.tur}</td><td>${r.tur==='Aidat' ? 'Daire ' + r.daire : r.kategori}</td><td>${r.ay}</td><td>${Number(r.tutar).toFixed(2)}</td>`;
    TBL.appendChild(tr);
  });
};
render();

document.getElementById('btnKaydet').addEventListener('click', ()=>{
  const tur = turSel.value;
  const ay = aySel.value;
  const daire = daireSel.value;
  const kategori = document.getElementById('kategori').value;
  const tutar = parseFloat(document.getElementById('tutar').value||'0');
  const rec = { tur, ay, tutar, daire: tur==='Aidat'? daire : '', kategori: tur==='Gider'? kategori : '' };
  const arr = load(); arr.push(rec); save(arr); render();
});

document.getElementById('btnTemizle').addEventListener('click', ()=>{
  if(confirm('Tüm kayıtlar silinsin mi?')){ localStorage.removeItem('records'); render(); }
});

document.getElementById('btnExcel').addEventListener('click', ()=>{
  const records = load();

  // Sheet 1: Aidat Çizelgesi (Daire x Ay pivot)
  const header = ['Daire', ...aylar];
  const table = [header];
  for(let d=1; d<=44; d++){
    const row = [d];
    aylar.forEach(ay => {
      const sum = records
        .filter(r => r.tur==='Aidat' && String(r.daire)===String(d) && r.ay===ay)
        .reduce((a,b)=>a+Number(b.tutar||0), 0);
      row.push(sum);
    });
    table.push(row);
  }
  const totalRow = ['Aylık Toplam'];
  for(let j=1;j<header.length;j++){
    let s = 0;
    for(let i=1;i<table.length;i++){ s += Number(table[i][j]||0); }
    totalRow.push(s);
  }
  table.push(totalRow);
  const wsAidat = XLSX.utils.aoa_to_sheet(table);

  // Sheet 2: Gelir-Gider Tablosu
  const gelirGider = [['Tür','Daire/Kategori','Ay','Tutar']];
  let toplamGelir = 0, toplamGider = 0;
  records.forEach(r=>{
    gelirGider.push([r.tur, r.tur==='Aidat' ? ('Daire ' + r.daire) : r.kategori, r.ay, Number(r.tutar||0)]);
    if(r.tur==='Aidat'){ toplamGelir += Number(r.tutar||0); }
    if(r.tur==='Gider'){ toplamGider += Number(r.tutar||0); }
  });
  gelirGider.push([]);
  gelirGider.push(['Toplam Gelir','','',toplamGelir]);
  gelirGider.push(['Toplam Gider','','',toplamGider]);
  gelirGider.push(['Kasa (Gelir-Gider)','','',toplamGelir-toplamGider]);
  const wsGG = XLSX.utils.aoa_to_sheet(gelirGider);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsAidat, 'Aidat Cizelgesi');
  XLSX.utils.book_append_sheet(wb, wsGG, 'Gelir Gider');
  XLSX.writeFile(wb, 'SiteRapor.xlsx');
});
