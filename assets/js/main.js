// assets/js/main.js
import * as UI from './ui.js';
import * as IO from './dataio.js';

const els = {
  tx:  document.getElementById('txFile'),
  ln:  document.getElementById('lnFile'),
  pm:  document.getElementById('pmFile'),
  out: document.getElementById('tableContainer'),
};

function showPreview(label, result){
  const headers = result?.meta?.fields ?? (result.data.length ? Object.keys(result.data[0]) : []);
  const rows = (result.data || []).slice(0,10).map(obj => headers.map(h=> obj[h]));
  UI.renderTable(els.out, rows, headers);
}

['tx','ln','pm'].forEach(key=>{
  const input = els[key];
  if(!input) return;
  input.addEventListener('change', async (e)=>{
    const file = e.target.files && e.target.files[0];
    if(!file) return;
    try{
      const res = await IO.parseCSVFile(file);
      showPreview(key.toUpperCase(), res);
    }catch(err){
      console.error('CSV parse error:', err);
      alert('CSVの読み込みに失敗しました');
    }finally{
      // allow re-selecting the same file
      e.target.value = '';
    }
  });
});
