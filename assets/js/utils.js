// assets/js/utils.js
export const toNarrow = s => (s ?? '').replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0));
export const toNumber = v => {
  if (v == null || v === '') return 0;
  const n = Number(String(v).replace(/,/g,''));
  return Number.isFinite(n) ? n : 0;
};
export function parseDate(s){
  if(!s) return null;
  const t = new Date(s);
  return isNaN(t) ? null : t;
}
export function downloadCSV(filename, csvText){
  const blob = new Blob([csvText], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement('a'), {href:url, download:filename});
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}
