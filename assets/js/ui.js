// assets/js/ui.js
export const qs  = (sel, el=document)=> el.querySelector(sel);
export const qsa = (sel, el=document)=> [...el.querySelectorAll(sel)];

export function renderTable(container, rows, headers){
  const el = (typeof container==='string') ? qs(container) : container;
  if (!el) return;
  const thead = headers?.length ? `<thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead>` : '';
  const tbody = `<tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c ?? ''}</td>`).join('')}</tr>`).join('')}</tbody>`;
  el.innerHTML = `<table>${thead}${tbody}</table>`;
}
