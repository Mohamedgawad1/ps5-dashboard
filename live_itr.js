(() => {
  const LS = 'live_itr_state';
  function fmt(n){ return (n||0).toLocaleString('en-US'); }
  function build(update){
    try{ syncCards(update); }catch(e){}
    const e = document.querySelector('#itr-live-badge');
    if(!e) return;
    const cb = update.closed, eb = (update.closed_by_discipline||{}).E||0,
          ib = (update.closed_by_discipline||{}).I||0, tb = (update.closed_by_discipline||{}).T||0;
    e.innerHTML =
      '<div class="itr-badge-title">ITR Closed &middot; Live</div>' +
      '<div class="itr-badge-big">' + fmt(cb) + '</div>' +
      '<div class="itr-badge-sub">E ' + fmt(eb) + ' &middot; I ' + fmt(ib) + ' &middot; T ' + fmt(tb) +
      ' &middot; Total E&I&T ' + fmt(update.eit_total) + '</div>' +
      '<div class="itr-badge-time">Updated ' + (update.updated||'') + '</div>';
  }
  function fetchItr(){
    fetch('itr_live_state.json?v=' + Math.floor(Date.now()/120000), {cache:'no-store'})
      .then(r => r.json())
      .then(u => { try{ localStorage.setItem(LS, JSON.stringify(u)); }catch(e){} build(u); })
      .catch(() => { try{ const o=localStorage.getItem(LS); if(o) build(JSON.parse(o)); }catch(e){} });
  }
  if(!document.querySelector('#itr-live-css')) {
    const st = document.createElement('style'); st.id='itr-live-css';
    st.textContent = '#itr-live-badge{position:fixed;right:16px;bottom:16px;z-index:99999;background:#0b2f56;color:#fff;border:1px solid #38bdf8;border-radius:12px;padding:10px 14px;font-family:Segoe UI,Arial,sans-serif;box-shadow:0 6px 18px rgba(0,0,0,.35);min-width:210px}'
      + '.itr-badge-title{font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#7dd3fc}'
      + '.itr-badge-big{font-size:30px;font-weight:800;line-height:1.1}'
      + '.itr-badge-sub{font-size:12px;color:#cbd5e1;margin-top:2px}'
      + '.itr-badge-time{font-size:10px;color:#94a3b8;margin-top:4px}';
    document.head.appendChild(st);
  }
  if(!document.querySelector('#itr-live-badge')){
    const d = document.createElement('div'); d.id='itr-live-badge';
    d.innerHTML = '<div class="itr-badge-title">ITR Closed &middot; Live</div><div class="itr-badge-big">&hellip;</div>';
    document.body.appendChild(d);
  }
  fetchItr();
  setInterval(fetchItr, 120000);

  function syncCards(u){
    if(!u) return;
    const closed = u.closed||0, eit = u.eit_total||0;
    const pct = eit>0 ? (100*closed/eit) : 0, pctTxt = pct.toFixed(2);
    const fmt = n => (n||0).toLocaleString('en-US');
    const set = (id,val) => { const el = document.getElementById(id); if(el) el.textContent = val; };
    set('totalClosed', fmt(closed));
    set('totalPct', pctTxt);
    const bar = document.getElementById('totalPctBar') || document.querySelector('.progress-fill');
    if(bar) bar.style.width = Math.min(100, pct) + '%';
    document.querySelectorAll('.kpi').forEach(c=>{
      const lbl = c.querySelector('.lbl'), v = c.querySelector('.val');
      if(!lbl || !v) return;
      const t = (lbl.textContent||'').trim();
      if(/CPP AGI EIT|Total ITRs Closed/i.test(t) && !/Today/i.test(t)){
        v.textContent = fmt(closed) + (eit>0 ? ' / '+fmt(eit) : '');
      } else if(/Progress/i.test(t)){ v.textContent = pctTxt + '%'; }
    });
  }
  try{ syncCards(JSON.parse(localStorage.getItem('live_itr_state')||'{}')); }catch(e){}

})();
