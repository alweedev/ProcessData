(function(){
  // Signal enhanced flow and isolate namespace
  try { window.INATIVACAO_ENHANCED = true; } catch(e) {}

  const $ = (id) => document.getElementById(id);
  const showToast = (msg, type='info') => {
    try { window.showToast && window.showToast(msg, type); } catch(_) { console.log(msg); }
  };
  function setBannerContent(elem, html){
    if(!elem) return;
    elem.classList.remove('upload-animate-out');
    if(html){
      elem.innerHTML = html;
      elem.classList.remove('upload-animate-in');
      void elem.offsetWidth;
      elem.classList.add('upload-animate-in');
      const end = (e)=>{ if(e.animationName==='bannerIn'){ elem.classList.remove('upload-animate-in'); elem.removeEventListener('animationend', end);} };
      elem.addEventListener('animationend', end);
    } else {
      elem.innerHTML='';
    }
  }
  function clearBanner(elem){
    if(!elem || !elem.innerHTML) return;
    elem.classList.remove('upload-animate-in');
    elem.classList.add('upload-animate-out');
    const end = (e)=>{ if(e.animationName==='bannerOut'){ elem.innerHTML=''; elem.classList.remove('upload-animate-out'); elem.removeEventListener('animationend', end);} };
    elem.addEventListener('animationend', end);
  }

  // (Lista via upload removida) – sem necessidade de carregar XLSX aqui

  // Basic state
  const state = {
    rawItems: [],
    validCpfs: [],
    validNames: [],
    validEmails: [],
    duplicates: [],
    results: [],
    filtered: [],
    page: 1,
    pageSize: 10,
  };

  // Utils
  const escapeHtml = (s)=> String(s??'').replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));
  const formatCpf = (c)=> c && c.length===11 ? `${c.slice(0,3)}.${c.slice(3,6)}.${c.slice(6,9)}-${c.slice(9)}` : (c||'');
  const isValidFullName = (s)=>{
    const norm = String(s).normalize('NFKD').replace(/[\u0300-\u036f]/g,'');
    const parts = norm.trim().split(/\s+/);
    return parts.length >= 2 && norm.trim().length >= 3;
  };

  function parseTextarea(text){
    const lines = (text||'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
    state.rawItems = lines;
    const cpfRaw = [], nameRaw = [], emailRaw = [];
    const emailPat = /^[^@\s]+@[^@\s]+\.[^@\s]+$/i;
    for(const line of lines){
      const digits = line.replace(/\D/g,'');
      if (emailPat.test(line)) {
        emailRaw.push(line);
      } else if(digits.length===11) {
        cpfRaw.push(digits);
      } else {
        nameRaw.push(line);
      }
    }
    const seen = new Set(); const dups = []; const validCpfs = [];
    for(const c of cpfRaw){ if(seen.has(c)){ if(!dups.includes(c)) dups.push(c); } else { seen.add(c); validCpfs.push(c); } }
    state.validCpfs = validCpfs;
    state.duplicates = dups;
    state.validNames = nameRaw.filter(isValidFullName);
    state.validEmails = emailRaw;
    return { totalValid: validCpfs.length + state.validNames.length + state.validEmails.length, dupCount: dups.length, totalLines: lines.length };
  }

  function renderCounts(){
    const summaryEl = $('lista_valid_summary');
    const dupWarnEl = $('lista_duplicates_warning');
    const total = (state.validCpfs?.length||0) + (state.validNames?.length||0) + (state.validEmails?.length||0);
    if (summaryEl) summaryEl.innerHTML = `<span class="badge bg-primary">${total}</span> itens válidos (CPF, Nome Completo ou E-mail)`;
    if (dupWarnEl) dupWarnEl.textContent = state.duplicates.length ? `CPFs duplicados: ${state.duplicates.join(', ')}` : '';
  }

  function renderTable(){
    const body = $('results_body'); const pageInfo = $('page_info');
    if(!body || !pageInfo) return;
    const start = (state.page-1)*state.pageSize;
    const rows = state.filtered.slice(start, start+state.pageSize);
    body.innerHTML = rows.map(r=>{
      let cls = '';
      const status = (r.status_atual || '').toString().trim().toUpperCase();
      if (!r.found) {
        cls = 'table-warning';
      } else if (status === 'ATIVO') {
        cls = 'status-ativo';
      } else if (status && status !== 'ATIVO') {
        // qualquer coisa diferente de ATIVO consideramos como inativo/bloqueado
        cls = 'status-inativo';
      }
      return `<tr class="${cls}" data-cpf="${escapeHtml(r.cpf||'')}">
        <td>${escapeHtml(r.nome||'')}</td>
        <td><code>${escapeHtml(formatCpf(r.cpf||''))}</code></td>
        <td>${escapeHtml(r.email||'')}</td>
        <td class="status-cell">${escapeHtml(r.status_atual||'')}</td>
      </tr>`;
    }).join('');
    pageInfo.textContent = `Página ${state.page} de ${Math.max(1, Math.ceil(state.filtered.length/state.pageSize))}`;
  }

  function wire(){
    const baseInput = $('inativacao_base');
    const clearBaseBtn = $('inativacao_clear_base_btn');
    const baseFeedback = $('inativacao_base_feedback');
    const textarea = $('lista_text');
    const runBtn = $('inativacao_btn');
    const generateBtn = $('inativacao_generate_btn');
    const resultsWrap = $('inativacao_results');
    const resultsControls = $('inativacao_results_controls');
    const pagination = $('results_pagination');
    const searchInput = $('result_search');
    const statusEl = $('inativacao_status');
    const debugEl = $('inativacao_debug');
    const progressWrap = $('inativacao_progress');
    const progressBar = $('inativacao_progressBar');

    if(!textarea || !runBtn) return;

    // Parse initial
    parseTextarea(textarea.value);
    renderCounts();
    runBtn.disabled = !(state.validCpfs.length || state.validNames.length || state.validEmails.length) || !baseInput?.files?.[0];

    textarea.addEventListener('input', ()=>{ parseTextarea(textarea.value); renderCounts(); runBtn.disabled = !(state.validCpfs.length || state.validNames.length || state.validEmails.length) || !baseInput.files[0]; });
    textarea.addEventListener('paste', ()=> setTimeout(()=>{ parseTextarea(textarea.value); renderCounts(); runBtn.disabled = !(state.validCpfs.length || state.validNames.length || state.validEmails.length) || !baseInput.files[0]; }, 0));
    baseInput?.addEventListener('change', ()=>{ runBtn.disabled = !(state.validCpfs.length || state.validNames.length || state.validEmails.length) || !baseInput.files[0]; });
    baseInput?.addEventListener('change', ()=>{
      if (baseInput.files && baseInput.files[0]) {
        if (baseFeedback) setBannerContent(baseFeedback, `<span class="upload-check" aria-hidden="true"><i class="bi bi-check-lg text-success" style="font-size:1.15rem; line-height:1; display:inline-block;"></i></span>${escapeHtml(baseInput.files[0].name)}`);
      } else if (baseFeedback) clearBanner(baseFeedback);
    });

    // Clear buttons for attachments
    clearBaseBtn && clearBaseBtn.addEventListener('click', ()=>{
      try {
        if (baseInput) baseInput.value = '';
        if (baseFeedback) clearBanner(baseFeedback);
        // Reset results UI if based on wrong base
        resultsWrap?.classList.add('d-none');
        resultsControls?.classList.add('d-none');
        pagination?.classList.add('d-none');
        const body = $('results_body'); if (body) body.innerHTML = '';
        statusEl && (statusEl.innerHTML = '');
        debugEl && (debugEl.textContent = '');
        progressWrap?.classList.add('d-none');
        progressBar && (progressBar.style.width='0%');
        progressBar && (progressBar.textContent='');
        runBtn && (runBtn.disabled = !(state.validCpfs.length || state.validNames.length || state.validEmails.length));
        showToast('Base limpa. Selecione um novo arquivo.', 'info');
      } catch(_){}
    });
    // Upload de lista removido – entrada somente via textarea

    // Buscar (preview) usando /api/inativacao/buscar
    runBtn.addEventListener('click', async ()=>{
      statusEl.innerHTML = '';
      debugEl.textContent = '';
      if(!baseInput.files[0]){ showToast('Envie a base.','danger'); return; }
      if(!(state.validCpfs.length || state.validNames.length || state.validEmails.length)) { showToast('Nenhum CPF, Nome ou E-mail válido para buscar.','danger'); return; }
      runBtn.disabled = true; runBtn.classList.add('btn-processing');
      progressWrap.classList.remove('d-none');
      progressBar.style.width = '15%'; progressBar.textContent = 'Buscando...';
      statusEl.innerHTML = '<span class="spinner-border spinner-border-sm me-2 text-primary"></span>Buscando usuários na base...';

      // skeleton rows
      try {
        const body = $('results_body');
        if (body) body.innerHTML = Array.from({length:10}).map(()=>'<tr class="skeleton-row">'+['','','',''].map(()=>'<td></td>').join('')+'</tr>').join('');
        $('results_table')?.setAttribute('aria-busy','true');
      } catch(_){}

      try {
        const fd = new FormData();
        fd.append('base', baseInput.files[0]);
        fd.append('itens', JSON.stringify(state.validCpfs.concat(state.validNames).concat(state.validEmails)));
        const resp = await fetch('/api/inativacao/buscar', { method: 'POST', body: fd });
        const data = await resp.json();
        if(!resp.ok) throw new Error(data.error || 'Falha na busca');
        state.results = Array.isArray(data.items) ? data.items : [];
        state.filtered = state.results.slice();
        state.page = 1;
        renderTable();
        resultsWrap.classList.remove('d-none');
        resultsControls.classList.remove('d-none');
        pagination.classList.remove('d-none');
        generateBtn?.classList.remove('d-none');
        showToast(`Busca concluída: ${state.results.length} itens.`, 'success');
      } catch(err){
        debugEl.textContent = 'Erro: ' + (err.message||err);
        showToast('Erro na busca: ' + (err.message||err), 'danger');
      } finally {
        runBtn.disabled = false; runBtn.classList.remove('btn-processing');
        progressWrap.classList.add('d-none'); progressBar.style.width='0%'; progressBar.textContent=''; statusEl.innerHTML='';
        try{ $('results_table')?.removeAttribute('aria-busy'); }catch(_){ }
      }
    });

    // Paginação e filtro
    $('page_prev')?.addEventListener('click', ()=>{ if(state.page>1){ state.page--; renderTable(); } });
    $('page_next')?.addEventListener('click', ()=>{ const max = Math.ceil(state.filtered.length/state.pageSize); if(state.page<max){ state.page++; renderTable(); } });
    searchInput?.addEventListener('input', ()=>{
      const term = (searchInput.value||'').trim().toLowerCase();
      if(!term){ state.filtered = state.results.slice(); state.page=1; renderTable(); return; }
      state.filtered = state.results.filter(r=> (String(r.nome||'').toLowerCase().includes(term) || String(r.cpf||'').includes(term)) );
      state.page = 1; renderTable();
    });

    // Geração final (Excel) em /api/process_inativacao
    generateBtn && generateBtn.addEventListener('click', async ()=>{
      if(!baseInput.files[0]){ showToast('Envie a base para gerar a inativação.','danger'); return; }
      const listaText = (textarea && textarea.value) ? textarea.value : (state.validCpfs.concat(state.validNames)).join('\n');
      progressWrap.classList.remove('d-none');
      const xhr = new XMLHttpRequest();
      xhr.upload.addEventListener('progress', (e)=>{
        if(e.lengthComputable){ const pct=(e.loaded/e.total)*100; progressBar.style.width=pct.toFixed(0)+'%'; progressBar.textContent=pct.toFixed(0)+'%'; }
      });
      xhr.open('POST', '/api/process_inativacao', true);
      xhr.responseType = 'blob';
      xhr.onload = ()=>{
        if(xhr.status>=200 && xhr.status<300){
          const blob = xhr.response; if(!blob || blob.size===0){ showToast('Arquivo gerado inválido.','danger'); return; }
          const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='saida_inativacao.xlsx'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
          showToast('Inativação processada.','success');
        } else {
          try{ const reader=new FileReader(); reader.onload=()=>{ try{ const obj=JSON.parse(reader.result||'{}'); showToast(obj.error||('Erro '+xhr.status),'danger'); }catch(_){ showToast(reader.result||('Erro '+xhr.status),'danger'); } }; reader.readAsText(xhr.response); } catch(_){ showToast('Erro ao gerar inativação.','danger'); }
        }
        progressWrap.classList.add('d-none'); progressBar.style.width='0%'; progressBar.textContent='';
      };
      xhr.onerror = ()=>{ showToast('Erro de rede ao gerar inativação.','danger'); progressWrap.classList.add('d-none'); };
      const fd = new FormData(); fd.append('base', baseInput.files[0]); fd.append('lista_text', listaText);
      xhr.send(fd);
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire); else wire();
})();
