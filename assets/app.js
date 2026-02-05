
/* FinDash (2026) - lê data/data.json (gerado pelo Power Automate → GitHub) */
(() => {
  const DATA_URL = './data/data.json';

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  const fmtBRL = (n) => {
    const v = Number(n || 0);
    return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  };

  const fmtPct = (n) => {
    if (n === null || n === undefined || Number.isNaN(Number(n))) return '—';
    return (Number(n) * 100).toFixed(0) + '%';
  };

  const monthLabel = (m) => m; // already PT month name

  const state = {
    data: null,
    activeMetrics: new Set(['entrada','saida','liquido','diff_m1','crescimento']),
    monthKey: null,
    charts: {}
  };

  function loadTheme(){
    const saved = localStorage.getItem('findash-theme') || 'dark';
    document.documentElement.dataset.theme = saved;
  }
  function toggleTheme(){
    const cur = document.documentElement.dataset.theme || 'dark';
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    localStorage.setItem('findash-theme', next);
    // redraw charts to fit theme contrast
    renderAll();
  }

  async function fetchJSON(noCache=false){
    const url = noCache ? (DATA_URL + '?ts=' + Date.now()) : DATA_URL;
    const res = await fetch(url, { cache: 'no-store' });
    if(!res.ok) throw new Error('Falha ao buscar data.json (' + res.status + ')');
    return await res.json();
  }

  function setUpdatedAt(){
    const pill = $('#updatedAtPill');
    if(!pill || !state.data) return;
    const dt = state.data.generatedAt ? new Date(state.data.generatedAt) : null;
    pill.textContent = dt && !isNaN(dt) ? ('Atualizado: ' + dt.toLocaleString('pt-BR')) : 'Atualizado: —';
  }

  function mountMonthSelect(selectEl, onChange){
    if(!selectEl || !state.data) return;
    selectEl.innerHTML = '';
    state.data.months.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m.key;
      opt.textContent = monthLabel(m.name);
      selectEl.appendChild(opt);
    });

    const defaultKey = state.monthKey || state.data.months[0]?.key;
    if(defaultKey){
      selectEl.value = defaultKey;
      state.monthKey = defaultKey;
    }

    selectEl.addEventListener('change', (e) => {
      state.monthKey = e.target.value;
      onChange?.();
    });
  }

  function getMonth(){
    if(!state.data) return null;
    return state.data.months.find(m => m.key === state.monthKey) || state.data.months[0];
  }

  function categoriesFromSaidas(monthObj){
    const map = new Map();
    (monthObj?.details?.saidas || []).forEach(it => {
      const k = (it.descricao || '').trim() || 'Outros';
      map.set(k, (map.get(k) || 0) + Number(it.valor || 0));
    });
    const arr = Array.from(map.entries()).map(([k,v]) => ({categoria:k, valor:v}));
    arr.sort((a,b) => b.valor - a.valor);
    return arr;
  }

  function destroyChart(id){
    const ch = state.charts[id];
    if(ch){ ch.destroy(); delete state.charts[id]; }
  }

  function chartDefaults(){
    // Keep charts readable in both themes without manually choosing colors.
    const light = document.documentElement.dataset.theme === 'light';
    Chart.defaults.color = light ? 'rgba(10,12,20,0.75)' : 'rgba(255,255,255,0.78)';
    Chart.defaults.borderColor = light ? 'rgba(10,12,20,0.10)' : 'rgba(255,255,255,0.10)';
    Chart.defaults.font.family = 'Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif';
  }

  function renderDashboard(){
    const m = getMonth();
    if(!m) return;

    // KPI
    const s = m.summary || {};
    const elEntrada = $('#kpiEntrada');
    const elSaida = $('#kpiSaida');
    const elLiquido = $('#kpiLiquido');
    const elDiff = $('#kpiDiff');
    const elCresc = $('#kpiCresc');

    if(elEntrada) elEntrada.textContent = fmtBRL(s.entrada);
    if(elSaida) elSaida.textContent = fmtBRL(s.saida);
    if(elLiquido) elLiquido.textContent = fmtBRL(s.liquido);
    if(elDiff) elDiff.textContent = fmtBRL(s.diff_m1);
    if(elCresc) elCresc.textContent = fmtPct(s.crescimento);

    // Summary table (Tabela1)
    const tbody = $('#summaryTable tbody');
    if(tbody){
      tbody.innerHTML = '';
      state.data.months.forEach(mm => {
        const tr = document.createElement('tr');
        const isActive = mm.key === state.monthKey;
        if(isActive) tr.style.background = 'rgba(14,165,233,0.10)';
        tr.innerHTML = `
          <td><b>${monthLabel(mm.name)}</b></td>
          <td data-col="entrada" class="right">${fmtBRL(mm.summary?.entrada)}</td>
          <td data-col="saida" class="right">${fmtBRL(mm.summary?.saida)}</td>
          <td data-col="liquido" class="right">${fmtBRL(mm.summary?.liquido)}</td>
          <td data-col="diff_m1" class="right">${fmtBRL(mm.summary?.diff_m1)}</td>
          <td data-col="crescimento" class="right">${mm.summary?.crescimento === null || mm.summary?.crescimento === undefined ? '—' : fmtPct(mm.summary?.crescimento)}</td>
        `;
        tbody.appendChild(tr);
      });

      // show/hide columns based on chips
      const cols = ['entrada','saida','liquido','diff_m1','crescimento'];
      cols.forEach(c => {
        const show = state.activeMetrics.has(c);
        $$('#summaryTable [data-col="'+c+'"]').forEach(td => {
          td.style.display = show ? '' : 'none';
        });
      });
    }

    // Charts
    chartDefaults();

    // Trend chart
    const trendCanvas = $('#trendChart');
    if(trendCanvas){
      destroyChart('trendChart');
      const labels = state.data.months.map(mm => monthLabel(mm.name));
      const entrada = state.data.months.map(mm => Number(mm.summary?.entrada || 0));
      const saida = state.data.months.map(mm => Number(mm.summary?.saida || 0));
      const liquido = state.data.months.map(mm => Number(mm.summary?.liquido || 0));
      state.charts.trendChart = new Chart(trendCanvas, {
        type: 'line',
        data: {
          labels,
          datasets: [
            { label:'Entrada', data: entrada, tension: 0.35 },
            { label:'Saída', data: saida, tension: 0.35 },
            { label:'Líquido', data: liquido, tension: 0.35 }
          ]
        },
        options: {
          responsive:true,
          maintainAspectRatio:false,
          plugins: { legend: { position:'bottom' }, tooltip: { callbacks: { label: (ctx)=> `${ctx.dataset.label}: ${fmtBRL(ctx.parsed.y)}` } } },
          scales: {
            y: { ticks: { callback: (v)=> Number(v).toLocaleString('pt-BR') } }
          }
        }
      });
    }

    // Top categories chart (from monthly saidas details)
    const topCanvas = $('#topCatsChart');
    if(topCanvas){
      destroyChart('topCatsChart');
      const cats = categoriesFromSaidas(m).slice(0,8);
      const labels = cats.map(x => x.categoria);
      const values = cats.map(x => x.valor);
      state.charts.topCatsChart = new Chart(topCanvas, {
        type:'doughnut',
        data:{ labels, datasets:[{ label:'Saídas', data: values }]},
        options:{
          responsive:true,
          maintainAspectRatio:false,
          plugins:{ legend:{ position:'bottom' }, tooltip:{ callbacks:{ label:(ctx)=> `${ctx.label}: ${fmtBRL(ctx.parsed)}` } } }
        }
      });
    }

    // Nubank & Santander cards (Tabela2/3)
    const nubankTotal = state.data.months.reduce((a,mm)=>a+Number(mm.accounts?.nubank_saida||0),0);
    const santTotal = state.data.months.reduce((a,mm)=>a+Number(mm.accounts?.santander_saida||0),0);
    if($('#nubankTotal')) $('#nubankTotal').textContent = fmtBRL(nubankTotal);
    if($('#santanderTotal')) $('#santanderTotal').textContent = fmtBRL(santTotal);

    const nubCanvas = $('#nubankChart');
    if(nubCanvas){
      destroyChart('nubankChart');
      state.charts.nubankChart = new Chart(nubCanvas, {
        type:'bar',
        data:{
          labels: state.data.months.map(mm => monthLabel(mm.name)),
          datasets:[{ label:'Nubank (Saída)', data: state.data.months.map(mm=>Number(mm.accounts?.nubank_saida||0)) }]
        },
        options:{
          responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label:(ctx)=> fmtBRL(ctx.parsed.y) } } },
          scales:{ y:{ ticks:{ callback:(v)=> Number(v).toLocaleString('pt-BR') } } }
        }
      });
    }

    const sanCanvas = $('#santanderChart');
    if(sanCanvas){
      destroyChart('santanderChart');
      state.charts.santanderChart = new Chart(sanCanvas, {
        type:'bar',
        data:{
          labels: state.data.months.map(mm => monthLabel(mm.name)),
          datasets:[{ label:'Santander (Saída)', data: state.data.months.map(mm=>Number(mm.accounts?.santander_saida||0)) }]
        },
        options:{
          responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label:(ctx)=> fmtBRL(ctx.parsed.y) } } },
          scales:{ y:{ ticks:{ callback:(v)=> Number(v).toLocaleString('pt-BR') } } }
        }
      });
    }
  }

  function renderDetail(){
    const m = getMonth();
    if(!m) return;
    chartDefaults();

    const entradas = (m.details?.entradas || []).slice().sort((a,b)=> (b.valor||0)-(a.valor||0));
    const saidas = (m.details?.saidas || []).slice().sort((a,b)=> (b.valor||0)-(a.valor||0));

    const entradaTotal = entradas.reduce((a,x)=>a+Number(x.valor||0),0);
    const saidaTotal = saidas.reduce((a,x)=>a+Number(x.valor||0),0);
    const liquido = Number(m.summary?.liquido ?? (entradaTotal - saidaTotal));

    if($('#detailEntradaTotal')) $('#detailEntradaTotal').textContent = 'Total: ' + fmtBRL(entradaTotal);
    if($('#detailSaidaTotal')) $('#detailSaidaTotal').textContent = 'Saída: ' + fmtBRL(saidaTotal);
    if($('#detailLiquido')) $('#detailLiquido').textContent = 'Líquido: ' + fmtBRL(liquido);

    const et = $('#entradaTable tbody');
    if(et){
      et.innerHTML = '';
      entradas.forEach(x=>{
        const tr=document.createElement('tr');
        tr.innerHTML = `<td>${escapeHtml(x.descricao||'')}</td><td class="right">${fmtBRL(x.valor)}</td>`;
        et.appendChild(tr);
      });
    }

    const st = $('#saidaTable tbody');
    if(st){
      st.innerHTML = '';
      saidas.forEach(x=>{
        const tr=document.createElement('tr');
        tr.innerHTML = `<td>${escapeHtml(x.descricao||'')}</td><td class="right">${fmtBRL(x.valor)}</td>`;
        st.appendChild(tr);
      });
    }

    // donut categories
    const catsCanvas = $('#detailCatsChart');
    if(catsCanvas){
      destroyChart('detailCatsChart');
      const cats = categoriesFromSaidas(m).slice(0,10);
      state.charts.detailCatsChart = new Chart(catsCanvas, {
        type:'doughnut',
        data:{ labels: cats.map(x=>x.categoria), datasets:[{ data: cats.map(x=>x.valor) }] },
        options:{
          responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{ position:'bottom' }, tooltip:{ callbacks:{ label:(ctx)=> `${ctx.label}: ${fmtBRL(ctx.parsed)}` } } }
        }
      });
    }

    // bar Entrada vs Saída
    const barCanvas = $('#detailBarChart');
    if(barCanvas){
      destroyChart('detailBarChart');
      state.charts.detailBarChart = new Chart(barCanvas, {
        type:'bar',
        data:{
          labels:['Entrada','Saída'],
          datasets:[{ label: monthLabel(m.name), data:[entradaTotal, saidaTotal] }]
        },
        options:{
          responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label:(ctx)=> fmtBRL(ctx.parsed.y) } } },
          scales:{ y:{ ticks:{ callback:(v)=> Number(v).toLocaleString('pt-BR') } } }
        }
      });
    }
  }

  function escapeHtml(str){
    return String(str)
      .replaceAll('&','&amp;')
      .replaceAll('<','&lt;')
      .replaceAll('>','&gt;')
      .replaceAll('"','&quot;')
      .replaceAll("'","&#039;");
  }

  function wireChips(){
    $$('.chip').forEach(btn=>{
      btn.addEventListener('click', ()=>{
        const metric = btn.dataset.metric;
        const isOn = btn.classList.toggle('active');
        if(isOn) state.activeMetrics.add(metric);
        else state.activeMetrics.delete(metric);
        renderDashboard();
      });
    });
  }

  function renderAll(){
    setUpdatedAt();
    // year label
    $$('#yearLabel').forEach(el => { if(state.data?.year) el.textContent = state.data.year; });

    // which page?
    const onDetail = !!$('#monthSelectDetail');
    if(onDetail){
      mountMonthSelect($('#monthSelectDetail'), renderDetail);
      renderDetail();
    }else{
      mountMonthSelect($('#monthSelect'), renderDashboard);
      wireChips();
      renderDashboard();
    }
  }

  async function init(){
    loadTheme();

    const refreshBtn = $('#refreshBtn');
    if(refreshBtn){
      refreshBtn.addEventListener('click', async ()=>{
        try{
          state.data = await fetchJSON(true);
          // keep selected month if still exists
          if(state.monthKey && !state.data.months.some(m=>m.key===state.monthKey)){
            state.monthKey = state.data.months[0]?.key;
          }
          renderAll();
        }catch(err){
          alert('Não foi possível carregar os dados.\n' + err.message);
        }
      });
    }

    const themeBtn = $('#themeBtn');
    if(themeBtn) themeBtn.addEventListener('click', toggleTheme);

    try{
      state.data = await fetchJSON(false);
      state.monthKey = state.data.months[0]?.key;
      renderAll();
    }catch(err){
      alert('Não foi possível carregar data/data.json.\n' +
        'Verifique se o arquivo existe no repositório e se o GitHub Pages está apontando para a branch correta.\n\n' +
        err.message);
    }
  }

  window.addEventListener('DOMContentLoaded', init);
})();
