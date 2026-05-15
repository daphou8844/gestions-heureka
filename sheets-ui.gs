/**
 * ============================================================
 *  sheets-ui.gs — Gestions Heureka
 *  Interface visuelle sombre · Menu enrichi · Dashboard PDF
 * ============================================================
 *
 *  INSTALLATION (dans le même projet que sheets-script.gs) :
 *  1. Apps Script → + Nouveau fichier → Script → "sheets-ui"
 *  2. Coller ce code et sauvegarder (Ctrl+S)
 *  3. IMPORTANT : supprimer ou commenter la fonction onOpen()
 *     dans sheets-script.gs (ce fichier la remplace)
 *  4. Recharger le Google Sheets → le menu 🏗 Heureka apparaît
 *  5. Exécuter initSheetsUi() une première fois pour créer
 *     l'onglet Dashboard et installer les triggers
 * ============================================================
 */

// ============================================================
//  PALETTE NOIR ET OR
// ============================================================
const UI_C = {
  bg:    '#0d0d0d',
  bg2:   '#1a1a1a',
  bg3:   '#212121',
  bg4:   '#2a2a2a',
  gold:  '#D4AF37',
  gold2: '#B8960C',
  white: '#FFFFFF',
  txt2:  '#999999',
  ok:    '#4CAF50',
  warn:  '#FF9800',
  err:   '#F44336',
};

// ============================================================
//  1. MENU PRINCIPAL
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Heureka')
    .addItem('📊 Dashboard visuel',           'showDashboard')
    .addItem('🔄 Actualiser onglet Dashboard', 'refreshDashboardSheet')
    .addSeparator()
    .addItem('💰 Exporter Paie — PDF',         'exportPayrollPdf')
    .addItem('📋 Rapport hebdomadaire',         'weeklyPayrollReport')
    .addItem('📅 Planning de la semaine',       'showWeeklySchedule')
    .addSeparator()
    .addItem('🗑 Vider Vue Live',               'clearLive')
    .addItem('⚙ Initialiser / Réparer',        'initSheetsUi')
    .addItem('ℹ À propos',                     'showAbout')
    .addToUi();
}

// ============================================================
//  2. DASHBOARD — BARRE LATÉRALE HTML COMPLÈTE
// ============================================================
function showDashboard() {
  const data = getDashboardData_();
  const html = HtmlService.createHtmlOutput(buildDashboardHtml_(data))
    .setTitle('🏗 Dashboard Heureka')
    .setWidth(460);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Appelée par le bouton "Actualiser" dans la sidebar via google.script.run
function refreshDashboardSidebar() {
  showDashboard();
}

// Retourne l'URL Drive du PDF — appelé depuis la sidebar
function exportPayrollPdfGetUrl() {
  return exportPayrollPdf_(false);
}

// ============================================================
//  COLLECTE DES DONNÉES DU DASHBOARD
// ============================================================
function getDashboardData_() {
  const tz  = 'America/Toronto';
  const now = new Date();
  const today   = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeNow = Utilities.formatDate(now, tz, 'HH:mm');

  // Bornes semaine (lun–dim)
  const dow = now.getDay() || 7;
  const wS  = new Date(now); wS.setDate(now.getDate() - (dow - 1));
  const wE  = new Date(now); wE.setDate(now.getDate() + (7 - dow));
  const weekStart = Utilities.formatDate(wS, tz, 'yyyy-MM-dd');
  const weekEnd   = Utilities.formatDate(wE, tz, 'yyyy-MM-dd');

  // Lecture onglets (réutilise les fonctions de sheets-script.gs)
  const employees = readSheet(SHEETS.EQUIPE,   employeeFromRow).filter(e => e.active === '1');
  const jobs      = readSheet(SHEETS.JOBS,     jobFromRow);
  const allPunchs = readSheet(SHEETS.PUNCHS,   punchFromRow);
  const live      = readSheet(SHEETS.LIVE,     liveFromRow);
  const horaires  = readSheet(SHEETS.HORAIRES, horaireFromRow);

  const punchsToday = allPunchs.filter(p => p.date === today);
  const punchsWeek  = allPunchs.filter(p => p.date >= weekStart && p.date <= weekEnd && p.punchOut);

  // Heures par employé cette semaine
  const heuresSemaine = employees.map(e => {
    const eps = punchsWeek.filter(p => p.empId === e.id);
    const hrs = eps.reduce((s, p) => s + (p.netHours || 0), 0);
    return { id: e.id, fname: e.fname, lname: e.lname, role: e.role, rate: e.rate, heures: hrs };
  }).filter(e => e.heures > 0).sort((a, b) => b.heures - a.heures);

  // Chantiers actifs enrichis
  const chantiersActifs = jobs.filter(j => j.status === 'actif').map(j => {
    const empActifs  = live.filter(l => l.jobId === j.id && (l.status === 'in' || l.status === 'paused')).length;
    const hrsWeek    = punchsWeek.filter(p => p.jobId === j.id).reduce((s, p) => s + (p.netHours || 0), 0);
    return { ...j, empActifs, hrsWeek };
  });

  // Employés actifs maintenant
  const empActifs = live.filter(l => l.status === 'in' || l.status === 'paused');

  // Alertes
  const alerts = [];
  allPunchs.filter(p => !p.punchOut && p.date < today).forEach(p => {
    alerts.push({ type: 'err', msg: `Punch OUT oublié — ${p.empName} (${p.date} in: ${p.punchIn})` });
  });
  const todayPunchedIds = new Set(punchsToday.map(p => p.empId));
  const seenAbsent      = new Set();
  horaires.filter(h => h.date === today).forEach(h => {
    (h.empIds || []).forEach(eid => {
      if (!todayPunchedIds.has(eid) && !seenAbsent.has(eid)) {
        seenAbsent.add(eid);
        const emp = employees.find(e => e.id === eid);
        if (emp) alerts.push({ type: 'warn', msg: `Absent — ${emp.fname} ${emp.lname} (horaire prévu)` });
      }
    });
  });

  // KPIs
  const totalHresToday = punchsToday.filter(p => p.punchOut).reduce((s, p) => s + (p.netHours || 0), 0);
  const totalHrsWeek   = punchsWeek.reduce((s, p) => s + (p.netHours || 0), 0);
  const totalPayWeek   = punchsWeek.reduce((s, p) => {
    const emp = employees.find(e => e.id === p.empId);
    return s + (p.netHours || 0) * ((emp && emp.rate) || 0);
  }, 0);

  return {
    today, timeNow, weekStart, weekEnd,
    empActifs, punchsToday, heuresSemaine, chantiersActifs, alerts,
    totalHresToday, totalHrsWeek, totalPayWeek,
    nbEmpActifs:       empActifs.length,
    nbChantiersActifs: chantiersActifs.length,
  };
}

// ============================================================
//  BUILDER HTML DU DASHBOARD
// ============================================================
function buildDashboardHtml_(d) {
  const fmtM  = n => '$ ' + Number(n || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  const ini   = name => (name || '?').split(' ').map(w => w[0] || '').join('').toUpperCase().slice(0, 2);
  const avCls = name => (name || '').charCodeAt(0) % 2 === 0 ? 'av-g' : 'av-d';

  /* ── KPI CARDS ── */
  const kpiHtml = `
    <div class="kpi-grid">
      <div class="kpi ok">
        <div class="kpi-val">${d.nbEmpActifs}</div>
        <div class="kpi-lbl">Actifs maintenant</div>
      </div>
      <div class="kpi">
        <div class="kpi-val">${d.totalHresToday.toFixed(1)}h</div>
        <div class="kpi-lbl">Heures aujourd'hui</div>
      </div>
      <div class="kpi">
        <div class="kpi-val">${d.totalHrsWeek.toFixed(1)}h</div>
        <div class="kpi-lbl">Heures — semaine</div>
      </div>
      <div class="kpi warn">
        <div class="kpi-val">${fmtM(d.totalPayWeek)}</div>
        <div class="kpi-lbl">Paie — semaine</div>
      </div>
    </div>`;

  /* ── ALERTES ── */
  const alertsHtml = d.alerts.length ? `
    <div class="sec-title">⚠ Alertes (${d.alerts.length})</div>
    ${d.alerts.map(a => `
      <div class="alert alert-${a.type}">
        <span class="alert-ico">${a.type === 'err' ? '⏰' : '👤'}</span>
        <span>${escHtml_(a.msg)}</span>
      </div>`).join('')}` : '';

  /* ── VUE LIVE ── */
  const liveHtml = d.empActifs.length
    ? `<table class="tbl">
        <thead><tr><th>Employé</th><th>Chantier</th><th>In</th><th>Statut</th></tr></thead>
        <tbody>${d.empActifs.map(l => `
          <tr>
            <td><div style="display:flex;align-items:center;gap:6px">
              <div class="av ${avCls(l.empName)}">${ini(l.empName)}</div>
              <span style="font-weight:700">${escHtml_((l.empName || '').split(' ')[0])}</span>
            </div></td>
            <td class="txt-gold s11">${escHtml_((l.jobName || '—').slice(0, 16))}</td>
            <td class="mono">${l.punchIn || '—'}</td>
            <td><span class="badge b-${l.status === 'in' ? 'ok' : 'warn'}">${l.status === 'in' ? 'En travail' : 'En pause'}</span></td>
          </tr>`).join('')}
        </tbody></table>`
    : '<div class="empty">Aucun employé actif en ce moment</div>';

  /* ── PUNCHS DU JOUR ── */
  const punchsHtml = d.punchsToday.length
    ? `<table class="tbl">
        <thead><tr><th>Employé</th><th>Chantier</th><th>In</th><th>Out</th><th>Hrs</th></tr></thead>
        <tbody>${d.punchsToday.map(p => `
          <tr>
            <td style="font-weight:700">${escHtml_((p.empName || '—').split(' ')[0])}</td>
            <td class="txt2 s11">${escHtml_((p.jobName || '—').slice(0, 14))}</td>
            <td class="mono">${p.punchIn || '—'}</td>
            <td class="mono">${p.punchOut ? p.punchOut : '<span class="badge b-warn">actif</span>'}</td>
            <td class="txt-gold fw8">${p.punchOut ? p.netHours.toFixed(1) + 'h' : '—'}</td>
          </tr>`).join('')}
        </tbody></table>`
    : '<div class="empty">Aucun punch aujourd\'hui</div>';

  /* ── HEURES SEMAINE (barres) ── */
  const maxH = d.heuresSemaine.length ? d.heuresSemaine[0].heures : 1;
  const barsHtml = d.heuresSemaine.length
    ? `<div style="margin-bottom:4px;font-size:10px;color:#999;text-align:right;font-weight:700">${d.totalHrsWeek.toFixed(1)}h total</div>
       ${d.heuresSemaine.map(e => {
         const pct = Math.round((e.heures / maxH) * 100);
         const pay = fmtM(e.heures * (e.rate || 0));
         return `<div class="bar-row">
           <div class="bar-top">
             <span class="bar-name">${escHtml_(e.fname + ' ' + e.lname)}</span>
             <span style="display:flex;gap:8px;align-items:baseline">
               <span class="bar-val">${e.heures.toFixed(1)}h</span>
               <span style="font-size:10px;color:#666">${pay}</span>
             </span>
           </div>
           <div class="bar-bg"><div class="bar-fill" style="width:${pct}%"></div></div>
         </div>`;
       }).join('')}`
    : '<div class="empty">Aucune heure cette semaine</div>';

  /* ── CHANTIERS ACTIFS ── */
  const chantiersHtml = d.chantiersActifs.length
    ? d.chantiersActifs.map(j => `
        <div class="chantier">
          <div>
            <div class="ch-name">${escHtml_(j.name)}</div>
            <div class="ch-sub">${escHtml_(j.client || '—')} · ${j.hrsWeek.toFixed(1)}h cette sem.</div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            ${j.empActifs > 0
              ? `<span class="badge b-ok">▶ ${j.empActifs} actif${j.empActifs > 1 ? 's' : ''}</span>`
              : '<span class="badge b-gray">En attente</span>'}
          </div>
        </div>`).join('')
    : '<div class="empty">Aucun chantier actif</div>';

  return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Google Sans',Arial,sans-serif;background:#0d0d0d;color:#fff;font-size:13px;padding-bottom:24px}
::-webkit-scrollbar{width:4px}
::-webkit-scrollbar-track{background:#1a1a1a}
::-webkit-scrollbar-thumb{background:#D4AF37;border-radius:2px}

/* HEADER */
.hdr{background:linear-gradient(135deg,#1a1a1a,#212121);border-bottom:1px solid rgba(212,175,55,.3);padding:14px 16px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:10}
.hdr-logo{width:38px;height:38px;background:linear-gradient(135deg,#D4AF37,#B8960C);border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:#000;flex-shrink:0;letter-spacing:-0.5px}
.hdr-brand{font-size:13px;font-weight:800;color:#D4AF37;letter-spacing:.2px}
.hdr-sub{font-size:10px;color:#777;text-transform:uppercase;letter-spacing:1px;margin-top:1px}
.hdr-time{margin-left:auto;font-size:20px;font-weight:900;color:#D4AF37;font-variant-numeric:tabular-nums;flex-shrink:0}

/* SECTIONS */
.section{padding:12px 16px}
.sec-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:1.5px;color:#888;margin-bottom:10px;display:flex;align-items:center;gap:7px}
.sec-title::after{content:'';flex:1;height:1px;background:rgba(212,175,55,.12)}

/* KPIs */
.kpi-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.kpi{background:#1a1a1a;border:1px solid rgba(212,175,55,.14);border-radius:10px;padding:12px 14px}
.kpi-val{font-size:22px;font-weight:900;color:#D4AF37;line-height:1;letter-spacing:-0.5px}
.kpi-lbl{font-size:10px;color:#888;margin-top:4px;text-transform:uppercase;letter-spacing:.5px}
.kpi.ok .kpi-val{color:#4CAF50}
.kpi.warn .kpi-val{color:#FF9800}

/* TABLES */
.tbl{width:100%;border-collapse:collapse}
.tbl th{font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:#D4AF37;padding:6px 8px;border-bottom:1px solid rgba(212,175,55,.2);text-align:left}
.tbl td{padding:7px 8px;font-size:12px;border-bottom:1px solid rgba(255,255,255,.04);vertical-align:middle}
.tbl tr:last-child td{border-bottom:none}
.tbl tr:hover td{background:rgba(212,175,55,.04)}

/* BADGES */
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap}
.b-ok{background:rgba(76,175,80,.15);color:#4CAF50;border:1px solid rgba(76,175,80,.3)}
.b-warn{background:rgba(255,152,0,.15);color:#FF9800;border:1px solid rgba(255,152,0,.3)}
.b-gold{background:rgba(212,175,55,.13);color:#D4AF37;border:1px solid rgba(212,175,55,.3)}
.b-gray{background:rgba(150,150,150,.1);color:#888;border:1px solid rgba(150,150,150,.2)}
.b-err{background:rgba(244,67,54,.12);color:#F44336;border:1px solid rgba(244,67,54,.3)}

/* ALERTES */
.alert{border-radius:8px;padding:9px 12px;margin-bottom:7px;font-size:11px;font-weight:600;display:flex;align-items:flex-start;gap:8px;line-height:1.4}
.alert-err{background:rgba(244,67,54,.07);border-left:3px solid #F44336;color:#F44336}
.alert-warn{background:rgba(255,152,0,.07);border-left:3px solid #FF9800;color:#FF9800}
.alert-ico{flex-shrink:0;margin-top:1px}

/* BARRES */
.bar-row{margin-bottom:9px}
.bar-top{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:3px}
.bar-name{font-size:12px;font-weight:700}
.bar-val{font-size:12px;color:#D4AF37;font-weight:800}
.bar-bg{height:6px;background:#212121;border-radius:3px;overflow:hidden}
.bar-fill{height:100%;background:linear-gradient(90deg,#D4AF37,#B8960C);border-radius:3px}

/* CHANTIERS */
.chantier{background:#1a1a1a;border:1px solid rgba(212,175,55,.13);border-radius:9px;padding:10px 13px;margin-bottom:7px;display:flex;align-items:center;justify-content:space-between;gap:10px}
.ch-name{font-size:12px;font-weight:800;color:#fff}
.ch-sub{font-size:10px;color:#888;margin-top:2px}

/* AVATARS */
.av{width:24px;height:24px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:9px;font-weight:800;flex-shrink:0}
.av-g{background:linear-gradient(135deg,#D4AF37,#B8960C);color:#000}
.av-d{background:linear-gradient(135deg,#2a2a2a,#111);color:#D4AF37;border:1px solid rgba(212,175,55,.35)}

/* BOUTONS */
.btn{display:block;width:calc(100% - 32px);margin:0 16px;background:linear-gradient(135deg,#D4AF37,#B8960C);color:#000;border:none;border-radius:9px;padding:14px;font-size:13px;font-weight:800;cursor:pointer;text-align:center;letter-spacing:.3px;font-family:inherit;transition:opacity .2s}
.btn:hover{opacity:.88}
.btn:active{transform:scale(.98)}
.btn-sec{background:transparent;color:#D4AF37;border:1px solid rgba(212,175,55,.35)}
.btn-sec:hover{background:rgba(212,175,55,.07)}
.spinner{display:none;width:16px;height:16px;border:2px solid rgba(0,0,0,.3);border-top:2px solid #000;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto}
@keyframes spin{to{transform:rotate(360deg)}}

/* UTILITAIRES */
.txt-gold{color:#D4AF37}.fw8{font-weight:800}.mono{font-variant-numeric:tabular-nums}.txt2{color:#888}.s11{font-size:11px}
.empty{text-align:center;padding:20px;color:#555;font-size:12px}
.gap{height:8px}
.msg{font-size:11px;min-height:18px;text-align:center;padding:4px 16px;font-weight:600}
.msg.ok{color:#4CAF50}.msg.err{color:#F44336}
</style>
</head>
<body>

<!-- EN-TÊTE -->
<div class="hdr">
  <div class="hdr-logo">GH</div>
  <div>
    <div class="hdr-brand">Gestions Heureka</div>
    <div class="hdr-sub">Dashboard · ${d.today}</div>
  </div>
  <div class="hdr-time">${d.timeNow}</div>
</div>

<!-- KPI -->
<div class="section">
  <div class="sec-title">📊 Résumé du jour</div>
  ${kpiHtml}
</div>

<!-- ALERTES -->
${d.alerts.length ? `<div class="section" style="padding-top:0">
  <div class="sec-title">⚠ Alertes (${d.alerts.length})</div>
  ${alertsHtml}
</div>` : ''}

<!-- VUE LIVE -->
<div class="section">
  <div class="sec-title">🟢 Vue live (${d.nbEmpActifs})</div>
  ${liveHtml}
</div>

<!-- PUNCHS DU JOUR -->
<div class="section">
  <div class="sec-title">⏱ Punchs du jour (${d.punchsToday.length})</div>
  ${punchsHtml}
</div>

<!-- HEURES SEMAINE -->
<div class="section">
  <div class="sec-title">📅 Heures semaine (${d.weekStart} → ${d.weekEnd})</div>
  ${barsHtml}
</div>

<!-- CHANTIERS ACTIFS -->
<div class="section">
  <div class="sec-title">🏗 Chantiers actifs (${d.nbChantiersActifs})</div>
  ${chantiersHtml}
</div>

<!-- ACTIONS -->
<div style="padding:12px 0 4px">
  <button class="btn" id="pdf-btn" onclick="exportPdf()">💰 Exporter Paie PDF — Semaine</button>
  <div id="pdf-msg" class="msg"></div>
  <div class="gap"></div>
  <button class="btn btn-sec" onclick="actualisertableau()">🔄 Actualiser onglet Dashboard</button>
  <div id="tbl-msg" class="msg"></div>
</div>

<div style="font-size:10px;color:#444;text-align:center;padding:14px 0 6px;letter-spacing:.5px">
  GESTIONS HEUREKA · ${d.today} ${d.timeNow}
</div>

<script>
function exportPdf() {
  const btn = document.getElementById('pdf-btn');
  const msg = document.getElementById('pdf-msg');
  btn.textContent = '⏳ Génération en cours...';
  btn.disabled = true;
  google.script.run
    .withSuccessHandler(url => {
      btn.textContent = '✓ PDF prêt — Cliquer pour ouvrir';
      btn.disabled = false;
      btn.onclick = () => window.open(url, '_blank');
      msg.textContent = 'Fichier sauvegardé dans Google Drive';
      msg.className = 'msg ok';
    })
    .withFailureHandler(err => {
      btn.textContent = '💰 Exporter Paie PDF — Semaine';
      btn.disabled = false;
      msg.textContent = 'Erreur: ' + err;
      msg.className = 'msg err';
    })
    .exportPayrollPdfGetUrl();
}
function actualisertableau() {
  const msg = document.getElementById('tbl-msg');
  msg.textContent = 'Actualisation...';
  msg.className = 'msg';
  google.script.run
    .withSuccessHandler(() => {
      msg.textContent = '✓ Tableau Dashboard mis à jour';
      msg.className = 'msg ok';
    })
    .withFailureHandler(err => {
      msg.textContent = 'Erreur: ' + err;
      msg.className = 'msg err';
    })
    .refreshDashboardSheet();
}
</script>
</body>
</html>`;
}

// ============================================================
//  3. ONGLET DASHBOARD FORMATÉ (noir et or dans Sheets)
// ============================================================
function refreshDashboardSheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Dashboard';
  let dash   = ss.getSheetByName(name);
  if (!dash) { dash = ss.insertSheet(name, 0); }

  dash.clear();
  dash.clearFormats();
  dash.setTabColor(UI_C.gold);

  const data = getDashboardData_();
  const tz   = 'America/Toronto';
  const now  = new Date();
  const ts   = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');

  // Largeurs de colonnes
  const colW = [30, 180, 140, 120, 120, 120, 140, 140, 30];
  colW.forEach((w, i) => dash.setColumnWidth(i + 1, w));

  // Fond global noir
  dash.getRange(1, 1, 300, colW.length).setBackground(UI_C.bg).setFontColor(UI_C.white).setFontFamily('Arial');

  let r = 1;

  // ── TITRE ──────────────────────────────────────────────
  dash.setRowHeight(r, 72);
  dash.getRange(r, 1, 1, 9).merge()
    .setValue('🏗  GESTIONS HEUREKA — Dashboard  ·  ' + ts)
    .setBackground(UI_C.bg2)
    .setFontColor(UI_C.gold)
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  addBorderBottom_(dash, r, 1, 9, UI_C.gold2);
  r++;

  // ── KPI ────────────────────────────────────────────────
  dash.setRowHeight(r, 16); r++;  // spacer

  dash.setRowHeight(r, 66);

  setCell_(dash, r, 2, 2, data.nbEmpActifs + '\nActifs maintenant',
           '#0d2b1a', UI_C.ok);
  setCell_(dash, r, 4, 2, data.totalHresToday.toFixed(1) + 'h\nHeures aujourd\'hui',
           UI_C.bg2, UI_C.gold);
  setCell_(dash, r, 6, 2, data.totalHrsWeek.toFixed(1) + 'h\nHeures — semaine',
           UI_C.bg2, UI_C.gold);
  setCell_(dash, r, 8, 1, fmtM_(data.totalPayWeek) + '\nPaie — semaine',
           '#1f1600', UI_C.warn);
  r++;

  // ── ALERTES ────────────────────────────────────────────
  if (data.alerts.length) {
    r++;
    secTitle_(dash, r, '⚠  ALERTES (' + data.alerts.length + ')', UI_C.err);
    r++;
    data.alerts.forEach(a => {
      dash.setRowHeight(r, 28);
      dash.getRange(r, 2, 1, 7).merge()
        .setValue((a.type === 'err' ? '⏰  ' : '👤  ') + a.msg)
        .setBackground(a.type === 'err' ? '#1f0808' : '#1f1400')
        .setFontColor(a.type === 'err' ? UI_C.err : UI_C.warn)
        .setFontSize(10).setFontWeight('bold')
        .setVerticalAlignment('middle');
      r++;
    });
  }

  // ── VUE LIVE ───────────────────────────────────────────
  r++;
  secTitle_(dash, r, '🟢  VUE LIVE  (' + data.nbEmpActifs + ')');
  r++;
  tableHeader_(dash, r, ['Employé', 'Chantier', 'Punch In', 'Statut', 'Dernière MAJ'], 2);
  r++;
  if (data.empActifs.length) {
    data.empActifs.forEach(l => {
      dash.setRowHeight(r, 24);
      const isIn = l.status === 'in';
      dash.getRange(r, 1, 1, 9).setBackground(isIn ? '#0d2b1a' : '#201400');
      setDataRow_(dash, r, [l.empName||'—', l.jobName||'—', l.punchIn||'—',
                             isIn ? '▶ En travail' : '⏸ En pause', l.lastMaj||'—']);
      dash.getRange(r, 5).setFontColor(isIn ? UI_C.ok : UI_C.warn).setFontWeight('bold');
      r++;
    });
  } else {
    emptyRow_(dash, r, 'Aucun employé actif en ce moment'); r++;
  }

  // ── PUNCHS DU JOUR ─────────────────────────────────────
  r++;
  secTitle_(dash, r, '⏱  PUNCHS DU JOUR  (' + data.punchsToday.length + ')');
  r++;
  tableHeader_(dash, r, ['Employé', 'Chantier', 'Punch In', 'Punch Out', 'Heures nettes'], 2);
  r++;
  if (data.punchsToday.length) {
    data.punchsToday.forEach(p => {
      dash.setRowHeight(r, 24);
      const done = !!p.punchOut;
      dash.getRange(r, 1, 1, 9).setBackground(done ? '#0d1a0d' : '#0d0d1f');
      setDataRow_(dash, r, [p.empName||'—', p.jobName||'—', p.punchIn||'—',
                             p.punchOut||'(actif)', done ? p.netHours.toFixed(2)+'h' : '—']);
      if (done) dash.getRange(r, 6).setFontColor(UI_C.gold).setFontWeight('bold');
      r++;
    });
  } else {
    emptyRow_(dash, r, 'Aucun punch enregistré aujourd\'hui'); r++;
  }

  // ── HEURES SEMAINE ─────────────────────────────────────
  r++;
  secTitle_(dash, r, '📅  HEURES PAR EMPLOYÉ — SEMAINE  (' + data.weekStart + ' → ' + data.weekEnd + ')');
  r++;
  tableHeader_(dash, r, ['Employé', 'Rôle', 'Heures nettes', 'Taux/h', 'Montant semaine'], 2);
  r++;
  if (data.heuresSemaine.length) {
    data.heuresSemaine.forEach(e => {
      dash.setRowHeight(r, 24);
      dash.getRange(r, 1, 1, 9).setBackground(UI_C.bg2);
      const pay = e.heures * (e.rate || 0);
      setDataRow_(dash, r, [e.fname + ' ' + e.lname, e.role,
                             e.heures.toFixed(2) + 'h', fmtM_(e.rate), fmtM_(pay)]);
      dash.getRange(r, 4).setFontColor(UI_C.gold).setFontWeight('bold');
      dash.getRange(r, 6).setFontColor(UI_C.gold).setFontWeight('bold');
      r++;
    });
    // Total
    const totalPay = data.heuresSemaine.reduce((s, e) => s + e.heures * (e.rate || 0), 0);
    dash.setRowHeight(r, 28);
    dash.getRange(r, 2, 1, 3).merge()
      .setValue('TOTAL')
      .setBackground(UI_C.bg4).setFontColor(UI_C.gold)
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('right');
    dash.getRange(r, 5, 1, 2).merge()
      .setValue(data.totalHrsWeek.toFixed(2) + 'h');
    dash.getRange(r, 7).setValue(fmtM_(totalPay))
      .setBackground(UI_C.bg4).setFontColor(UI_C.gold)
      .setFontWeight('bold').setFontSize(11);
    r++;
  } else {
    emptyRow_(dash, r, 'Aucune heure enregistrée cette semaine'); r++;
  }

  // ── CHANTIERS ACTIFS ───────────────────────────────────
  r++;
  secTitle_(dash, r, '🏗  CHANTIERS ACTIFS  (' + data.nbChantiersActifs + ')');
  r++;
  tableHeader_(dash, r, ['Chantier', 'Client', 'Début', 'Fin', 'Actifs / Hrs sem.'], 2);
  r++;
  if (data.chantiersActifs.length) {
    data.chantiersActifs.forEach(j => {
      dash.setRowHeight(r, 24);
      dash.getRange(r, 1, 1, 9).setBackground(UI_C.bg2);
      setDataRow_(dash, r, [j.name, j.client||'—', j.start||'—', j.end||'—',
                             j.empActifs + ' actif(s) · ' + j.hrsWeek.toFixed(1) + 'h sem.']);
      if (j.empActifs > 0) dash.getRange(r, 6).setFontColor(UI_C.ok).setFontWeight('bold');
      r++;
    });
  } else {
    emptyRow_(dash, r, 'Aucun chantier actif'); r++;
  }

  // ── PIED DE PAGE ───────────────────────────────────────
  r++;
  dash.getRange(r, 1, 1, 9).merge()
    .setValue('Gestions Heureka · Actualisé le ' + ts + ' · America/Toronto')
    .setBackground(UI_C.bg2).setFontColor('#555555')
    .setFontSize(9).setHorizontalAlignment('center').setItalic(true);

  dash.setFrozenRows(1);
  ss.setActiveSheet(dash);
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard actualisé ✓', '🏗 Heureka', 4);
}

// ============================================================
//  4. EXPORT PAIE — PDF (semaine courante)
// ============================================================

// Version menu (affiche alerte)
function exportPayrollPdf() {
  const url = exportPayrollPdf_(true);
  if (url) {
    SpreadsheetApp.getUi().alert(
      '✅ PDF créé avec succès!',
      'Fichier disponible dans Google Drive.\n\nOuvrir depuis Drive ou via le Dashboard.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Version sidebar (retourne l'URL)
function exportPayrollPdf_(showAlert) {
  const tz  = 'America/Toronto';
  const now = new Date();
  const dow = now.getDay() || 7;
  const wS  = new Date(now); wS.setDate(now.getDate() - (dow - 1));
  const wE  = new Date(now); wE.setDate(now.getDate() + (7 - dow));
  const from = Utilities.formatDate(wS, tz, 'yyyy-MM-dd');
  const to   = Utilities.formatDate(wE, tz, 'yyyy-MM-dd');

  const employees = readSheet(SHEETS.EQUIPE, employeeFromRow).filter(e => e.active === '1');
  const punchs    = getPunchs({ from, to });

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const tmpName   = '_Paie_Export_Tmp';
  let   tmp       = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  buildPayrollSheet_(tmp, employees, punchs, from, to);

  // Export PDF via URL API (meilleur rendu que getAs)
  const ssId    = ss.getId();
  const gid     = tmp.getSheetId();
  const pdfUrl  = 'https://docs.google.com/spreadsheets/d/' + ssId +
    '/export?format=pdf&gid=' + gid +
    '&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false' +
    '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';

  const token   = ScriptApp.getOAuthToken();
  const pdfBlob = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: 'Bearer ' + token }
  }).getBlob().setName('Paie_Heureka_' + from + '_' + to + '.pdf');

  const file = DriveApp.createFile(pdfBlob);
  const url  = file.getUrl();

  // Supprimer onglet temporaire
  ss.deleteSheet(tmp);

  return url;
}

function buildPayrollSheet_(sheet, employees, punchs, from, to) {
  const OT_HOURS = 40, OT_RATE = 1.5;
  const COLS     = 9;

  sheet.getRange(1, 1, 200, COLS).setBackground(UI_C.bg).setFontColor(UI_C.white).setFontFamily('Arial');

  [40, 170, 100, 90, 80, 100, 100, 100, 100].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Titre
  sheet.setRowHeight(1, 70);
  sheet.getRange(1, 1, 1, COLS).merge()
    .setValue('GESTIONS HEUREKA — Rapport de Paie\n' + from + '  au  ' + to)
    .setBackground(UI_C.bg2).setFontColor(UI_C.gold)
    .setFontSize(15).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  // En-têtes
  sheet.setRowHeight(2, 30);
  const hdrs = ['Employé', 'Rôle', 'Hrs total', 'Hrs rég.', 'Hrs sup.', 'Taux/h', 'Paie rég.', 'Paie sup.', 'TOTAL'];
  sheet.getRange(2, 1, 1, COLS).setValues([hdrs])
    .setBackground(UI_C.bg4).setFontColor(UI_C.gold)
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(2);

  let r = 3, grandTotal = 0;
  const rows = [];

  employees.forEach(emp => {
    const eps = punchs.filter(p => p.empId === emp.id && p.punchOut);
    const tot = eps.reduce((s, p) => s + (p.netHours || 0), 0);
    if (!tot) return;
    const reg   = Math.min(tot, OT_HOURS);
    const ot    = Math.max(0, tot - OT_HOURS);
    const rate  = parseFloat(emp.rate) || 0;
    const total = reg * rate + ot * rate * OT_RATE;
    grandTotal += total;
    rows.push([
      emp.fname + ' ' + emp.lname, emp.role,
      parseFloat(tot.toFixed(2)),
      parseFloat(reg.toFixed(2)),
      parseFloat(ot.toFixed(2)),
      parseFloat(rate.toFixed(2)),
      parseFloat((reg * rate).toFixed(2)),
      parseFloat((ot * rate * OT_RATE).toFixed(2)),
      parseFloat(total.toFixed(2)),
    ]);
  });

  if (rows.length) {
    const dataRange = sheet.getRange(r, 1, rows.length, COLS);
    dataRange.setValues(rows)
      .setBackground(UI_C.bg2).setFontColor(UI_C.white).setFontSize(10);
    // Alternance de lignes
    rows.forEach((_, i) => {
      if (i % 2 === 0) sheet.getRange(r + i, 1, 1, COLS).setBackground(UI_C.bg3);
    });
    // Colonnes montants en or
    sheet.getRange(r, 7, rows.length, 3).setFontColor(UI_C.gold).setFontWeight('bold');
    r += rows.length;
  }

  // Ligne total
  sheet.setRowHeight(r, 36);
  sheet.getRange(r, 1, 1, 8).merge()
    .setValue('GRAND TOTAL')
    .setBackground(UI_C.bg4).setFontColor(UI_C.gold)
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(r, 9)
    .setValue(parseFloat(grandTotal.toFixed(2)))
    .setBackground(UI_C.gold).setFontColor('#000')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  r++;

  // Pied
  sheet.setRowHeight(r, 24);
  sheet.getRange(r, 1, 1, COLS).merge()
    .setValue('Gestions Heureka · Généré le ' + Utilities.formatDate(new Date(), 'America/Toronto', 'dd/MM/yyyy HH:mm'))
    .setBackground(UI_C.bg2).setFontColor('#555')
    .setFontSize(9).setItalic(true).setHorizontalAlignment('center');
}

// ============================================================
//  5. PLANNING SEMAINE — BOÎTE DE DIALOGUE HTML
// ============================================================
function showWeeklySchedule() {
  const tz  = 'America/Toronto';
  const now = new Date();
  const dow = now.getDay() || 7;
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const DAYS  = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];

  const weekDates = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(now); d.setDate(now.getDate() - (dow - 1) + i);
    weekDates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }

  const employees = readSheet(SHEETS.EQUIPE,   employeeFromRow).filter(e => e.active === '1');
  const shifts    = readSheet(SHEETS.HORAIRES, horaireFromRow);
  const punchs    = readSheet(SHEETS.PUNCHS,   punchFromRow);

  // Construire la grille
  const gridRows = employees.map(e => {
    const cells = weekDates.map(date => {
      const dayShifts = shifts.filter(s => s.date === date && (s.empIds || []).includes(e.id));
      const hasPunch  = punchs.some(p => p.empId === e.id && p.date === date);
      return { shifts: dayShifts, hasPunch, isToday: date === today };
    });
    return { emp: e, cells };
  });

  let gridHtml = `
    <div style="display:grid;grid-template-columns:130px repeat(7,1fr);gap:0;border:1px solid rgba(212,175,55,.2);border-radius:8px;overflow:hidden;font-size:11px">
      <div style="background:#1a1a1a;padding:8px 10px;font-size:10px;font-weight:800;color:#D4AF37;text-transform:uppercase;letter-spacing:1px;border-right:1px solid rgba(212,175,55,.15)">Employé</div>
      ${DAYS.map((d, i) => `
        <div style="background:#1a1a1a;padding:8px 6px;text-align:center;font-size:10px;font-weight:800;color:${weekDates[i] === today ? '#D4AF37' : '#999'};text-transform:uppercase;letter-spacing:.8px;border-right:1px solid rgba(212,175,55,.1);${weekDates[i] === today ? 'border-top:2px solid #D4AF37' : ''}">
          ${d}<br><span style="font-size:13px;font-weight:900">${weekDates[i].slice(-2)}</span>
        </div>`).join('')}
      ${gridRows.map(row => `
        <div style="background:#0d0d0d;padding:8px 10px;border-top:1px solid rgba(255,255,255,.04);border-right:1px solid rgba(212,175,55,.15);display:flex;align-items:center;gap:7px">
          <div style="width:22px;height:22px;border-radius:50%;background:${row.emp.fname.charCodeAt(0)%2===0?'linear-gradient(135deg,#D4AF37,#B8960C)':'linear-gradient(135deg,#2a2a2a,#111)'};display:flex;align-items:center;justify-content:center;font-size:8px;font-weight:800;color:${row.emp.fname.charCodeAt(0)%2===0?'#000':'#D4AF37'};flex-shrink:0">${row.emp.fname.slice(0,1).toUpperCase()+(row.emp.lname||'').slice(0,1).toUpperCase()}</div>
          <span style="font-weight:700;font-size:11px;color:#fff">${row.emp.fname}</span>
        </div>
        ${row.cells.map(c => `
          <div style="background:${c.isToday?'rgba(212,175,55,.04)':'#0d0d0d'};border-top:1px solid rgba(255,255,255,.04);border-right:1px solid rgba(212,175,55,.08);padding:5px;min-height:52px;${c.isToday?'border-top:1px solid rgba(212,175,55,.25)':''}">
            ${c.shifts.map(s => `<div style="background:rgba(212,175,55,.12);border:1px solid rgba(212,175,55,.25);border-radius:4px;padding:2px 5px;margin-bottom:2px;font-size:9px;font-weight:700;color:#D4AF37;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml_(s.jobName||'—')}">${s.start}–${s.end}</div>`).join('')}
            ${c.hasPunch ? '<div style="font-size:8px;color:#4CAF50;font-weight:800;margin-top:2px">✓ punchté</div>' : ''}
          </div>`).join('')}`).join('')}
    </div>`;

  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<style>
  body{background:#0d0d0d;color:#fff;font-family:'Google Sans',Arial,sans-serif;margin:0;padding:16px;font-size:12px}
  h2{color:#D4AF37;margin-bottom:4px;font-size:15px}
  .sub{color:#888;font-size:11px;margin-bottom:16px}
  ::-webkit-scrollbar{height:4px;width:4px}
  ::-webkit-scrollbar-thumb{background:#D4AF37;border-radius:2px}
  ::-webkit-scrollbar-track{background:#1a1a1a}
</style></head><body>
<h2>📅 Planning de la semaine</h2>
<div class="sub">${weekDates[0]} — ${weekDates[6]}</div>
<div style="overflow-x:auto">${gridHtml}</div>
</body></html>`)
    .setTitle('📅 Planning — Gestions Heureka')
    .setWidth(920).setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, '📅 Planning de la semaine');
}

// ============================================================
//  6. À PROPOS
// ============================================================
function showAbout() {
  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  body{background:#0d0d0d;color:#fff;font-family:'Google Sans',Arial,sans-serif;padding:24px;font-size:13px;text-align:center}
  .logo{width:64px;height:64px;background:linear-gradient(135deg,#D4AF37,#B8960C);border-radius:16px;display:flex;align-items:center;justify-content:center;font-size:24px;font-weight:900;color:#000;margin:0 auto 16px}
  h2{color:#D4AF37;font-size:18px;margin-bottom:6px}
  p{color:#999;font-size:12px;line-height:1.6}
  .sep{height:1px;background:rgba(212,175,55,.15);margin:16px 0}
  .item{color:#fff;font-size:12px;padding:5px 0;display:flex;justify-content:space-between}
  .val{color:#D4AF37;font-weight:700}
</style></head>
<body>
  <div class="logo">GH</div>
  <h2>Gestions Heureka</h2>
  <p>Système de gestion de chantier<br>Construction · Québec, Canada</p>
  <div class="sep"></div>
  <div class="item"><span>Fuseau horaire</span><span class="val">America/Toronto</span></div>
  <div class="item"><span>Version script</span><span class="val">2.0</span></div>
  <div class="item"><span>Onglets actifs</span><span class="val">Équipe · Jobs · Punchs · Vue Live · Horaires · Dashboard</span></div>
  <div class="sep"></div>
  <p style="color:#555;font-size:11px">Toutes les heures et dates utilisent le fuseau America/Toronto (Québec).<br>Données stockées dans Google Sheets, synchronisées depuis admin.html et punch.html.</p>
</body></html>`)
    .setWidth(380).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'ℹ À propos — Gestions Heureka');
}

// ============================================================
//  7. INITIALISATION
// ============================================================
function initSheetsUi() {
  // Créer / réparer l'onglet Dashboard
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Dashboard';
  let   dash = ss.getSheetByName(name);
  if (!dash) {
    dash = ss.insertSheet(name, 0);
  }
  dash.setTabColor(UI_C.gold);

  // S'assurer que tous les onglets de données existent
  if (typeof initSheets === 'function') initSheets();

  // Actualiser le dashboard maintenant
  refreshDashboardSheet();

  SpreadsheetApp.getUi().alert(
    '✅ Interface Heureka prête!',
    'Onglet Dashboard créé et actualisé.\n\n' +
    'Utilisez le menu 🏗 Heureka pour :\n' +
    '  · Ouvrir le Dashboard visuel (barre latérale)\n' +
    '  · Exporter la paie en PDF\n' +
    '  · Voir le planning de la semaine\n\n' +
    'Le Dashboard se trouve en premier onglet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
//  UTILITAIRES PRIVÉS
// ============================================================

function setCell_(sheet, row, col, span, value, bg, fg) {
  const range = span > 1 ? sheet.getRange(row, col, 1, span).merge() : sheet.getRange(row, col);
  range.setValue(value)
    .setBackground(bg || UI_C.bg2)
    .setFontColor(fg || UI_C.gold)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function secTitle_(sheet, row, label, color) {
  sheet.setRowHeight(row, 30);
  sheet.getRange(row, 1, 1, 9).merge()
    .setValue('  ' + label)
    .setBackground(UI_C.bg3)
    .setFontColor(color || UI_C.gold)
    .setFontSize(11).setFontWeight('bold')
    .setVerticalAlignment('middle');
  addBorderBottom_(sheet, row, 1, 9, color || UI_C.gold2);
}

function tableHeader_(sheet, row, labels, startCol) {
  sheet.setRowHeight(row, 26);
  labels.forEach((lbl, i) => {
    sheet.getRange(row, startCol + i)
      .setValue(lbl)
      .setBackground(UI_C.bg4).setFontColor(UI_C.gold)
      .setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
}

function setDataRow_(sheet, row, values, startCol) {
  startCol = startCol || 2;
  sheet.setRowHeight(row, 24);
  values.forEach((v, i) => {
    sheet.getRange(row, startCol + i)
      .setValue(v)
      .setFontSize(10).setVerticalAlignment('middle');
  });
}

function emptyRow_(sheet, row, msg) {
  sheet.setRowHeight(row, 30);
  sheet.getRange(row, 2, 1, 7).merge()
    .setValue(msg)
    .setFontColor('#444444').setFontSize(10).setItalic(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function addBorderBottom_(sheet, row, col, numCols, color) {
  try {
    sheet.getRange(row, col, 1, numCols)
      .setBorder(null, null, true, null, null, null, color || UI_C.gold2,
                 SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) { /* ignore si la version Sheets ne supporte pas la couleur de bordure */ }
}

function fmtM_(n) {
  return '$ ' + Number(n || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

function escHtml_(str) {
  return (str || '').toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
