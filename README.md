<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover"/>
  <title>Quality | Хаб Качества 360°</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      /* ===== Background ===== */
      --bg: #ffffff;
      --bg-card: #ffffff;
      --bg-brand: #fdd835;
      --bg-brand-hover: #fdc435;
      --bg-brand-pressed: #fdb435;
      --bg-backdrop: rgba(0,0,0,0.48);
      /* $color-status-neutral-background — Light Status/Neutral background #f1f1f3 */
      /* $color-palette-amber-600 — DS Amber #ff9419 (ярко оранжевый) */
      --color-orange: #ff9419;
      /* $color-palette-green-600 — DS Green #2ca853 */
      --color-green: #2ca853;
      --bg-neutral: #f1f1f3;

      /* ===== Text ===== */
      --text-primary: rgba(9,11,22,0.94);
      --text-secondary: rgba(25,28,52,0.7);
      --text-muted: rgba(25,28,52,0.48);
      --text-inverse: #ffffff;
      --text-dark-active: rgba(0,0,0,0.87);
      --text-dark-inactive: rgba(0,0,0,0.6);
      --text-dark-disabled: rgba(0,0,0,0.38);
      --text-light-active: rgba(255,255,255,0.87);
      --text-light-inactive: rgba(255,255,255,0.6);
      --text-light-disabled: rgba(255,255,255,0.38);

      /* ===== Border ===== */
      --border: rgba(25,28,52,0.18);
      --border-focused: #202123;
      --border-focus: #212121;
      --border-error: #ff5555;
      --border-width: 1px;

      /* ===== Accent (DS) ===== */
      /* $color-background-brand-hover — Const Background/Brand hover #fdc435 */
      --accent: #fdc435;
      --accent-hover: #fdb435;
      --accent-soft: #fff7d7;
      --accent-purple: #7e00ed;
      --accent-purple-soft: #f3e7fe;
      /* $color-accent-lemon-background — Light Accent/Lemon background #fff7d7 */
      --bg-lemon: #fff7d7;

      /* ===== Status ===== */
      --status-error: #ff5555;
      --status-error-bg: #ffecef;
      --status-success: #10B981;

      /* ===== Radius ===== */
      --radius-2: 2px;
      --radius-4: 4px;
      --radius-6: 6px;
      --radius-8: 8px;
      --radius-10: 10px;
      --radius-12: 12px;
      --radius-16: 16px;
      --radius-20: 20px;
      --radius-24: 24px;
      --radius-30: 30px;
      --radius-circle: 50%;

      /* ===== Icon size ===== */
      --icon-sm: 18px;
      --icon-md: 20px;
      --icon-lg: 24px;

      /* ===== Control height ===== */
      --control-h-sm: 40px;
      --control-h-md: 48px;
      --control-h-lg: 56px;

      /* ===== Spacing ===== */
      --sp-1: 4px;  --sp-2: 8px;  --sp-3: 12px;  --sp-4: 16px;  --sp-5: 20px;
      --sp-6: 24px; --sp-7: 28px; --sp-8: 32px;  --sp-9: 36px;  --sp-10: 40px;

      /* ===== Font ===== */
      --font-text: 'Beeline Sans','Helvetica Neue',Helvetica,Arial,sans-serif;
      --font-header: 'Beeline Sans','Helvetica Neue',Helvetica,Arial,sans-serif;
      --font-code: 'Roboto Mono',monospace;

      /* ===== Shadow ===== */
      --shadow-sm: 0 1px 2px rgba(0,0,0,0.06);
      --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
      --shadow-lg: 0 12px 32px rgba(0,0,0,0.12);
      --shadow-accent: 0 8px 24px rgba(253,196,53,0.20);

      /* ===== Градиенты (бренд) ===== */
      --grad-hero: linear-gradient(135deg, #fdd835 0%, #000000 100%);
      --grad-btn: #000000;
      --grad-card: linear-gradient(160deg, rgba(253,216,53,0.06) 0%, rgba(0,0,0,0.02) 100%);

      /* ===== Прочее ===== */
      --ease: cubic-bezier(0.4, 0, 0.2, 1);
      --icon-color: var(--text-secondary);
      --sidebar-w: 240px;
      --sidebar-w-collapsed: 64px;
    }

    /* ===== DS yellowbe icon font (BeelineIcons.woff2 — ligature approach: <i class="beeline-icons">ligature_name</i>) ===== */
    @font-face {
      font-family: 'BeelineIcons';
      src: url('fonts/BeelineIcons.woff2') format('woff2');
      font-weight: normal;
      font-style: normal;
      font-display: block;
    }
    /* ===== BeelineSans для логотипа "билайн" ===== */
    @font-face {
      font-family: 'BeelineSans';
      src: url('fonts/BeelineSans-Regular.woff2') format('woff2');
      font-weight: 400;
      font-style: normal;
      font-display: swap;
    }
    .beeline-icons {
      font-family: 'BeelineIcons';
      font-weight: bold;
      font-style: normal;
      line-height: 1;
      letter-spacing: normal;
      text-transform: none;
      display: inline-block;
      white-space: nowrap;
      word-wrap: normal;
      direction: ltr;
      -webkit-font-feature-settings: 'liga';
      -webkit-font-smoothing: antialiased;
      font-feature-settings: 'liga';
    }

    body.dark {
      --bg: #121212;
      --bg-card: #1E1E1E;
      --text-primary: rgba(255,255,255,0.94);
      --text-secondary: rgba(255,255,255,0.7);
      --text-muted: rgba(255,255,255,0.48);
      --border: rgba(255,255,255,0.12);
      --border-focused: #ffffff;
      --border-focus: #ffffff;
      --accent-soft: rgba(253,196,53,0.16);
      --accent-purple-soft: rgba(126,0,237,0.18);
      --bg-lemon: rgba(253,196,53,0.14);
      --bg-neutral: rgba(255,255,255,0.10);
      --status-error-bg: rgba(255,85,85,0.14);
      --shadow-sm: 0 1px 2px rgba(0,0,0,0.4);
      --shadow-md: 0 4px 12px rgba(0,0,0,0.4);
      --shadow-lg: 0 12px 32px rgba(0,0,0,0.5);
      --shadow-accent: 0 8px 24px rgba(253,196,53,0.25);
      --grad-card: linear-gradient(160deg, rgba(253,216,53,0.08) 0%, rgba(0,0,0,0.04) 100%);
    }
    body.dark .nav-item { color: rgba(255,255,255,0.85); }
    body.dark .nav-item i { color: rgba(255,255,255,0.55); }
    body.dark .nav-item:hover { background: transparent; color: rgba(255,255,255,0.85); }
    body.dark .nav-item:hover i { color: rgba(255,255,255,0.55); }
    body.dark .nav-item.active { background: var(--bg-neutral); color: var(--text-primary); }
    body.dark .nav-item.active i { color: var(--accent); }
    body.dark .page-title { color: rgba(255,255,255,0.94); }
    body.dark .sidebar { background: #1A1A1A; }
    body.dark .topbar { background: #1A1A1A; }
    body.dark .dir-name { color: rgba(255,255,255,0.94); }
    body.dark .dir-sub { color: rgba(255,255,255,0.6); }
    body.dark .sidebar-divider { opacity: 0.3; }
    * { margin:0; padding:0; box-sizing:border-box; }
    html, body { height: 100%; }
    body {
      font-family: var(--font-text);
      background: var(--bg);
      color: var(--text-primary);
      transition: background 0.3s var(--ease), color 0.3s var(--ease);
      font-weight: 400;
      letter-spacing: 0.2px;
      -webkit-font-smoothing: antialiased;
      overflow-x: hidden;
      text-shadow: none;
    }
    /* ===== Aurora background (entire page) ===== */
    .bg-aurora {
      position: fixed; top: 0; left: 0; right: 0; bottom: 0;
      pointer-events: none; z-index: 0; overflow: hidden;
    }
    .bg-blob {
      position: absolute; border-radius: 50%; filter: blur(80px);
      opacity: 0.35; will-change: transform;
    }
    .bg-blob--1 { width: 500px; height: 500px; top: -120px; left: -100px; background: radial-gradient(circle, rgba(253,216,53,0.6) 0%, transparent 70%); animation: bgBlob1 22s ease-in-out infinite; }
    .bg-blob--2 { width: 450px; height: 450px; top: 10%; right: -120px; background: radial-gradient(circle, rgba(255,148,25,0.45) 0%, transparent 70%); animation: bgBlob2 26s ease-in-out infinite; }
    .bg-blob--3 { width: 420px; height: 420px; bottom: -100px; left: 20%; background: radial-gradient(circle, rgba(126,0,237,0.3) 0%, transparent 70%); animation: bgBlob3 20s ease-in-out infinite; }
    .bg-blob--4 { width: 380px; height: 380px; bottom: 15%; right: 10%; background: radial-gradient(circle, rgba(253,196,53,0.35) 0%, transparent 70%); animation: bgBlob4 24s ease-in-out infinite; }

    body.dark .bg-blob { opacity: 0.5; }
    body.dark .bg-blob--1 { background: radial-gradient(circle, rgba(253,216,53,0.5) 0%, transparent 70%); }
    body.dark .bg-blob--2 { background: radial-gradient(circle, rgba(255,148,25,0.4) 0%, transparent 70%); }
    body.dark .bg-blob--3 { background: radial-gradient(circle, rgba(126,0,237,0.35) 0%, transparent 70%); }
    body.dark .bg-blob--4 { background: radial-gradient(circle, rgba(253,196,53,0.3) 0%, transparent 70%); }

    @keyframes bgBlob1 { 0%,100%{transform:translate(0,0) scale(1)} 33%{transform:translate(80px,40px) scale(1.15)} 66%{transform:translate(-40px,70px) scale(0.9)} }
    @keyframes bgBlob2 { 0%,100%{transform:translate(0,0) scale(1)} 33%{transform:translate(-70px,50px) scale(0.85)} 66%{transform:translate(40px,-40px) scale(1.2)} }
    @keyframes bgBlob3 { 0%,100%{transform:translate(0,0) scale(1)} 50%{transform:translate(50px,-80px) scale(1.1)} }
    @keyframes bgBlob4 { 0%,100%{transform:translate(0,0) scale(1)} 33%{transform:translate(-60px,-30px) scale(1.1)} 66%{transform:translate(30px,60px) scale(0.9)} }

    .main > * { position: relative; z-index: 1; }
    /* ===== App layout ===== */
    .app { display: flex; min-height: 100vh; position: relative; z-index: 1; }
    .sidebar {
      position: fixed; top: 0; left: 0; height: 100vh; width: var(--sidebar-w);
      background: var(--bg-card); border-right: var(--border-width) solid var(--border);
      display: flex; flex-direction: column; z-index: 200;
      transition: transform 0.3s var(--ease); overflow-y: auto;
      box-shadow: 2px 0 12px rgba(0,0,0,0.03);
      padding-top: 20px;
      padding-bottom: 64px;
    }
    body.dark .sidebar { box-shadow: 2px 0 16px rgba(0,0,0,0.3); }
    .sidebar-logo-img { height: 40px; width: auto; object-fit: contain; transition: transform 0.2s var(--ease); display: block; }
    .sidebar-logo-img--dark { display: none; }
    body.dark .sidebar-logo-img--light { display: none; }
    body.dark .sidebar-logo-img--dark { display: block; }
    .sidebar-section-label { padding: var(--sp-4) var(--sp-5) var(--sp-2); font-size: 11px; font-weight: 500; text-transform: uppercase; letter-spacing: 0.08em; color: var(--text-muted); }
    .sidebar-nav { padding: 0 var(--sp-3); display: flex; flex-direction: column; gap: 3px; }
    .nav-item {
      display: flex; align-items: center; gap: var(--sp-3);
      padding: 0 var(--sp-4); height: var(--control-h-md);
      background: transparent; border: none; border-radius: var(--radius-10);
      color: var(--text-secondary); font-family: var(--font-text); font-size: 15px; font-weight: 500;
      cursor: pointer; transition: all 0.15s var(--ease);
      text-align: left; width: 100%; position: relative; text-decoration: none;
    }
    .nav-item i { font-size: 20px; font-weight: 700; width: 20px; text-align: center; flex-shrink: 0; color: var(--text-muted); transition: color 0.15s var(--ease); }
    .nav-item span { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .nav-item:hover { background: transparent; color: var(--text-primary); }
    .nav-item:hover i { color: var(--text-muted); }
    .nav-item.active { background: var(--bg-neutral); color: var(--text-primary); font-weight: 500; }
    .nav-item.active i { color: var(--accent); }
    .nav-item.active::before { content: ''; position: absolute; left: 0; top: 6px; bottom: 6px; width: 4px; border-radius: 0 4px 4px 0; background: var(--bg-brand); }
    .sidebar-divider { height: var(--border-width); background: var(--border); margin: var(--sp-4) var(--sp-5); }
    .main { flex: 1; min-width: 0; margin-left: var(--sidebar-w); display: flex; flex-direction: column; min-height: 100vh; }
    .topbar {
      position: sticky; top: 0; z-index: 50; background: var(--bg-card);
      border-bottom: var(--border-width) solid var(--border);
      padding: 16px 24px; display: flex; align-items: center; gap: var(--sp-4); min-height: 64px;
      backdrop-filter: blur(8px);
      box-shadow: none;
    }
    /* ===== Sidebar collapse / Topbar logo + search (redesign) ===== */
    .sidebar-collapse-btn {
      display: flex; align-items: center; justify-content: center;
      height: var(--control-h-md);
      width: var(--sidebar-w);
      background: var(--bg-card); border: none; border-top: var(--border-width) solid var(--border); cursor: pointer;
      color: var(--text-secondary);
      position: fixed; bottom: 0; left: 0; z-index: 201;
      transition: background 0.15s var(--ease), color 0.15s var(--ease), width 0.3s var(--ease);
    }
    .sidebar-collapse-btn:hover { background: var(--accent-soft); color: var(--accent); }
    .sidebar-collapse-btn i { font-size: 20px; transition: none; }
    body.sidebar-collapsed { --sidebar-w: 64px; }
    body.sidebar-collapsed .nav-item span { display: none; }
    body.sidebar-collapsed .sidebar-section-label { visibility: hidden; }
    body.sidebar-collapsed .nav-item { justify-content: center; padding: 0; }
    body.sidebar-collapsed .nav-item i { margin: 0; }
    .sidebar-tt {
      position: fixed; z-index: 400; pointer-events: none;
      background: var(--bg-card); color: var(--text-primary);
      border: 1px solid var(--border); border-radius: var(--radius-8);
      padding: 6px 12px; font-size: 13px; font-weight: 500; white-space: nowrap;
      box-shadow: var(--shadow-md);
      opacity: 0; visibility: hidden;
      transition: opacity 0.15s var(--ease), visibility 0.15s var(--ease);
      transform: translateY(-50%);
    }
    .sidebar-tt.show { opacity: 1; visibility: visible; }
    @media (max-width: 900px) {
      .sidebar-collapse-btn { position: static !important; width: 100% !important; margin-top: auto; }
    }
    .topbar-logo { height: 40px; display: flex; align-items: center; gap: 6px; flex-shrink: 0; }
    .logo-sphere { height: 40px; flex-shrink: 0; display: flex; align-items: center; }
    .logo-sphere .sidebar-logo-img { height: 32px; width: auto; object-fit: contain; }
    .logo-sphere .sidebar-logo-img--dark { display: none; }
    .logo-sphere .sidebar-logo-img--light { display: block; }
    body.dark .logo-sphere .sidebar-logo-img--light { display: none; }
    body.dark .logo-sphere .sidebar-logo-img--dark { display: block; }
    /* light/dark переключение сохраняется через существующие .sidebar-logo-img--light/dark правила */
    .topbar-search {
      display: flex; align-items: center; gap: 8px;
      padding: 8px 16px; border-radius: var(--radius-30);
      background: var(--bg); border: 1px solid var(--border);
      width: 280px; height: 40px; transition: all 0.2s var(--ease);
    }
    .topbar-search:focus-within { border-color: var(--border-focused); box-shadow: 0 0 0 3px rgba(33,33,33,0.12); }
    .topbar-search i { color: var(--text-muted); font-size: 16px; }
    .topbar-search input { border: none; background: transparent; outline: none; flex: 1; font-family: var(--font-text); font-size: 14px; color: var(--text-primary); }
    .topbar-search input::placeholder { color: var(--text-muted); }
    @media (max-width: 900px) { .topbar-search { display: none; } }
    :focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
    .topbar .crumb { font-size: 12px; color: var(--text-muted); font-weight: 500; }
    .page-title { font-size: 26px; font-weight: 500; letter-spacing: -0.02em; color: var(--text-primary); }
    .topbar-spacer { flex: 1; }
    .topbar-actions { display: flex; gap: var(--sp-2); align-items: center; }
    #sync-badge {
      display: none; align-items: center; gap: 6px;
      padding: 6px 12px; border-radius: var(--radius-30);
      background: var(--accent-soft); color: var(--accent);
      font-size: 12px; font-weight: 500;
      animation: pulse 1.2s ease infinite;
    }
    #sync-badge i { font-size: 12px; }
    @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.6; } }
    .content { padding: var(--sp-8); flex: 1; }
    .content-inner { max-width: 1200px; margin: 0 auto; }
    .menu-toggle { display: none; background: var(--bg-card); border: var(--border-width) solid var(--border); width: var(--control-h-md); height: var(--control-h-md); border-radius: var(--radius-10); cursor: pointer; align-items: center; justify-content: center; color: var(--text-secondary); flex-shrink: 0; transition: all 0.15s var(--ease); }
    .menu-toggle:hover { color: var(--accent); border-color: var(--accent); background: var(--accent-soft); }
    .menu-toggle i { font-size: 20px; }
    .sidebar-overlay { display: none; position: fixed; inset: 0; background: var(--bg-backdrop); z-index: 150; opacity: 0; visibility: hidden; transition: opacity 0.3s ease, visibility 0.3s ease; backdrop-filter: blur(2px); }
    .sidebar-overlay.open { opacity: 1; visibility: visible; }
    .fb-overlay {
      position: fixed; inset: 0; background: var(--bg-backdrop);
      z-index: 2000; display: flex; align-items: center; justify-content: center;
      opacity: 0; visibility: hidden; transition: opacity 0.25s ease, visibility 0.25s ease;
      backdrop-filter: blur(4px); padding: var(--sp-4);
    }
    .fb-overlay.open { opacity: 1; visibility: visible; }
    .fb-modal {
      background: var(--bg-card); border: var(--border-width) solid var(--border);
      border-radius: var(--radius-20); padding: var(--sp-8);
      max-width: 440px; width: 100%; text-align: center;
      box-shadow: var(--shadow-lg);
      transform: translateY(20px) scale(0.97);
      transition: transform 0.25s var(--ease);
      position: relative;
    }
    .fb-overlay.open .fb-modal { transform: translateY(0) scale(1); }
    .fb-modal h3 { font-size: 18px; margin-bottom: var(--sp-2); font-weight: 500; color: var(--text-primary); display: flex; align-items: center; justify-content: center; gap: var(--sp-2); }
    .fb-modal p { font-size: 13px; color: var(--text-muted); }
    .fb-close {
      position: absolute; top: var(--sp-4); right: var(--sp-4);
      background: none; border: none; color: var(--text-muted);
      cursor: pointer; font-size: 20px; transition: color 0.2s ease;
      width: 32px; height: 32px; border-radius: var(--radius-8);
      display: flex; align-items: center; justify-content: center;
    }
    .fb-close:hover { color: var(--status-error); background: var(--status-error-bg); }
    @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
    .section-title {
      font-size: 26px; font-weight: 500;
      margin: 0 0 var(--sp-5); color: var(--text-primary);
      display: flex; align-items: center; gap: var(--sp-2);
    }
    .section-title i { color: var(--accent); font-size: 28px; }
    .section-title--sub {
      font-size: 16px; font-weight: 400;
      color: var(--text-secondary);
      margin-top: var(--sp-4);
    }
    .section-title--sub i { color: var(--text-secondary); font-size: 18px; }
    .arrow-down { display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; color: var(--accent); animation: bounceDown 1.5s ease-in-out infinite; }
    .arrow-down i { font-size: 22px; }
    @keyframes bounceDown { 0%,100%{transform:translateY(0)} 50%{transform:translateY(8px)} }
    /* ===== HERO SECTION — всегда тёмная плашка (#121212), градиент совпадает с тёмной темой ===== */
    .hero { position: relative; border-radius: var(--radius-24); overflow: hidden; margin-bottom: var(--sp-8); min-height: 560px; display: flex; align-items: center; animation: fadeUp 0.7s ease both; background: #121212; }
    .hero-bg { position: absolute; inset: 0; background: linear-gradient(135deg, rgba(253,216,53,0.15) 0%, rgba(18,18,18,0.5) 35%, #121212 70%); z-index: 0; }
    .hero-bg::after { content: ''; position: absolute; inset: 0; background-image: linear-gradient(rgba(255,255,255,0.02) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.02) 1px, transparent 1px); background-size: 40px 40px; mask-image: linear-gradient(180deg, transparent 0%, rgba(0,0,0,0.3) 50%, transparent 100%); -webkit-mask-image: linear-gradient(180deg, transparent 0%, rgba(0,0,0,0.3) 50%, transparent 100%); }
    /* Тёмная тема — фон hero прозрачный, aurora просвечивает, граница исчезает */
    body.dark .hero { border-radius: 0; border: none; box-shadow: none; margin-bottom: var(--sp-6); background: transparent; }
    body.dark .hero-bg { background: transparent; }
    body.dark .hero-bg::after { background-image: none; }
    /* Светлая тема — фон hero как в тёмной теме (#121212), без рамки и тени */
    body:not(.dark) .hero { border: none; box-shadow: none; }

    /* ===== HERO-BAND: светлая тема — тёмная плашка продлена до низа карточек направлений ===== */
    body:not(.dark) .hero-band {
      background: #121212;
      border-radius: var(--radius-24);
      overflow: hidden;
      padding: 0 var(--sp-8) var(--sp-8);
      margin-bottom: var(--sp-8);
    }
    body:not(.dark) .hero-band .hero { border-radius: 0; margin-bottom: 0; background: transparent; border: none; box-shadow: none; outline: none; }
    body:not(.dark) .hero-band .hero-bg { background: transparent; }
    body:not(.dark) .hero-band .hero-bg::after { background-image: none; }
    body:not(.dark) .hero-band .section-title,
    body:not(.dark) .hero-band .section-title--sub { color: rgba(255,255,255,0.94); }
    body:not(.dark) .hero-band .section-title--sub { color: rgba(255,255,255,0.72); }
    body:not(.dark) .hero-band .section-title i,
    body:not(.dark) .hero-band .section-title--sub i { color: var(--accent); }
    body:not(.dark) .hero-band .arrow-down { color: var(--accent); }
    /* Карточки внутри band — как в тёмной теме */
    body:not(.dark) .hero-band .info-banner,
    body:not(.dark) .hero-band .role-card,
    body:not(.dark) .hero-band .dir-card { background: rgba(255,255,255,0.10); }
    body:not(.dark) .hero-band .info-banner-text,
    body:not(.dark) .hero-band .role-desc,
    body:not(.dark) .hero-band .dir-sub { color: rgba(255,255,255,0.7); }
    body:not(.dark) .hero-band .info-banner-text strong,
    body:not(.dark) .hero-band .role-name,
    body:not(.dark) .hero-band .dir-name { color: rgba(255,255,255,0.94); }
    body:not(.dark) .hero-band .role-card:hover,
    body:not(.dark) .hero-band .dir-card:hover { background: #1E1E1E; }
    body:not(.dark) .hero-band .card-emoji,
    body:not(.dark) .hero-band .role-icon { box-shadow: 0 0 0 1px rgba(255,255,255,0.06); }

    /* ===== Мягкое фоновое свечение ===== */
    /* ===== Мягкое фоновое свечение ===== */

    /* ===== Яркие анимационные пятна ===== */
    .hero-glow { position: absolute; border-radius: 50%; filter: blur(80px); z-index: 0; pointer-events: none; }
    .hero-glow--1 { width: 450px; height: 450px; top: -120px; left: -100px; background: radial-gradient(circle, rgba(253,216,53,0.3) 0%, transparent 70%); animation: glowMove1 12s ease-in-out infinite; }
    .hero-glow--2 { width: 400px; height: 400px; bottom: -120px; right: 10%; background: radial-gradient(circle, rgba(255,148,25,0.2) 0%, transparent 70%); animation: glowMove2 14s ease-in-out infinite; }
    @keyframes glowMove1 { 0%,100% { transform: translate(0,0) scale(1); opacity: 0.3; } 33% { transform: translate(50px,30px) scale(1.15); opacity: 0.45; } 66% { transform: translate(-20px,60px) scale(0.9); opacity: 0.25; } }
    @keyframes glowMove2 { 0%,100% { transform: translate(0,0) scale(1); opacity: 0.2; } 33% { transform: translate(-40px,-20px) scale(1.1); opacity: 0.35; } 66% { transform: translate(30px,-50px) scale(0.85); opacity: 0.15; } }

    .hero-content { position: relative; z-index: 2; padding: var(--sp-10) var(--sp-8); max-width: 640px; }
    .hero-badge { display: inline-flex; align-items: center; gap: 8px; padding: 6px 16px; border-radius: var(--radius-30); background: rgba(253,216,53,0.12); border: 1px solid rgba(253,216,53,0.25); color: #fdd835; font-size: 13px; font-weight: 500; margin-bottom: var(--sp-5); animation: fadeIn 0.8s ease 0.2s both; }
    .hero-badge i { font-size: 14px; }
    .hero h1 { font-size: 104px; font-weight: 800; letter-spacing: -0.04em; color: #fff; margin-bottom: var(--sp-4); line-height: 1.0; animation: fadeUp 0.7s ease 0.1s both; }
    .hero h1 .accent-text { color: #fdd835; text-shadow: 0 0 40px rgba(253,216,53,0.6); }
    .hero-sub { font-size: 18px; color: rgba(255,255,255,0.65); font-weight: 400; max-width: 520px; margin-bottom: var(--sp-7); line-height: 1.55; animation: fadeUp 0.7s ease 0.2s both; }
    .hero-cta { display: flex; gap: var(--sp-3); flex-wrap: wrap; animation: fadeUp 0.7s ease 0.3s both; }

    /* ===== Сеть-созвездие (констелляция) — крупная, заполняет правую часть hero ===== */
    .hero-nav { position: absolute; top: 0; right: 0; width: 55%; height: 100%; z-index: 1; pointer-events: none; animation: fadeIn 1s ease 0.5s both; }
    .hero-nav svg { position: absolute; inset: 0; width: 100%; height: 100%; overflow: visible; }
    /* Пульсация центрального узла */
    .node-main { transform-origin: center; transform-box: fill-box; animation: mainPulse 3s ease-in-out infinite; }
    @keyframes mainPulse { 0%,100% { transform: scale(1); opacity: 0.8; } 50% { transform: scale(1.2); opacity: 1; } }
    /* Мягкое дыхание периферийных узлов */
    .node { transform-origin: center; transform-box: fill-box; animation: nodeBreathe 4s ease-in-out infinite; }
    .node--2 { animation-delay: 0.6s; }
    .node--3 { animation-delay: 1.2s; }
    .node--4 { animation-delay: 1.8s; }
    .node--5 { animation-delay: 2.4s; }
    .node--6 { animation-delay: 3.0s; }
    @keyframes nodeBreathe { 0%,100% { transform: scale(1); opacity: 0.5; } 50% { transform: scale(1.25); opacity: 1; } }
    /* Движение пакетов данных по связям */
    .data-dot { animation: dataTravel 4s linear infinite; }
    .data-dot--2 { animation-delay: 1.33s; }
    .data-dot--3 { animation-delay: 2.67s; }
    @keyframes dataTravel { 0% { offset-distance: 0%; opacity: 0; } 10% { opacity: 1; } 85% { opacity: 1; } 100% { offset-distance: 100%; opacity: 0; } }
    @media (max-width: 768px) { .hero-nav { width: 60%; } .hero { min-height: 360px; } .hero-content { padding: var(--sp-6) var(--sp-4); } .hero h1 { font-size: 48px; } .hero-sub { font-size: 15px; } }
    .hero-btn-primary { display: inline-flex; align-items: center; gap: var(--sp-2); padding: 0 var(--sp-6); height: var(--control-h-md); background: #fdd835; color: #000; border: none; border-radius: var(--radius-30); font-family: var(--font-text); font-size: 15px; font-weight: 600; cursor: pointer; transition: all 0.25s var(--ease); position: relative; overflow: hidden; }
    .hero-btn-primary::before { content: ''; position: absolute; top: 0; left: -100%; width: 100%; height: 100%; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent); transition: left 0.5s var(--ease); }
    .hero-btn-primary:hover::before { left: 100%; }
    .hero-btn-primary:hover { background: #fdc435; transform: translateY(-2px); box-shadow: 0 8px 24px rgba(253,216,53,0.5); }
    .hero-btn-primary:hover i { transform: translateX(4px); }
    .hero-btn-primary i { transition: transform 0.25s var(--ease); }

    /* ===== ROLE CARDS (для кого этот портал) — компактные ===== */
    .roles-section { margin-bottom: var(--sp-8); }
    .roles-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: var(--sp-3); margin-top: var(--sp-5); }
    @media (max-width: 768px) { .roles-grid { grid-template-columns: 1fr; } }
    .role-card { background: var(--bg-neutral); border: none; border-radius: var(--radius-12); padding: var(--sp-4); display: flex; flex-direction: column; gap: var(--sp-2); transition: all 0.25s var(--ease); box-shadow: none; position: relative; overflow: hidden; }
    .role-card::after { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, transparent, var(--accent), transparent); opacity: 0; transition: opacity 0.25s var(--ease); }
    .role-card:hover { background: var(--bg-card); transform: translateY(-3px); box-shadow: var(--shadow-md); }
    .role-card:hover::after { opacity: 1; }
    .role-icon { width: 38px; height: 38px; border-radius: var(--radius-10); display: flex; align-items: center; justify-content: center; font-size: 18px; margin-bottom: var(--sp-1); transition: transform 0.3s var(--ease); }
    .role-icon i { font-size: 18px; transition: transform 0.3s var(--ease); }
    .role-card:hover .role-icon { transform: scale(1.1) rotate(-5deg); }
    .role-card:hover .role-icon i { transform: scale(1.05); }
    .ri-amber { background: var(--accent-soft); color: var(--color-orange); }
    .ri-purple { background: var(--accent-purple-soft); color: var(--accent-purple); }
    .ri-green { background: #e7f6eb; color: var(--color-green); }
    .ri-blue { background: #e7f0ff; color: #2563eb; }
    .role-name { font-size: 15px; font-weight: 600; color: var(--text-primary); }
    .role-desc { font-size: 12px; color: var(--text-secondary); line-height: 1.4; }

    /* ===== INFO BANNER ===== */
    .info-banner { display: flex; align-items: center; gap: var(--sp-4); padding: var(--sp-6) var(--sp-8); background: var(--bg-neutral); border: none; border-radius: var(--radius-16); margin-bottom: var(--sp-8); }
    .info-banner i { font-size: 28px; color: var(--accent); flex-shrink: 0; }
    .info-banner-text { font-size: 20px; color: var(--text-secondary); line-height: 1.55; }
    .info-banner-text strong { color: var(--text-primary); font-weight: 600; font-size: 21px; }

    @keyframes fadeUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
    .directions-grid {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: var(--sp-4); margin-top: var(--sp-6);
    }
    .dir-card {
      background: var(--bg-neutral); border: none; border-radius: var(--radius-20); padding: 24px;
      text-align: left; cursor: pointer;
      transition: all 0.25s var(--ease);
      position: relative; overflow: hidden;
      display: flex; flex-direction: column; align-items: flex-start; gap: var(--sp-3);
      min-height: 140px;
      box-shadow: none;
    }
    .dir-card::after { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, transparent, var(--accent), transparent); opacity: 0; transition: opacity 0.25s var(--ease); }
    body.dark .dir-card {
      background: var(--bg-neutral);
    }
    .dir-card::before {
      content: ''; position: absolute; top: 0; left: 0; right: 0; height: 4px;
      background: linear-gradient(90deg, transparent, var(--accent), var(--accent), transparent);
      background-size: 200% 100%;
      opacity: 0; transition: opacity 0.3s var(--ease);
    }
    .dir-card:hover {
      background: var(--bg-card);
      transform: translateY(-3px);
      box-shadow: var(--shadow-md);
    }
    body.dark .dir-card:hover {
      box-shadow: var(--shadow-md);
    }
    .dir-card:hover::after {
      opacity: 1;
    }
    .dir-card:hover::before {
      opacity: 1;
      animation: shimmer 1.5s linear infinite;
    }
    @keyframes shimmer {
      0% { background-position: -200% 0; }
      100% { background-position: 200% 0; }
    }
    .card-emoji {
      font-size: 30px; margin-bottom: var(--sp-4); display: inline-flex;
      align-items: center; justify-content: center;
      width: 56px; height: 56px; border-radius: var(--radius-16);
      transition: transform 0.3s var(--ease);
    }
    .ce-teal { background: #e7f6eb; color: #2ca853; }
    .ce-purple { background: var(--accent-purple-soft); color: var(--accent-purple); }
    .ce-amber { background: #fae4f7; color: #bc00b8; }
    .dir-card:hover .card-emoji { transform: scale(1.1) rotate(-5deg); }
    .dir-name { font-size: 20px; font-weight: 500; position: relative; z-index: 1; color: var(--text-dark-active); letter-spacing: -0.01em; }
    .dir-sub { font-size: 14px; color: var(--text-dark-inactive); margin-top: var(--sp-2); position: relative; z-index: 1; }
    .drag-handle, .link-drag-handle {
      position: absolute; bottom: 8px; right: 8px;
      background: none; border: none;
      cursor: grab; z-index: 10; font-size: 0.6rem;
      color: var(--text-muted); opacity: 0.15;
      border-radius: var(--radius-2); padding: 3px;
      transition: all 0.2s ease;
    }
    .drag-handle:hover, .link-drag-handle:hover { opacity: 0.7; color: var(--accent); transform: scale(1.15); }
    .adm-icon {
      position: absolute; top: 10px; background: transparent !important; border: none !important;
      cursor: pointer; z-index: 10; font-size: 0.8rem; padding: var(--sp-1);
      transition: all 0.2s ease; opacity: 0.25;
    }
    .adm-icon:hover { opacity: 1; transform: scale(1.1); }
    .adm-icon.edit { right: 38px; color: var(--accent); }
    .adm-icon.delete { right: 10px; color: var(--status-error); }
    .links-grid {
      display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
      gap: var(--sp-3); margin-top: var(--sp-6);
    }
    .link-item {
      background: var(--bg-neutral); border: none; border-radius: var(--radius-12);
      padding: var(--sp-4) var(--sp-5);
      font-size: 15px; font-weight: 500; line-height: 1.4;
      text-align: left; cursor: pointer;
      transition: all 0.25s var(--ease);
      display: flex; flex-direction: column; align-items: flex-start; gap: var(--sp-2);
      position: relative; font-family: var(--font-text);
      color: var(--text-primary); box-shadow: none;
      min-height: auto; overflow: hidden;
    }
    .link-item::after { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, transparent, var(--accent), transparent); opacity: 0; transition: opacity 0.25s var(--ease); }
    .link-item:hover {
      background: var(--bg-card); color: var(--text-primary);
      box-shadow: var(--shadow-md);
      transform: translateY(-3px);
    }
    .link-item:hover::after { opacity: 1; }
    .link-item i.link-icon {
      width: auto; height: auto; border-radius: 0;
      display: flex; align-items: center; justify-content: flex-start;
      font-size: 24px; font-weight: 700; transition: transform 0.3s var(--ease);
      flex-shrink: 0; opacity: 1; margin-top: 0;
      background: transparent !important;
    }
    .link-item:hover i.link-icon { transform: scale(1.1) rotate(-5deg); }
    /* Цвета иконок по категориям — без фона, работают в тёмной/светлой теме */
    .link-icon--report { color: #4d8eff !important; }
    .link-icon--regulation { color: #a054ff !important; }
    .link-icon--news { color: #ffac3d !important; }
    .link-icon--qms { color: #3dcc66 !important; }
    .link-icon--management { color: #ff4d6d !important; }
    .link-icon--default { color: var(--accent) !important; }
    .link-item:hover i.link-icon { opacity: 1; }
    .link-item .link-text { flex: 1; font-size: 15px; font-weight: 500; }
    .link-item .link-desc { font-size: 12px; color: var(--text-secondary); line-height: 1.4; }
    /* Плашка для кнопок «Для руководителей» */
    .link-item--manager {
      background: var(--accent-soft) !important;
    }
    .link-item--manager:hover {
      background: var(--bg-card) !important;
      box-shadow: 0 8px 24px rgba(253,196,53,0.20);
    }
    .link-item--manager::after {
      background: linear-gradient(90deg, transparent, var(--accent-purple), transparent) !important;
    }
    .manager-badge {
      position: absolute; bottom: 8px; right: 10px;
      display: inline-flex; align-items: center; gap: 4px;
      padding: 3px 10px; border-radius: var(--radius-30);
      background: var(--accent-purple); color: #fff;
      font-size: 10px; font-weight: 600; letter-spacing: 0.03em;
      text-transform: uppercase; line-height: 1.2;
      z-index: 5;
    }
    .manager-badge i { font-size: 11px; }
    .dragging { opacity: 0.35; transform: scale(0.97); }
    .lock-badge {
      display: inline-flex; align-items: center; justify-content: center;
      flex-shrink: 0; margin-left: auto; padding-left: var(--sp-2);
      font-size: 16px; color: var(--text-muted); opacity: 0.7;
    }
    .link-item:hover .lock-badge { color: var(--accent); opacity: 1; }
    .tooltip-card {
      position: absolute;
      top: calc(100% + 10px);
      bottom: auto;
      left: 50%;
      transform: translateX(-50%) translateY(-6px);
      min-width: 300px; max-width: 400px;
      padding: var(--sp-3) var(--sp-4);
      background: var(--bg-card); border: var(--border-width) solid var(--accent);
      border-radius: var(--radius-6); font-size: 13px; font-weight: 400;
      line-height: 1.5; color: var(--text-secondary);
      pointer-events: none; opacity: 0; visibility: hidden;
      transition: opacity 0.2s ease, transform 0.2s ease, visibility 0.2s ease;
      z-index: 999; box-shadow: var(--shadow-lg);
    }
    .link-item.tt-open { z-index: 1000; }
    .link-item.tt-open .tooltip-card {
      opacity: 1; visibility: visible; pointer-events: auto;
      transform: translateX(-50%) translateY(0);
    }
    .tooltip-card .tt-copy-btn {
      pointer-events: auto;
      display: inline-flex; align-items: center; gap: 4px;
      background: var(--accent-soft); color: var(--accent);
      border: 1px solid var(--accent); border-radius: var(--radius-4);
      padding: 3px 8px; font-size: 11px; font-weight: 500;
      cursor: pointer; margin-top: 8px;
      transition: all 0.15s var(--ease);
    }
    .tooltip-card .tt-copy-btn:hover {
      background: var(--accent); color: var(--bg-card);
    }
    .tooltip-card .tt-copy-btn i { font-size: 12px; }
    .tooltip-card .tt-copy-done {
      background: var(--color-green); color: #fff; border-color: var(--color-green);
    }
    .tooltip-card::after {
      content: ''; position: absolute; bottom: 100%; left: 50%;
      transform: translateX(-50%); border: 6px solid transparent;
      border-bottom-color: var(--accent);
    }
    .tooltip-card .tt-badge {
      display: inline-flex; align-items: center; justify-content: center;
      width: 20px; height: 20px; border-radius: var(--radius-4);
      background: var(--accent-soft); color: var(--accent);
      font-size: 0.6rem; margin-right: 5px; flex-shrink: 0; vertical-align: middle;
    }
    .tooltip-card .tt-title {
      display: flex; align-items: center; margin-bottom: 6px;
      font-weight: 500; font-size: 12px; color: var(--accent);
      text-transform: uppercase; letter-spacing: 0.04em;
    }
    .tooltip-card strong { color: var(--accent); font-weight: 500; }
    textarea {
      width: 100%; max-width: 420px; padding: var(--sp-3) var(--sp-4);
      border-radius: var(--radius-10); border: var(--border-width) solid var(--border);
      background: var(--bg); font-family: var(--font-text);
      resize: vertical; margin: var(--sp-3) 0; outline: none;
      transition: all 0.2s var(--ease); font-size: 14px; color: inherit;
      min-height: 80px;
    }
    textarea:focus { border-color: var(--border-focused); box-shadow: 0 0 0 3px rgba(33,33,33,0.12); }
    textarea::placeholder { color: var(--text-muted); }
    .btn-send {
      background: var(--grad-btn); border: none; padding: 0 var(--sp-6);
      height: var(--control-h-md); border-radius: var(--radius-30); color: var(--text-light-active); font-weight: 500;
      cursor: pointer; transition: all 0.2s var(--ease);
      font-size: 15px; margin-top: var(--sp-2); font-family: var(--font-text);
      display: inline-flex; align-items: center; gap: var(--sp-2);
    }
    .btn-send:hover { box-shadow: var(--shadow-md); transform: translateY(-1px); }
    .tab-pane { display: none; }
    .tab-pane.active { display: block; animation: fadeIn 0.2s ease; }
    .analytics-wrap { max-width: 1280px; margin: 0 auto; padding: var(--sp-4); }
    .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: var(--sp-3); margin-bottom: var(--sp-7); }
    .metric-box {
      background: var(--bg-card); border-radius: var(--radius-16); padding: var(--sp-6);
      text-align: center; border: var(--border-width) solid var(--border);
      transition: all 0.2s var(--ease); box-shadow: var(--shadow-sm);
    }
    .metric-box:hover { border-color: var(--accent); transform: translateY(-2px); box-shadow: var(--shadow-md); }
    .metric-val { font-size: 36px; font-weight: 500; color: var(--accent); letter-spacing: -0.02em; }
    .metric-lbl { font-size: 13px; color: var(--text-muted); font-weight: 500; margin-top: var(--sp-2); }
    .table-wrap {
      overflow-x: auto; margin-top: var(--sp-4); border-radius: var(--radius-12);
      border: var(--border-width) solid var(--border); width: 100%;
    }
    table { width: 100%; border-collapse: collapse; background: var(--bg-card); overflow: hidden; table-layout: auto; }
    th, td { padding: var(--sp-3) var(--sp-4); text-align: left; border-bottom: var(--border-width) solid var(--border); white-space: normal; word-wrap: break-word; font-size: 14px; }
    th { background: var(--accent-soft); color: var(--accent); font-weight: 500; text-transform: uppercase; font-size: 11px; letter-spacing: 0.06em; position: sticky; top: 0; z-index: 10; }
    tbody tr:hover { background: var(--accent-soft); }
    .dl-btn {
      background: var(--grad-btn); border: none; height: var(--control-h-md);
      padding: 0 var(--sp-6); border-radius: var(--radius-30); cursor: pointer; font-weight: 500;
      color: var(--text-light-active); font-family: var(--font-text); font-size: 14px;
      transition: all 0.2s var(--ease);
      display: inline-flex; align-items: center; gap: var(--sp-2);
    }
    .dl-btn:hover { transform: translateY(-1px); box-shadow: var(--shadow-md); }
    .filter-row { display: flex; gap: var(--sp-2); flex-wrap: wrap; margin-bottom: var(--sp-4); }
    .filter-row select, .filter-row input {
      height: var(--control-h-md); padding: 0 var(--sp-4); border-radius: var(--radius-10);
      border: var(--border-width) solid var(--border); background: var(--bg-card);
      font-family: var(--font-text); color: inherit; cursor: pointer;
      font-weight: 500; outline: none; transition: all 0.2s var(--ease); font-size: 14px;
    }
    .filter-row select:focus, .filter-row input:focus { border-color: var(--border-focused); box-shadow: 0 0 0 3px rgba(33,33,33,0.12); }
    .sync-badge {
      display: inline-flex; align-items: center; gap: 5px;
      font-size: 11px; color: var(--text-muted); margin-left: var(--sp-2);
      font-weight: 500;
    }
    .sync-badge .dot {
      width: 6px; height: 6px; border-radius: 50%;
      background: var(--status-success); animation: pulse-dot 2s infinite;
    }
    @keyframes pulse-dot { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }
    @media (max-width: 900px) {
      .sidebar { transform: translateX(-100%); box-shadow: var(--shadow-lg); }
      .sidebar.open { transform: translateX(0); }
      .main { margin-left: 0; }
      .topbar { padding: var(--sp-3) var(--sp-4); min-height: 56px; }
      .page-title { font-size: 20px; }
      .menu-toggle { display: inline-flex; }
      .content { padding: var(--sp-4); }
      .directions-grid { grid-template-columns: 1fr; gap: var(--sp-3); }
      .links-grid { grid-template-columns: 1fr; }
      .hero h1 { font-size: 28px; }
      .hero-sub { font-size: 14px; }
      .info-banner { flex-direction: column; text-align: center; }
      .metric-grid { grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); }
      .metric-box { padding: var(--sp-4); }
      .metric-val { font-size: 28px; }
      .tooltip-card { min-width: 240px; max-width: 300px; left: 0; transform: translateX(0) translateY(-6px); }
      .link-item.tt-open .tooltip-card { transform: translateX(0) translateY(0); }
      .tooltip-card::after { left: 20px; transform: none; }
    }
  </style>
</head>
<body>
  <!-- Aurora-фон на весь экран -->
  <div class="bg-aurora" aria-hidden="true">
    <div class="bg-blob bg-blob--1"></div>
    <div class="bg-blob bg-blob--2"></div>
    <div class="bg-blob bg-blob--3"></div>
    <div class="bg-blob bg-blob--4"></div>
  </div>
  <div class="sidebar-overlay" id="sidebar-overlay"></div>
  <div class="app">
    <aside class="sidebar" id="sidebar">
      <nav class="sidebar-nav">
        <button class="nav-item" id="nav-home" data-label="Главная"><i class="beeline-icons">home</i><span>Главная</span></button>
        <button class="nav-item" id="fb-open-btn" data-label="Обратная связь"><i class="beeline-icons">chat</i><span>Обратная связь</span></button>
        <button class="nav-item" id="feedback-stats-btn" style="display:none;" data-label="Статистика"><i class="beeline-icons">reports</i><span>Статистика</span></button>
        <button class="nav-item" id="theme-toggle" data-label="Тема"><i class="beeline-icons">sun</i><span>Тема</span></button>
        <button class="nav-item" id="admin-toggle" data-label="Войти (админ)"><i class="beeline-icons">lock</i><span>Войти (админ)</span></button>
      </nav>
      <div class="sidebar-divider"></div>
      <div class="sidebar-section-label">Направления</div>
      <nav class="sidebar-nav" id="dir-nav"></nav>
      <button class="sidebar-collapse-btn" id="sidebar-collapse-btn" aria-label="Свернуть меню"><i class="beeline-icons">menu_open</i></button>
    </aside>
    <div class="sidebar-tt" id="sidebar-tt" role="tooltip"></div>
    <div class="main">
      <header class="topbar">
        <button class="menu-toggle" id="menu-toggle"><i class="beeline-icons">menu</i></button>
        <div class="topbar-logo">
          <div class="logo-sphere">
            <img src="Логотип для светлой темы.PNG" alt="" class="sidebar-logo-img sidebar-logo-img--light" />
            <img src="Логотип для темной темы.PNG" alt="" class="sidebar-logo-img sidebar-logo-img--dark" />
          </div>
        </div>
        <div class="topbar-spacer"></div>
        <div class="topbar-search">
          <i class="beeline-icons">search</i>
          <input type="search" placeholder="Поиск по разделам…" aria-label="Поиск по разделам" />
        </div>
        <div class="topbar-actions">
          <span class="sync-badge" id="sync-badge" style="display:none;"><span class="dot"></span></span>
        </div>
      </header>
      <div class="content">
        <div class="content-inner" id="main-content"></div>
      </div>
    </div>
  </div>
  <div class="fb-overlay" id="fb-overlay">
    <div class="fb-modal">
      <button class="fb-close" id="fb-close-btn"><i class="beeline-icons">close</i></button>
      <h3><i class="beeline-icons" style="color:var(--bg-brand)">magic</i> Обратная связь</h3>
      <p style="margin-bottom:4px;">Твоё мнение помогает становиться лучше</p>
      <textarea id="fb-text" rows="3" placeholder="Комментарий или предложение..."></textarea>
      <button class="btn-send" id="send-fb"><i class="beeline-icons">send</i> Отправить</button>
    </div>
  </div>
  <script>
    // Firebase удалён — все данные хранятся в localStorage.
    // Для multi-user sync заменить функции cloud* на fetch() к внутреннему backend API.
    const ORDER_KEYS = { mainButtons: 'main_buttons_order', directionLinks: (dir) => `direction_${dir}_links_order` };
    const CE = ['ce-teal', 'ce-purple', 'ce-amber'];
    const ADMIN_PASSWORD = 'beeline2025';

    const defaultMainButtons = [
      { id: 'service', text: 'Сервис', icon: 'headset_help', href: '?dir=service', sub: '' },
      { id: 'sales', text: 'Продажи', icon: 'graph_up', href: '?dir=sales', sub: '' },
      { id: 'retention', text: 'Сохранение', icon: 'star', href: '?dir=retention', sub: '' }
    ];

    const directionsDefaults = {
      service: { name: "Сервис", icon: "headset_help", links: [
        { text: "Регламенты оценки качества / Матрицы оценки / Карта фрода", href: "#", type: "link" },
        { text: "Описание скоринга / Новости / Процесс взаимодействия с учениками", href: "#", type: "link" },
        { text: "Стандарты обслуживания", href: "#", type: "link" },
        { text: "Чат бот", href: "#", type: "link" },
        { text: "Отчет по качеству", href: "#", type: "link" },
        { text: "Нормативы / Цели", href: "#", type: "link" },
        { text: "Контроль качества. Жалобы / Благодарности / Ошибки сотрудников", href: "#", type: "link" }
      ]},
      sales: { name: "Продажи", icon: "graph_up", links: [
        { text: "Регламенты оценки качества / Матрицы оценки / Карта фрода", href: "#", type: "link" },
        { text: "Описание скоринга / Новости / Процесс взаимодействия с учениками", href: "#", type: "link" },
        { text: "Стандарты обслуживания", href: "#", type: "link" },
        { text: "Отчет по качеству", href: "#", type: "link" },
        { text: "Нормативы / Цели", href: "#", type: "link" },
        { text: "Чат бот", href: "#", type: "link" },
        { text: "Контроль качества. Жалобы / Благодарности / Ошибки сотрудников", href: "#", type: "link" }
      ]},
      retention: { name: "Сохранение", icon: "star", links: [
        { text: "Регламенты оценки качества / Матрицы оценки / Карта фрода", href: "#", type: "link" },
        { text: "Описание скоринга / Новости / Процесс взаимодействия с учениками", href: "#", type: "link" },
        { text: "Стандарты обслуживания", href: "#", type: "link" },
        { text: "Чат бот", href: "#", type: "link" },
        { text: "Отчет по качеству", href: "#", type: "link" },
        { text: "Нормативы / Цели", href: "#", type: "link" },
        { text: "Контроль качества. Жалобы / Благодарности / Ошибки сотрудников", href: "#", type: "link" }
      ]}
    };

    let mainButtons = JSON.parse(JSON.stringify(defaultMainButtons));
    let directions = JSON.parse(JSON.stringify(directionsDefaults));

    // ===== ICON NAME MIGRATION (legacy FA → DS yellowbe ligature) =====
    // localStorage могут содержать старые FA иконки (fa-headset, fa-arrow-trend-up, ...).
    // Эта функция конвертирует их в DS ligature имена при загрузке данных.
    function migrateIconName(ic) {
      if (!ic || typeof ic !== 'string') return 'star';
      const map = {
        'fa-headset': 'headset_help',
        'fa-arrow-trend-up': 'graph_up',
        'fa-shield-heart': 'flash',
        'fa-bolt': 'flash',
        'fa-house': 'home',
        'fa-comment-dots': 'chat',
        'fa-comment': 'chat',
        'fa-magnifying-glass': 'search',
        'fa-sun': 'sun',
        'fa-moon': 'half_moon',
        'fa-lock': 'lock'
      };
      return map[ic] || ic;
    }
    function migrateMainButtons(arr) {
      if (!Array.isArray(arr)) return arr;
      arr.forEach(b => { if (b && b.icon) b.icon = migrateIconName(b.icon); });
      return arr;
    }
    function migrateDirections(obj) {
      if (!obj || typeof obj !== 'object') return obj;
      Object.keys(obj).forEach(k => {
        const d = obj[k];
        if (d && Array.isArray(d.links)) {
          if (d.icon) d.icon = migrateIconName(d.icon);
          d.links.forEach(ln => { if (ln && ln.icon) ln.icon = migrateIconName(ln.icon); });
        }
      });
      return obj;
    }
    let feedbackList = [];
    let isAdmin = localStorage.getItem('adminActive') === 'true';
    const urlParams = new URLSearchParams(window.location.search);
    // Если в URL нет ?dir=, восстанавливаем последнее выбранное направление из localStorage
    let dirKey = urlParams.get('dir');
    if (!dirKey) {
      const savedDir = localStorage.getItem('currentDirection');
      if (savedDir) dirKey = savedDir;
    }
    const feedbackStatsMode = urlParams.get('feedback');

    // ===== LOCAL STORAGE =====
    function saveOrder(k, a) { localStorage.setItem(k, JSON.stringify(a)); }
    function loadOrder(k, da) {
      const s = localStorage.getItem(k);
      if (s) { try { const o = JSON.parse(s); if (Array.isArray(o) && o.length === da.length) return o; } catch (e) {} }
      return da.map((_, i) => i);
    }
    function reorderArrayByIndices(a, oi) { const r = []; oi.forEach(i => { if (i >= 0 && i < a.length) r.push(a[i]); }); return r; }
    function saveMainButtons() { localStorage.setItem('mainButtons', JSON.stringify(mainButtons)); }
    function saveDirections() { localStorage.setItem('directionsData', JSON.stringify(directions)); }

    // ===== STORAGE (localStorage + optional backend API) =====
    // Автоопределение: если файл открыт через сервер (http/https) — используем API.
    // Если открыт локально (file://) или через Vite dev (5173) — только localStorage.
    const _port = String(window.location.port || '');
    const API_BASE = (window.location.protocol.startsWith('http') && _port !== '5173' && _port !== '4173')
      ? window.location.origin + '/api'
      : null;
    let adminToken = localStorage.getItem('adminToken') || null;

    async function apiGet(endpoint) {
      if (!API_BASE) return null;
      try {
        const headers = {};
        if (adminToken) headers['Authorization'] = `Bearer ${adminToken}`;
        const res = await fetch(`${API_BASE}/${endpoint}`, { headers, signal: AbortSignal.timeout ? AbortSignal.timeout(4000) : undefined });
        if (!res.ok) return null;
        return await res.json();
      } catch (e) { return null; }
    }
    async function apiPut(endpoint, data) {
      if (!API_BASE) return;
      try {
        await fetch(`${API_BASE}/${endpoint}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${adminToken}` },
          body: JSON.stringify(data)
        });
      } catch (e) { console.warn(`API PUT ${endpoint}:`, e); }
    }
    async function apiPost(endpoint, data) {
      if (!API_BASE) return;
      try {
        await fetch(`${API_BASE}/${endpoint}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(data)
        });
      } catch (e) { console.warn(`API POST ${endpoint}:`, e); }
    }

    // cloud-функции — мост между UI и хранилищем (API или localStorage)
    async function cloudSave(path, data) {
      if (path === 'app/directions') return apiPut('directions', data);
      if (path === 'app/mainButtons') return apiPut('mainButtons', data);
    }
    async function cloudLoad(path) {
      if (path === 'app/directions') return apiGet('directions');
      if (path === 'app/mainButtons') return apiGet('mainButtons');
      if (path === 'feedback') return apiGet('feedback');
      return null;
    }
    async function cloudPush(path, data) {
      if (path === 'feedback') return apiPost('feedback', data);
      if (path === 'clicks') return apiPost('clicks', data);
    }

    async function syncToCloud() {
      // Данные уже сохранены в localStorage (saveDirections/saveMainButtons).
      // Если backend доступен — отправляем на сервер для других пользователей.
      await apiPut('directions', JSON.parse(JSON.stringify(directions)));
      await apiPut('mainButtons', JSON.parse(JSON.stringify(mainButtons)));
    }

    async function syncFeedbackFromCloud() {
      if (!API_BASE || !isAdmin) return; // только админ может читать отзывы
      const cloudFB = await cloudLoad('feedback');
      if (!cloudFB || !Array.isArray(cloudFB)) return;
      const localRaw = localStorage.getItem('quality_feedback');
      let localItems = [];
      try { localItems = localRaw ? JSON.parse(localRaw) : []; } catch (e) {}
      const existingKeys = new Set(localItems.map(f => f.timestamp + '|' + f.rating + '|' + f.comment));
      let added = 0;
      cloudFB.forEach(item => {
        if (item && item.timestamp) {
          const key = item.timestamp + '|' + (item.rating || 0) + '|' + (item.comment || '');
          if (!existingKeys.has(key)) { localItems.push(item); existingKeys.add(key); added++; }
        }
      });
      if (added > 0) {
        localStorage.setItem('quality_feedback', JSON.stringify(localItems));
        if (feedbackStatsMode === 'true') { feedbackList = localItems; renderMainPage(); }
      }
    }

    async function syncFromCloud() {
      if (!API_BASE) return;
      const dirsData = await cloudLoad('app/directions');
      if (dirsData) {
        migrateDirections(dirsData);
        Object.keys(dirsData).forEach(k => {
          if (dirsData[k] && Array.isArray(dirsData[k].links)) {
            directions[k] = dirsData[k];
          }
        });
        saveDirections();
      }
      const btnsData = await cloudLoad('app/mainButtons');
      if (btnsData && Array.isArray(btnsData) && btnsData.length > 0) {
        migrateMainButtons(btnsData);
        mainButtons = btnsData;
        saveMainButtons();
      }
      await syncFeedbackFromCloud();
      renderSidebarNav();
      renderMainPage();
    }

    // ===== UI HELPERS =====
    function createAdminIcons(p, ia, od, oe) {
      if (!ia) return;
      const ei = document.createElement('i'); ei.className = 'beeline-icons adm-icon edit'; ei.textContent = 'edit'; ei.title = 'Редактировать'; ei.onclick = oe;
      const di = document.createElement('i'); di.className = 'beeline-icons adm-icon delete'; di.textContent = 'delete'; di.title = 'Удалить'; di.onclick = od;
      p.appendChild(ei); p.appendChild(di);
    }
    function createDragIcon(el, isL = false) {
      const d = document.createElement('i'); d.className = 'beeline-icons ' + (isL ? 'link-drag-handle' : 'drag-handle'); d.textContent = 'drag_indicator'; d.title = 'Перетащить'; el.appendChild(d); return d;
    }
    function addQualityTooltip(btn) {
      // Добавляем значок замка рядом с текстом кнопки
      const lockBadge = document.createElement('span');
      lockBadge.className = 'lock-badge';
      lockBadge.innerHTML = '<i class="beeline-icons">lock</i>';
      lockBadge.title = 'Требуется доступ — нажмите для информации';
      btn.appendChild(lockBadge);

      const tt = document.createElement('div'); tt.className = 'tooltip-card';
      const textContent = 'Доступ предоставляется по заявке QLIK Stream 02.027_A. Customer Care Qlik Sense. Доступ к Стримам роль Пользователь';
      tt.innerHTML = '<div class="tt-title"><span class="tt-badge"><i class="beeline-icons">lock</i></span> Требуется доступ</div>Доступ предоставляется по заявке <strong>QLIK Stream 02.027_A. Customer Care Qlik Sense. Доступ к Стримам</strong> роль <strong>Пользователь</strong><button class="tt-copy-btn" type="button"><i class="beeline-icons">copy</i> Копировать</button>';
      const copyBtn = tt.querySelector('.tt-copy-btn');
      copyBtn.addEventListener('click', function(e) {
        e.preventDefault(); e.stopPropagation();
        navigator.clipboard.writeText(textContent).then(function() {
          copyBtn.classList.add('tt-copy-done');
copyBtn.innerHTML = '<i class="beeline-icons">check</i> Скопировано';
          setTimeout(function() {
            copyBtn.classList.remove('tt-copy-done');
copyBtn.innerHTML = '<i class="beeline-icons">copy</i> Копировать';
          }, 2000);
        }).catch(function() {
          var ta = document.createElement('textarea');
          ta.value = textContent; document.body.appendChild(ta); ta.select();
          try { document.execCommand('copy'); copyBtn.innerHTML = '<i class="beeline-icons">check</i> Скопировано'; } catch(err) {}
          document.body.removeChild(ta);
          setTimeout(function() { copyBtn.innerHTML = '<i class="beeline-icons">copy</i> Копировать'; }, 2000);
        });
      });
      btn.appendChild(tt);
    }
    function showSyncBadge() {
      const b = document.getElementById('sync-badge'); if (!b) return;
      b.style.display = 'inline-flex';
      // Индикатор режима: online (API) или offline (localStorage)
      if (API_BASE) {
        b.title = 'Онлайн: правки админа синхронизируются между пользователями';
        b.querySelector('.dot').style.background = 'var(--status-success)';
      } else {
        b.title = 'Офлайн (localStorage): правки видны только в этом браузере';
        b.querySelector('.dot').style.background = 'var(--color-orange)';
      }
    }
    function renderSidebarNav() {
      const nav = document.getElementById('dir-nav'); if (!nav) return;
      const cb = [...mainButtons];
      nav.innerHTML = '';
      cb.forEach((btn) => {
        const a = document.createElement('button');
        a.className = 'nav-item' + (dirKey === btn.id ? ' active' : '');
        // DS yellowbe icon font (ligature); fallback на FA, если icon начинается с 'fa-' (напр. пользовательские 'fa-folder')
        const ic = btn.icon || 'star';
        a.setAttribute('data-label', btn.text);
        a.innerHTML = ic.startsWith('fa-') ? `<i class="fas ${ic}"></i><span>${btn.text}</span>` : `<i class="beeline-icons">${ic}</i><span>${btn.text}</span>`;
        a.onclick = () => { localStorage.setItem('currentDirection', btn.id); window.location.href = `?dir=${btn.id}`; };
        nav.appendChild(a);
      });
      const home = document.getElementById('nav-home');
      if (home) home.classList.toggle('active', !dirKey && !feedbackStatsMode);
      const stats = document.getElementById('feedback-stats-btn');
      if (stats) { stats.classList.toggle('active', feedbackStatsMode); stats.style.display = isAdmin ? 'flex' : 'none'; }
      const adminBtn = document.getElementById('admin-toggle');
      if (adminBtn) {
        adminBtn.querySelector('span').textContent = isAdmin ? 'Выйти (админ)' : 'Войти (админ)';
        // lock = DS yellowbe ligature; sign-out-alt = FA (нет подтверждённого DS equivalent)
        const ai = adminBtn.querySelector('i');
        if (isAdmin) { ai.className = 'beeline-icons'; ai.textContent = 'log_out'; }
        else { ai.className = 'beeline-icons'; ai.textContent = 'lock'; }
      }
    }
    function setHeader(title) {
      // page-title убран из topbar — заголовок виден в контенте (welcome h2 / section-title)
      const el = document.getElementById('page-title');
      if (el) el.textContent = title;
    }

    // ===== RENDER =====
    function renderMainPage() {
      if (feedbackStatsMode) { renderFeedbackStatsPage(); return; }
      if (dirKey && directions[dirKey]) { renderDirectionPage(dirKey); return; }
      renderHomePage();
    }

    function renderHomePage() {
      setHeader('Quality');
      // Скрываем поиск на главной странице
      const ts = document.querySelector('.topbar-search'); if (ts) ts.style.display = 'none';
      // Сброс строки поиска при переходе на главный экран
      const si = document.querySelector('.topbar-search input'); if (si) si.value = '';
      const m = document.getElementById('main-content');
      let cb = [...mainButtons];
      m.innerHTML = `
        <div class="hero-band">
        <div class="hero">
          <div class="hero-bg"></div>
          <div class="hero-glow hero-glow--1"></div>
          <div class="hero-glow hero-glow--2"></div>
          <div class="hero-nav">
            <svg viewBox="0 0 220 220" fill="none" xmlns="http://www.w3.org/2000/svg">
              <defs>
                <!-- Свечение для узлов -->
                <filter id="nodeGlow" x="-100%" y="-100%" width="300%" height="300%">
                  <feGaussianBlur stdDeviation="3" result="blur" />
                  <feMerge><feMergeNode in="blur" /><feMergeNode in="SourceGraphic" /></feMerge>
                </filter>
                <!-- Сильное свечение для центрального узла -->
                <filter id="centerGlow" x="-100%" y="-100%" width="300%" height="300%">
                  <feGaussianBlur stdDeviation="5" result="blur" />
                  <feMerge><feMergeNode in="blur" /><feMergeNode in="SourceGraphic" /></feMerge>
                </filter>
                <!-- Фоновая звёздная пыль -->
                <filter id="starDust" x="-50%" y="-50%" width="200%" height="200%">
                  <feGaussianBlur stdDeviation="1" result="blur" />
                  <feMerge><feMergeNode in="blur" /><feMergeNode in="SourceGraphic" /></feMerge>
                </filter>
              </defs>

              <!-- Тёмный космический фон -->
              <circle cx="110" cy="110" r="92" fill="rgba(0,0,0,0.45)" stroke="rgba(253,216,53,0.06)" stroke-width="1" />

              <!-- === СВЯЗИ (линии констелляции) === -->
              <!-- Радиальные от центра ко всем узлам -->
              <line x1="110" y1="110" x2="40" y2="45" stroke="rgba(253,216,53,0.08)" stroke-width="1" />
              <line x1="110" y1="110" x2="175" y2="35" stroke="rgba(253,216,53,0.08)" stroke-width="1" />
              <line x1="110" y1="110" x2="55" y2="165" stroke="rgba(253,216,53,0.08)" stroke-width="1" />
              <line x1="110" y1="110" x2="170" y2="150" stroke="rgba(253,216,53,0.08)" stroke-width="1" />
              <line x1="110" y1="110" x2="145" y2="80" stroke="rgba(253,216,53,0.08)" stroke-width="1" />
              <line x1="110" y1="110" x2="75" y2="105" stroke="rgba(253,216,53,0.08)" stroke-width="1" />

              <!-- Перемычки между периферийными узлами (образуют созвездие) -->
              <line x1="40" y1="45" x2="175" y2="35" stroke="rgba(253,216,53,0.06)" stroke-width="0.8" />
              <line x1="55" y1="165" x2="170" y2="150" stroke="rgba(253,216,53,0.06)" stroke-width="0.8" />
              <line x1="145" y1="80" x2="75" y2="105" stroke="rgba(253,216,53,0.06)" stroke-width="0.8" />
              <line x1="40" y1="45" x2="75" y2="105" stroke="rgba(253,216,53,0.05)" stroke-width="0.6" />
              <line x1="175" y1="35" x2="145" y2="80" stroke="rgba(253,216,53,0.05)" stroke-width="0.6" />
              <line x1="55" y1="165" x2="75" y2="105" stroke="rgba(253,216,53,0.05)" stroke-width="0.6" />
              <line x1="170" y1="150" x2="145" y2="80" stroke="rgba(253,216,53,0.05)" stroke-width="0.6" />

              <!-- === АНИМИРОВАННЫЕ ПАКЕТЫ ДАННЫХ ПО МАРШРУТАМ === -->
              <!-- Пакет 1: N1 → центр → N4 -->
              <circle r="3" fill="#fdd835" class="data-dot"
                style="offset-path:path('M40,45 L110,110 L170,150'); filter:drop-shadow(0 0 6px rgba(253,216,53,0.9));" />
              <!-- Пакет 2: N6 → центр → N2 -->
              <circle r="3" fill="#fdd835" class="data-dot data-dot--2"
                style="offset-path:path('M75,105 L110,110 L175,35'); filter:drop-shadow(0 0 6px rgba(253,216,53,0.9));" />
              <!-- Пакет 3: N3 → N4 -->
              <circle r="3" fill="#fdd835" class="data-dot data-dot--3"
                style="offset-path:path('M55,165 L170,150'); filter:drop-shadow(0 0 6px rgba(253,216,53,0.9));" />

              <!-- === УЗЛЫ (периферийные) === -->
              <g class="node"><circle cx="40" cy="45" r="3" fill="#fdd835" filter="url(#nodeGlow)" /></g>
              <g class="node node--2"><circle cx="175" cy="35" r="3" fill="#fdd835" filter="url(#nodeGlow)" /></g>
              <g class="node node--3"><circle cx="55" cy="165" r="3" fill="#fdd835" filter="url(#nodeGlow)" /></g>
              <g class="node node--4"><circle cx="170" cy="150" r="3" fill="#fdd835" filter="url(#nodeGlow)" /></g>
              <g class="node node--5"><circle cx="145" cy="80" r="2.5" fill="#fdd835" filter="url(#nodeGlow)" /></g>
              <g class="node node--6"><circle cx="75" cy="105" r="2.5" fill="#fdd835" filter="url(#nodeGlow)" /></g>

              <!-- === ЦЕНТРАЛЬНЫЙ УЗЕЛ === -->
              <g class="node-main">
                <!-- Внешняя оболочка -->
                <circle cx="110" cy="110" r="14" fill="rgba(253,216,53,0.04)" />
                <circle cx="110" cy="110" r="9" fill="rgba(253,216,53,0.08)" />
                <!-- Ядро -->
                <circle cx="110" cy="110" r="6" fill="#fdd835" filter="url(#centerGlow)" />
                <!-- Точка внутри -->
                <circle cx="110" cy="110" r="2.5" fill="#fff" />
              </g>

              <!-- === ЗВЁЗДНАЯ ПЫЛЬ (мелкие точки фона) === -->
              <circle cx="35" cy="80" r="0.8" fill="rgba(255,255,255,0.2)" filter="url(#starDust)" />
              <circle cx="180" cy="75" r="0.6" fill="rgba(255,255,255,0.15)" filter="url(#starDust)" />
              <circle cx="88" cy="28" r="0.5" fill="rgba(255,255,255,0.12)" />
              <circle cx="155" cy="175" r="0.7" fill="rgba(255,255,255,0.18)" filter="url(#starDust)" />
              <circle cx="60" cy="50" r="0.4" fill="rgba(255,255,255,0.1)" />
              <circle cx="165" cy="120" r="0.5" fill="rgba(255,255,255,0.12)" />
            </svg>
          </div>
          <div class="hero-content">
            <div class="hero-badge"><i class="fas fa-shield-halved"></i> Единое информационное пространство качества</div>
            <h1>Навигатор <span class="accent-text">Качества</span></h1>
            <div class="hero-cta">
              <button class="hero-btn-primary" id="hero-cta-start"><i class="fas fa-arrow-right-long"></i> Выбрать направление</button>
            </div>
          </div>
        </div>

        <div class="info-banner">
          <i class="fas fa-circle-info"></i>
          <div class="info-banner-text">
            <strong>Добро пожаловать в Навигатор Качества</strong> — единая точка доступа к регламентам, отчётам, стандартам и инструментам для всех участников процесса мониторинга качества
          </div>
        </div>

        <div class="roles-section">
          <div class="section-title"><i class="fas fa-users"></i> Для кого этот портал</div>
          <div class="roles-grid">
            <div class="role-card"><div class="role-icon ri-purple"><i class="fas fa-user-tie"></i></div><div class="role-name">Руководители и TL</div><div class="role-desc">Отчет по качеству, аналитика, регламенты, апелляции</div></div>
            <div class="role-card"><div class="role-icon ri-amber"><i class="fas fa-headset"></i></div><div class="role-name">Операторы</div><div class="role-desc">Отчет по качеству, стандарты, новости</div></div>
            <div class="role-card"><div class="role-icon ri-blue"><i class="fas fa-graduation-cap"></i></div><div class="role-name">Ученики и Новички</div><div class="role-desc">Карточки, процессы взаимодействия, стандарты</div></div>
          </div>
        </div>

        <div class="section-title section-title--sub"><span class="arrow-down"><i class="beeline-icons">nav_arrow_down</i></span> Выберите направление, чтобы продолжить</div>
        <div class="directions-grid" id="dir-grid"></div>
        </div>
      `;
      showSyncBadge();

      // Hero CTA handler — прокрутка к направлениям
      const heroStart = document.getElementById('hero-cta-start');
      if (heroStart) heroStart.addEventListener('click', () => { document.getElementById('dir-grid').scrollIntoView({ behavior: 'smooth' }); });
      const g = document.getElementById('dir-grid');
      function rb(buttons) {
        g.innerHTML = ''; const f = document.createDocumentFragment();
        buttons.forEach((btn, idx) => {
          const c = document.createElement('div'); c.className = 'dir-card'; c.setAttribute('data-id', btn.id);
          const cc = CE[idx % CE.length];
          c.innerHTML = `<div class="card-emoji ${cc}">${(btn.icon || 'star').startsWith('fa-') ? `<i class="fas ${btn.icon}"></i>` : `<i class="beeline-icons">${btn.icon || 'star'}</i>`}</div><div class="dir-name">${btn.text}</div>${btn.sub ? '<div class="dir-sub">' + btn.sub + '</div>' : ''}`;
          if (isAdmin) createDragIcon(c);
          c.onclick = (e) => { if (e.target.closest('.drag-handle') || e.target.closest('.adm-icon')) return; localStorage.setItem('currentDirection', btn.id); window.location.href = `?dir=${btn.id}`; };
          if (isAdmin) {
            createAdminIcons(c, isAdmin, async () => {
              if (confirm('Удалить?')) { mainButtons.splice(mainButtons.findIndex(b => b.id === btn.id), 1); saveMainButtons(); await syncToCloud(); renderSidebarNav(); rb(mainButtons); }
            }, async () => {
              const nt = prompt('Название:', btn.text); if (nt) btn.text = nt;
              const nh = prompt('Ссылка:', btn.href); if (nh) btn.href = nh;
              const ns = prompt('Подзаголовок:', btn.sub || ''); if (ns !== null) btn.sub = ns;
              saveMainButtons(); await syncToCloud(); renderSidebarNav(); rb(mainButtons);
            });
          }
          if (isAdmin) {
            c.draggable = true;
            c.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/plain', btn.id); c.classList.add('dragging'); });
            c.addEventListener('dragend', () => c.classList.remove('dragging'));
            c.addEventListener('dragover', (e) => e.preventDefault());
            c.addEventListener('drop', (e) => { e.preventDefault(); const fi = mainButtons.findIndex(b => b.id === e.dataTransfer.getData('text/plain')); const ti = mainButtons.findIndex(b => b.id === btn.id); if (fi === ti) return; const mv = mainButtons.splice(fi, 1)[0]; mainButtons.splice(ti, 0, mv); saveMainButtons(); syncToCloud(); renderSidebarNav(); rb(mainButtons); });
          }
          f.appendChild(c);
        });
        g.appendChild(f);
        if (isAdmin) {
          const ac = document.createElement('div'); ac.className = 'dir-card';
          ac.innerHTML = '<div class="card-emoji ce-teal"><i class="beeline-icons">add</i></div><div class="dir-name">Добавить</div>';
          ac.onclick = async () => { const t = prompt('Название:'); if (t) { const id = t.toLowerCase().replace(/\s/g, '_'); const h = prompt('Ссылка:', `?dir=${id}`); const s = prompt('Подзаголовок:', ''); mainButtons.push({ id, text: t, icon: 'fa-folder', href: h || `?dir=${id}`, sub: s || '' }); saveMainButtons(); if (!directions[id]) directions[id] = { name: t, icon: 'fa-folder', links: [] }; saveDirections(); await syncToCloud(); renderSidebarNav(); rb(mainButtons); } };
          g.appendChild(ac);
        }
      }
      rb(cb);
    }

    // ===== LINK ICON CATEGORIZATION =====
    // Возвращает { icon, colorClass } в зависимости от текста ссылки
    // Иконки — yellowbe DS ligatures (https://yellowbe.beeline.ru/backoffice/foundation/icons)
    function getLinkIcon(text) {
      const t = (text || '').toLowerCase();
      // Отчёты
      if (t.includes('отчет') || t.includes('отчёт')) return { icon: 'reports', colorClass: 'link-icon--report' };
      // Регламенты, схемы, стандарты
      if (t.includes('регламент') || t.includes('схема') || t.includes('стандарт') || t.includes('матриц') || t.includes('карта фрода')) return { icon: 'journal', colorClass: 'link-icon--regulation' };
      // Встречи, новости
      if (t.includes('встреч') || t.includes('новост') || t.includes('чат') || t.includes('бот')) return { icon: 'megaphone', colorClass: 'link-icon--news' };
      // QMS
      if (t.includes('qms')) return { icon: 'certificate', colorClass: 'link-icon--qms' };
      // Для руководителей, апелляции
      if (t.includes('руководител') || t.includes('апелляц') || t.includes('жалоб') || t.includes('благодарност') || t.includes('ошибк')) return { icon: 'suitcase', colorClass: 'link-icon--management' };
      // Дефолт
      return { icon: 'link', colorClass: 'link-icon--default' };
    }

    // ===== LINK CATEGORY PRIORITY (fixed order) =====
    // Порядок: Отчёты → Регламенты → Новости → Руководители → QMS → Остальное
    function getCategoryPriority(text) {
      const t = (text || '').toLowerCase();
      if (t.includes('отчет') || t.includes('отчёт')) return 0;
      if (t.includes('регламент') || t.includes('схема') || t.includes('стандарт') || t.includes('матриц') || t.includes('карта фрода')) return 1;
      if (t.includes('встреч') || t.includes('новост') || t.includes('чат') || t.includes('бот')) return 2;
      if (t.includes('руководител') || t.includes('апелляц') || t.includes('жалоб') || t.includes('благодарност') || t.includes('ошибк')) return 3;
      if (t.includes('qms')) return 4;
      return 5;
    }
    function sortLinksByCategory(links) {
      return [...links].sort((a, b) => {
        const pa = getCategoryPriority(a.text);
        const pb = getCategoryPriority(b.text);
        if (pa !== pb) return pa - pb;
        return (a.text || '').localeCompare(b.text || '', 'ru');
      });
    }

    function renderDirectionPage(dk) {
      const dir = directions[dk];
      setHeader(dir.name);
      // Показываем поиск на странице направления (скрыт на главной)
      const ts = document.querySelector('.topbar-search'); if (ts) ts.style.display = '';
      // Сброс строки поиска при переходе на страницу направления
      const si = document.querySelector('.topbar-search input'); if (si) si.value = '';
      const m = document.getElementById('main-content');
      m.innerHTML = `<div class="section-title">${(dir.icon || 'folder').startsWith('fa-') ? `<i class="fas ${dir.icon}"></i>` : `<i class="beeline-icons">${dir.icon || 'folder'}</i>`} ${dir.name}</div><div id="links-container" class="links-grid"></div>`;
      showSyncBadge();
      const ct = document.getElementById('links-container');
      // Порядок ссылок — как задал админ (сохраняется в dir.links и синхронизируется через backend API)
      const ol = dir.links;
      function rl(links) {
        ct.innerHTML = ''; const f = document.createDocumentFragment();
        links.forEach((link, idx) => {
          const b = document.createElement('button'); b.className = 'link-item'; b.setAttribute('data-idx', idx);
          const li = getLinkIcon(link.text);
          // Плашка «Для руководителей» — по флагу isManager (админ присваивает вручную)
          const isManager = !!link.isManager;
          if (isManager) b.classList.add('link-item--manager');
          b.innerHTML = `<i class="beeline-icons link-icon ${li.colorClass}">${li.icon}</i><span class="link-text">${link.text}</span>${isManager ? '<span class="manager-badge"><i class="beeline-icons">star</i> Для руководителей</span>' : ''}`;
          if (isAdmin) createDragIcon(b, true);
          const hasTooltip = link.text && link.text.includes('Отчет по качеству');
          if (hasTooltip) addQualityTooltip(b);
          b.onclick = (e) => {
            if (e.target.closest('.link-drag-handle') || e.target.closest('.tt-copy-btn')) return;
            if (e.target.closest('.tooltip-card')) return;
            // Переходим по ссылке (если есть)
            if (link.href && link.href !== '#') {
              // Статистика кликов — сохранение в localStorage + отправка на backend
              try {
                const clicks = JSON.parse(localStorage.getItem('quality_clicks') || '[]');
                const clickData = { timestamp: new Date().toISOString(), direction: dir.name, linkText: link.text };
                clicks.push(clickData);
                localStorage.setItem('quality_clicks', JSON.stringify(clicks));
                // Отправка на backend (если доступен)
                cloudPush('clicks', clickData);
              } catch (e) { /* ignore */ }
              window.open(link.href, '_blank');
            }
          };
          // Tooltip открывается при наведении (информационный)
          if (hasTooltip) {
            b.addEventListener('mouseenter', () => {
              document.querySelectorAll('.link-item.tt-open').forEach(el => el.classList.remove('tt-open'));
              b.classList.add('tt-open');
            });
            b.addEventListener('mouseleave', () => {
              b.classList.remove('tt-open');
            });
          }
          if (isAdmin) {
            createAdminIcons(b, isAdmin, async () => {
              if (confirm('Удалить?')) { dir.links.splice(dir.links.findIndex(l => l.text === link.text), 1); saveDirections(); await syncToCloud(); rl(dir.links); }
            }, async () => {
              const nt = prompt('Текст:', link.text); if (nt) link.text = nt;
              const nh = prompt('URL:', link.href); if (nh) link.href = nh;
              const isMgr = confirm('Отметить как «Для руководителей»?\n\n(ОК = да, Отмена = нет / снять метку)');
              link.isManager = isMgr;
              saveDirections(); await syncToCloud(); rl(dir.links);
            });
            b.draggable = true;
            b.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/plain', String(idx)); b.classList.add('dragging'); });
            b.addEventListener('dragend', () => b.classList.remove('dragging'));
            b.addEventListener('dragover', (e) => e.preventDefault());
            b.addEventListener('drop', (e) => {
              e.preventDefault();
              const fi = parseInt(e.dataTransfer.getData('text/plain'), 10);
              if (isNaN(fi) || fi === idx) return;
              const src = links[fi]; const dst = links[idx];
              const si = dir.links.findIndex(l => l.text === src.text);
              const di2 = dir.links.findIndex(l => l.text === dst.text);
              if (si === -1 || di2 === -1 || si === di2) return;
              const mv = dir.links.splice(si, 1)[0];
              dir.links.splice(di2, 0, mv);
              saveDirections(); syncToCloud(); rl(dir.links);
            });
          }
          f.appendChild(b);
        });
        ct.appendChild(f);
        if (isAdmin) {
          const ab = document.createElement('button'); ab.className = 'link-item';
          ab.innerHTML = '<i class="beeline-icons link-icon link-icon--default">add</i><span class="link-text">Добавить</span>';
          ab.onclick = async () => {
            const t = prompt('Текст:'); if (!t) return;
            const h = prompt('URL:', '#');
            const isManager = confirm('Отметить как «Для руководителей»?');
            dir.links.push({ text: t, href: h || '#', type: 'link', isManager: isManager });
            saveDirections(); await syncToCloud(); rl(dir.links);
          };
          ct.appendChild(ab);
        }
      }
      rl(ol);

      // Глобальный обработчик: клик вне tooltip — закрывает
      if (!window._ttCloseHandler) {
        window._ttCloseHandler = true;
        document.addEventListener('click', function(e) {
          if (!e.target.closest('.link-item.tt-open')) {
            document.querySelectorAll('.link-item.tt-open').forEach(el => el.classList.remove('tt-open'));
          }
        }, true);
      }
    }

    async function renderFeedbackStatsPage() {
      const m = document.getElementById('main-content');
      setHeader('Статистика');
      const ts = document.querySelector('.topbar-search'); if (ts) ts.style.display = 'none';

      // Загружаем отзывы: с сервера (если админ + API) или из localStorage
      let fb = JSON.parse(localStorage.getItem('quality_feedback') || '[]');
      if (API_BASE && isAdmin) {
        const serverFb = await apiGet('feedback');
        if (serverFb && Array.isArray(serverFb) && serverFb.length > 0) {
          // Мержим серверные + локальные, убирая дубли
          const seen = new Set(fb.map(f => f.timestamp + '|' + (f.direction||'') + '|' + (f.comment||'')));
          serverFb.forEach(item => {
            if (item && item.timestamp) {
              const key = item.timestamp + '|' + (item.direction||'') + '|' + (item.comment||'');
              if (!seen.has(key)) { fb.push(item); seen.add(key); }
            }
          });
          // Сортируем по дате (новые сверху)
          fb.sort((a, b) => {
            const da = new Date(a.createdAt || a.timestamp || 0);
            const db = new Date(b.createdAt || b.timestamp || 0);
            return db - da;
          });
          // Сохраняем в localStorage для кэша
          localStorage.setItem('quality_feedback', JSON.stringify(fb));
        }
      }
      feedbackList = fb;
      const t = fb.length;

      function prd(ds) { if (!ds) return null; const p = ds.split(',')[0].split('.'); if (p.length !== 3) return null; return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10)); }
      const mn = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
      const mnShort = ["Янв","Фев","Мар","Апр","Май","Июн","Июл","Авг","Сен","Окт","Ноя","Дек"];

      // Группировка данных по периоду
      function groupByPeriod(items, period) {
        const groups = {};
        items.forEach(f => {
          const d = prd(f.timestamp);
          if (!d) return;
          let key, label;
          if (period === 'year') { key = String(d.getFullYear()); label = String(d.getFullYear()); }
          else if (period === 'month') { key = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0'); label = mnShort[d.getMonth()] + ' ' + d.getFullYear(); }
          else if (period === 'week') {
            const onejan = new Date(d.getFullYear(),0,1);
            const week = Math.ceil(((d - onejan) / 86400000 + onejan.getDay() + 1) / 7);
            key = d.getFullYear() + '-W' + String(week).padStart(2,'0'); label = 'Нед. ' + week + ' (' + d.getFullYear() + ')';
          }
          if (!groups[key]) groups[key] = { label, count: 0, items: [] };
          groups[key].count++;
          groups[key].items.push(f);
        });
        return Object.keys(groups).sort().map(k => groups[k]);
      }

      // Направления для фильтра
      const dirSet = new Set(); fb.forEach(f => { if (f.direction) dirSet.add(f.direction); });
      const dirList = Array.from(dirSet).sort();

      // Загружаем счётчики заходов с сервера
      let userVisits = 0, adminVisits = 0, lastVisits = [];
      if (API_BASE) {
        const visitsData = await apiGet('visits');
        if (visitsData) {
          userVisits = visitsData.userVisits || 0;
          adminVisits = visitsData.adminVisits || 0;
          lastVisits = visitsData.lastVisits || [];
        }
      }

      m.innerHTML = `
        <div class="section-title"><i class="beeline-icons">reports</i> Статистика обратной связи</div>
        <div class="analytics-wrap">
          <div class="metric-grid">
            <div class="metric-box"><div class="metric-val">${t}</div><div class="metric-lbl">Всего отзывов</div></div>
            <div class="metric-box"><div class="metric-val">${dirList.length}</div><div class="metric-lbl">Направлений</div></div>
            <div class="metric-box"><div class="metric-val">${userVisits}</div><div class="metric-lbl">Заходы пользователей</div></div>
            <div class="metric-box"><div class="metric-val">${adminVisits}</div><div class="metric-lbl">Заходы админа</div></div>
            <div class="metric-box"><div class="metric-val">${fb.length > 0 ? (function() { var maxDate = 0; fb.forEach(function(f) { var d = prd(f.timestamp); if (d && d.getTime() > maxDate) maxDate = d.getTime(); }); return maxDate > 0 ? new Date(maxDate).toLocaleDateString('ru-RU') : '—'; })() : '—'}</div><div class="metric-lbl">Последний отзыв</div></div>
          </div>

          <div class="filter-row" style="margin-bottom:16px;">
            <select id="period-select">
              <option value="month">По месяцам</option>
              <option value="week">По неделям</option>
              <option value="year">По годам</option>
            </select>
            <select id="dir-filter">
              <option value="">Все направления</option>
              ${dirList.map(d => '<option value="' + d + '">' + d + '</option>').join('')}
            </select>
            <select id="yf">
              <option value="">Все годы</option>
            </select>
          </div>

          <div style="background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius-12);padding:20px;margin-bottom:20px;">
            <div style="font-size:14px;font-weight:500;color:var(--text-secondary);margin-bottom:12px;">Динамика отзывов</div>
            <div style="height:180px;position:relative;"><canvas id="dyn-chart"></canvas></div>
          </div>

          <div class="table-wrap" style="max-height:600px;overflow-y:auto;">
            <table id="ft">
              <thead><th>Период</th><th>Дата</th><th>Направление</th><th>Комментарий</th></thead>
              <tbody></tbody>
            </table>
          </div>

          <button class="dl-btn" id="efb" style="margin-top:10px;"><i class="beeline-icons">download</i> Скачать все отзывы (.xlsx)</button>

          <div style="margin-top:32px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">
              <div style="font-size:14px;font-weight:500;color:var(--text-secondary);">Статистика заходов</div>
              <button class="dl-btn" id="dl-visits" style="height:36px;padding:0 16px;font-size:13px;"><i class="beeline-icons">download</i> Скачать заходы (.xlsx)</button>
            </div>
            <div class="table-wrap" style="max-height:400px;overflow-y:auto;">
              <table id="visits-table">
                <thead><th>Дата и время</th><th>Тип пользователя</th></thead>
                <tbody></tbody>
              </table>
            </div>
          </div>
        </div>`;

      let chartInstance = null;

      function renderChart(period, dirFilter, yearFilter) {
        let filtered = [...fb];
        if (dirFilter) filtered = filtered.filter(f => f.direction === dirFilter);
        if (yearFilter) filtered = filtered.filter(f => { const d = prd(f.timestamp); return d && String(d.getFullYear()) === yearFilter; });

        const grouped = groupByPeriod(filtered, period);
        const labels = grouped.map(g => g.label);
        const counts = grouped.map(g => g.count);

        const ctx = document.getElementById('dyn-chart');
        if (!ctx) return;

        if (chartInstance) chartInstance.destroy();

        const isDark = document.body.classList.contains('dark');
        const tickColor = isDark ? '#999' : '#666';
        const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(128,128,128,0.1)';
        const tooltipBg = isDark ? 'rgba(30,30,30,0.95)' : 'rgba(0,0,0,0.8)';
        const tooltipText = isDark ? '#eee' : '#fff';

        chartInstance = new Chart(ctx, {
          type: period === 'year' ? 'bar' : 'line',
          data: {
            labels: labels,
            datasets: [{
              label: 'Количество отзывов',
              data: counts,
              backgroundColor: period === 'year' ? 'rgba(253,196,53,0.5)' : 'rgba(253,196,53,0.15)',
              borderColor: '#fdc435',
              borderWidth: 2,
              tension: 0.35,
              fill: period !== 'year',
              pointBackgroundColor: '#fdc435',
              pointBorderColor: isDark ? '#222' : '#fff',
              pointBorderWidth: 2,
              pointRadius: 4,
              pointHoverRadius: 6
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: {
                backgroundColor: tooltipBg,
                titleColor: tooltipText,
                bodyColor: tooltipText,
                padding: 12,
                cornerRadius: 8
              }
            },
            scales: {
              x: { grid: { display: false }, ticks: { color: tickColor, font: { size: 11 } } },
              y: { beginAtZero: true, grid: { color: gridColor }, ticks: { color: tickColor, stepSize: 1 } }
            }
          }
        });

        // Единая таблица: Период | Дата | Направление | Комментарий
        const tb = document.querySelector('#ft tbody');
        if (tb) {
          tb.innerHTML = '';
          // Сортируем: новые сверху
          const sortedFb = [...filtered].sort((a, b) => {
            const da = new Date(a.createdAt || a.timestamp || 0);
            const db = new Date(b.createdAt || b.timestamp || 0);
            return db - da;
          });
          if (sortedFb.length === 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = '<td colspan="4" style="text-align:center;color:var(--text-muted);padding:20px;">Нет отзывов</td>';
            tb.appendChild(tr);
          } else {
            // Группируем по периоду, выводим каждый отзыв отдельной строкой
            const grouped = groupByPeriod(sortedFb, period);
            grouped.reverse().forEach(g => {
              g.items.forEach((f) => {
                const tr = document.createElement('tr');
                const periodLabel = g.label;
                const date = f.timestamp || '—';
                const dir = f.direction || '—';
                const comment = f.comment || 'Без комментария';
                tr.innerHTML = '<td style="white-space:nowrap;color:var(--accent);font-weight:500;">' + periodLabel + '</td><td style="white-space:nowrap;">' + date + '</td><td>' + dir + '</td><td style="max-width:500px;word-wrap:break-word;">' + comment + '</td>';
                tb.appendChild(tr);
              });
            });
          }
        }
      }

      // Заполняем фильтр годов
      const ys = new Set(); fb.forEach(f => { const d = prd(f.timestamp); if (d) ys.add(d.getFullYear()); });
      const sy = Array.from(ys).sort((a, b) => b - a);
      const yfSel = document.getElementById('yf');
      sy.forEach(y => { const opt = document.createElement('option'); opt.value = y; opt.textContent = y; yfSel.appendChild(opt); });

      function updateAll() {
        renderChart(document.getElementById('period-select').value, document.getElementById('dir-filter').value, document.getElementById('yf').value);
      }

      document.getElementById('period-select').addEventListener('change', updateAll);
      document.getElementById('dir-filter').addEventListener('change', updateAll);
      document.getElementById('yf').addEventListener('change', updateAll);

      document.getElementById('efb').addEventListener('click', () => {
        // Экспортируем ВСЕ отзывы из таблицы (Период + Дата + Направление + Комментарий)
        const period = document.getElementById('period-select').value;
        const exportData = feedbackList.map(f => {
          const d = prd(f.timestamp);
          let periodLabel = '—';
          if (d) {
            if (period === 'year') periodLabel = String(d.getFullYear());
            else if (period === 'month') periodLabel = mnShort[d.getMonth()] + ' ' + d.getFullYear();
            else if (period === 'week') {
              const onejan = new Date(d.getFullYear(), 0, 1);
              const week = Math.ceil(((d - onejan) / 86400000 + onejan.getDay() + 1) / 7);
              periodLabel = 'Нед. ' + week + ' (' + d.getFullYear() + ')';
            }
          }
          return {
            'Период': periodLabel,
            'Дата': f.timestamp || '',
            'Направление': f.direction || '',
            'Комментарий': f.comment || ''
          };
        });
        const ws = XLSX.utils.json_to_sheet(exportData);
        ws['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 18 }, { wch: 60 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Отзывы');
        XLSX.writeFile(wb, 'отзывы_' + new Date().toISOString().slice(0, 10) + '.xlsx');
      });

      // ===== Заполнение таблицы заходов =====
      const visitsTb = document.querySelector('#visits-table tbody');
      if (visitsTb) {
        visitsTb.innerHTML = '';
        if (lastVisits.length === 0) {
          const tr = document.createElement('tr');
          tr.innerHTML = '<td colspan="2" style="text-align:center;color:var(--text-muted);padding:20px;">Нет данных</td>';
          visitsTb.appendChild(tr);
        } else {
          lastVisits.forEach(v => {
            const tr = document.createElement('tr');
            const dt = v.timestamp ? new Date(v.timestamp).toLocaleString('ru-RU') : '—';
            const userType = v.isAdmin ? 'Админ' : 'Пользователь';
            const userColor = v.isAdmin ? 'var(--accent-purple)' : 'var(--accent)';
            tr.innerHTML = '<td style="white-space:nowrap;">' + dt + '</td><td style="font-weight:600;color:' + userColor + ';">' + userType + '</td>';
            visitsTb.appendChild(tr);
          });
        }
      }

      // ===== Скачивание таблицы заходов =====
      document.getElementById('dl-visits').addEventListener('click', () => {
        const exportData = lastVisits.map(v => ({
          'Дата и время': v.timestamp ? new Date(v.timestamp).toLocaleString('ru-RU') : '',
          'Тип пользователя': v.isAdmin ? 'Админ' : 'Пользователь'
        }));
        const ws = XLSX.utils.json_to_sheet(exportData);
        ws['!cols'] = [{ wch: 28 }, { wch: 20 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Заходы');
        XLSX.writeFile(wb, 'заходы_' + new Date().toISOString().slice(0, 10) + '.xlsx');
      });

      updateAll();
    }

    let fbModalInited = false;
    function initFeedbackModal() {
      if (fbModalInited) return;
      fbModalInited = true;
      const overlay = document.getElementById('fb-overlay');
      const openBtn = document.getElementById('fb-open-btn');
      const closeBtn = document.getElementById('fb-close-btn');
      const sendBtn = document.getElementById('send-fb');
      const fbText = document.getElementById('fb-text');
      if (!overlay || !openBtn) return;
      openBtn.addEventListener('click', () => overlay.classList.add('open'));
      closeBtn.addEventListener('click', () => overlay.classList.remove('open'));
      overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.classList.remove('open'); });
      sendBtn.addEventListener('click', () => {
        const currentDir = dirKey && directions[dirKey] ? directions[dirKey].name : '';
        const fbItem = { timestamp: new Date().toLocaleString(), direction: currentDir, rating: 0, comment: fbText.value || 'Без комментария' };
        const list = JSON.parse(localStorage.getItem('quality_feedback') || '[]');
        list.push(fbItem);
        localStorage.setItem('quality_feedback', JSON.stringify(list));
        // Отправка на backend (если доступен) — для multi-user
        cloudPush('feedback', { ...fbItem, createdAt: new Date().toISOString() });
        alert('Спасибо за отзыв!');
        fbText.value = '';
        overlay.classList.remove('open');
      });
    }

    function addAdminFAB() { /* no-op: admin теперь в sidebar */ }

    // ===== SIDEBAR / NAV HANDLERS =====
    const sidebar = document.getElementById('sidebar');
    const sidebarOverlay = document.getElementById('sidebar-overlay');
    const menuToggle = document.getElementById('menu-toggle');
    function closeSidebar() { sidebar.classList.remove('open'); sidebarOverlay.classList.remove('open'); }
    if (menuToggle) menuToggle.addEventListener('click', () => { sidebar.classList.toggle('open'); sidebarOverlay.classList.toggle('open'); });
    if (sidebarOverlay) sidebarOverlay.addEventListener('click', closeSidebar);

    // ===== SIDEBAR COLLAPSE (persist in localStorage) =====
    const collapseBtn = document.getElementById('sidebar-collapse-btn');
    if (localStorage.getItem('sidebarCollapsed') === 'true') document.body.classList.add('sidebar-collapsed');
    if (collapseBtn) {
      collapseBtn.addEventListener('click', () => {
        const isCollapsed = document.body.classList.toggle('sidebar-collapsed');
        localStorage.setItem('sidebarCollapsed', isCollapsed);
        // DS yellowbe icon: одна ligature 'menu_open' для обоих состояний (rotate через CSS при желании)
        const ci = collapseBtn.querySelector('i');
        if (ci) { ci.className = 'beeline-icons'; ci.textContent = 'menu_open'; }
      });
      // init icon state
      if (document.body.classList.contains('sidebar-collapsed')) {
        const ci = collapseBtn.querySelector('i');
        if (ci) { ci.className = 'beeline-icons'; ci.textContent = 'menu_open'; }
      }
    }

    // ===== TOOLTIP для свёрнутой панели (показывать название при наведении) =====
    const sidebarTT = document.getElementById('sidebar-tt');
    const TT_OFFSET = 14;
    if (sidebarTT) {
      document.addEventListener('mouseover', (e) => {
        const item = e.target.closest('.nav-item');
        const isCollapsed = document.body.classList.contains('sidebar-collapsed');
        if (!item || !item.getAttribute('data-label') || !isCollapsed) { hideSidebarTT(); return; }
        const label = item.getAttribute('data-label');
        const span = item.querySelector('span');
        // Показывать тултип, только когда название скрыто (свёрнутое меню)
        if (span && span.offsetParent !== null) { hideSidebarTT(); return; }
        const r = item.getBoundingClientRect();
        sidebarTT.textContent = label;
        sidebarTT.style.left = (r.right + TT_OFFSET) + 'px';
        sidebarTT.style.top = (r.top + r.height / 2) + 'px';
        sidebarTT.classList.add('show');
      });
      document.addEventListener('mouseout', (e) => {
        if (e.target.closest('.nav-item')) hideSidebarTT();
      });
    }
    function hideSidebarTT() { const t = document.getElementById('sidebar-tt'); if (t) t.classList.remove('show'); }

    // ===== SEARCH FILTER (фильтрация dir-card на главной и link-item на странице направления) =====
    const searchInput = document.querySelector('.topbar-search input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        const q = e.target.value.trim().toLowerCase();
        // На главной: фильтровать .dir-card по .dir-name и .dir-sub
        const dirCards = document.querySelectorAll('.dir-card');
        if (dirCards.length > 0) {
          dirCards.forEach(card => {
            const name = (card.querySelector('.dir-name')?.textContent || '').toLowerCase();
            const sub = (card.querySelector('.dir-sub')?.textContent || '').toLowerCase();
            if (!q || name.includes(q) || sub.includes(q)) {
              card.style.display = '';
            } else {
              card.style.display = 'none';
            }
          });
        }
        // На странице направления: фильтровать .link-item по .link-text
        const linkItems = document.querySelectorAll('.link-item');
        if (linkItems.length > 0) {
          linkItems.forEach(item => {
            const text = (item.querySelector('.link-text')?.textContent || '').toLowerCase();
            if (!q || text.includes(q)) {
              item.style.display = '';
            } else {
              item.style.display = 'none';
            }
          });
        }
      });
    }

    const adminBtn = document.getElementById('admin-toggle');
    if (adminBtn) adminBtn.addEventListener('click', async () => {
      if (isAdmin) {
        localStorage.removeItem('adminActive');
        localStorage.removeItem('adminToken');
        adminToken = null;
        location.reload();
      } else {
        const p = prompt('Пароль:');
        if (!p) return;
        if (API_BASE) {
          // Авторизация через backend
          try {
            const res = await fetch(`${API_BASE}/login`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ password: p })
            });
            const data = await res.json();
            if (data.success && data.token) {
              adminToken = data.token;
              localStorage.setItem('adminToken', adminToken);
              localStorage.setItem('adminActive', 'true');
              location.reload();
            } else {
              alert('Неверный пароль');
            }
          } catch (e) {
            alert('Ошибка подключения к серверу');
          }
        } else {
          // Локальный режим (без backend)
          if (p === ADMIN_PASSWORD) { localStorage.setItem('adminActive', 'true'); location.reload(); }
          else alert('Неверный пароль');
        }
      }
    });
    const statsBtn = document.getElementById('feedback-stats-btn');
    if (statsBtn) statsBtn.addEventListener('click', () => { window.location.href = '?feedback=true'; });
    const homeBtn = document.getElementById('nav-home');
    if (homeBtn) homeBtn.addEventListener('click', () => { localStorage.removeItem('currentDirection'); window.location.href = window.location.pathname; });

    // ===== THEME =====
    const themeToggle = document.getElementById('theme-toggle');
    if (localStorage.getItem('dark-theme-v2') === 'true') document.body.classList.add('dark');
    // init theme icon based on current theme (DS yellowbe: sun = light, half_moon = dark — обе ligatures подтверждены)
    const themeIcon = themeToggle.querySelector('i');
    if (themeIcon) {
      const isDark = document.body.classList.contains('dark');
      themeIcon.className = 'beeline-icons';
      themeIcon.textContent = isDark ? 'half_moon' : 'sun';
    }
    themeToggle.addEventListener('click', () => {
      document.body.classList.toggle('dark');
      const isD = document.body.classList.contains('dark');
      localStorage.setItem('dark-theme-v2', isD);
      const ti = themeToggle.querySelector('i');
      ti.className = 'beeline-icons';
      ti.textContent = isD ? 'half_moon' : 'sun';
      themeToggle.querySelector('span').textContent = 'Тема';
      // Перерисовка графика на странице статистики при смене темы
      const ps = document.getElementById('period-select');
      if (ps) ps.dispatchEvent(new Event('change'));
    });

    // ===== LOAD FROM LOCAL STORAGE (instant) =====
    const savedData = localStorage.getItem('directionsData');
    if (savedData) { try { const l = JSON.parse(savedData); migrateDirections(l); Object.keys(directions).forEach(k => { if (l[k] && Array.isArray(l[k].links)) directions[k] = l[k]; }); } catch (e) {} }
    const savedMB = localStorage.getItem('mainButtons');
    if (savedMB) { try { const p = JSON.parse(savedMB); if (Array.isArray(p) && p.length > 0) { migrateMainButtons(p); mainButtons = p; } } catch (e) {} }

    // ===== RENDER INSTANTLY =====
    // feedbackList загружается внутри renderFeedbackStatsPage (с сервера для админа)
    renderSidebarNav();
    renderMainPage();
    initFeedbackModal();

    // ===== VISIT TRACKING =====
    // Отправляем заход на сервер (отдельно пользователь, отдельно админ)
    if (API_BASE) {
      fetch(`${API_BASE}/visits`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ isAdmin: isAdmin })
      }).catch(() => {});
    }

    // ===== API SYNC (если backend доступен) =====
    if (API_BASE) {
      // Первичная загрузка с сервера
      syncFromCloud();
      // Polling каждые 30 сек как fallback
      setInterval(syncFromCloud, 30000);

      // SSE — мгновенные обновления при правках админа
      if (typeof EventSource !== 'undefined') {
        try {
          const evtSource = new EventSource(`${API_BASE}/events`);
          evtSource.addEventListener('directions-updated', () => syncFromCloud());
          evtSource.addEventListener('mainButtons-updated', () => syncFromCloud());
          evtSource.addEventListener('feedback-added', () => syncFeedbackFromCloud());
        } catch (e) { console.warn('SSE unavailable, polling only:', e); }
      }
    }
  </script>
</body>
</html>
