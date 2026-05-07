<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover"/>
  <title>Quality | Хаб Качества 360°</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --accent: #3B82F6;
      --accent-hover: #60A5FA;
      --accent-soft: rgba(59,130,246,0.1);
      --accent-glow: rgba(59,130,246,0.3);
      --teal: #14B8A6;
      --teal-soft: rgba(20,184,166,0.1);
      --rose: #FB7185;
      --rose-soft: rgba(251,113,133,0.1);
      --amber: #FBBF24;
      --amber-soft: rgba(251,191,36,0.1);
      --violet: #A78BFA;
      --violet-soft: rgba(167,139,250,0.1);
      --grad-hero: linear-gradient(135deg, #3B82F6 0%, #8B5CF6 50%, #EC4899 100%);
      --grad-btn: linear-gradient(135deg, #3B82F6 0%, #8B5CF6 100%);
      --grad-card: linear-gradient(160deg, rgba(59,130,246,0.06) 0%, rgba(139,92,246,0.04) 100%);
      --bg: #F8FAFC;
      --bg-card: #FFFFFF;
      --bg-card-hover: #F1F5F9;
      --text-primary: #0F172A;
      --text-secondary: #475569;
      --text-muted: #94A3B8;
      --border: #E2E8F0;
      --border-hover: #CBD5E1;
      --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
      --shadow-md: 0 4px 12px rgba(0,0,0,0.06);
      --shadow-lg: 0 12px 40px rgba(0,0,0,0.08);
      --shadow-accent: 0 8px 24px rgba(59,130,246,0.2);
      --r-sm: 8px;
      --r-md: 12px;
      --r-lg: 16px;
      --r-xl: 20px;
      --r-full: 9999px;
      --ease: cubic-bezier(0.4, 0, 0.2, 1);
      --bee-yellow: #FFD600;
      --bee-yellow-glow: rgba(255,214,0,0.4);
    }
    body.dark {
      --bg: #0B1120;
      --bg-card: rgba(30,41,59,0.7);
      --bg-card-hover: rgba(51,65,85,0.6);
      --text-primary: #F1F5F9;
      --text-secondary: #94A3B8;
      --text-muted: #64748B;
      --border: rgba(148,163,184,0.1);
      --border-hover: rgba(148,163,184,0.2);
      --shadow-sm: 0 1px 2px rgba(0,0,0,0.2);
      --shadow-md: 0 4px 12px rgba(0,0,0,0.3);
      --shadow-lg: 0 12px 40px rgba(0,0,0,0.4);
      --shadow-accent: 0 8px 24px rgba(59,130,246,0.15);
      --grad-card: linear-gradient(160deg, rgba(59,130,246,0.08) 0%, rgba(139,92,246,0.05) 100%);
    }
    * { margin:0; padding:0; box-sizing:border-box; }
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      background: var(--bg);
      min-height: 100vh;
      color: var(--text-primary);
      transition: background 0.4s var(--ease), color 0.4s var(--ease);
      font-weight: 500;
      -webkit-font-smoothing: antialiased;
      overflow-x: hidden;
    }
    .bg-grid {
      position: fixed; inset: 0; z-index: 0; pointer-events: none;
      background-size: 60px 60px;
      background-image:
        linear-gradient(to right, var(--border) 1px, transparent 1px),
        linear-gradient(to bottom, var(--border) 1px, transparent 1px);
      opacity: 0.3;
    }
    body.dark .bg-grid { opacity: 0.06; }
    .bg-glow {
      position: fixed; border-radius: 50%; pointer-events: none; z-index: 0;
      filter: blur(120px); opacity: 0.12;
    }
    .bg-glow-1 { width: 700px; height: 700px; top: -15%; right: -10%; background: radial-gradient(circle, #3B82F6, transparent 70%); }
    .bg-glow-2 { width: 500px; height: 500px; bottom: -10%; left: -5%; background: radial-gradient(circle, #8B5CF6, transparent 70%); }
    body.dark .bg-glow { opacity: 0.08; }
    .wrapper { max-width: 1280px; margin: 0 auto; padding: 2rem 2rem 4rem; position: relative; z-index: 1; }
    header { position: relative; text-align: center; margin-bottom: 1rem; z-index: 2; }
    .bee-logo-wrap { display: flex; justify-content: center; margin-bottom: 1.25rem; }
    .bee-logo-wrap svg { height: 32px; width: auto; filter: drop-shadow(0 2px 8px rgba(255,214,0,0.2)); transition: filter 0.3s ease; }
    .bee-logo-wrap svg:hover { filter: drop-shadow(0 4px 16px rgba(255,214,0,0.5)); }
    .theme-toggle {
      position: absolute; top: 0; right: 0;
      background: var(--bg-card); border: 1px solid var(--border);
      color: var(--text-muted); border-radius: var(--r-full);
      padding: 8px 18px; font-size: 0.78rem; font-weight: 600;
      cursor: pointer; transition: all 0.2s var(--ease);
      display: flex; align-items: center; gap: 7px; font-family: inherit;
    }
    .theme-toggle:hover { border-color: var(--accent); color: var(--accent); }
    h1 {
      font-size: 3.2rem; font-weight: 900;
      letter-spacing: -0.04em; margin-bottom: 0;
      line-height: 1; animation: slideUp 0.6s var(--ease);
      background: var(--grad-hero);
      -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
    }
    @keyframes slideUp { from { opacity:0; transform:translateY(24px); } to { opacity:1; transform:translateY(0); } }
    @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
    @keyframes popIn { from { opacity:0; transform:scale(0.95) translateY(8px); } to { opacity:1; transform:scale(1) translateY(0); } }
    .section-title {
      text-align: center; font-size: 1.2rem; font-weight: 700;
      margin: 2.5rem 0 1.5rem; color: var(--text-secondary);
      display: flex; align-items: center; justify-content: center; gap: 10px;
      width: 100%;
    }
    .section-title::before, .section-title::after {
      content: ''; flex: 0 0 32px; height: 2px; border-radius: 1px;
      background: var(--grad-hero); opacity: 0.3;
    }
    .section-title i { color: var(--accent); font-size: 0.9rem; }
    .directions-grid {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 1rem; margin: 0;
    }
    .dir-card {
      background: var(--bg-card); border: 1px solid var(--border);
      border-radius: var(--r-xl); padding: 2.25rem 1.75rem 1.75rem;
      text-align: center; cursor: pointer;
      transition: all 0.25s var(--ease);
      position: relative; overflow: hidden;
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      min-height: 200px; box-shadow: var(--shadow-sm);
    }
    .dir-card::after {
      content: ''; position: absolute; inset: 0;
      background: var(--grad-card); opacity: 0;
      transition: opacity 0.3s ease;
    }
    .dir-card:hover { border-color: var(--accent); transform: translateY(-3px); box-shadow: var(--shadow-lg); }
    .dir-card:hover::after { opacity: 1; }
    .dir-card:active { transform: translateY(-1px); }
    .card-emoji {
      font-size: 1.4rem; margin-bottom: 0.75rem; display: inline-flex;
      align-items: center; justify-content: center;
      width: 52px; height: 52px; border-radius: 14px;
      transition: all 0.25s var(--ease);
    }
    .ce-blue { background: var(--accent-soft); color: var(--accent); }
    .ce-teal { background: var(--teal-soft); color: var(--teal); }
    .ce-rose { background: var(--rose-soft); color: var(--rose); }
    .ce-amber { background: var(--amber-soft); color: var(--amber); }
    .ce-violet { background: var(--violet-soft); color: var(--violet); }
    .dir-card:hover .card-emoji { transform: scale(1.1) rotate(-3deg); }
    .dir-name { font-size: 1.1rem; font-weight: 700; position: relative; z-index: 1; color: var(--text-primary); letter-spacing: -0.01em; }
    .dir-sub { font-size: 0.75rem; color: var(--text-muted); margin-top: 4px; position: relative; z-index: 1; }
    .drag-handle, .link-drag-handle {
      position: absolute; bottom: 8px; right: 8px;
      background: none; border: none;
      cursor: grab; z-index: 10; font-size: 0.6rem;
      color: var(--text-muted); opacity: 0.15;
      border-radius: 4px; padding: 3px;
      transition: all 0.2s ease;
    }
    .drag-handle:hover, .link-drag-handle:hover { opacity: 0.7; color: var(--accent); transform: scale(1.15); }
    .adm-icon {
      position: absolute; top: 10px; background: transparent !important; border: none !important;
      cursor: pointer; z-index: 10; font-size: 0.8rem; padding: 4px;
      transition: all 0.2s ease; opacity: 0.25;
    }
    .adm-icon:hover { opacity: 1; transform: scale(1.1); }
    .adm-icon.edit { right: 38px; color: var(--accent); }
    .adm-icon.delete { right: 10px; color: var(--rose); }
    .links-grid {
      display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
      gap: 0.6rem; margin-top: 1.5rem;
    }
    .link-item {
      background: var(--bg-card); border: 1px solid var(--border);
      padding: 1rem 1.15rem; border-radius: var(--r-lg);
      font-size: 0.85rem; font-weight: 500;
      text-align: left; cursor: pointer;
      transition: all 0.2s var(--ease);
      display: flex; align-items: flex-start; gap: 10px;
      position: relative; font-family: inherit;
      color: var(--text-primary); box-shadow: var(--shadow-sm);
    }
    .link-item:hover {
      border-color: var(--accent);
      background: var(--accent-soft); color: var(--accent);
      transform: translateY(-1px);
      box-shadow: var(--shadow-md);
    }
    .link-item i.link-icon {
      font-size: 0.8rem; color: var(--accent); transition: all 0.2s var(--ease);
      flex-shrink: 0; opacity: 0.6; margin-top: 2px;
    }
    .link-item:hover i.link-icon { opacity: 1; transform: translateX(2px); }
    .link-item .link-text { flex: 1; }
    .link-adm-icon {
      position: absolute; top: 8px; background: transparent !important; border: none !important;
      cursor: pointer; z-index: 10; font-size: 0.7rem; padding: 3px 5px;
      transition: all 0.2s ease; opacity: 0.25;
    }
    .link-adm-icon:hover { opacity: 1; transform: scale(1.1); }
    .link-adm-icon.edit { right: 62px; color: var(--accent); }
    .link-adm-icon.delete { right: 32px; color: var(--rose); }
    .dragging { opacity: 0.35; transform: scale(0.97); }
    #back-btn {
      position: fixed; top: 1.5rem; left: 1.5rem; z-index: 100;
      background: var(--bg-card); border: 1px solid var(--border);
      padding: 8px 18px; border-radius: var(--r-full); font-weight: 600;
      cursor: pointer; transition: all 0.2s var(--ease);
      display: inline-flex; align-items: center; gap: 7px;
      color: var(--text-secondary); font-family: inherit; font-size: 0.78rem;
      box-shadow: var(--shadow-md);
    }
    #back-btn:hover { border-color: var(--accent); color: var(--accent); transform: translateX(-2px); }
    .tooltip-card {
      position: absolute;
      bottom: calc(100% + 10px);
      left: 50%;
      transform: translateX(-50%) translateY(6px);
      min-width: 300px; max-width: 400px;
      padding: 12px 16px;
      background: var(--bg-card); border: 1px solid var(--accent);
      border-radius: var(--r-md); font-size: 0.8rem; font-weight: 400;
      line-height: 1.5; color: var(--text-secondary);
      pointer-events: none; opacity: 0; visibility: hidden;
      transition: opacity 0.2s ease, transform 0.2s ease, visibility 0.2s ease;
      z-index: 999; box-shadow: var(--shadow-accent);
    }
    .tooltip-card::after {
      content: ''; position: absolute; top: 100%; left: 50%;
      transform: translateX(-50%); border: 6px solid transparent;
      border-top-color: var(--accent);
    }
    .tooltip-card .tt-badge {
      display: inline-flex; align-items: center; justify-content: center;
      width: 20px; height: 20px; border-radius: 5px;
      background: var(--accent-soft); color: var(--accent);
      font-size: 0.6rem; margin-right: 5px; flex-shrink: 0; vertical-align: middle;
    }
    .tooltip-card .tt-title {
      display: flex; align-items: center; margin-bottom: 6px;
      font-weight: 700; font-size: 0.75rem; color: var(--accent);
      text-transform: uppercase; letter-spacing: 0.04em;
    }
    .tooltip-card strong { color: var(--accent); font-weight: 700; }
    .link-item:hover .tooltip-card {
      opacity: 1; visibility: visible;
      transform: translateX(-50%) translateY(0);
    }
    .feedback-card {
      margin: 4rem auto 0; padding: 1.5rem;
      background: var(--bg-card); border: 1px solid var(--border);
      border-radius: var(--r-xl); text-align: center;
      max-width: 440px; position: relative;
      box-shadow: var(--shadow-sm);
      animation: popIn 0.35s var(--ease);
    }
    .feedback-card h3 { font-size: 0.92rem; margin-bottom: 0.6rem; font-weight: 700; color: var(--text-primary); }
    .feedback-card p { font-size: 0.76rem; color: var(--text-muted); }
    .stars-row { display: flex; gap: 10px; justify-content: center; margin: 0.75rem 0; direction: ltr; }
    .stars-row i { font-size: 1.4rem; cursor: pointer; transition: all 0.15s ease; color: var(--border); }
    .stars-row i:hover, .stars-row i.active {
      color: var(--bee-yellow) !important; transform: scale(1.18);
      filter: drop-shadow(0 0 6px var(--bee-yellow-glow));
    }
    textarea {
      width: 100%; max-width: 420px; padding: 0.7rem 0.85rem;
      border-radius: var(--r-md); border: 1px solid var(--border);
      background: var(--bg); font-family: inherit;
      resize: vertical; margin: 0.5rem 0; outline: none;
      transition: all 0.2s var(--ease); font-size: 0.84rem; color: inherit;
    }
    textarea:focus { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
    .btn-send {
      background: var(--grad-btn); border: none; padding: 8px 22px;
      border-radius: var(--r-full); color: #fff; font-weight: 600;
      cursor: pointer; transition: all 0.2s var(--ease);
      font-size: 0.78rem; margin-top: 4px; font-family: inherit;
    }
    .btn-send:hover { transform: translateY(-1px); box-shadow: var(--shadow-accent); opacity: 0.92; }
    .fab {
      position: fixed; bottom: 22px; right: 22px;
      width: 50px; height: 50px; border-radius: 50%;
      background: var(--grad-btn); border: none; color: #fff;
      font-size: 1.1rem; cursor: pointer;
      box-shadow: var(--shadow-accent);
      transition: all 0.2s var(--ease); z-index: 1000;
      display: flex; align-items: center; justify-content: center;
    }
    .fab:hover { transform: scale(1.06); box-shadow: 0 10px 30px rgba(59,130,246,0.3); }
    footer {
      text-align: center; margin-top: 4rem; padding-top: 1.5rem;
      font-size: 0.72rem; color: var(--text-muted); border-top: 1px solid var(--border);
    }
    .analytics-wrap { max-width: 1280px; margin: 0 auto; padding: 16px; }
    .pill-tabs {
      display: inline-flex; margin-bottom: 20px; gap: 4px;
      background: var(--bg-card); padding: 4px; border-radius: var(--r-full);
      border: 1px solid var(--border);
    }
    .pill-tab {
      padding: 7px 16px; font-size: 0.8rem; font-weight: 600;
      background: transparent; border: none;
      border-radius: var(--r-full); cursor: pointer;
      transition: all 0.2s var(--ease);
      color: var(--text-muted); font-family: inherit;
    }
    .pill-tab.active { background: var(--accent); color: #fff; }
    .pill-tab:hover:not(.active) { color: var(--text-primary); }
    .tab-pane { display: none; }
    .tab-pane.active { display: block; animation: fadeIn 0.2s ease; }
    .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 0.65rem; margin-bottom: 1.75rem; }
    .metric-box {
      background: var(--bg-card); border-radius: var(--r-lg); padding: 1rem;
      text-align: center; border: 1px solid var(--border);
      transition: all 0.2s var(--ease); box-shadow: var(--shadow-sm);
    }
    .metric-box:hover { border-color: var(--accent); }
    .metric-val { font-size: 1.8rem; font-weight: 900; color: var(--accent); letter-spacing: -0.02em; }
    .metric-lbl { font-size: 0.7rem; color: var(--text-muted); font-weight: 500; margin-top: 2px; }
    .table-wrap {
      overflow-x: auto; margin-top: 14px; border-radius: var(--r-md);
      border: 1px solid var(--border); width: 100%;
    }
    table { width: 100%; border-collapse: collapse; background: var(--bg-card); overflow: hidden; table-layout: auto; }
    th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid var(--border); white-space: normal; word-wrap: break-word; font-size: 0.82rem; }
    th { background: var(--accent-soft); color: var(--accent); font-weight: 600; text-transform: uppercase; font-size: 0.65rem; letter-spacing: 0.06em; position: sticky; top: 0; z-index: 10; }
    .dl-btn {
      background: var(--grad-btn); border: none; padding: 8px 18px;
      border-radius: var(--r-full); cursor: pointer; font-weight: 600;
      color: #fff; font-family: inherit; font-size: 0.8rem;
      transition: all 0.2s var(--ease);
    }
    .dl-btn:hover { opacity: 0.9; transform: translateY(-1px); box-shadow: var(--shadow-accent); }
    .filter-row { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
    .filter-row select, .filter-row input {
      padding: 7px 12px; border-radius: var(--r-full);
      border: 1px solid var(--border); background: var(--bg-card);
      font-family: inherit; color: inherit; cursor: pointer;
      font-weight: 500; outline: none; transition: all 0.2s var(--ease); font-size: 0.82rem;
    }
    .filter-row select:focus, .filter-row input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
    @media (max-width: 768px) {
      .wrapper { padding: 1rem; }
      h1 { font-size: 2.2rem; }
      .directions-grid { gap: 0.75rem; }
      #back-btn { top: 1rem; left: 1rem; padding: 6px 12px; font-size: 0.72rem; }
      header { margin-top: 2rem; }
      .pill-tabs { gap: 2px; padding: 3px; }
      .pill-tab { padding: 5px 10px; font-size: 0.72rem; }
      .feedback-card { padding: 1.15rem; margin-top: 3rem; }
      .metric-grid { grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); }
      .bee-logo-wrap svg { height: 26px; }
      .tooltip-card { min-width: 240px; max-width: 300px; left: 0; transform: translateX(0) translateY(6px); }
      .link-item:hover .tooltip-card { transform: translateX(0) translateY(0); }
      .tooltip-card::after { left: 20px; transform: none; }
    }
  </style>
</head>
<body>
  <div class="bg-grid"></div>
  <div class="bg-glow bg-glow-1"></div>
  <div class="bg-glow bg-glow-2"></div>
  <div class="wrapper">
    <header>
      <button id="theme-toggle" class="theme-toggle"><i class="fas fa-moon"></i> <span>Тема</span></button>
      <div class="bee-logo-wrap">
        <svg viewBox="0 0 320 60" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect class="bee-pill" x="1" y="1" width="318" height="58" rx="29" fill="#FFD600"/>
          <text class="bee-text" x="160" y="40" text-anchor="middle" font-family="'Inter','Manrope',sans-serif" font-weight="800" font-size="28" fill="#000000" letter-spacing="-0.5">bee&#x2060;line</text>
          <circle cx="235" cy="16" r="3" fill="#000000"/>
        </svg>
      </div>
      <h1>Quality</h1>
    </header>
    <main id="main-content"></main>
    <footer><p>© Quality · Управление качеством</p></footer>
  </div>
  <script>
    const ORDER_KEYS={mainButtons:'main_buttons_order',directionLinks:(dir)=>`direction_${dir}_links_order`};
    const CE=['ce-blue','ce-teal','ce-rose','ce-amber','ce-violet'];
    const defaultMainButtons=[
      {id:'service',text:'Сервис',icon:'fa-headset',href:'?dir=service',sub:'Оценка и стандарты'},
      {id:'sales',text:'Продажи',icon:'fa-arrow-trend-up',href:'?dir=sales',sub:'Метрики и аналитика'},
      {id:'retention',text:'Сохранение',icon:'fa-shield-heart',href:'?dir=retention',sub:'Удержание клиентов'}
    ];
    const directionsDefaults={
      service:{name:"Сервис",icon:"fa-headset",links:[
        {text:"Регламенты оценки качества / Матрицы оценки / Карта фрода",href:"#",type:"link"},
        {text:"Описание скоринга / Новости / Процесс взаимодействия с учениками",href:"#",type:"link"},
        {text:"Стандарты обслуживания",href:"#",type:"link"},
        {text:"Чат бот",href:"#",type:"link"},
        {text:"Отчет по качеству",href:"#",type:"link"},
        {text:"Нормативы / Цели",href:"#",type:"link"},
        {text:"Контроль качества. Жалобы / Благодарности / Ошибки сотрудников",href:"#",type:"link"}
      ]},
      sales:{name:"Продажи",icon:"fa-arrow-trend-up",links:[
        {text:"Регламенты оценки качества / Матрицы оценки / Карта фрода",href:"#",type:"link"},
        {text:"Описание скоринга / Новости / Процесс взаимодействия с учениками",href:"#",type:"link"},
        {text:"Стандарты обслуживания",href:"#",type:"link"},
        {text:"Отчет по качеству",href:"#",type:"link"},
        {text:"Нормативы / Цели",href:"#",type:"link"},
        {text:"Чат бот",href:"#",type:"link"},
        {text:"Контроль качества. Жалобы / Благодарности / Ошибки сотрудников",href:"#",type:"link"}
      ]},
      retention:{name:"Сохранение",icon:"fa-shield-heart",links:[
        {text:"Регламенты оценки качества / Матрицы оценки / Карта фрода",href:"#",type:"link"},
        {text:"Описание скоринга / Новости / Процесс взаимодействия с учениками",href:"#",type:"link"},
        {text:"Стандарты обслуживания",href:"#",type:"link"},
        {text:"Чат бот",href:"#",type:"link"},
        {text:"Отчет по качеству",href:"#",type:"link"},
        {text:"Нормативы / Цели",href:"#",type:"link"},
        {text:"Контроль качества. Жалобы / Благодарности / Ошибки сотрудников",href:"#",type:"link"}
      ]}
    };
    function fixReportsLinks(dirs){Object.keys(dirs).forEach(k=>{if(dirs[k]&&Array.isArray(dirs[k].links)){dirs[k].links.forEach(link=>{if(link.text&&link.text.includes('Отчет по качеству')&&link.type!=='link'){link.type='link';if(!link.href||link.href==='reports'||link.href==='reports-sales'||link.href==='reports-retention'){link.href='#';}}});}});}
    function saveOrder(k,a){localStorage.setItem(k,JSON.stringify(a));}
    function loadOrder(k,da){const s=localStorage.getItem(k);if(s){try{const o=JSON.parse(s);if(Array.isArray(o)&&o.length===da.length)return o;}catch(e){}}return da.map((_,i)=>i);}
    function reorderArrayByIndices(a,oi){const r=[];oi.forEach(i=>{if(i>=0&&i<a.length)r.push(a[i]);});return r;}
    function createAdminIcons(p,ia,od,oe){if(!ia)return;const ei=document.createElement('i');ei.className='fas fa-pencil-alt adm-icon edit';ei.title='Редактировать';ei.onclick=oe;const di=document.createElement('i');di.className='fas fa-trash-alt adm-icon delete';di.title='Удалить';di.onclick=od;p.appendChild(ei);p.appendChild(di);}
    function createDragIcon(el,isL=false){const d=document.createElement('i');d.className=isL?'fas fa-grip-vertical link-drag-handle':'fas fa-grip-vertical drag-handle';d.title='Перетащить';el.appendChild(d);return d;}
    function saveMainButtons(){localStorage.setItem('mainButtons',JSON.stringify(mainButtons));}
    function addQualityTooltip(btn){const tt=document.createElement('div');tt.className='tooltip-card';tt.innerHTML='<div class="tt-title"><span class="tt-badge"><i class="fas fa-lock"></i></span> Требуется доступ</div>Доступ предоставляется по заявке <strong>QLIK Stream 02.027_A. Customer Care Qlik Sense. Доступ к Стримам</strong> роль <strong>Пользователь</strong>';btn.appendChild(tt);}
    function renderMainPage(){const m=document.getElementById('main-content');if(feedbackStatsMode==='true'){renderFeedbackStatsPage();return;}if(dirKey&&directions[dirKey]){renderDirectionPage(dirKey);return;}renderMainButtonsPage();}
    function renderMainButtonsPage(){const m=document.getElementById('main-content');m.innerHTML=`<div class="section-title"><i class="fas fa-compass"></i> Выбери направление</div><div class="directions-grid" id="dir-grid"></div>`;const g=document.getElementById('dir-grid');let cb=[...mainButtons];const so=loadOrder(ORDER_KEYS.mainButtons,cb);const ob=reorderArrayByIndices(cb,so);function rb(buttons){g.innerHTML='';const f=document.createDocumentFragment();buttons.forEach((btn,idx)=>{const c=document.createElement('div');c.className='dir-card';c.setAttribute('data-id',btn.id);const cc=CE[idx%CE.length];c.innerHTML=`<div class="card-emoji ${cc}"><i class="fas ${btn.icon||'fa-star'}"></i></div><div class="dir-name">${btn.text}</div>${btn.sub?'<div class="dir-sub">'+btn.sub+'</div>':''}`;createDragIcon(c);c.onclick=(e)=>{if(e.target.closest('.drag-handle')||e.target.closest('.adm-icon'))return;localStorage.setItem('currentDirection',btn.id);window.location.href=`?dir=${btn.id}`;};if(isAdmin){createAdminIcons(c,isAdmin,()=>{if(confirm('Удалить?')){mainButtons.splice(mainButtons.findIndex(b=>b.id===btn.id),1);saveMainButtons();rb(reorderArrayByIndices(mainButtons,loadOrder(ORDER_KEYS.mainButtons,mainButtons)));}},()=>{const nt=prompt('Название:',btn.text);if(nt)btn.text=nt;const nh=prompt('Ссылка:',btn.href);if(nh)btn.href=nh;saveMainButtons();rb(reorderArrayByIndices(mainButtons,loadOrder(ORDER_KEYS.mainButtons,mainButtons)));});}c.draggable=true;c.addEventListener('dragstart',(e)=>{e.dataTransfer.setData('text/plain',btn.id);c.classList.add('dragging');});c.addEventListener('dragend',()=>c.classList.remove('dragging'));c.addEventListener('dragover',(e)=>e.preventDefault());c.addEventListener('drop',(e)=>{e.preventDefault();const fi=mainButtons.findIndex(b=>b.id===e.dataTransfer.getData('text/plain'));const ti=mainButtons.findIndex(b=>b.id===btn.id);if(fi===ti)return;const mv=mainButtons.splice(fi,1)[0];mainButtons.splice(ti,0,mv);saveMainButtons();saveOrder(ORDER_KEYS.mainButtons,mainButtons.map((_,i)=>i));rb(mainButtons);});f.appendChild(c);});g.appendChild(f);if(isAdmin){const ac=document.createElement('div');ac.className='dir-card';ac.innerHTML='<div class="card-emoji ce-blue"><i class="fas fa-plus"></i></div><div class="dir-name">Добавить</div>';ac.onclick=()=>{const t=prompt('Название:');if(t){const id=t.toLowerCase().replace(/\s/g,'_');const h=prompt('Ссылка:',`?dir=${id}`);mainButtons.push({id,text:t,icon:'fa-folder',href:h||`?dir=${id}`,sub:''});saveMainButtons();if(!directions[id])directions[id]={name:t,icon:'fa-folder',links:[]};localStorage.setItem('directionsData',JSON.stringify(directions));rb(reorderArrayByIndices(mainButtons,loadOrder(ORDER_KEYS.mainButtons,mainButtons)));}};g.appendChild(ac);}}rb(ob);addAdminFAB();}
    function renderDirectionPage(dk){const dir=directions[dk];const m=document.getElementById('main-content');m.innerHTML=`<button id="back-btn"><i class="fas fa-arrow-left"></i> Назад</button><div class="section-title"><i class="fas ${dir.icon||'fa-folder'}"></i> ${dir.name}</div><div id="links-container" class="links-grid"></div>`;const ct=document.getElementById('links-container');const so=loadOrder(ORDER_KEYS.directionLinks(dk),dir.links);const ol=reorderArrayByIndices(dir.links,so);function rl(links){ct.innerHTML='';const f=document.createDocumentFragment();links.forEach((link,idx)=>{const b=document.createElement('button');b.className='link-item';b.setAttribute('data-idx',idx);b.innerHTML=`<i class="fas fa-arrow-up-right-from-square link-icon"></i><span class="link-text">${link.text}</span>`;createDragIcon(b,true);if(link.text&&link.text.includes('Отчет по качеству')){addQualityTooltip(b);}b.onclick=(e)=>{if(e.target.closest('.link-drag-handle')||e.target.closest('.link-adm-icon')||e.target.closest('.tooltip-card'))return;if(link.href&&link.href!=='#'){const cl=JSON.parse(localStorage.getItem('quality_clicks')||'[]');cl.push({timestamp:new Date().toISOString(),direction:dir.name,linkText:link.text});localStorage.setItem('quality_clicks',JSON.stringify(cl));window.open(link.href,'_blank');}else{alert('Ссылка временно недоступна');}};if(isAdmin){createAdminIcons(b,isAdmin,()=>{if(confirm('Удалить?')){dir.links.splice(dir.links.findIndex(l=>l.text===link.text),1);localStorage.setItem('directionsData',JSON.stringify(directions));rl(reorderArrayByIndices(dir.links,loadOrder(ORDER_KEYS.directionLinks(dk),dir.links)));}},()=>{const nt=prompt('Текст:',link.text);if(nt)link.text=nt;const nh=prompt('URL:',link.href);if(nh)link.href=nh;localStorage.setItem('directionsData',JSON.stringify(directions));rl(reorderArrayByIndices(dir.links,loadOrder(ORDER_KEYS.directionLinks(dk),dir.links)));});}b.draggable=true;b.addEventListener('dragstart',(e)=>{e.dataTransfer.setData('text/plain',idx);b.classList.add('dragging');});b.addEventListener('dragend',()=>b.classList.remove('dragging'));b.addEventListener('dragover',(e)=>e.preventDefault());b.addEventListener('drop',(e)=>{e.preventDefault();const fi=parseInt(e.dataTransfer.getData('text/plain'),10);if(fi===idx)return;const mv=dir.links.splice(fi,1)[0];dir.links.splice(idx,0,mv);localStorage.setItem('directionsData',JSON.stringify(directions));const oi=dir.links.map((_,i)=>i);saveOrder(ORDER_KEYS.directionLinks(dk),oi);rl(reorderArrayByIndices(dir.links,oi));});f.appendChild(b);});ct.appendChild(f);if(isAdmin){const ab=document.createElement('button');ab.className='link-item';ab.innerHTML='<i class="fas fa-plus link-icon"></i><span class="link-text">Добавить</span>';ab.onclick=()=>{const t=prompt('Текст:');if(t){const h=prompt('URL:','#');dir.links.push({text:t,href:h||'#',type:'link'});localStorage.setItem('directionsData',JSON.stringify(directions));rl(reorderArrayByIndices(dir.links,loadOrder(ORDER_KEYS.directionLinks(dk),dir.links)));}};ct.appendChild(ab);}}rl(ol);addFeedbackSection(dir.name);document.getElementById('back-btn').onclick=()=>window.location.href=window.location.pathname;}
    function renderFeedbackStatsPage(){const fb=JSON.parse(localStorage.getItem('quality_feedback')||'[]');const m=document.getElementById('main-content');const rc={1:0,2:0,3:0,4:0,5:0};fb.forEach(f=>{if(f.rating>=1&&f.rating<=5)rc[f.rating]++;});const t=fb.length;function prd(ds){if(!ds)return null;const p=ds.split(',')[0].split('.');if(p.length!==3)return null;return new Date(parseInt(p[2],10),parseInt(p[1],10)-1,parseInt(p[0],10));}const mn=["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];const ys=new Set();const ms=new Set();fb.forEach(f=>{const d=prd(f.timestamp);if(d){ys.add(d.getFullYear());ms.add(d.getMonth());}});const sy=Array.from(ys).sort((a,b)=>b-a);const sm=Array.from(ms).sort((a,b)=>a-b);const md={};fb.forEach(f=>{const d=prd(f.timestamp);if(d){const ym=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;if(!md[ym])md[ym]={month:mn[d.getMonth()]+' '+d.getFullYear(),ratings:[0,0,0,0,0],count:0,avgRating:0};md[ym].ratings[f.rating-1]++;md[ym].count++;md[ym].avgRating+=f.rating;}});const sym=Object.keys(md).sort();const cl=sym.map(ym=>md[ym].month);const cd={labels:cl,datasets:[{label:'Средний балл',data:sym.map(ym=>{const d=md[ym];return d.count>0?parseFloat((d.avgRating/d.count).toFixed(2)):0;}),backgroundColor:'#3B82F6',borderColor:'#3B82F6',borderWidth:2,tension:0.4,yAxisID:'y'},{label:'Количество отзывов',data:sym.map(ym=>md[ym].count),backgroundColor:'rgba(20,184,166,0.2)',borderColor:'#14B8A6',borderWidth:2,type:'bar',yAxisID:'y1'}]};m.innerHTML=`<button id="back-btn"><i class="fas fa-arrow-left"></i> Назад</button><div class="section-title"><i class="fas fa-chart-column"></i> Статистика обратной связи</div><div class="analytics-wrap"><div class="metric-grid"><div class="metric-box"><div class="metric-val">${t}</div><div class="metric-lbl">Всего отзывов</div></div>${[1,2,3,4,5].map(r=>'<div class="metric-box"><div class="metric-val">'+rc[r]+'</div><div class="metric-lbl">'+'★'.repeat(r)+'☆'.repeat(5-r)+' ('+(t?((rc[r]/t)*100).toFixed(1):0)+'%)</div></div>').join('')}</div><div class="pill-tabs"><button class="pill-tab active" data-tab="reviews"><i class="fas fa-comment-dots" style="margin-right:5px;"></i>Отзывы</button><button class="pill-tab" data-tab="chart"><i class="fas fa-chart-pie" style="margin-right:5px;"></i>Распределение</button><button class="pill-tab" data-tab="trend"><i class="fas fa-chart-line" style="margin-right:5px;"></i>Динамика</button></div><div id="reviews-tab" class="tab-pane active"><div class="filter-row"><select id="yf"><option value="">Все годы</option>${sy.map(y=>'<option value="'+y+'">'+y+'</option>').join('')}</select><select id="mf"><option value="">Все месяцы</option>${sm.map(mo=>'<option value="'+mo+'">'+mn[mo]+'</option>').join('')}</select><select id="rf"><option value="">Оценка</option><option value="5">★★★★★</option><option value="4">★★★★☆</option><option value="3">★★★☆☆</option><option value="2">★★☆☆☆</option><option value="1">★☆☆☆☆</option></select></div><div class="table-wrap"><table id="ft"><thead><th>Дата</th><th>Направление</th><th>Оценка</th><th>Отзыв</th></thead><tbody></tbody></table></div><button class="dl-btn" id="efb" style="margin-top:10px;"><i class="fas fa-download"></i> Скачать отзывы</button></div><div id="chart-tab" class="tab-pane"><div style="height:380px;"><canvas id="fc"></canvas></div></div><div id="trend-tab" class="tab-pane"><div style="height:380px;"><canvas id="rtc"></canvas></div></div></div>`;function rft(){const y=document.getElementById('yf').value;const mo=document.getElementById('mf').value;const ra=document.getElementById('rf').value;let fl=[...fb];if(y)fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getFullYear()==y;});if(mo!=="")fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getMonth()==mo;});if(ra)fl=fl.filter(f=>f.rating===parseInt(ra));const tb=document.querySelector('#ft tbody');if(tb){tb.innerHTML='';fl.forEach(f=>{const d=prd(f.timestamp);let fd='—';if(d)fd=String(d.getDate()).padStart(2,'0')+'.'+String(d.getMonth()+1).padStart(2,'0')+'.'+d.getFullYear();const tr=document.createElement('tr');tr.innerHTML='<td>'+fd+'</td><td>'+f.direction+'</td><td>'+'★'.repeat(f.rating)+'☆'.repeat(5-f.rating)+'</td><td>'+(f.comment||'')+'</td>';tb.appendChild(tr);});}}function rrc(){const ctx=document.getElementById('fc').getContext('2d');if(ctx){if(window.fbC)window.fbC.destroy();window.fbC=new Chart(ctx,{type:'bar',data:{labels:['★☆☆☆☆','★★☆☆☆','★★★☆☆','★★★★☆','★★★★★'],datasets:[{label:'Отзывы',data:[rc[1],rc[2],rc[3],rc[4],rc[5]],backgroundColor:'#3B82F6',borderRadius:6}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{color:'#94A3B8'}}},scales:{y:{beginAtZero:true,ticks:{color:'#94A3B8'},grid:{color:'rgba(148,163,184,0.08)'}},x:{ticks:{color:'#94A3B8'},grid:{color:'rgba(148,163,184,0.08)'}}}}});}}function rrtc(){const ctx=document.getElementById('rtc').getContext('2d');if(ctx&&cl.length>0){if(window.fTC)window.fTC.destroy();window.fTC=new Chart(ctx,{type:'line',data:cd,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{color:'#94A3B8'}}},scales:{y:{type:'linear',display:true,position:'left',beginAtZero:true,max:5,ticks:{color:'#94A3B8'},grid:{color:'rgba(148,163,184,0.08)'},title:{display:true,text:'Средний балл',color:'#94A3B8'}},y1:{type:'linear',display:true,position:'right',beginAtZero:true,ticks:{color:'#94A3B8'},grid:{color:'rgba(148,163,184,0.08)'},title:{display:true,text:'Отзывы',color:'#94A3B8'}}}}});}}document.getElementById('yf').addEventListener('change',rft);document.getElementById('mf').addEventListener('change',rft);document.getElementById('rf').addEventListener('change',rft);document.getElementById('efb').addEventListener('click',()=>{const y=document.getElementById('yf').value;const mo=document.getElementById('mf').value;const ra=document.getElementById('rf').value;let fl=[...fb];if(y)fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getFullYear()==y;});if(mo!=="")fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getMonth()==mo;});if(ra)fl=fl.filter(f=>f.rating===parseInt(ra));const ws=XLSX.utils.json_to_sheet(fl.map(f=>({Дата:f.timestamp,Направление:f.direction,Оценка:f.rating+' звезд',Отзыв:f.comment})));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Отзывы');XLSX.writeFile(wb,'отзывы_'+new Date().toISOString().slice(0,10)+'.xlsx');});rft();rrc();rrtc();const tabs=document.querySelectorAll('.pill-tab');tabs.forEach(tab=>{tab.addEventListener('click',()=>{tabs.forEach(x=>x.classList.remove('active'));tab.classList.add('active');document.querySelectorAll('.tab-pane').forEach(c=>c.classList.remove('active'));document.getElementById(tab.dataset.tab+'-tab').classList.add('active');if(tab.dataset.tab==='trend')rrtc();else if(tab.dataset.tab==='chart')rrc();});});document.getElementById('back-btn').onclick=()=>window.location.href=window.location.pathname;}
    function addFeedbackSection(dn){const m=document.getElementById('main-content');const ex=document.querySelector('.feedback-card');if(ex)ex.remove();const d=document.createElement('div');d.className='feedback-card';d.innerHTML=`<h3><i class="fas fa-sparkles" style="color:var(--bee-yellow)"></i> Обратная связь</h3><p style="margin-bottom:4px;">Твоё мнение помогает становиться лучше</p><div class="stars-row" id="stars-row">${[1,2,3,4,5].map(i=>'<i class="far fa-star" data-val="'+i+'"></i>').join('')}</div><textarea id="fb-text" rows="3" placeholder="Комментарий или предложение..."></textarea><button class="btn-send" id="send-fb"><i class="fas fa-paper-plane"></i> Отправить</button>`;m.appendChild(d);let rating=0;const stars=d.querySelectorAll('.stars-row i');const offColor='var(--border)';stars.forEach(s=>{s.addEventListener('click',()=>{rating=parseInt(s.dataset.val);stars.forEach(ss=>{if(parseInt(ss.dataset.val)<=rating)ss.className='fas fa-star active';else ss.className='far fa-star';});});s.addEventListener('mouseenter',()=>{const hv=parseInt(s.dataset.val);stars.forEach(ss=>{if(parseInt(ss.dataset.val)<=hv)ss.style.color='#FFD600';else ss.style.color=offColor;});});s.addEventListener('mouseleave',()=>{stars.forEach(ss=>{if(parseInt(ss.dataset.val)<=rating)ss.style.color='#FFD600';else ss.style.color=offColor;});});});d.querySelector('#send-fb').onclick=()=>{if(rating===0)return alert('Поставьте оценку');const fb={timestamp:new Date().toLocaleString(),direction:dn,rating:rating,comment:d.querySelector('#fb-text').value||'Без комментария'};const list=JSON.parse(localStorage.getItem('quality_feedback')||'[]');list.push(fb);localStorage.setItem('quality_feedback',JSON.stringify(list));alert('Спасибо за отзыв!');rating=0;stars.forEach(s=>{s.className='far fa-star';s.style.color=offColor;});d.querySelector('#fb-text').value='';};}
    function addAdminFAB(){const b=document.createElement('button');b.className='fab';b.innerHTML=isAdmin?'<i class="fas fa-sign-out-alt"></i>':'<i class="fas fa-lock"></i>';b.onclick=()=>{if(isAdmin){localStorage.removeItem('adminActive');location.reload();}else{const p=prompt('Пароль:');if(p==='beeline2025'){localStorage.setItem('adminActive','true');location.reload();}else alert('Неверный пароль');}};document.body.appendChild(b);if(isAdmin){const sb=document.createElement('button');sb.className='fab';sb.style.bottom='84px';sb.style.background='linear-gradient(135deg,#EF4444,#DC2626)';sb.innerHTML='<i class="fas fa-chart-column"></i>';sb.onclick=()=>{window.location.href='?feedback=true';};document.body.appendChild(sb);}}
    const themeToggle=document.getElementById('theme-toggle');if(localStorage.getItem('dark-theme-v2')==='true')document.body.classList.add('dark');themeToggle.addEventListener('click',()=>{document.body.classList.toggle('dark');const isD=document.body.classList.contains('dark');localStorage.setItem('dark-theme-v2',isD);themeToggle.querySelector('i').className=isD?'fas fa-moon':'fas fa-sun';themeToggle.querySelector('span').textContent='Тема';});
    let directions=JSON.parse(JSON.stringify(directionsDefaults));let mainButtons=[...defaultMainButtons];
    const savedData=localStorage.getItem('directionsData');if(savedData){try{const l=JSON.parse(savedData);Object.keys(directions).forEach(k=>{if(l[k]&&Array.isArray(l[k].links))directions[k].links=l[k].links;});}catch(e){}}
    fixReportsLinks(directions);localStorage.setItem('directionsData',JSON.stringify(directions));
    const savedMB=localStorage.getItem('mainButtons');if(savedMB){try{const p=JSON.parse(savedMB);if(Array.isArray(p)&&p.length>0)mainButtons=p;}catch(e){}}
    const isAdmin=localStorage.getItem('adminActive')==='true';const urlParams=new URLSearchParams(window.location.search);const dirKey=urlParams.get('dir');const feedbackStatsMode=urlParams.get('feedback');
    renderMainPage();
  </script>
</body>
</html>
