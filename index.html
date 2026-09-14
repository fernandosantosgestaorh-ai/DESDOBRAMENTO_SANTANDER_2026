<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Desdobramento Santander 2026</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    
    <style>
        :root {
            --chama: #DA5322;
            --azeitona: #484540;
            --azeitona-clara: #6b6760;
            --bg-page: #f4f7f6;
            --bg-card: #ffffff;
            --verde-sucesso: #e2efda;
            --verde-texto: #385723;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: var(--azeitona);
            color: var(--azeitona);
        }
        
        header {
            background-color: var(--chama);
            color: #ffffff;
            padding: 15px 30px;
            display: grid;
            grid-template-columns: auto 1fr auto;
            align-items: center;
            border-bottom: 3px solid #ffffff;
            gap: 20px;
        }

        .header-left {
            display: flex;
            align-items: center;
            gap: 15px;
            grid-column: 1;
        }
        .header-logo {
            height: 42px;
            width: auto;
            max-width: 140px;
            object-fit: contain;
            background-color: #ffffff;
            padding: 4px 8px;
            border-radius: 6px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }
        
        .filters {
            display: flex;
            gap: 10px;
            align-items: center;
            flex-wrap: wrap;
        }
        .filter-group {
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .filter-group label {
            color: #ffffff;
            font-size: 12px;
            font-weight: 800;
            text-transform: uppercase;
        }
        .filters select, .filters input {
            padding: 8px 12px;
            border-radius: 6px;
            border: 2px solid rgba(255,255,255,0.8);
            font-weight: bold;
            color: var(--azeitona);
            background-color: #ffffff;
            cursor: pointer;
            font-size: 13px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filters input[type="number"] {
            width: 50px;
            text-align: center;
            padding: 8px;
        }

        header h1 {
            margin: 0;
            font-size: 22px;
            font-weight: 800;
            letter-spacing: 0.5px;
            text-align: center;
            grid-column: 2;
            white-space: nowrap;
        }

        .header-right {
            grid-column: 3;
        }

        .upload-toolbar {
            background-color: #2b2926;
            padding: 12px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid var(--chama);
            flex-wrap: wrap;
            gap: 15px;
        }
        .upload-toolbar-title {
            color: #ffffff;
            font-size: 13px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .upload-buttons-group {
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
        }
        .btn-upload-card {
            background: #ffffff;
            color: var(--chama);
            font-size: 12px;
            font-weight: 800;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            text-transform: uppercase;
            box-shadow: 0 3px 6px rgba(0,0,0,0.2);
            border: 2px solid #ffffff;
            transition: all 0.2s ease;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        .btn-upload-card:hover {
            background-color: #fdf2ee;
            transform: translateY(-2px);
            box-shadow: 0 5px 10px rgba(0,0,0,0.3);
        }
        .btn-upload-card.vigente {
            background-color: var(--chama);
            color: #ffffff;
            border-color: #ffffff;
        }
        .btn-upload-card.vigente:hover {
            background-color: #c2431b;
        }
        .file-input { display: none; }

        .ticker-container {
            background-color: #383532;
            color: #ffffff;
            padding: 8px 0;
            overflow: hidden;
            white-space: nowrap;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.4);
            border-bottom: 2px solid rgba(255,255,255,0.1);
            font-size: 13px;
            font-weight: 600;
        }
        .ticker-text {
            display: inline-block;
            padding-left: 100%;
            animation: marquee 35s linear infinite;
        }
        .ticker-text span {
            margin-right: 50px;
            color: #ffb099;
        }
        .ticker-text strong {
            color: #ffffff;
        }
        @keyframes marquee {
            0% { transform: translate(0, 0); }
            100% { transform: translate(-100%, 0); }
        }

        .container {
            padding: 25px 30px;
            max-width: 1650px;
            margin: auto;
        }

        .insights {
            background-color: var(--bg-card);
            padding: 15px 20px;
            border-radius: 6px;
            margin-bottom: 20px;
            border-left: 6px solid var(--chama);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }
        .insights h3 { margin-top: 0; color: var(--chama); margin-bottom: 6px; font-size: 16px; }
        .insights p { margin: 0; color: var(--azeitona); font-size: 14px; font-weight: 600;}
        
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        .kpi-card {
            background: var(--bg-card);
            padding: 16px;
            border-radius: 6px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            border-top: 4px solid var(--chama);
            text-align: center;
        }
        .kpi-card h3 {
            margin: 0 0 8px 0;
            font-size: 11px;
            color: var(--azeitona);
            text-transform: uppercase;
            font-weight: 800;
        }
        .kpi-card p {
            margin: 0;
            font-size: 20px;
            font-weight: 900;
            color: var(--chama);
        }

        .charts-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 25px;
        }
        .chart-card {
            background: var(--bg-card);
            padding: 20px;
            border-radius: 6px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }
        .chart-card h4 {
            margin-top: 0;
            color: var(--chama);
            text-align: center;
            font-weight: 800;
        }

        .table-container {
            background: var(--bg-card);
            padding: 20px;
            border-radius: 6px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            text-align: center;
            white-space: nowrap;
        }
        th, td {
            padding: 8px 6px;
            border: 1px solid var(--azeitona);
        }
        
        #mainThead th {
            background-color: var(--chama);
            color: #ffffff;
            font-size: 11px;
            text-transform: uppercase;
        }
        
        .top-summary-row { font-weight: bold; background-color: var(--azeitona); color: #ffffff; }
        .top-summary-row td { color: #ffffff; border-color: #6b6760; }
        .meta-row { font-weight: bold; background-color: #ffffff; }
        .meta-pct { font-weight: bold; background-color: #d9e1f2; color: #000000; }
        .meta-pct-warning { font-weight: bold; background-color: var(--verde-sucesso) !important; color: var(--verde-texto) !important; }
        
        .meta-atingida {
            background-color: var(--verde-sucesso) !important;
            color: var(--verde-texto) !important;
            font-weight: bold;
        }

        .spacer-row td {
            height: 16px;
            border: none !important;
            background-color: var(--bg-card) !important;
        }
        
        .sec-header {
            font-weight: bold;
            font-size: 11px;
            text-transform: uppercase;
        }
        .bg-azeitona { background-color: var(--azeitona) !important; color: #ffffff !important; }
        .bg-azeitona-clara { background-color: var(--azeitona-clara) !important; color: #ffffff !important; }
        
        .text-azeitona { color: var(--azeitona); font-weight: 700; }
        .text-chama { color: var(--chama); font-weight: 800; background-color: rgba(218, 83, 34, 0.08); }

        #toast {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: #2b2926;
            color: #fff;
            padding: 12px 20px;
            border-radius: 6px;
            border-left: 4px solid var(--chama);
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            display: none;
            z-index: 1000;
            font-weight: 600;
            font-size: 13px;
        }
    </style>
</head>
<body>

    <header>
        <div class="header-left">
            <img src="logo.png" alt="Logo" class="header-logo" onerror="this.src='data:image/svg+xml;utf8,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 200 200\'><path d=\'M30 40 C70 40, 160 40, 160 100 C160 140, 110 140, 75 140 L55 180 L55 140 C35 140, 30 120, 30 100 Z\' fill=\'%23DA5322\'/><path d=\'M75 110 C105 110, 175 110, 175 155 C175 185, 125 185, 95 185 L80 210 L80 185 C55 185, 65 160, 75 155 Z\' fill=\'%23484540\'/></svg>'">
            <h1>Dashboard Desdobramento Santander 2026</h1>
        </div>
        
        <div class="header-right">
            <div class="filters">
                <div class="filter-group">
                    <label for="inputDiasUteis">Dias Úteis</label>
                    <input type="number" id="inputDiasUteis" value="21" min="1" max="31" onchange="updateDashboard()" title="Multiplicador para cálculo da Meta Geral">
                </div>
                <select id="monthFilter" onchange="updateDashboard()">
                    <option value="AGO">Mês Anterior: Agosto 2026</option>
                    <option value="SET" selected>Mês Vigente: Setembro 2026</option>
                </select>
                <select id="portfolioFilter" onchange="updateDashboard()">
                    <option value="PF">Carteira: PF</option>
                    <option value="PJ_B1">Carteira: PJ_B1</option>
                    <option value="PJ_B2">Carteira: PJ_B2</option>
                </select>
            </div>
        </div>
    </header>

    <div class="upload-toolbar">
        <div class="upload-toolbar-title">
            📂 Central de Arquivos Excel
        </div>
        <div class="upload-buttons-group">
            <label for="uploadAnterior" class="btn-upload-card">
                📅 Carregar Mês Anterior
            </label>
            <input type="file" id="uploadAnterior" class="file-input" accept=".xlsx, .xls" onchange="processarArquivo(this, 'AGO')" />

            <label for="uploadVigente" class="btn-upload-card vigente">
                🔥 Carregar Mês Vigente
            </label>
            <input type="file" id="uploadVigente" class="file-input" accept=".xlsx, .xls" onchange="processarArquivo(this, 'SET')" />
        </div>
    </div>

    <div class="ticker-container">
        <div class="ticker-text" id="tickerContent">
            <span>⚡ <strong>Painel Executivo:</strong> Carregando dados operacionais...</span>
            <span>🎯 <strong>Métricas & Hit Rates:</strong> Médias calculadas dia a dia em tempo real por carteira.</span>
        </div>
    </div>

    <div class="container">
        <div class="insights">
            <h3>💡 Diagnóstico do Analista (IA)</h3>
            <p id="insightText">Aguardando carregamento dos dados...</p>
        </div>

        <div class="kpi-grid">
            <div class="kpi-card">
                <h3>Contatos Úteis (CPC)</h3>
                <p id="kpi-cpc">0</p>
            </div>
            <div class="kpi-card">
                <h3>Taxa de CPC</h3>
                <p id="kpi-taxacpc">0,0%</p>
            </div>
            <div class="kpi-card">
                <h3>Propostas Emitidas</h3>
                <p id="kpi-prop">0</p>
            </div>
            <div class="kpi-card">
                <h3>Taxa de Conversão</h3>
                <p id="kpi-conversao" style="color: var(--chama);">0,0%</p>
            </div>
            <div class="kpi-card">
                <h3>CASH Recuperado ($)</h3>
                <p id="kpi-cash" style="color: var(--azeitona);">R$ 0,00</p>
            </div>
            <div class="kpi-card">
                <h3>Efetividade</h3>
                <p id="kpi-efet">0,0%</p>
            </div>
        </div>

        <div class="charts-container">
            <div class="chart-card">
                <h4>Funil de Conversão: CPC vs Proposta Diária</h4>
                <canvas id="chartFunil" height="100"></canvas>
            </div>
            <div class="chart-card">
                <h4>Performance Financeira: ($) CASH vs ($) CONTÁBIL</h4>
                <canvas id="chartFinanceiro" height="100"></canvas>
            </div>
        </div>

        <div class="table-container">
            <table id="dataTable">
                <thead id="mainThead">
                    <tr>
                        <th>DIA</th>
                        <th>CAPACITY</th>
                        <th>TEMPO DISP</th>
                        <th>DISCAGEM</th>
                        <th>ATENDIDA</th>
                        <th>CPC</th>
                        <th>CPCA</th>
                        <th>PROPOSTA</th>
                        <th>PAGAMENTO</th>
                        <th>($) CASH</th>
                        <th>($) TKT CASH</th>
                        <th>($) CONTABIL</th>
                        <th>($) TKT CONTABIL</th>
                        <th>(%) CONVERSÃO</th>
                        <th>(%) CPC</th>
                        <th>(%) EFETIVIDADE</th>
                    </tr>
                </thead>
                <tbody id="metaBody">
                </tbody>
                <tbody id="secHeaderBody">
                    <tr class="sec-header">
                        <td class="bg-azeitona">DATA</td>
                        <td class="bg-azeitona-clara">REALIZADO</td>
                        <td class="bg-azeitona-clara">TEMPO REALIZADO</td>
                        <td class="bg-azeitona-clara">DISCAGEM</td>
                        <td class="bg-azeitona-clara">ATENDIDA</td>
                        <td class="bg-azeitona">CPC</td>
                        <td class="bg-azeitona">CPCA</td>
                        <td class="bg-azeitona">PROPOSTA</td>
                        <td class="bg-azeitona">PAGAMENTO</td>
                        <td class="bg-azeitona">($) CASH</td>
                        <td class="bg-azeitona">($) TKT CASH</td>
                        <td class="bg-azeitona">($) CONTABIL</td>
                        <td class="bg-azeitona">($) TKT CONTABIL</td>
                        <td class="bg-azeitona">(%) CONVERSÃO</td>
                        <td class="bg-azeitona">(%) CPC</td>
                        <td class="bg-azeitona">(%) EFETIVIDADE</td>
                    </tr>
                </tbody>
                <tbody id="dataBody">
                </tbody>
            </table>
        </div>
    </div>

    <div id="toast">Notificação do sistema</div>

    <script>
        let db = {
            "AGO": { "PF": null, "PJ_B1": null, "PJ_B2": null },
            "SET": { "PF": null, "PJ_B1": null, "PJ_B2": null }
        };

        let chartFunil, chartFinan;

        function showToast(msg) {
            const toast = document.getElementById('toast');
            toast.innerText = msg;
            toast.style.display = 'block';
            setTimeout(() => { toast.style.display = 'none'; }, 4000);
        }

        function formatCurrency(val) { 
            if(val === "-" || val === undefined || val === null || isNaN(val)) return "-";
            return Number(val).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); 
        }
        function formatPercent(val) { 
            if(val === "-" || val === undefined || val === null || isNaN(val)) return "-";
            return (Number(val) * 100).toFixed(2) + '%'; 
        }
        function formatNumber(val, decimals=2) { 
            if(val === "-" || val === undefined || val === null || isNaN(val)) return "-";
            return Number(val).toLocaleString('pt-BR', {maximumFractionDigits: decimals}); 
        }

        // Converte string HH:MM:SS para total de segundos para cálculo matemático
        function timeToSeconds(timeStr) {
            if(!timeStr || timeStr === '-' || typeof timeStr !== 'string') return 0;
            let parts = timeStr.split(':');
            if(parts.length !== 3) return 0;
            return (parseInt(parts[0])||0)*3600 + (parseInt(parts[1])||0)*60 + (parseInt(parts[2])||0);
        }

        // Converte total de segundos de volta para string HH:MM:SS
        function secondsToTime(totalSecs) {
            if(isNaN(totalSecs) || totalSecs < 0) return '-';
            totalSecs = Math.round(totalSecs);
            let h = Math.floor(totalSecs/3600);
            let m = Math.floor((totalSecs%3600)/60);
            let s = totalSecs%60;
            return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
        }

        async function carregarDadosIniciais() {
            const savedDb = localStorage.getItem('santander_dashboard_db');
            if (savedDb) {
                try {
                    db = JSON.parse(savedDb);
                } catch(e) {
                    console.error("Erro ao carregar dados salvos", e);
                }
            }

            try {
                const [respAgo, respSet] = await Promise.all([
                    fetch('ACOMPANHAMENTO DESDOBRAMENTO AGOSTO 2026.xlsx').catch(() => null),
                    fetch('ACOMPANHAMENTO DESDOBRAMENTO SETEMBRO 2026.xlsx').catch(() => null)
                ]);

                if (respAgo && respAgo.ok) {
                    const bufAgo = await respAgo.arrayBuffer();
                    processarBufferExcel(bufAgo, 'AGO', false);
                }
                if (respSet && respSet.ok) {
                    const bufSet = await respSet.arrayBuffer();
                    processarBufferExcel(bufSet, 'SET', false);
                }
            } catch(err) {
                console.log("Modo manual ativo.");
            }

            updateDashboard();
        }

        function processarBufferExcel(arrayBuffer, targetMonth, showAlert = true) {
            try {
                const data = new Uint8Array(arrayBuffer);
                const workbook = XLSX.read(data, {type: 'array'});
                let foundData = false;

                ['PF', 'PJ_B1', 'PJ_B2'].forEach(sheetName => {
                    if(!workbook.Sheets[sheetName]) return;
                    
                    const ws = workbook.Sheets[sheetName];
                    const rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true});
                    
                    let dataRowIdx = -1;
                    let metaGeralIdx = -1;
                    for(let i=0; i<rows.length; i++) {
                        if(rows[i] && typeof rows[i][1] === 'string') {
                            let txt = rows[i][1].trim();
                            if(txt === 'DATA') dataRowIdx = i;
                            if(txt === 'META GERAL') metaGeralIdx = i;
                        }
                    }

                    if(dataRowIdx !== -1) {
                        const excelKeys = ['DATA', 'REALIZADO', 'TEMPO REALIZADO', 'DISCAGEM', 'ATENDIDA', 'CPC', 'CPCA', 'PROPOSTA', 'PAGAMENTO', '($) CASH', '($) TKT CASH', '($) CONTABIL', '($) TKT CONTABIL', '(%) CONVERSÃO', '(%) CPC', '(%) EFETIVIDADE'];
                        
                        let topSummaryIdx = (metaGeralIdx > 2) ? metaGeralIdx - 2 : 5;
                        let top_summary = {}; // Aqui fica gravada a META DIÁRIA bruta (Linha DIA)
                        if(rows[topSummaryIdx]) {
                            excelKeys.forEach((key, c_i) => {
                                let val = rows[topSummaryIdx][c_i + 1];
                                if(val === undefined || val === null || val === "") val = "-";
                                if(typeof val === 'number' && key === 'TEMPO REALIZADO') {
                                    let totalSecs = Math.round(val * 24 * 3600);
                                    val = `${String(Math.floor(totalSecs/3600)).padStart(2,'0')}:${String(Math.floor((totalSecs%3600)/60)).padStart(2,'0')}:${String(totalSecs%60).padStart(2,'0')}`;
                                }
                                top_summary[key] = val;
                            });
                        }

                        function getRowByLabel(lbl) {
                            for(let i=0; i<dataRowIdx; i++) {
                                if(rows[i] && rows[i][1] && rows[i][1].toString().trim() === lbl) {
                                    let res = {};
                                    excelKeys.forEach((key, c_i) => {
                                        let val = rows[i][c_i + 1];
                                        if(val === undefined || val === null || val === "") val = "-";
                                        if(typeof val === 'number' && key === 'TEMPO REALIZADO') {
                                            let totalSecs = Math.round(val * 24 * 3600);
                                            val = `${String(Math.floor(totalSecs/3600)).padStart(2,'0')}:${String(Math.floor((totalSecs%3600)/60)).padStart(2,'0')}:${String(totalSecs%60).padStart(2,'0')}`;
                                        }
                                        res[key] = val;
                                    });
                                    return res;
                                }
                            }
                            return {};
                        }

                        let daily = [];
                        for(let i = dataRowIdx + 1; i < rows.length; i++) {
                            let d_val = rows[i][1];
                            if(d_val !== undefined && d_val !== null && d_val !== "") {
                                let rowRes = {};
                                excelKeys.forEach((key, c_i) => {
                                    let val = rows[i][c_i + 1];
                                    if(val === undefined || val === null || val === "") val = "-";
                                    if(typeof val === 'number' && key === 'TEMPO REALIZADO') {
                                        let totalSecs = Math.round(val * 24 * 3600);
                                        val = `${String(Math.floor(totalSecs/3600)).padStart(2,'0')}:${String(Math.floor((totalSecs%3600)/60)).padStart(2,'0')}:${String(totalSecs%60).padStart(2,'0')}`;
                                    }
                                    if(typeof val === 'number' && key === 'DATA') {
                                        const dateObj = new Date(Math.round((val - 25569)*86400*1000));
                                        dateObj.setMinutes(dateObj.getMinutes() + dateObj.getTimezoneOffset());
                                        val = `${String(dateObj.getDate()).padStart(2, '0')}/${String(dateObj.getMonth() + 1).padStart(2, '0')}/${dateObj.getFullYear()}`;
                                    }
                                    rowRes[key] = val;
                                });
                                daily.push(rowRes);
                            }
                        }

                        db[targetMonth][sheetName] = {
                            "top_summary": top_summary, // META DIARIA
                            "meta_geral": getRowByLabel('META GERAL'), // Valor base do excel
                            "realizado": getRowByLabel('REALIZADO'),
                            "meta_relogio": getRowByLabel('META RELOGIO'),
                            "pct_meta_realizado": getRowByLabel('(%) META REALIZADO'),
                            "pct_meta_relogio": getRowByLabel('(%) META RELOGIO'),
                            "daily": daily
                        };
                        foundData = true;
                    }
                });

                if(foundData) {
                    localStorage.setItem('santander_dashboard_db', JSON.stringify(db));
                    if(showAlert) {
                        showToast(`✅ Planilha do ${targetMonth === 'AGO' ? 'Mês Anterior' : 'Mês Vigente'} carregada com sucesso!`);
                        document.getElementById('monthFilter').value = targetMonth;
                        updateDashboard();
                    }
                }
            } catch(err) {
                console.error(err);
            }
        }

        function processarArquivo(inputElement, targetMonth) {
            const file = inputElement.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(event) {
                processarBufferExcel(event.target.result, targetMonth, true);
                updateDashboard();
            };
            reader.readAsArrayBuffer(file);
        }

        function updateDashboard() {
            const month = document.getElementById('monthFilter').value;
            const port = document.getElementById('portfolioFilter').value;
            const diasUteis = parseInt(document.getElementById('inputDiasUteis').value) || 21;
            
            const portfolioData = (db[month] && db[month][port]) ? db[month][port] : null;
            const data = portfolioData ? portfolioData.daily : [];

            let sDisc = 0, sAtend = 0, sCpc = 0, sProp = 0, sCash = 0, sEfet = 0;
            if(portfolioData && portfolioData.realizado) {
                sDisc = Number(portfolioData.realizado['DISCAGEM']) || 0;
                sAtend = Number(portfolioData.realizado['ATENDIDA']) || 0;
                sCpc = Number(portfolioData.realizado['CPC']) || 0;
                sProp = Number(portfolioData.realizado['PROPOSTA']) || 0;
                sCash = Number(portfolioData.realizado['($) CASH']) || 0;
                sEfet = Number(portfolioData.realizado['(%) EFETIVIDADE']) || 0;
            }

            const taxaCpcCalc = sAtend > 0 ? (sCpc / sAtend) : 0;
            const taxaConvCalc = sCpc > 0 ? (sProp / sCpc) : 0;

            document.getElementById('kpi-cpc').innerText = formatNumber(sCpc);
            document.getElementById('kpi-taxacpc').innerText = formatPercent(taxaCpcCalc);
            document.getElementById('kpi-prop').innerText = formatNumber(sProp);
            document.getElementById('kpi-conversao').innerText = formatPercent(taxaConvCalc);
            document.getElementById('kpi-cash').innerText = formatCurrency(sCash);
            document.getElementById('kpi-efet').innerText = formatPercent(sEfet);

            let sumHitRateAtend = 0;
            let sumHitRateCpc = 0;
            let sumConversion = 0;
            let validDaysAtend = 0;
            let validDaysCpc = 0;
            let validDaysConv = 0;

            data.forEach(d => {
                let disc = Number(d['DISCAGEM']) || 0;
                let atend = Number(d['ATENDIDA']) || 0;
                let cpc = Number(d['CPC']) || 0;
                let prop = Number(d['PROPOSTA']) || 0;

                if (disc > 0) {
                    sumHitRateAtend += (atend / disc);
                    validDaysAtend++;
                    sumHitRateCpc += (cpc / disc);
                    validDaysCpc++;
                }
                if (cpc > 0) {
                    sumConversion += (prop / cpc);
                    validDaysConv++;
                }
            });

            let avgHitRateAtend = validDaysAtend > 0 ? (sumHitRateAtend / validDaysAtend) : 0;
            let avgHitRateCpc = validDaysCpc > 0 ? (sumHitRateCpc / validDaysCpc) : 0;
            let avgConversion = validDaysConv > 0 ? (sumConversion / validDaysConv) : 0;

            let tktCash = portfolioData && portfolioData.top_summary ? portfolioData.top_summary['($) TKT CASH'] : 0;
            let tktContabil = portfolioData && portfolioData.top_summary ? portfolioData.top_summary['($) TKT CONTABIL'] : 0;

            const tickerEl = document.getElementById('tickerContent');
            tickerEl.innerHTML = `
                <span>🔥 <strong>Carteira ${port} (${month === 'AGO' ? 'Agosto' : 'Setembro'}):</strong> Discagens: ${formatNumber(sDisc, 0)} | CPCs: ${formatNumber(sCpc, 0)} | CASH: ${formatCurrency(sCash)}</span>
                <span>🎯 <strong>Hit Rate Atendimento Médio:</strong> ${formatPercent(avgHitRateAtend)} | <strong>Hit Rate CPC Médio:</strong> ${formatPercent(avgHitRateCpc)}</span>
                <span>📈 <strong>Taxa de Conversão Média:</strong> ${formatPercent(avgConversion)}</span>
                <span>💰 <strong>Tickets Médios:</strong> Tkt Cash: ${formatCurrency(tktCash)} | Tkt Contábil: ${formatCurrency(tktContabil)}</span>
                <span>⭐ <strong>Status:</strong> Monitoramento diário ativo.</span>
            `;

            if(!portfolioData || data.length === 0) {
                document.getElementById('insightText').innerHTML = `Aguardando leitura dos dados de <b>${port}</b> para o mês de ${month === 'AGO' ? 'Agosto' : 'Setembro'}. Se necessário, carregue o arquivo Excel na barra superior.`;
            } else {
                let msg = `Exibindo carteira <b>${port}</b> (${month === 'AGO' ? 'Agosto' : 'Setembro'}). `;
                if(sEfet < 0.3) {
                    msg += `⚠️ Alerta Gerencial: Efetividade abaixo de 30%. Recomendável impulsionar acordos firmados.`;
                } else {
                    msg += `⭐ Operação estável com boa conversão no funil de negociação.`;
                }
                document.getElementById('insightText').innerHTML = msg;
            }

            const labels = data.map(d => typeof d['DATA'] === 'string' ? d['DATA'].substring(0,5) : '');
            const ctxFunil = document.getElementById('chartFunil').getContext('2d');
            if(chartFunil) chartFunil.destroy();
            chartFunil = new Chart(ctxFunil, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        { label: 'Contatos Úteis (CPC)', data: data.map(d => Number(d['CPC'])||0), borderColor: '#484540', backgroundColor: 'rgba(72,69,64,0.1)', fill: true, tension: 0.3 },
                        { label: 'Propostas', data: data.map(d => Number(d['PROPOSTA'])||0), borderColor: '#DA5322', borderWidth: 3, tension: 0.3 }
                    ]
                },
                options: { responsive: true, interaction: { mode: 'index', intersect: false } }
            });

            const ctxFinan = document.getElementById('chartFinanceiro').getContext('2d');
            if(chartFinan) chartFinan.destroy();
            chartFinan = new Chart(ctxFinan, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [
                        { label: 'Recuperado - CASH', data: data.map(d => Number(d['($) CASH'])||0), backgroundColor: '#DA5322', yAxisID: 'y' },
                        { label: 'Dívida - CONTÁBIL', data: data.map(d => Number(d['($) CONTABIL'])||0), type: 'line', borderColor: '#484540', borderDash: [5, 5], tension: 0.1, yAxisID: 'y1' }
                    ]
                },
                options: {
                    responsive: true, interaction: { mode: 'index', intersect: false },
                    scales: {
                        y: { type: 'linear', display: true, position: 'left' },
                        y1: { type: 'linear', display: true, position: 'right', grid: {drawOnChartArea: false} }
                    }
                }
            });

            const tbodyMeta = document.getElementById('metaBody');
            
            if(portfolioData && portfolioData.top_summary) {
                const ts = portfolioData.top_summary; // Meta Diária Original
                const rz = portfolioData.realizado; // Realizado Original
                const mr = portfolioData.meta_relogio;
                const pde = portfolioData.pct_meta_relogio;
                
                let mgCalc = {}; // Meta Geral Dinâmica
                let prCalc = {}; // (%) Meta Realizado Dinâmico

                const keysToMultiply = ['DISCAGEM', 'ATENDIDA', 'CPC', 'CPCA', 'PROPOSTA', 'PAGAMENTO', '($) CASH', '($) CONTABIL'];

                for(let k in ts) {
                    if(k === 'TEMPO REALIZADO') {
                        let secs = timeToSeconds(ts[k]);
                        mgCalc[k] = secondsToTime(secs * diasUteis);
                        let rzSecs = rz ? timeToSeconds(rz[k]) : 0;
                        prCalc[k] = (secs * diasUteis) > 0 ? (rzSecs / (secs * diasUteis)) : "-";
                    } else if(keysToMultiply.includes(k)) {
                        let val = (Number(ts[k]) || 0) * diasUteis;
                        mgCalc[k] = val;
                        prCalc[k] = val > 0 ? ((Number(rz?.[k]) || 0) / val) : "-";
                    } else {
                        let val = ts[k];
                        mgCalc[k] = val;
                        if(k === 'REALIZADO' || k === 'CAPACITY') {
                            prCalc[k] = "-"; 
                        } else {
                            let mVal = Number(val) || 0;
                            prCalc[k] = mVal !== 0 ? ((Number(rz?.[k]) || 0) / mVal) : "-";
                        }
                    }
                }

                tbodyMeta.innerHTML = `
                    <tr class="top-summary-row">
                        <td class="text-chama" style="color: #fff; text-align: center;">META DIÁRIA</td>
                        <td>${formatNumber(ts['REALIZADO'])}</td>
                        <td>${ts['TEMPO REALIZADO']}</td>
                        <td>${formatNumber(ts['DISCAGEM'], 0)}</td>
                        <td>${formatNumber(ts['ATENDIDA'])}</td>
                        <td>${formatNumber(ts['CPC'])}</td>
                        <td>${formatNumber(ts['CPCA'])}</td>
                        <td>${formatNumber(ts['PROPOSTA'])}</td>
                        <td>${formatNumber(ts['PAGAMENTO'])}</td>
                        <td>${formatCurrency(ts['($) CASH'])}</td>
                        <td>${formatCurrency(ts['($) TKT CASH'])}</td>
                        <td>${formatCurrency(ts['($) CONTABIL'])}</td>
                        <td>${formatCurrency(ts['($) TKT CONTABIL'])}</td>
                        <td>${formatPercent(ts['(%) CONVERSÃO'])}</td>
                        <td>${formatPercent(ts['(%) CPC'])}</td>
                        <td>${formatPercent(ts['(%) EFETIVIDADE'])}</td>
                    </tr>
                    <tr class="meta-row">
                        <td class="text-azeitona">META GERAL (${diasUteis}d)</td>
                        <td>${formatNumber(mgCalc['REALIZADO'])}</td>
                        <td>${mgCalc['TEMPO REALIZADO']}</td>
                        <td>${formatNumber(mgCalc['DISCAGEM'], 0)}</td>
                        <td>${formatNumber(mgCalc['ATENDIDA'])}</td>
                        <td>${formatNumber(mgCalc['CPC'])}</td>
                        <td>${formatNumber(mgCalc['CPCA'])}</td>
                        <td>${formatNumber(mgCalc['PROPOSTA'])}</td>
                        <td>${formatNumber(mgCalc['PAGAMENTO'])}</td>
                        <td>${formatCurrency(mgCalc['($) CASH'])}</td>
                        <td>${formatCurrency(mgCalc['($) TKT CASH'])}</td>
                        <td>${formatCurrency(mgCalc['($) CONTABIL'])}</td>
                        <td>${formatCurrency(mgCalc['($) TKT CONTABIL'])}</td>
                        <td>${formatPercent(mgCalc['(%) CONVERSÃO'])}</td>
                        <td>${formatPercent(mgCalc['(%) CPC'])}</td>
                        <td>${formatPercent(mgCalc['(%) EFETIVIDADE'])}</td>
                    </tr>
                    <tr class="meta-row">
                        <td class="text-azeitona">REALIZADO</td>
                        <td>${formatNumber(rz['REALIZADO'])}</td>
                        <td>${rz['TEMPO REALIZADO']}</td>
                        <td>${formatNumber(rz['DISCAGEM'], 0)}</td>
                        <td>${formatNumber(rz['ATENDIDA'])}</td>
                        <td>${formatNumber(rz['CPC'])}</td>
                        <td>${formatNumber(rz['CPCA'])}</td>
                        <td>${formatNumber(rz['PROPOSTA'])}</td>
                        <td>${formatNumber(rz['PAGAMENTO'])}</td>
                        <td>${formatCurrency(rz['($) CASH'])}</td>
                        <td>${formatCurrency(rz['($) TKT CASH'])}</td>
                        <td>${formatCurrency(rz['($) CONTABIL'])}</td>
                        <td>${formatCurrency(rz['($) TKT CONTABIL'])}</td>
                        <td>${formatPercent(rz['(%) CONVERSÃO'])}</td>
                        <td>${formatPercent(rz['(%) CPC'])}</td>
                        <td>${formatPercent(rz['(%) EFETIVIDADE'])}</td>
                    </tr>
                    <tr class="meta-row">
                        <td class="text-azeitona">META RELOGIO</td>
                        <td>${formatNumber(mr['REALIZADO'])}</td>
                        <td>${mr['TEMPO REALIZADO']}</td>
                        <td>${formatNumber(mr['DISCAGEM'], 0)}</td>
                        <td>${formatNumber(mr['ATENDIDA'])}</td>
                        <td>${formatNumber(mr['CPC'])}</td>
                        <td>${formatNumber(mr['CPCA'])}</td>
                        <td>${formatNumber(mr['PROPOSTA'])}</td>
                        <td>${formatNumber(mr['PAGAMENTO'])}</td>
                        <td>${formatCurrency(mr['($) CASH'])}</td>
                        <td>${formatCurrency(mr['($) TKT CASH'])}</td>
                        <td>${formatCurrency(mr['($) CONTABIL'])}</td>
                        <td>${formatCurrency(mr['($) TKT CONTABIL'])}</td>
                        <td>${formatPercent(mr['(%) CONVERSÃO'])}</td>
                        <td>${formatPercent(mr['(%) CPC'])}</td>
                        <td>${formatPercent(mr['(%) EFETIVIDADE'])}</td>
                    </tr>
                    
                    <!-- Respiro logo abaixo da Meta Relógio -->
                    <tr class="spacer-row"><td colspan="16"></td></tr>

                    <tr class="meta-pct">
                        <td colspan="3">(%) META REALIZADO</td>
                        <td>${formatPercent(prCalc['DISCAGEM'])}</td>
                        <td>${formatPercent(prCalc['ATENDIDA'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['CPC'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['CPCA'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['PROPOSTA'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['PAGAMENTO'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['($) CASH'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['($) TKT CASH'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['($) CONTABIL'])}</td>
                        <td class="meta-pct-warning">${formatPercent(prCalc['($) TKT CONTABIL'])}</td>
                        <td>${formatPercent(prCalc['(%) CONVERSÃO'])}</td>
                        <td>${formatPercent(prCalc['(%) CPC'])}</td>
                        <td>${formatPercent(prCalc['(%) EFETIVIDADE'])}</td>
                    </tr>
                    
                    <!-- Respiro logo abaixo de (%) Meta Realizado -->
                    <tr class="spacer-row"><td colspan="16"></td></tr>

                    <tr class="meta-pct" style="background-color: #fff2cc;">
                        <td colspan="3">(%) META RELOGIO</td>
                        <td>${formatPercent(pde['DISCAGEM'])}</td>
                        <td>${formatPercent(pde['ATENDIDA'])}</td>
                        <td>${formatPercent(pde['CPC'])}</td>
                        <td>${formatPercent(pde['CPCA'])}</td>
                        <td>${formatPercent(pde['PROPOSTA'])}</td>
                        <td>${formatPercent(pde['PAGAMENTO'])}</td>
                        <td>${formatPercent(pde['($) CASH'])}</td>
                        <td>${formatPercent(pde['($) TKT CASH'])}</td>
                        <td>${formatPercent(pde['($) CONTABIL'])}</td>
                        <td>${formatPercent(pde['($) TKT CONTABIL'])}</td>
                        <td>${formatPercent(pde['(%) CONVERSÃO'])}</td>
                        <td>${formatPercent(pde['(%) CPC'])}</td>
                        <td>${formatPercent(pde['(%) EFETIVIDADE'])}</td>
                    </tr>
                    
                    <!-- Respiro antes do cabeçalho da data -->
                    <tr class="spacer-row"><td colspan="16"></td></tr>
                `;
            } else { tbodyMeta.innerHTML = ''; }

            const tbodyData = document.getElementById('dataBody');
            tbodyData.innerHTML = '';
            
            const target = portfolioData && portfolioData.top_summary ? portfolioData.top_summary : {};

            data.forEach(r => {
                const tr = document.createElement('tr');
                
                function checkMeta(key, val) {
                    if(target[key] !== undefined && target[key] !== "-" && !isNaN(target[key])) {
                        return Number(val) >= Number(target[key]) ? 'meta-atingida' : '';
                    }
                    return '';
                }

                tr.innerHTML = `
                    <td class="text-azeitona">${r['DATA']}</td>
                    <td class="text-azeitona">${formatNumber(r['REALIZADO'])}</td>
                    <td class="text-azeitona">${r['TEMPO REALIZADO'] || '-'}</td>
                    <td class="text-azeitona ${checkMeta('DISCAGEM', r['DISCAGEM'])}">${formatNumber(r['DISCAGEM'], 0)}</td>
                    <td class="text-azeitona ${checkMeta('ATENDIDA', r['ATENDIDA'])}">${formatNumber(r['ATENDIDA'])}</td>
                    <td class="text-chama ${checkMeta('CPC', r['CPC'])}">${formatNumber(r['CPC'])}</td>
                    <td class="text-chama ${checkMeta('CPCA', r['CPCA'])}">${formatNumber(r['CPCA'])}</td>
                    <td class="text-chama ${checkMeta('PROPOSTA', r['PROPOSTA'])}">${formatNumber(r['PROPOSTA'])}</td>
                    <td class="text-azeitona ${checkMeta('PAGAMENTO', r['PAGAMENTO'])}">${formatNumber(r['PAGAMENTO'])}</td>
                    <td class="text-chama ${checkMeta('($) CASH', r['($) CASH'])}">${formatCurrency(r['($) CASH'])}</td>
                    <td class="text-azeitona">${formatCurrency(r['($) TKT CASH'])}</td>
                    <td class="text-azeitona ${checkMeta('($) CONTABIL', r['($) CONTABIL'])}">${formatCurrency(r['($) CONTABIL'])}</td>
                    <td class="text-azeitona">${formatCurrency(r['($) TKT CONTABIL'])}</td>
                    <td class="text-chama ${checkMeta('(%) CONVERSÃO', r['(%) CONVERSÃO'])}">${formatPercent(r['(%) CONVERSÃO'])}</td>
                    <td class="text-chama ${checkMeta('(%) CPC', r['(%) CPC'])}">${formatPercent(r['(%) CPC'])}</td>
                    <td class="text-chama ${checkMeta('(%) EFETIVIDADE', r['(%) EFETIVIDADE'])}">${formatPercent(r['(%) EFETIVIDADE'])}</td>
                `;
                tbodyData.appendChild(tr);
            });
        }

        window.onload = async function() {
            carregarDadosIniciais();
        };
    </script>
</body>
</html>
