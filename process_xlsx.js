
const $ = (id) => document.getElementById(id);
const statusEl = $("status");
// const previewBody = $("preview").querySelector("tbody");

let DATA = [];
const loading = document.getElementById('loading');
const container = document.getElementById('container');
const newFile = document.getElementById('adicionar-arquivo');
const EPS = 0.005;        // valores em moeda próximos de zero serão tratados como 0
const BASE_MIN_PCT = 1;   // valor mínimo (R$) para calcular variação em %

function setStatus(msg) { statusEl.textContent = msg; }

function parseNumberCell(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    // tenta converter "1.234,56" -> 1234.56
    const s = String(v).trim().replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

function detectHeaderRow(rows) {
    // Retorna índice da linha de cabeçalho, tentando achar "Conta Contábil" e colunas "mm yyyy"
    const rxMesAno = /^\s*\d{2}\s+\d{4}\s*$/;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const r = rows[i].map(c => String(c || '').trim());
        const hasConta = r.some(c => c.toLowerCase().includes("conta contábil") || c.toLowerCase().includes("conta contabil"));
        const monthCols = r.filter(c => rxMesAno.test(c));
        if (hasConta && monthCols.length >= 1) return i;
    }
    // fallback comum: linha 1 (zero-based) como nos exemplos
    return 1;
}

function processSheet(name, sheet) {
    // Converte para matriz de linhas
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    if (!rows.length) return [];
    const hdrIdx = detectHeaderRow(rows);
    const header = rows[hdrIdx].map(c => String(c || '').trim());
    const dataRows = rows.slice(hdrIdx + 1);

    // Detecta colunas
    const colConta = header.findIndex(c => c.toLowerCase().includes("conta contábil") || c.toLowerCase().includes("conta contabil"));
    const colDesc = header.findIndex(c => c.toLowerCase().includes("descrição") || c.toLowerCase().includes("descricao"));
    const rxMesAno = /^\s*\d{2}\s+\d{4}\s*$/;
    const monthIdxs = header.map((c, i) => rxMesAno.test(c) ? i : -1).filter(i => i >= 0);

    if (colConta < 0 || monthIdxs.length === 0) {
        console.warn("Cabeçalho não identificado na aba", name, header);
        return [];
    }

    // Monta registros
    const out = [];
    for (const r of dataRows) {
        const conta = r[colConta];
        if (conta == null || String(conta).trim() === "") continue;
        for (const i of monthIdxs) {
            const label = String(header[i]).trim(); // "mm yyyy"
            const mm = parseInt(label.split(/\s+/)[0], 10);
            const yyyy = parseInt(label.split(/\s+/)[1], 10);
            const valor = parseNumberCell(r[i]);
            out.push({
                centro_custo: name,
                conta_contabil: String(conta).trim(),
                ano: yyyy,
                mes: mm,
                valor: valor
            });
        }
    }
    return out;
}

// function renderPreview(data) {
//     // limpa
//     previewBody.innerHTML = "";
//     data.slice(0, 200).forEach(rec => {
//         const tr = document.createElement("tr");
//         tr.innerHTML = `<td>${rec.centro_custo}</td><td>${rec.conta_contabil}</td><td>${rec.ano}</td><td>${rec.mes}</td><td>${rec.valor}</td>`;
//         previewBody.appendChild(tr);
//     });
// }

$("btnParse").addEventListener("click", async () => {
    const file = $("file").files[0];

    if (!file) {
        setStatus("Selecione um arquivo primeiro.");
        alert("Selecione um arquivo primeiro.")
        return;
    }

    loading.style.display = 'flex';
    container.style.display = 'none';
    setStatus("Lendo arquivo…");

    const buf = await file.arrayBuffer();
    let wb;

    try {
        wb = XLSX.read(buf, { type: "array", cellDates: false, cellNF: false, cellText: false });
    } catch (e) {
        console.error(e);
        setStatus("Não foi possível ler o arquivo. Verifique o formato.");
        container.style.display = 'block';
        loading.style.display = 'none';
        return;
    }

    const ccSheets = wb.SheetNames.filter(n => n.toUpperCase().startsWith("CC_"));
    if (ccSheets.length === 0) {
        setStatus("Nenhuma aba CC_ encontrada.");
        container.style.display = 'block';
        loading.style.display = 'none';
        return;
    }

    setStatus("Processando abas: " + ccSheets.join(", "));
    let all = [];
    for (const sname of ccSheets) {
        const sheet = wb.Sheets[sname];
        const rows = processSheet(sname, sheet);
        all = all.concat(rows);
    }
    if (all.length === 0) {
        setStatus("Nenhum dado encontrado nas abas CC_.");
        container.style.display = 'block';
        loading.style.display = 'none';
        return;
    }

    // Ordena por ano, mes (opcional)
    all.sort((a, b) => (a.ano - b.ano) || (a.mes - b.mes) || a.centro_custo.localeCompare(b.centro_custo) || a.conta_contabil.localeCompare(b.conta_contabil));

    DATA = all;
    // Atualiza UI
    // $("ccSheets").textContent = ccSheets.join(", ");
    // $("rowsCount").textContent = all.length.toLocaleString('pt-BR');
    // const anos = [...new Set(all.map(r => r.ano))].sort((a, b) => a - b);
    // $("anos").textContent = anos.join(", ");
    // $("jsonOut").value = JSON.stringify(all, null, 2);
    // $("btnDownload").disabled = false;
    // renderPreview(all);
    setStatus("Pronto! dados lidos com sucesso.");

    setTimeout(() => {
        loading.style.display = 'none';
        renderFilters(DATA);
        render(DATA);
    }, 1500);
});

// $("btnDownload").addEventListener("click", () => {
//     const blob = new Blob([$("jsonOut").value], { type: "application/json;charset=utf-8" });
//     const url = URL.createObjectURL(blob);
//     const a = document.createElement("a");
//     a.href = url; a.download = "data.json"; a.click();
//     setTimeout(() => URL.revokeObjectURL(url), 2000);
// });

// $("btnCopy").addEventListener("click", async () => {
//     try {
//         await navigator.clipboard.writeText($("jsonOut").value);
//         $("btnCopy").textContent = "copiado!";
//         setTimeout(() => $("btnCopy").textContent = "copiar", 1500);
//     } catch (e) {
//         $("btnCopy").textContent = "falhou :(";
//         setTimeout(() => $("btnCopy").textContent = "copiar", 1500);
//     }
// });


// gera o dashboard

// Helpers
const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

function fmtBR(v) {
    v = Number(v || 0);
    if (Math.abs(v) < EPS) v = 0;
    return "R$ " + v.toLocaleString('pt-BR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
};

function fmtP(v) {
    if (!isFinite(v)) return '0.00';
    if (Math.abs(v) < 0.005) v = 0;
    return v.toFixed(2) + '%';
};

function unique(arr) {
    return [...new Set(arr)];
}

function byFilters(rec, cc, conta) {
    return (!cc || cc === '*' || rec.centro_custo === cc) && (!conta || conta === '*' || rec.conta_contabil === conta);
}

// Tooltip helpers
const TIP = document.getElementById('tooltip');
function showTip(text, evt) {
    TIP.textContent = text;
    TIP.style.display = 'block';
    moveTip(evt);
}

function moveTip(evt) {
    TIP.style.left = (evt.clientX + 12) + 'px';
    TIP.style.top = (evt.clientY + 12) + 'px';
}

function hideTip() {
    TIP.style.display = 'none';
}

// Populate selects
function fillSelect(sel, opts, withAll = true) {
    const el = document.getElementById(sel);
    el.innerHTML = "";
    if (withAll) el.append(new Option("Todos", "*"));
    opts.forEach(o => el.append(new Option(o, o)));
}

function renderFilters(data) {
    // Derived lists
    const anos = unique(data.map(d => d.ano)).sort();
    const ccs = unique(data.map(d => d.centro_custo)).sort();
    const contas = unique(data.map(d => d.conta_contabil)).sort();

    fillSelect("selCC", ccs);
    fillSelect("selConta", contas);
    fillSelect("selAnoA", anos, false);
    fillSelect("selAnoB", anos, false);
    document.getElementById("selAnoA").value = anos[0];
    document.getElementById("selAnoB").value = anos[anos.length - 1];
}

// adicionar novo arquivo para análise
newFile.addEventListener("click", () => {
    container.style.display = 'block';
    document.getElementById("container-dashboard").style.display = 'none';
});

// Tabs
document.querySelectorAll(".tab-btn").forEach(btn => btn.addEventListener("click", (e) => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    e.currentTarget.classList.add("active");
    const tab = e.currentTarget.dataset.tab;
    document.getElementById("tab-anual").classList.toggle("hidden", tab !== "anual");
    document.getElementById("tab-mensal").classList.toggle("hidden", tab !== "mensal");
    render(DATA);
}));

// SVG utilities (gridlines; NO fixed labels on bars)
function clearSVG(id) {
    const s = document.getElementById(id);
    while (s.firstChild)
        s.removeChild(s.firstChild);
    return s;
}
function addText(svg, x, y, text, size = 10, anchor = 'middle', fill = '#111827') {
    const t = document.createElementNS("http://www.w3.org/2000/svg", "text");
    t.setAttribute("x", x); t.setAttribute("y", y);
    t.setAttribute("font-size", size);
    t.setAttribute("text-anchor", anchor);
    t.setAttribute("fill", fill);
    t.textContent = text; svg.appendChild(t);
}

function barChart(id, labels, seriesA, seriesB, colors, labelA, labelB) {
    const svg = clearSVG(id); const W = 800, H = 320, padL = 56, pad = 40;
    const max = Math.max(1, ...seriesA, ...seriesB);
    const n = Math.max(labels.length, 1); const groupW = (W - padL - pad) / n; const barW = groupW / 3;
    // gridlines + y-axis labels (5 ticks)
    for (let i = 0; i <= 5; i++) {
        const y = H - pad - (H - 2 * pad) * i / 5;
        const val = max * i / 5;
        const ln = document.createElementNS("http://www.w3.org/2000/svg", "line");
        ln.setAttribute("x1", padL); ln.setAttribute("y1", y);
        ln.setAttribute("x2", W - pad);
        ln.setAttribute("y2", y);
        ln.setAttribute("stroke", "#e5e7eb");

        svg.appendChild(ln);
        addText(svg, padL - 6, y + 3, fmtBR(val), 9, 'end', '#6b7280');
    }
    // x-axis
    const ax = document.createElementNS("http://www.w3.org/2000/svg", "line");
    ax.setAttribute("x1", padL);
    ax.setAttribute("y1", H - pad);
    ax.setAttribute("x2", W - pad);
    ax.setAttribute("y2", H - pad);
    ax.setAttribute("stroke", "#e5e7eb");
    svg.appendChild(ax);

    labels.forEach((lab, i) => {
        const xCenter = padL + i * groupW + groupW / 2;
        addText(svg, xCenter, H - 10, lab, 10, 'middle', '#6b7280');
        const v1 = seriesA[i] || 0, v2 = seriesB[i] || 0;
        const h1 = (H - 2 * pad) * (v1 / max); const h2 = (H - 2 * pad) * (v2 / max);
        const r1 = document.createElementNS("http://www.w3.org/2000/svg", "rect");

        r1.setAttribute("x", xCenter - barW - 2);
        r1.setAttribute("y", H - pad - h1);
        r1.setAttribute("width", barW);
        r1.setAttribute("height", h1);
        r1.setAttribute("fill", colors[0]);

        svg.appendChild(r1);
        const r2 = document.createElementNS("http://www.w3.org/2000/svg", "rect");

        r2.setAttribute("x", xCenter + 2);
        r2.setAttribute("y", H - pad - h2);
        r2.setAttribute("width", barW);
        r2.setAttribute("height", h2);
        r2.setAttribute("fill", colors[1]);
        svg.appendChild(r2);

        // tooltips on hover (valor completo BRL + labels)
        [[r1, v1, labelA], [r2, v2, labelB]].forEach(([rect, val, labY]) => {
            rect.addEventListener('mouseenter', (ev) => showTip(lab + " • " + labY + ": " + fmtBR(val), ev));
            rect.addEventListener('mousemove', (ev) => moveTip(ev));
            rect.addEventListener('mouseleave', hideTip);
            rect.setAttribute('title', fmtBR(val));
        });
    });
}

function lineChart(id, xs, ys, color) {
    const svg = clearSVG(id); const W = 800, H = 320, padL = 56, pad = 40; const max = Math.max(1, ...ys);
    // gridlines
    for (let i = 0; i <= 5; i++) {
        const y = H - pad - (H - 2 * pad) * i / 5;
        const val = max * i / 5;
        const ln = document.createElementNS("http://www.w3.org/2000/svg", "line");
        ln.setAttribute("x1", padL);
        ln.setAttribute("y1", y);
        ln.setAttribute("x2", W - pad);
        ln.setAttribute("y2", y);
        ln.setAttribute("stroke", "#e5e7eb");

        svg.appendChild(ln);
        addText(svg, padL - 6, y + 3, fmtBR(val), 9, 'end', '#6b7280');
    }
    // axis
    const ax = document.createElementNS("http://www.w3.org/2000/svg", "line");
    ax.setAttribute("x1", padL); ax.setAttribute("y1", H - pad);
    ax.setAttribute("x2", W - pad); ax.setAttribute("y2", H - pad);
    ax.setAttribute("stroke", "#e5e7eb");
    svg.appendChild(ax);
    const pts = [];

    xs.forEach((lab, i) => {
        const x = padL + (xs.length === 1 ? 0 : i * (W - padL - pad) / (xs.length - 1));
        const y = H - pad - (H - 2 * pad) * (ys[i] / max); pts.push(x + "," + y);
        addText(svg, x, H - 10, lab, 10, 'middle', '#6b7280');
    });

    const poly = document.createElementNS("http://www.w3.org/2000/svg", "polyline");
    poly.setAttribute("points", pts.join(" "));
    poly.setAttribute("fill", "none");
    poly.setAttribute("stroke", color);
    poly.setAttribute("stroke-width", "2");
    svg.appendChild(poly);

    // points
    xs.forEach((lab, i) => {
        const x = padL + (xs.length === 1 ? 0 : i * (W - padL - pad) / (xs.length - 1));
        const y = H - pad - (H - 2 * pad) * (ys[i] / max);
        const c = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        c.setAttribute("cx", x);
        c.setAttribute("cy", y);
        c.setAttribute("r", 3);
        c.setAttribute("fill", color);
        svg.appendChild(c);
    });
}

function pieChart(id, values, labels) {
    const svg = clearSVG(id); const W = 800, H = 320; const cx = 200, cy = H / 2, r = 100;
    const sum = values.reduce((a, b) => a + b, 0) || 1; let angle = 0;
    values.forEach((v, i) => {
        const a1 = angle; const a2 = angle + 2 * Math.PI * (v / sum); angle = a2;
        const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
        const x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
        const large = (a2 - a1) > Math.PI ? 1 : 0;
        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        const d = "M " + cx + " " + cy + " L " + x1 + " " + y1 + " A " + r + " " + r + " 0 " + large + " 1 " + x2 + " " + y2 + " Z";
        path.setAttribute("d", d);
        path.setAttribute("fill", ["#60a5fa", "#34d399", "#f59e0b", "#ef4444", "#a78bfa", "#10b981", "#fb7185", "#22d3ee"][i % 8]);
        svg.appendChild(path);
        const lx = 420, ly = 40 + i * 18;
        const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        rect.setAttribute("x", lx);
        rect.setAttribute("y", ly - 10);
        rect.setAttribute("width", 10);
        rect.setAttribute("height", 10);
        rect.setAttribute("fill", ["#60a5fa", "#34d399", "#f59e0b", "#ef4444", "#a78bfa", "#10b981", "#fb7185", "#22d3ee"][i % 8]);
        svg.appendChild(rect);
        const txt = document.createElementNS("http://www.w3.org/2000/svg", "text");
        txt.setAttribute("x", lx + 16);
        txt.setAttribute("y", ly);
        txt.setAttribute("font-size", "12");
        txt.textContent = labels[i] + " - " + (v / sum * 100).toFixed(1) + "%";
        svg.appendChild(txt);
    });
}

// Aggregations
function totalByYear(data, ano, cc, conta) {
    return data.filter(d => d.ano === ano && byFilters(d, cc, conta)).reduce((s, d) => s + d.valor, 0);
}

function groupByCC(data, ano, cc, conta) {
    const m = new Map();
    data.filter(d => d.ano === ano && byFilters(d, cc, conta)).forEach(d => {
        m.set(d.centro_custo, (m.get(d.centro_custo) || 0) + d.valor);
    });
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
}
function monthlySeries(data, ano, cc, conta) {
    const arr = new Array(12).fill(0);
    data.filter(d => d.ano === ano && byFilters(d, cc, conta)).forEach(d => {
        arr[d.mes - 1] += d.valor;
    });
    return arr;
}
function tableRows(data, anoA, anoB, cc, conta, search) {
    const q = (search || "").toLowerCase();
    // group by CC+Conta
    const map = new Map();
    data.filter(d => byFilters(d, cc, conta)).forEach(d => {
        const key = d.centro_custo + "|" + d.conta_contabil;
        if (!map.has(key))
            map.set(key, { cc: d.centro_custo, conta: d.conta_contabil, A: 0, B: 0 });
        const obj = map.get(key);
        if (d.ano === anoA) obj.A += d.valor;
        if (d.ano === anoB) obj.B += d.valor;
    });
    return [...map.values()]
        .filter(o => o.cc.toLowerCase().includes(q) || o.conta.toLowerCase().includes(q))
        .map(o => ({ cc: o.cc, conta: o.conta, A: o.A, B: o.B, dif: o.B - o.A, var: Math.abs(o.A) < BASE_MIN_PCT ? null : (o.B - o.A) / o.A * 100 }))
        .sort((a, b) => b.dif - a.dif);
}

// Render
function render(data) {
    document.getElementById("container-dashboard").style.display = 'block';
    const cc = document.getElementById("selCC").value;
    const conta = document.getElementById("selConta").value;
    const anoA = +document.getElementById("selAnoA").value;
    const anoB = +document.getElementById("selAnoB").value;
    const topN = +document.getElementById("selTopN").value;
    const search = document.getElementById("txtSearch").value || "";

    // Labels
    document.getElementById("lblAnoA").textContent = "Total " + anoA;
    document.getElementById("lblAnoB").textContent = "Total " + anoB;
    document.getElementById("lgAnoA").textContent = anoA;
    document.getElementById("lgAnoB").textContent = anoB;
    document.getElementById("thAnoA").textContent = anoA;
    document.getElementById("thAnoB").textContent = anoB;

    // Cards
    const totA = totalByYear(data, anoA, cc, conta);
    const totB = totalByYear(data, anoB, cc, conta);
    document.getElementById("cardAnoA").textContent = fmtBR(totA);
    document.getElementById("cardAnoB").textContent = fmtBR(totB);
    document.getElementById("cardDiff").textContent = fmtBR(totB - totA);
    document.getElementById("cardVar").textContent = fmtP(totA === 0 ? NaN : (totB - totA) / totA * 100);

    // Charts (Anual) with Top N
    const gA = groupByCC(data, anoA, cc === '*' ? null : cc, conta === '*' ? null : conta);
    const gB = groupByCC(data, anoB, cc === '*' ? null : cc, conta === '*' ? null : conta);
    let labels = unique([...gA.map(x => x[0]), ...gB.map(x => x[0])]);
    if (topN > 0) labels = labels.slice(0, topN);
    const sA = labels.map(lab => (gA.find(x => x[0] === lab)?.[1]) || 0);
    const sB = labels.map(lab => (gB.find(x => x[0] === lab)?.[1]) || 0);
    barChart("chartComparativo", labels, sA, sB, ["#2563eb", "#10b981"], "Ano " + anoA, "Ano " + anoB);
    document.getElementById("ttlComparativo").textContent = "Comparativo " + anoA + " vs " + anoB + " (Top " + (topN || labels.length) + ")";

    // Pie (Ano B)
    const gb = gB.slice(0, topN > 0 ? topN : gB.length);
    pieChart("chartPizza", gb.map(x => x[1]), gb.map(x => x[0]));
    document.getElementById("ttlDistribuicao").textContent = "Distribuição por Centro de Custo (" + anoB + ") – Top " + (topN || gb.length);

    // Linha evolução (Ano B)
    const serieB = monthlySeries(data, anoB, cc === '*' ? null : cc, conta === '*' ? null : conta);
    lineChart("chartLinha", monthNames, serieB, "#10b981");
    document.getElementById("ttlEvolucao").textContent = "Evolução Mensal (" + anoB + ")";

    // Tabela com busca textual
    const rows = tableRows(data, anoA, anoB, cc === '*' ? null : cc, conta === '*' ? null : conta, search);
    const tbody = document.querySelector("#tbl tbody");
    tbody.innerHTML = "";
    rows.forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = "<td>" + r.cc + "</td><td>" + r.conta + "</td>"
            + "<td class='num'>" + fmtBR(r.A) + "</td><td class='num'>" + fmtBR(r.B) + "</td>"
            + "<td class='num'>" + fmtBR(r.dif) + "</td><td class='num'>" + (r.var == null ? '0.00%' : fmtP(r.var)) + "</td>";
        tbody.appendChild(tr);
    });

    // Mensal: barras por mês (Ano A vs Ano B)
    const serieA = monthlySeries(data, anoA, cc === '*' ? null : cc, conta === '*' ? null : conta);
    barChart("chartMensal", monthNames, serieA, serieB, ["#2563eb", "#10b981"], "Ano " + anoA, "Ano " + anoB);
    document.getElementById("ttlMensal").textContent = "Valores por Mês - " + anoA + " (azul) vs " + anoB + " (verde)";
}

// Reactivity
["selCC", "selConta", "selAnoA", "selAnoB", "selTopN"].forEach(id => document.getElementById(id).addEventListener("change", () => render(DATA)));
let searchTimer = null;

document.getElementById("txtSearch").addEventListener("input", (e) => { clearTimeout(searchTimer); searchTimer = setTimeout(() => render(DATA), 250); });