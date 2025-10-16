let secoes = [];
let dadosPorSecao = [];
let colunasPorSecao = [];
let colunasSelecionadasPorSecao = [];
let nomesArquivos = [];

const inputExcel = document.getElementById("input-excel");
const secoesContainer = document.getElementById("secoes");
const abasContainer = document.getElementById("abas-container");

// 📂 Importar múltiplos arquivos
inputExcel.addEventListener("change", async function (e) {
    for (const file of e.target.files) {
        await carregarArquivo(file);
    }
    atualizarAbas();
});

// 🔹 Criar nova seção para cada arquivo
async function carregarArquivo(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const dados = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            const colunas = Object.keys(dados[0] || {});
            const id = secoes.length;

            // 🔤 Nome limpo (sem extensão nem caracteres especiais)
            const nomeLimpo = file.name
                .replace(/\.[^/.]+$/, "")
                .replace(/[^a-zA-Z0-9À-ÿ_\- ]/g, "")
                .trim();

            nomesArquivos.push(nomeLimpo);

            const secao = document.createElement("div");
            secao.classList.add("secao");
            secao.id = `secao-${id}`;
            secao.innerHTML = `
                <div id="colunas-${id}"></div>
                <div class="botoes">
                    <button id="gerar-${id}">Gerar Tabela</button>
                    <button id="exportar-${id}" style="display:none;">Exportar Excel</button>
                    <button id="imprimir-${id}" style="display:none;">🖨️ Imprimir</button>
                </div>
                <div id="cabecalho-relatorio">
                    <h2>Relatório - ${nomeLimpo}</h2>
                    <p id="data-relatorio-${id}"></p>
                </div>
                <table id="tabela-${id}"></table>
            `;

            secoesContainer.appendChild(secao);
            secoes.push(secao);
            dadosPorSecao.push(dados);
            colunasPorSecao.push(colunas);
            colunasSelecionadasPorSecao.push(colunas);

            mostrarCheckboxes(id);
            configurarBotoes(id);
            resolve();
        };
        reader.readAsArrayBuffer(file);
    });
}

// 🧩 Mostrar checkboxes
function mostrarCheckboxes(id) {
    const div = document.getElementById(`colunas-${id}`);
    div.innerHTML = "<h3>Selecione as colunas:</h3>";
    colunasPorSecao[id].forEach((col) => {
        div.innerHTML += `
          <label>
            <input type="checkbox" class="coluna-check-${id}" value="${col}" checked> ${col}
          </label>
        `;
    });
}

// ⚙️ Configurar botões
function configurarBotoes(id) {
    document.getElementById(`gerar-${id}`).addEventListener("click", () => {
        const checks = [...document.querySelectorAll(`.coluna-check-${id}:checked`)];
        colunasSelecionadasPorSecao[id] = checks.map((c) => c.value);
        gerarTabela(id);
        document.getElementById(`exportar-${id}`).style.display = "inline-block";
        document.getElementById(`imprimir-${id}`).style.display = "inline-block";
    });

    document.getElementById(`exportar-${id}`).addEventListener("click", () => {
        exportarExcel(id);
    });

    document.getElementById(`imprimir-${id}`).addEventListener("click", () => {
        const dataAtual = new Date();
        const formatado = dataAtual.toLocaleDateString("pt-BR") + " " + dataAtual.toLocaleTimeString("pt-BR");
        document.getElementById(`data-relatorio-${id}`).textContent = "Gerado em: " + formatado;
        selecionarAba(id);
        window.print();
    });
}

// 🧱 Gerar tabela
function gerarTabela(id) {
    const tabela = document.getElementById(`tabela-${id}`);
    tabela.innerHTML = "";

    const colunas = colunasSelecionadasPorSecao[id];
    const dados = dadosPorSecao[id];

    const thead = document.createElement("thead");
    const trHead = document.createElement("tr");
    colunas.forEach((c) => {
        const th = document.createElement("th");
        th.textContent = c;
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    tabela.appendChild(thead);

    const tbody = document.createElement("tbody");
    dados.forEach((linha) => {
        const tr = document.createElement("tr");
        colunas.forEach((c) => {
            const td = document.createElement("td");
            td.textContent = linha[c];
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    tabela.appendChild(tbody);

    ativarSelecaoExcel(tabela);
}

// 📤 Exportar Excel
function exportarExcel(id) {
    const colunas = colunasSelecionadasPorSecao[id];
    const dados = dadosPorSecao[id];
    const filtrado = dados.map((linha) => {
        const obj = {};
        colunas.forEach((c) => (obj[c] = linha[c]));
        return obj;
    });

    const ws = XLSX.utils.json_to_sheet(filtrado);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Filtrado");
    XLSX.writeFile(wb, `${nomesArquivos[id] || "Tabela"}_Filtrada.xlsx`);
}

// 🔄 Atualiza abas
function atualizarAbas() {
    abasContainer.innerHTML = "";
    secoes.forEach((secao, i) => {
        const aba = document.createElement("div");
        aba.classList.add("aba");
        aba.innerHTML = `
            <span class="nome">${nomesArquivos[i] || `Planilha ${i + 1}`}</span>
            <span class="fechar" title="Fechar">❌</span>
        `;
        aba.querySelector(".nome").addEventListener("click", () => selecionarAba(i));
        aba.querySelector(".fechar").addEventListener("click", (e) => {
            e.stopPropagation();
            removerAba(i);
        });
        abasContainer.appendChild(aba);
    });
    if (secoes.length > 0) selecionarAba(secoes.length - 1);
}

// ❌ Remover aba e planilha
function removerAba(id) {
    secoes[id].remove();
    secoes.splice(id, 1);
    dadosPorSecao.splice(id, 1);
    colunasPorSecao.splice(id, 1);
    colunasSelecionadasPorSecao.splice(id, 1);
    nomesArquivos.splice(id, 1);
    atualizarAbas();
}

// 🟢 Selecionar aba
function selecionarAba(id) {
    secoes.forEach((s, i) => {
        s.classList.toggle("ativa", i === id);
    });
    document.querySelectorAll(".aba").forEach((a, i) => {
        a.classList.toggle("ativa", i === id);
    });
}

// ✴️ Seleção estilo Excel
function ativarSelecaoExcel(tabela) {
    let selecionando = false;
    let colunaInicial = null;

    tabela.querySelectorAll("td").forEach((td) => {
        td.addEventListener("mousedown", (e) => {
            if (e.ctrlKey || e.shiftKey) return;
            e.preventDefault();
            selecionando = true;
            colunaInicial = td.cellIndex;
            limparSelecao(tabela);
            td.classList.add("selecionado");
        });

        td.addEventListener("mouseenter", (e) => {
            if (selecionando && td.cellIndex === colunaInicial) {
                td.classList.add("selecionado");
            }
        });

        td.addEventListener("mouseup", () => {
            selecionando = false;
        });
    });

    document.addEventListener("mouseup", () => (selecionando = false));
}

function limparSelecao(tabela) {
    tabela.querySelectorAll(".selecionado").forEach((td) => td.classList.remove("selecionado"));
    tabela.querySelectorAll(".linha-selecionada").forEach((tr) => tr.classList.remove("linha-selecionada"));
}
