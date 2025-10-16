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
    // Usa colunasPorSecao[id] como fonte de verdade para as colunas originais.
    colunasPorSecao[id].forEach((col) => {
        // Verifica se a coluna ainda está na lista de colunas selecionadas para marcar o checkbox
        const isChecked = colunasSelecionadasPorSecao[id].includes(col) ? 'checked' : '';
        div.innerHTML += `
          <label>
            <input type="checkbox" class="coluna-check-${id}" value="${col}" ${isChecked}> ${col}
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
        // Coloca o foco na tabela (opcional, mas bom para garantir a impressão)
        document.getElementById(`secao-${id}`).scrollIntoView({ behavior: 'smooth' });
        window.print();
    });
}

// 🧱 Gerar tabela
function gerarTabela(id) {
    const tabela = document.getElementById(`tabela-${id}`);
    tabela.innerHTML = "";

    // Usa sempre a lista de colunas SELECIONADAS para renderizar a tabela
    const colunas = colunasSelecionadasPorSecao[id];
    const dados = dadosPorSecao[id];

    // Cria o cabeçalho (thead)
    const thead = document.createElement("thead");
    const trHead = document.createElement("tr");

    colunas.forEach((c, colIndex) => {
        const th = document.createElement("th");
        th.textContent = c;
        th.dataset.coluna = c; // Nome da coluna
        th.dataset.colIndex = colIndex; // Índice de exibição

        // Adiciona evento de clique duplo para selecionar/remover a coluna
        th.addEventListener("dblclick", () => selecionarColuna(tabela, colIndex));

        // Adiciona ícone de remoção (X)
        const removeIcon = document.createElement("span");
        removeIcon.classList.add("remover-coluna");
        removeIcon.textContent = "❌";
        removeIcon.title = `Remover coluna ${c}`;
        removeIcon.addEventListener("click", (e) => {
            e.stopPropagation();
            // Passamos o ID da seção e o NOME da coluna
            removerColuna(id, c);
        });
        th.appendChild(removeIcon);

        trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    tabela.appendChild(thead);

    // Cria o corpo (tbody)
    const tbody = document.createElement("tbody");
    dados.forEach((linha, rowIndex) => {
        const tr = document.createElement("tr");
        tr.dataset.linha = rowIndex; // Índice real no array de dados

        colunas.forEach((c, colIndex) => {
            const td = document.createElement("td");
            // Armazena o valor real em um atributo de dados para cópia limpa
            td.dataset.value = linha[c];
            td.textContent = linha[c];

            // Adiciona evento de clique duplo na primeira célula para copiar a linha.
            if (colIndex === 0) {
                td.addEventListener("dblclick", () => selecionarLinhaCompleta(tr));
            }

            // Adiciona ícone de remoção (X) na primeira célula
            if (colIndex === 0) {
                const removeIcon = document.createElement("span");
                removeIcon.classList.add("remover-linha");
                removeIcon.textContent = "❌";
                removeIcon.title = "Remover esta linha";
                removeIcon.addEventListener("click", (e) => {
                    e.stopPropagation();
                    // Passamos o ID da seção e o ÍNDICE REAL da linha no array de dados
                    removerLinha(id, rowIndex);
                });
                td.prepend(removeIcon);
            }

            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    tabela.appendChild(tbody);

    ativarSelecaoExcel(tabela);
}

// 🗑️ Remover coluna da tabela e dos dados
function removerColuna(id, nomeColuna) {
    if (!confirm(`Tem certeza que deseja remover a coluna "${nomeColuna}"?`)) return;

    // 1. Remove a coluna dos dados brutos (apaga a chave de cada objeto)
    dadosPorSecao[id].forEach(linha => delete linha[nomeColuna]);

    // 2. Remove a coluna das colunas selecionadas
    colunasSelecionadasPorSecao[id] = colunasSelecionadasPorSecao[id].filter(c => c !== nomeColuna);

    // 3. Remove a coluna da lista de todas as colunas disponíveis
    colunasPorSecao[id] = colunasPorSecao[id].filter(c => c !== nomeColuna);

    // 4. Re-renderiza tudo
    mostrarCheckboxes(id);
    gerarTabela(id);
}

// 🗑️ Remover linha da tabela e dos dados
function removerLinha(id, rowIndex) {
    if (!confirm("Tem certeza que deseja remover esta linha?")) return;
    // Remove a linha do array de dados brutos usando o índice
    dadosPorSecao[id].splice(rowIndex, 1);

    // Re-renderiza a tabela (O índice `rowIndex` foi passado corretamente e agora o array foi modificado)
    gerarTabela(id);
}

// 👆 Selecionar coluna inteira (em clique duplo no TH)
function selecionarColuna(tabela, colIndex) {
    limparSelecao(tabela);
    // Seleciona todas as células na coluna
    const cells = tabela.querySelectorAll(`tr > *:nth-child(${colIndex + 1})`);

    // Mapeia para o texto (usando data-value para ignorar o ícone na primeira célula)
    const textoColuna = Array.from(cells)
        .filter(cell => cell.tagName !== 'TH') // Ignora o cabeçalho
        .map(cell => cell.dataset.value || cell.textContent.trim())
        .join('\n');

    cells.forEach(cell => cell.classList.add("selecionado"));

    // Tenta copiar para o clipboard
    navigator.clipboard.writeText(textoColuna).then(() => {
        console.log("Coluna copiada para a área de transferência!");
        alert("Coluna copiada! (Cole onde desejar)");
    }).catch(err => {
        console.error("Erro ao copiar: ", err);
    });
}

// 👆 Selecionar linha inteira (em clique duplo no TD)
function selecionarLinhaCompleta(tr) {
    limparSelecao(tr.closest('table'));

    const cells = tr.querySelectorAll("td");

    // Mapeia para o texto (usando data-value para ignorar o ícone na primeira célula)
    const textoLinha = Array.from(cells)
        .map(cell => cell.dataset.value || cell.textContent.trim())
        .join('\t');

    tr.classList.add("linha-selecionada");

    // Tenta copiar para o clipboard
    navigator.clipboard.writeText(textoLinha).then(() => {
        console.log("Linha copiada para a área de transferência!");
        alert("Linha copiada! (Cole onde desejar)");
    }).catch(err => {
        console.error("Erro ao copiar: ", err);
    });
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
