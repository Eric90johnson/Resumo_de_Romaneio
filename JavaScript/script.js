const LOCAL_STORAGE_KEY = 'resumoDeCodigosDados_v2';
let dadosAgregados = new Map();
let limparDadosAntesDeCarregar = true;
let dadosIniciaisCarregados = false; 

const inputArquivo = document.getElementById('inputCSV');
const btnCarregar = document.getElementById('btnCarregar');
const btnAdicionar = document.getElementById('btnAdicionar');
const btnExcluir = document.getElementById('btnExcluir');
const btnConfirmarExclusaoFinal = document.getElementById('btnConfirmarExclusaoFinal');
const btnAbrirModalImpressao = document.getElementById('btnAbrirModalImpressao');
const btnConfirmarImpressao = document.getElementById('btnConfirmarImpressao');

// Referência para os checkboxes
const checkCargaFracionada = document.getElementById('checkCargaFracionada');
const checkCargaFechada = document.getElementById('checkCargaFechada');

const modalAviso = new bootstrap.Modal(document.getElementById('modalAvisoCarregamento'));
const modalExclusao = new bootstrap.Modal(document.getElementById('modalConfirmarExclusao'));
const modalAvisoInserir = new bootstrap.Modal(document.getElementById('modalAvisoInserir'));
const modalImpressao = new bootstrap.Modal(document.getElementById('modalImpressao'));

// Inicialização do novo modal de alerta
const modalAvisoRomaneio = new bootstrap.Modal(document.getElementById('modalAvisoRomaneio'));

const btnCopiar65 = document.getElementById('btnCopiar65');
const btnCopiar4 = document.getElementById('btnCopiar4');
const btnCopiarRefrigeracao = document.getElementById('btnCopiarRefrigeracao');

const todosOsContainers = document.querySelectorAll('.tabela-container');
const corpoTabela65 = document.querySelector('#tabelaFogao65 tbody');
const spanTotal65 = document.getElementById('totalFogao65');
const corpoTabela4 = document.querySelector('#tabelaFogao4 tbody');
const spanTotal4 = document.getElementById('totalFogao4');
const corpoTabelaRefrigeracao = document.querySelector('#tabelaRefrigeracao tbody');
const spanTotalRefrigeracao = document.getElementById('totalRefrigeracao');

const spanTotalGeral = document.getElementById('totalGeral');

btnAbrirModalImpressao.addEventListener('click', () => { modalImpressao.show(); });
btnConfirmarImpressao.addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('#modalImpressao .form-check-input');
    modalImpressao.hide();
    todosOsContainers.forEach(container => {
        container.classList.add('d-print-none');
    });
    checkboxes.forEach(check => {
        if (check.checked) {
            const containerId = check.value;
            document.getElementById(containerId).classList.remove('d-print-none');
            document.getElementById(containerId).classList.add('imprimir-bloco');
        }
    });
    setTimeout(() => { window.print(); }, 500);
});
window.onafterprint = function() {
    todosOsContainers.forEach(container => {
        container.classList.remove('d-print-none');
        container.classList.remove('imprimir-bloco');
    });
};

// Event listener do botão "Carregar Dados" com a nova validação
btnCarregar.addEventListener('click', () => {
    if (dadosIniciaisCarregados) {
        modalAviso.show();
        return;
    }
    // Validação dos checkboxes
    if (!checkCargaFracionada.checked && !checkCargaFechada.checked) {
        modalAvisoRomaneio.show();
        return;
    }
    limparDadosAntesDeCarregar = true;
    inputArquivo.click();
});

// Event listener do botão "Adicionar Mais Dados" com a nova validação
btnAdicionar.addEventListener('click', () => {
    if (!dadosIniciaisCarregados) {
        modalAvisoInserir.show();
        return;
    }
    // Validação dos checkboxes
    if (!checkCargaFracionada.checked && !checkCargaFechada.checked) {
        modalAvisoRomaneio.show();
        return;
    }
    limparDadosAntesDeCarregar = false;
    inputArquivo.click();
});

btnExcluir.addEventListener('click', () => { modalExclusao.show(); });
btnConfirmarExclusaoFinal.addEventListener('click', () => { dadosAgregados.clear(); localStorage.removeItem(LOCAL_STORAGE_KEY); dadosIniciaisCarregados = false; renderizarTabelas(); modalExclusao.hide(); });
btnCopiar65.addEventListener('click', () => { const textoParaCopiar = gerarTextoTabela(corpoTabela65); copiarParaClipboard(textoParaCopiar, btnCopiar65); });
btnCopiar4.addEventListener('click', () => { const textoParaCopiar = gerarTextoTabela(corpoTabela4); copiarParaClipboard(textoParaCopiar, btnCopiar4); });
btnCopiarRefrigeracao.addEventListener('click', () => { const textoParaCopiar = gerarTextoTabela(corpoTabelaRefrigeracao); copiarParaClipboard(textoParaCopiar, btnCopiarRefrigeracao); });

inputArquivo.addEventListener('change', (evento) => { 
    const arquivo = evento.target.files[0]; if (!arquivo) return;
    const nomeArquivo = arquivo.name.toLowerCase();
    const leitor = new FileReader();
    leitor.onerror = function(e) { console.error("ERRO GRAVE: Ocorreu um erro ao tentar ler o arquivo.", e); alert("Não foi possível ler o arquivo. Ele pode estar protegido ou corrompido."); };
    if (nomeArquivo.endsWith('.csv')) {
        leitor.onload = (e) => processarArquivoCSV(e.target.result);
        leitor.readAsText(arquivo, 'ISO-8859-1'); 
    } else if (nomeArquivo.endsWith('.xlsx') || nomeArquivo.endsWith('.xls')) {
        leitor.onload = (e) => processarArquivoExcel(e.target.result);
        leitor.readAsArrayBuffer(arquivo);
    } else {
        alert('Formato de arquivo não suportado.');
    }
    inputArquivo.value = '';
});

// Lógica para que apenas um checkbox possa ser selecionado por vez
checkCargaFracionada.addEventListener('change', () => {
    if (checkCargaFracionada.checked) {
        checkCargaFechada.checked = false;
    }
});

checkCargaFechada.addEventListener('change', () => {
    if (checkCargaFechada.checked) {
        checkCargaFracionada.checked = false;
    }
});

function gerarTextoTabela(tbody) {
    const linhas = [];
    tbody.querySelectorAll('tr').forEach(linha => {
        const colunas = [];
        linha.querySelectorAll('td').forEach(coluna => {
            colunas.push(coluna.textContent);
        });
        linhas.push(colunas.join('\t'));
    });
    return linhas.join('\n');
}

function copiarParaClipboard(texto, botao) {
    if (!texto) {
        alert('Não há dados para copiar.');
        return;
    }
    navigator.clipboard.writeText(texto).then(() => {
        const textoOriginal = botao.textContent;
        botao.classList.remove('btn-primary');
        botao.classList.add('btn-success');
        botao.textContent = 'Copiado!';
        setTimeout(() => {
            botao.textContent = textoOriginal;
            botao.classList.remove('btn-success');
            botao.classList.add('btn-primary');
        }, 2000);
    }).catch(err => {
        console.error('Falha ao copiar texto: ', err);
        alert('Não foi possível copiar os dados.');
    });
}

function processarArquivoCSV(textoCsv) { 
  const separador = ';'; 
  const linhas = textoCsv.split('\n'); 
  const dados = linhas.map(linha => linha.split(separador)); 
  aplicarLogicaDeNegocio(dados); 
}

function processarArquivoExcel(arrayBuffer) {
    let workbook;
    try {
        workbook = XLSX.read(arrayBuffer, {type: 'buffer'});
    } catch (e) {
        console.error("Falha ao ler o arquivo Excel.", e);
        alert('Não foi possível ler o arquivo Excel. O arquivo pode estar corrompido ou em um formato não suportado.');
        return;
    }
    const nomePlanilha = workbook.SheetNames[0];
    const planilha = workbook.Sheets[nomePlanilha];
    const dados = XLSX.utils.sheet_to_json(planilha, {header: 1});
    aplicarLogicaDeNegocio(dados);
}

function aplicarLogicaDeNegocio(linhas) {
  if (limparDadosAntesDeCarregar) { dadosAgregados.clear(); }
  let codigoAtual = null, descricaoAtual = '';
  for (const colunas of linhas) {
    try {
      if (!Array.isArray(colunas) || colunas.length === 0) continue;
      
      const primeiraCelula = String(colunas[0] || '').toUpperCase();
      if (primeiraCelula.includes("RELATÓRIO EXPEDIÇÃO") || primeiraCelula.includes("TOTAL GERAL")) {
          console.log(`Palavra-chave de parada '${primeiraCelula}' encontrada. Finalizando a leitura.`);
          break; 
      }

      const possivelCodigo = String(colunas[0] || '').trim();
      if (possivelCodigo && !isNaN(possivelCodigo) && possivelCodigo.length > 3) {
        codigoAtual = possivelCodigo;
        descricaoAtual = String(colunas[1] || '').trim();
      }
      
      const isTotalRow = colunas.slice(0, 5).some(cell => typeof cell === 'string' && cell.trim().toUpperCase() === 'TOTAL');
      if (isTotalRow && codigoAtual) {
        const valorTotal = colunas[9] ? parseInt(String(colunas[9]).trim(), 10) : 0;
        if (!isNaN(valorTotal)) {
            const tipoProduto = getTipoProduto(descricaoAtual);
            const itemExistente = dadosAgregados.get(codigoAtual) || { total: 0, tipo: tipoProduto };
            dadosAgregados.set(codigoAtual, { total: itemExistente.total + valorTotal, tipo: tipoProduto });
        }
        codigoAtual = null;
        descricaoAtual = '';
      }
    } catch (error) {
        console.warn("Uma linha com formato inesperado foi ignorada. Erro:", error, "Dados da linha:", colunas);
        continue; 
    }
  }
  if(dadosAgregados.size > 0) { dadosIniciaisCarregados = true; }
  salvarDados();
  renderizarTabelas();
}

function getTipoProduto(descricao) { 
  const descUpper = String(descricao || '').toUpperCase();
  const keywordsRefrigeracao = ["BEBEDOURO", "REFRIGERADOR", "CONSERVADOR", "PURIFICADOR", "FREEZER", "VITRINE"];
  if (keywordsRefrigeracao.some(keyword => descUpper.includes(keyword))) { return 'Refrigeracao'; }
  const match = descUpper.match(/\b(\d{4})\b/); 
  if (match) { 
    const primeiroDigito = match[1][0]; 
    if (['5', '6'].includes(primeiroDigito)) return '65Q'; 
    if (primeiroDigito === '4') return '4Q'; 
  } 
  return 'outro'; 
}

function salvarDados() { 
  const dadosArray = [...dadosAgregados.entries()]; 
  localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(dadosArray)); 
}

function carregarDados() {
  const dadosSalvos = localStorage.getItem(LOCAL_STORAGE_KEY);
  if (dadosSalvos) {
    const dadosArray = JSON.parse(dadosSalvos);
    dadosAgregados = new Map(dadosArray);
    if(dadosAgregados.size > 0) { dadosIniciaisCarregados = true; }
    renderizarTabelas();
  }
}

function renderizarTabelas() {
  corpoTabela65.innerHTML = '';
  corpoTabela4.innerHTML = '';
  corpoTabelaRefrigeracao.innerHTML = '';
  let total65 = 0, total4 = 0, totalRefrigeracao = 0;
  
  const dadosOrdenados = new Map([...dadosAgregados.entries()].sort());
  
  dadosOrdenados.forEach((item, codigo) => {
    const novaLinha = document.createElement('tr');
    novaLinha.innerHTML = `<td>0${codigo}</td><td>${item.total}</td>`;
    if (item.tipo === '65Q') {
      corpoTabela65.appendChild(novaLinha);
      total65 += item.total;
    } else if (item.tipo === '4Q') {
      corpoTabela4.appendChild(novaLinha);
      total4 += item.total;
    } else if (item.tipo === 'Refrigeracao') {
      corpoTabelaRefrigeracao.appendChild(novaLinha);
      totalRefrigeracao += item.total;
    }
  });
  
  spanTotal65.textContent = total65;
  spanTotal4.textContent = total4;
  spanTotalRefrigeracao.textContent = totalRefrigeracao;
  spanTotalGeral.textContent = total65 + total4 + totalRefrigeracao;
}

// Carrega os dados do localStorage ao iniciar a página
carregarDados();