"use strict";

// ============================================================
// 1. ELEMENTOS DA PÁGINA
// ============================================================

const campoBusca = document.querySelector("#campo-busca");
const filtroSituacao = document.querySelector("#filtro-situacao");
const filtroLocal = document.querySelector("#filtro-local");
const filtroResponsavel = document.querySelector("#filtro-responsavel");

const corpoFerramentas = document.querySelector("#corpo-ferramentas");
const cartoesResumo = document.querySelector("#cartoes-resumo");
const resultadoBusca = document.querySelector("#resultado-busca");
const legendaTabela = document.querySelector("#legenda-tabela");

// Elementos responsáveis pela paginação.
const itensPorPagina = document.querySelector("#itens-por-pagina");
const botaoPrimeira = document.querySelector("#pagina-primeira");
const botaoAnterior = document.querySelector("#pagina-anterior");
const botaoProxima = document.querySelector("#pagina-proxima");
const botaoUltima = document.querySelector("#pagina-ultima");
const informacaoPaginacao = document.querySelector("#informacao-paginacao");
const botaoExportar = document.querySelector("#botao-exportar");

// Armazena todos os registros recebidos da API.
let ferramentas = [];

// Controla a navegação da tabela.
let paginaAtual = 1;
let quantidadePorPagina = 10;

// ============================================================
// 2. FUNÇÕES AUXILIARES
// ============================================================

// Converte valores para texto e evita erros com campos vazios.
function obterTexto(valor) {
  return String(valor ?? "").trim();
}

// Normaliza textos para comparações.
function normalizar(texto) {
  return obterTexto(texto)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleUpperCase("pt-BR");
}

// Protege o HTML contra caracteres especiais presentes na planilha.
function escaparHTML(valor) {
  return obterTexto(valor).replace(/[&<>"']/g, caractere => {
    const entidades = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    };

    return entidades[caractere];
  });
}

// Formata números no padrão brasileiro.
function formatarNumero(numero) {
  return Number(numero || 0).toLocaleString("pt-BR");
}

// Verifica se uma ferramenta possui patrimônio válido.
// Campo vazio e hífen representam ausência de patrimônio.
function possuiPatrimonio(ferramenta) {
  const patrimonio = obterTexto(ferramenta.patrimonio);

  return patrimonio !== "" && patrimonio !== "-";
}

// Retorna uma situação legível, inclusive quando não informada.
function obterSituacao(ferramenta) {
  return obterTexto(ferramenta.situacao) || "NÃO INFORMADA";
}

// Classifica visualmente as situações conhecidas.
function obterClasseSituacao(situacao) {
  const texto = normalizar(situacao);

  if (texto.includes("DISPON")) return "situacao--disponivel";
  if (texto.includes("CAMPO")) return "situacao--campo";
  if (texto.includes("MANUT")) return "situacao--manutencao";

  return "situacao--outra";
}

// ============================================================
// 3. CARREGAMENTO DOS DADOS DA API
// ============================================================

async function carregarFerramentas() {
  try {
    resultadoBusca.textContent = "Carregando inventário...";

    const resposta = await fetch("/api/ferramentas", {
      cache: "no-store"
    });

    if (!resposta.ok) {
      throw new Error(`Erro HTTP ${resposta.status}`);
    }

    const dados = await resposta.json();

    if (!Array.isArray(dados.ferramentas)) {
      throw new Error("A API não retornou uma lista válida de ferramentas.");
    }

    ferramentas = dados.ferramentas;
    paginaAtual = 1;

    // Libera a exportação após o carregamento do inventário.
    if (botaoExportar) botaoExportar.disabled = false;

    // Atualiza os componentes do dashboard.
    atualizarResumo(dados);
    preencherFiltros();
    atualizarTabela();

    if (legendaTabela) {
      legendaTabela.textContent =
        "Dados carregados diretamente da planilha Excel.";
    }

  } catch (erro) {
    console.error("Erro ao carregar ferramentas:", erro);

    if (botaoExportar) botaoExportar.disabled = true;

    cartoesResumo.innerHTML = `
      <article class="cartao cartao--erro">
        <p class="cartao__titulo">Falha ao carregar os dados</p>
        <strong class="cartao__numero">!</strong>
        <p class="cartao__descricao">Verifique o servidor e a API.</p>
      </article>
    `;

    corpoFerramentas.innerHTML = `
      <tr>
        <td colspan="6">Não foi possível carregar as ferramentas.</td>
      </tr>
    `;

    resultadoBusca.textContent = "Erro ao consultar o inventário.";

    atualizarControlesPaginacao(0, 0);
  }
}

// ============================================================
// 4. INDICADORES DO DASHBOARD
// ============================================================

function atualizarResumo(dados) {
  const total = ferramentas.length;

  const comPatrimonio = ferramentas.filter(possuiPatrimonio).length;
  const semPatrimonio = total - comPatrimonio;

  // Agrupa as ferramentas de acordo com a situação da planilha.
  const agrupamento = new Map();

  ferramentas.forEach(ferramenta => {
    const situacao = obterSituacao(ferramenta);
    const chave = normalizar(situacao);

    if (!agrupamento.has(chave)) {
      agrupamento.set(chave, {
        nome: situacao,
        quantidade: 0
      });
    }

    agrupamento.get(chave).quantidade++;
  });

  // Indicadores gerais.
  let html = `
    <article class="cartao">
      <p class="cartao__titulo">Total de ferramentas</p>
      <strong class="cartao__numero">${formatarNumero(total)}</strong>
      <p class="cartao__descricao">Registros encontrados na planilha</p>
    </article>

    <article class="cartao">
      <p class="cartao__titulo">Com patrimônio</p>
      <strong class="cartao__numero">${formatarNumero(comPatrimonio)}</strong>
      <p class="cartao__descricao">Ferramentas identificadas</p>
    </article>

    <article class="cartao">
      <p class="cartao__titulo">Sem patrimônio</p>
      <strong class="cartao__numero">${formatarNumero(semPatrimonio)}</strong>
      <p class="cartao__descricao">Sem identificação patrimonial</p>
    </article>
  `;

  // Indicadores por situação.
  for (const grupo of agrupamento.values()) {
    html += `
      <article class="cartao">
        <p class="cartao__titulo">${escaparHTML(grupo.nome)}</p>
        <strong class="cartao__numero">
          ${formatarNumero(grupo.quantidade)}
        </strong>
        <p class="cartao__descricao">Ferramentas nesta situação</p>
      </article>
    `;
  }

  cartoesResumo.innerHTML = html;
}

// ============================================================
// 5. PREENCHIMENTO AUTOMÁTICO DOS FILTROS
// ============================================================

function preencherFiltros() {
  preencherSelect(
    filtroSituacao,
    ferramentas.map(f => obterSituacao(f)),
    "Todas"
  );

  preencherSelect(
    filtroLocal,
    ferramentas.map(f => obterTexto(f.local) || "NÃO INFORMADO"),
    "Todos"
  );

  preencherSelect(
    filtroResponsavel,
    ferramentas.map(f => obterTexto(f.responsavel) || "NÃO INFORMADO"),
    "Todos"
  );
}

// Cria opções sem duplicar valores.
function preencherSelect(elemento, valores, textoPadrao) {
  const selecionado = elemento.value;

  const valoresUnicos = [...new Map(
    valores.map(valor => [normalizar(valor), valor])
  ).values()].sort((a, b) => a.localeCompare(b, "pt-BR"));

  elemento.innerHTML = `
    <option value="">${textoPadrao}</option>
    ${valoresUnicos.map(valor => `
      <option value="${escaparHTML(valor)}">
        ${escaparHTML(valor)}
      </option>
    `).join("")}
  `;

  // Mantém a seleção anterior, se ainda existir.
  if (valoresUnicos.some(valor => valor === selecionado)) {
    elemento.value = selecionado;
  }
}

// ============================================================
// 6. PESQUISA E FILTROS
// ============================================================

function atualizarTabela() {
  // Reinicia a paginação quando a pesquisa ou algum filtro muda.
  paginaAtual = 1;

  // Usa a mesma filtragem da navegação e da exportação.
  renderizarTabela(obterFerramentasFiltradas());
}

// ============================================================
// 7. RENDERIZAÇÃO DA TABELA COM PAGINAÇÃO
// ============================================================

function renderizarTabela(lista) {
  const totalEncontrado = lista.length;

  const totalPaginas = Math.ceil(totalEncontrado / quantidadePorPagina);

  if (totalPaginas === 0) {
    paginaAtual = 1;
  } else if (paginaAtual > totalPaginas) {
    paginaAtual = totalPaginas;
  }

  const indiceInicial = (paginaAtual - 1) * quantidadePorPagina;
  const indiceFinal = Math.min(
    indiceInicial + quantidadePorPagina,
    totalEncontrado
  );

  const ferramentasDaPagina = lista.slice(indiceInicial, indiceFinal);

  if (ferramentasDaPagina.length === 0) {
    corpoFerramentas.innerHTML = `
      <tr>
        <td colspan="6">Nenhuma ferramenta encontrada.</td>
      </tr>
    `;
  } else {
    corpoFerramentas.innerHTML = ferramentasDaPagina.map(ferramenta => {
      const situacao = obterSituacao(ferramenta);
      const classe = obterClasseSituacao(situacao);

      return `
        <tr>
          <td>${escaparHTML(ferramenta.item)}</td>
          <td>${escaparHTML(ferramenta.patrimonio || "-")}</td>
          <td>${escaparHTML(ferramenta.descricao)}</td>
          <td>
            <span class="situacao ${classe}">
              ${escaparHTML(situacao)}
            </span>
          </td>
          <td>${escaparHTML(ferramenta.local || "-")}</td>
          <td>${escaparHTML(ferramenta.responsavel || "-")}</td>
        </tr>
      `;
    }).join("");
  }

  if (totalEncontrado === 0) {
    resultadoBusca.textContent = "Nenhuma ferramenta encontrada.";
  } else {
    resultadoBusca.textContent =
      `Exibindo ${formatarNumero(indiceInicial + 1)}–${formatarNumero(indiceFinal)} de ${formatarNumero(totalEncontrado)} ferramentas encontradas.`;
  }

  atualizarControlesPaginacao(totalPaginas, totalEncontrado);
}

// ============================================================
// 8. CONTROLES DA PAGINAÇÃO
// ============================================================

function atualizarControlesPaginacao(totalPaginas, totalEncontrado) {
  const possuiResultados = totalEncontrado > 0;

  // Exibe a página atual e o total de páginas.
  informacaoPaginacao.textContent =
    `Página ${paginaAtual} de ${Math.max(totalPaginas, 1)}`;

  // Desabilita os botões quando não há páginas ou quando
  // o usuário já está no início ou no fim da navegação.
  botaoPrimeira.disabled = !possuiResultados || paginaAtual === 1;
  botaoAnterior.disabled = !possuiResultados || paginaAtual === 1;

  botaoProxima.disabled =
    !possuiResultados || paginaAtual >= totalPaginas;

  botaoUltima.disabled =
    !possuiResultados || paginaAtual >= totalPaginas;
}

// ============================================================
// 9. EVENTOS DOS FILTROS E DA PAGINAÇÃO
// ============================================================

// Atualiza a pesquisa quando o usuário digita.
campoBusca.addEventListener("input", atualizarTabela);

// Atualiza a pesquisa quando um filtro é alterado.
filtroSituacao.addEventListener("change", atualizarTabela);
filtroLocal.addEventListener("change", atualizarTabela);
filtroResponsavel.addEventListener("change", atualizarTabela);

// Altera a quantidade de registros exibidos por página.
itensPorPagina.addEventListener("change", () => {
  quantidadePorPagina = Number(itensPorPagina.value);
  paginaAtual = 1;

  atualizarTabela();
});

// Navega para a primeira página.
botaoPrimeira.addEventListener("click", () => {
  paginaAtual = 1;
  atualizarTabelaSemReiniciarPagina();
});

// Navega para a página anterior.
botaoAnterior.addEventListener("click", () => {
  if (paginaAtual > 1) {
    paginaAtual--;
    atualizarTabelaSemReiniciarPagina();
  }
});

// Navega para a próxima página.
botaoProxima.addEventListener("click", () => {
  paginaAtual++;
  atualizarTabelaSemReiniciarPagina();
});

// Navega para a última página.
botaoUltima.addEventListener("click", () => {
  const totalEncontrado = obterFerramentasFiltradas().length;

  paginaAtual = Math.max(
    Math.ceil(totalEncontrado / quantidadePorPagina),
    1
  );

  atualizarTabelaSemReiniciarPagina();
});

// ============================================================
// 10. NAVEGAÇÃO SEM REINICIAR A PÁGINA ATUAL
// ============================================================

// Retorna os registros que correspondem aos filtros atuais.
// Essa função é usada pela navegação para não voltar à página 1.

function obterFerramentasFiltradas() {
  const termo = normalizar(campoBusca.value);

  const situacaoSelecionada = normalizar(filtroSituacao.value);
  const localSelecionado = normalizar(filtroLocal.value);
  const responsavelSelecionado = normalizar(filtroResponsavel.value);

  // Identifica se a pesquisa contém somente números.
  const pesquisaNumerica = /^\d+$/.test(termo);

  return ferramentas.filter(ferramenta => {
    const patrimonio = normalizar(ferramenta.patrimonio);

    let correspondeBusca;

    if (termo === "") {
      correspondeBusca = true;

    } else if (pesquisaNumerica) {
      // Pesquisa numérica: considera exclusivamente o patrimônio.
      // Não pesquisa item, descrição ou outros campos.
      correspondeBusca = patrimonio === termo;

    } else {
      // Pesquisa textual: procura nos campos disponíveis.
      const textoCompleto = normalizar([
        ferramenta.item,
        ferramenta.patrimonio,
        ferramenta.descricao,
        ferramenta.ordemCompra,
        ferramenta.situacao,
        ferramenta.local,
        ferramenta.responsavel,
        ferramenta.supervisor
      ].join(" "));

      correspondeBusca = textoCompleto.includes(termo);
    }

    const correspondeSituacao =
      !situacaoSelecionada ||
      normalizar(obterSituacao(ferramenta)) === situacaoSelecionada;

    const correspondeLocal =
      !localSelecionado ||
      normalizar(obterTexto(ferramenta.local) || "NÃO INFORMADO") === localSelecionado;

    const correspondeResponsavel =
      !responsavelSelecionado ||
      normalizar(obterTexto(ferramenta.responsavel) || "NÃO INFORMADO") === responsavelSelecionado;

    return (
      correspondeBusca &&
      correspondeSituacao &&
      correspondeLocal &&
      correspondeResponsavel
    );
  });
}

// ============================================================
// 11. EXPORTAÇÃO DO INVENTÁRIO PARA EXCEL
// ============================================================

async function exportarExcel() {
  try {
    // Recupera todos os registros que atendem à pesquisa e aos filtros.
    // A paginação não é aplicada aqui: o relatório inclui todas as páginas.
    const registros = obterFerramentasFiltradas();

    if (registros.length === 0) {
      alert("Não existem ferramentas para exportar com os filtros atuais.");
      return;
    }

    // ExcelJS é carregado pelo HTML antes deste arquivo JavaScript.
    if (typeof ExcelJS === "undefined") {
      throw new Error(
        "A biblioteca ExcelJS não foi carregada. Verifique a conexão com a internet e a inclusão do script no index.html."
      );
    }

    const workbook = new ExcelJS.Workbook();

    workbook.creator = "Sistema de Almoxarifado";
    workbook.subject = "Relatório de ferramentas cadastradas";
    workbook.title = "Inventário de Ferramentas";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Inventário", {
      views: [{ state: "frozen", ySplit: 1 }]
    });

    // Define as colunas que serão incluídas no relatório.
    worksheet.columns = [
      { header: "Item", key: "item", width: 12 },
      { header: "Patrimônio", key: "patrimonio", width: 18 },
      { header: "Descrição", key: "descricao", width: 42 },
      { header: "Ordem de Compra", key: "ordemCompra", width: 20 },
      { header: "Situação", key: "situacao", width: 22 },
      { header: "Local", key: "local", width: 28 },
      { header: "Responsável", key: "responsavel", width: 30 },
      { header: "Supervisor", key: "supervisor", width: 30 }
    ];

    // Mantém patrimônios como texto para preservar zeros à esquerda e hífens.
    registros.forEach(ferramenta => {
      worksheet.addRow({
        item: obterTexto(ferramenta.item),
        patrimonio: obterTexto(ferramenta.patrimonio) || "-",
        descricao: obterTexto(ferramenta.descricao),
        ordemCompra: obterTexto(ferramenta.ordemCompra),
        situacao: obterSituacao(ferramenta),
        local: obterTexto(ferramenta.local) || "-",
        responsavel: obterTexto(ferramenta.responsavel) || "-",
        supervisor: obterTexto(ferramenta.supervisor) || "-"
      });
    });

    // Formata o cabeçalho do relatório.
    const cabecalho = worksheet.getRow(1);

    cabecalho.font = {
      bold: true,
      color: { argb: "FFFFFFFF" }
    };

    cabecalho.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF176B45" }
    };

    cabecalho.alignment = {
      vertical: "middle",
      horizontal: "center"
    };

    cabecalho.height = 25;

    // Ativa os filtros no cabeçalho e mantém a primeira linha fixa.
    worksheet.autoFilter = {
      from: "A1",
      to: "H1"
    };

    // Alinha os dados verticalmente e permite quebra de linha nas descrições.
    worksheet.eachRow((linha, numeroLinha) => {
      if (numeroLinha > 1) {
        linha.alignment = {
          vertical: "middle",
          wrapText: true
        };
      }
    });

    const conteudo = await workbook.xlsx.writeBuffer();

    const arquivo = new Blob([conteudo], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    // Gera um nome único com data e horário locais.
    const agora = new Date();

    const data = agora.toLocaleDateString("pt-BR").replace(/\//g, "-");

    const horario = agora.toTimeString().slice(0, 8).replace(/:/g, "-");

    const nomeArquivo = `inventario-ferramentas-${data}-${horario}.xlsx`;

    // Inicia o download do arquivo gerado.
    const url = URL.createObjectURL(arquivo);

    const link = document.createElement("a");

    link.href = url;
    link.download = nomeArquivo;

    document.body.appendChild(link);

    link.click();
    link.remove();

    // Libera a URL temporária após o navegador iniciar o download.
    setTimeout(() => URL.revokeObjectURL(url), 1000);

    console.info(`${registros.length} ferramentas exportadas para ${nomeArquivo}.`);

  } catch (erro) {
    console.error("Erro ao exportar Excel:", erro);

    alert(
      "Não foi possível gerar o relatório Excel. Verifique a conexão com a internet e o console do navegador."
    );
  }
}

// Registra o evento do botão, caso ele exista no HTML.
if (botaoExportar) {
  botaoExportar.addEventListener("click", exportarExcel);
  botaoExportar.disabled = true;
}

// Atualiza a tabela preservando a página escolhida na navegação.
function atualizarTabelaSemReiniciarPagina() {
  renderizarTabela(obterFerramentasFiltradas());
}

// ============================================================
// 13. INICIALIZAÇÃO
// ============================================================

// Busca os dados assim que a página é carregada.
carregarFerramentas();