
// ============================================================
// 1. IMPORTAÇÕES
// ============================================================

// Importa os recursos necessários do Node.js.
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { extname, resolve, sep } from "node:path";

// Importa a biblioteca responsável pela leitura de arquivos Excel.
import ExcelJS from "exceljs";


// ============================================================
// 2. CONFIGURAÇÕES DO SERVIDOR
// ============================================================

// Define o diretório principal do projeto.
const diretorioProjeto = resolve(
  fileURLToPath(new URL(".", import.meta.url))
);

// Define o caminho da planilha Excel.

const caminhoPlanilha = resolve(
  diretorioProjeto,
  "..",
  "CONTROLE DE FERRAMENTAS.xlsx"
);

// Define a porta do servidor.
const porta = Number(process.env.PORT || 3000);

// Define os tipos de conteúdo dos arquivos disponibilizados.
const tiposDeConteudo = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};


// ============================================================
// 3. FUNÇÃO PARA RESPONDER EM JSON
// ============================================================

// Envia uma resposta no formato JSON.
function responderJson(resposta, codigo, dados) {
  resposta.writeHead(codigo, {
    "Content-Type": "application/json; charset=utf-8"
  });

  resposta.end(JSON.stringify(dados));
}


// ============================================================
// 4. CONVERSÃO DOS DADOS DA PLANILHA
// ============================================================

// Converte o conteúdo de uma célula em texto.
function converterParaTexto(linha, coluna) {
  const valor = linha.getCell(coluna).text;

  return valor === null || valor === undefined
    ? ""
    : valor.trim();
}


// ============================================================
// 5. LEITURA DA PLANILHA EXCEL
// ============================================================

// Lê todas as ferramentas diretamente da planilha Excel.
async function lerFerramentasDaPlanilha() {
  console.log("Iniciando leitura da planilha com ExcelJS...");

  // Cria uma instância do ExcelJS.
  const workbook = new ExcelJS.Workbook();

  // Abre o arquivo XLSX diretamente, sem iniciar o Microsoft Excel.
  await workbook.xlsx.readFile(caminhoPlanilha);

  // Localiza a aba CONTROLE.
  const worksheet = workbook.getWorksheet("CONTROLE");

  // Verifica se a aba existe.
  if (!worksheet) {
    throw new Error("A aba CONTROLE não foi encontrada na planilha.");
  }

  // Exibe informações para diagnóstico.
  console.log("Total de linhas:", worksheet.rowCount);
  console.log("Linhas efetivamente preenchidas:", worksheet.actualRowCount);

  // Armazena todos os registros encontrados.
  const ferramentas = [];


// ======================================================
// CONFERÊNCIA INDEPENDENTE DA QUANTIDADE DE ITENS
// Conta os registros diretamente nas células do Excel.
// ======================================================

let totalColunaA = 0;
let totalColunasABC = 0;
let totalDesdeLinha4 = 0;
let totalDesdeLinha1 = 0;

const registrosEncontrados = [];

for (let numeroLinha = 1; numeroLinha <= worksheet.rowCount; numeroLinha++) {
    const linha = worksheet.getRow(numeroLinha);

    const item = linha.getCell(1).text.trim();
    const patrimonio = linha.getCell(2).text.trim();
    const descricao = linha.getCell(3).text.trim();

    const temItem = item !== "";
    const temIdentificacao = Boolean(item || patrimonio || descricao);

    if (temItem) totalColunaA++;

    if (temIdentificacao) {
        totalDesdeLinha1++;

        if (numeroLinha >= 4) {
            totalDesdeLinha4++;
        }

        registrosEncontrados.push({
            linha: numeroLinha,
            item,
            patrimonio,
            descricao
        });
    }

    if (numeroLinha >= 4 && temIdentificacao) {
        totalColunasABC++;
    }
}

console.log("\n========== CONFERÊNCIA DOS REGISTROS ==========");
console.log("Registros com item na coluna A:", totalColunaA);
console.log("Registros com identificação desde a linha 1:", totalDesdeLinha1);
console.log("Registros com identificação desde a linha 4:", totalDesdeLinha4);
console.log("Total de registros A/B/C a partir da linha 4:", totalColunasABC);
console.log("Primeiras linhas encontradas:", registrosEncontrados.slice(0, 5));
console.log("Últimas linhas encontradas:", registrosEncontrados.slice(-5));
console.log("===============================================\n");


console.log("\n--- VERIFICAÇÃO DA SEQUÊNCIA DOS ITENS ---");

const itensNumericos = [];

for (let numeroLinha = 4; numeroLinha <= worksheet.rowCount; numeroLinha++) {
    const valor = worksheet.getRow(numeroLinha).getCell(1).text.trim();
    const numero = Number(valor);

    if (valor !== "" && Number.isInteger(numero)) {
        itensNumericos.push(numero);
    }
}

const itensExistentes = new Set(itensNumericos);
const itensAusentes = [];

for (let numero = 1; numero <= 898; numero++) {
    if (!itensExistentes.has(numero)) {
        itensAusentes.push(numero);
    }
}

console.log("Quantidade de itens numéricos:", itensNumericos.length);
console.log("Primeiro item:", Math.min(...itensNumericos));
console.log("Último item:", Math.max(...itensNumericos));
console.log("Itens ausentes:", itensAusentes);
console.log("------------------------------------------\n");


console.log("\n--- ANÁLISE DE DUPLICIDADES E SALTOS ---");

const ocorrencias = new Map();

for (let numeroLinha = 4; numeroLinha <= worksheet.rowCount; numeroLinha++) {
    const linha = worksheet.getRow(numeroLinha);

    const item = linha.getCell(1).text.trim();
    const patrimonio = linha.getCell(2).text.trim();
    const descricao = linha.getCell(3).text.trim();

    if (!/^\d+$/.test(item)) continue;

    const numero = Number(item);

    if (!ocorrencias.has(numero)) {
        ocorrencias.set(numero, []);
    }

    ocorrencias.get(numero).push({
        linha: numeroLinha,
        item,
        patrimonio,
        descricao
    });
}

// Identifica itens duplicados.
const duplicados = [];

for (const [numero, registros] of ocorrencias) {
    if (registros.length > 1) {
        duplicados.push({ numero, registros });
    }
}

console.log("Itens duplicados:", duplicados);


console.log("\n--- INSPEÇÃO DAS LINHAS FÍSICAS ---");

const intervalos = [
    { inicio: 510, fim: 522 },
    { inicio: 628, fim: 635 }
];

for (const intervalo of intervalos) {
    console.log(`\nVerificando linhas ${intervalo.inicio} até ${intervalo.fim}`);

    for (
        let numeroLinha = intervalo.inicio;
        numeroLinha <= intervalo.fim;
        numeroLinha++
    ) {
        const linha = worksheet.getRow(numeroLinha);

        const dados = [];

        for (let coluna = 1; coluna <= 11; coluna++) {
            dados.push(linha.getCell(coluna).text.trim());
        }

        console.log({
            linha: numeroLinha,
            oculta: linha.hidden,
            item: dados[0],
            patrimonio: dados[1],
            descricao: dados[2],
            demaisColunas: dados.slice(3)
        });
    }
}

console.log("--- FIM DA INSPEÇÃO ---\n");


// Mostra os registros próximos aos saltos identificados.
const numerosAusentes = [513, 515, 519, 520, 521, 522, 523, 636];

for (const numero of numerosAusentes) {
    console.log(`\nVerificando proximidade do item ${numero}:`);

    for (const vizinho of [numero - 1, numero + 1]) {
        const registros = ocorrencias.get(vizinho) || [];

        console.log(`Item ${vizinho}:`, registros);
    }
}

console.log("\nTotal de números distintos:", ocorrencias.size);
console.log("Total de ocorrências numéricas:", 
    [...ocorrencias.values()].reduce((total, registros) => total + registros.length, 0)
);

console.log("--- FIM DA ANÁLISE ---\n");


  // Percorre todas as linhas preenchidas da planilha.
  worksheet.eachRow({ includeEmpty: false }, (linha, numeroLinha) => {

    // Os registros começam na linha 4.
    if (numeroLinha < 4) {
      return;
    }

    // Extrai os dados das colunas correspondentes.
    const item = converterParaTexto(linha, 1);
    const patrimonio = converterParaTexto(linha, 2);
    const descricao = converterParaTexto(linha, 3);

    // Ignora somente linhas sem qualquer identificação.
    // O patrimônio não é obrigatório para cadastrar uma ferramenta.
    if (!item && !patrimonio && !descricao) {
      return;
    }

    // Adiciona a ferramenta mesmo que não possua patrimônio.
    ferramentas.push({
      item: item,
      patrimonio: patrimonio,
      descricao: descricao,
      ordemCompra: converterParaTexto(linha, 4),
      situacao: converterParaTexto(linha, 7),
      local: converterParaTexto(linha, 8),
      responsavel: converterParaTexto(linha, 9),
      supervisor: converterParaTexto(linha, 10)
    });
  });

  
  // Considera patrimônio válido somente quando o campo
  // não está vazio e não contém apenas um hífen.
  const comPatrimonio = ferramentas.filter(f => {
      const patrimonio = f.patrimonio.trim();

      return patrimonio !== "" && patrimonio !== "-";
  }).length;

  // Todas as ferramentas sem patrimônio.
  const semPatrimonio = ferramentas.length - comPatrimonio;

  // Exibe o resultado da leitura no terminal.
  console.log("--------------------------------------");
  console.log("LEITURA DA PLANILHA CONCLUÍDA");
  console.log("Total de ferramentas:", ferramentas.length);
  console.log("Com patrimônio:", comPatrimonio);
  console.log("Sem patrimônio:", semPatrimonio);
  console.log("--------------------------------------");

  // Retorna todos os registros organizados.
  return {
    total: ferramentas.length,
    comPatrimonio: comPatrimonio,
    semPatrimonio: semPatrimonio,
    ferramentas: ferramentas
  };
}


// ============================================================
// 6. ENTREGA DOS ARQUIVOS ESTÁTICOS
// ============================================================

// Entrega os arquivos HTML, CSS, JavaScript e imagens ao navegador.
async function enviarArquivo(caminhoRequisitado, resposta) {

  // A página inicial será exibida quando o usuário acessar a raiz.
  const caminhoArquivo = caminhoRequisitado === "/"
    ? "/index.html"
    : caminhoRequisitado;

  // Resolve o caminho absoluto do arquivo solicitado.
  const caminhoCompleto = resolve(
    diretorioProjeto,
    `.${caminhoArquivo}`
  );

  // Impede que uma requisição acesse arquivos fora do projeto.
  if (
    caminhoCompleto !== diretorioProjeto &&
    !caminhoCompleto.startsWith(diretorioProjeto + sep)
  ) {
    resposta.writeHead(403);
    resposta.end("Acesso negado.");
    return;
  }

  try {
    // Lê o arquivo solicitado.
    const conteudo = await readFile(caminhoCompleto);

    // Identifica a extensão do arquivo.
    const extensao = extname(caminhoCompleto).toLowerCase();

    // Define o tipo de conteúdo correspondente.
    const tipoConteudo = tiposDeConteudo[extensao]
      || "application/octet-stream";

    // Envia o arquivo ao navegador.
    resposta.writeHead(200, {
      "Content-Type": tipoConteudo
    });

    resposta.end(conteudo);

  } catch (erro) {

    // Trata arquivos que não foram encontrados.
    if (erro.code === "ENOENT" || erro.code === "EISDIR") {
      resposta.writeHead(404, {
        "Content-Type": "text/plain; charset=utf-8"
      });

      resposta.end("Arquivo não encontrado.");
      return;
    }

    // Registra outros erros no terminal.
    console.error("Erro ao enviar arquivo:", erro);

    resposta.writeHead(500, {
      "Content-Type": "text/plain; charset=utf-8"
    });

    resposta.end("Erro interno ao carregar o arquivo.");
  }
}


// ============================================================
// 7. CRIAÇÃO DO SERVIDOR HTTP
// ============================================================

// Cria o servidor responsável por receber as requisições.
const servidor = createServer(async (requisicao, resposta) => {

  try {

    // Identifica a URL solicitada pelo navegador.
    const url = new URL(
      requisicao.url,
      `http://${requisicao.headers.host || "localhost"}`
    );


    // --------------------------------------------------------
    // ROTA: API DE FERRAMENTAS
    // --------------------------------------------------------

    // Endpoint responsável por disponibilizar as ferramentas.
    if (url.pathname === "/api/ferramentas") {

      // Aceita somente requisições GET.
      if (requisicao.method !== "GET") {
        resposta.writeHead(405, {
          "Content-Type": "application/json; charset=utf-8",
          "Allow": "GET"
        });

        resposta.end(JSON.stringify({
          erro: "Método não permitido."
        }));

        return;
      }

      try {

        // Aguarda a leitura completa da planilha.
        const resultado = await lerFerramentasDaPlanilha();

        // Retorna todos os registros em formato JSON.
        responderJson(resposta, 200, resultado);

      } catch (erro) {

        // Registra o erro para facilitar a identificação do problema.
        console.error("Erro ao ler planilha:", erro);

        // Informa ao cliente que ocorreu uma falha.
        responderJson(resposta, 500, {
          erro: "Não foi possível ler a planilha."
        });
      }

      return;
    }


    // --------------------------------------------------------
    // ROTAS DOS ARQUIVOS ESTÁTICOS
    // --------------------------------------------------------

    // Disponibiliza os arquivos estáticos somente por GET.
    if (requisicao.method !== "GET") {
      resposta.writeHead(405, {
        "Content-Type": "text/plain; charset=utf-8",
        "Allow": "GET"
      });

      resposta.end("Método não permitido.");
      return;
    }

    // Decodifica o caminho solicitado.
    const caminhoRequisitado = decodeURIComponent(url.pathname);

    // Entrega o arquivo solicitado ao navegador.
    await enviarArquivo(caminhoRequisitado, resposta);

  } catch (erro) {

    // Trata erros inesperados durante o processamento da requisição.
    console.error("Erro no servidor:", erro);

    if (!resposta.headersSent) {
      responderJson(resposta, 500, {
        erro: "Erro interno do servidor."
      });
    } else {
      resposta.end();
    }
  }
});


// ============================================================
// 8. INICIALIZAÇÃO DO SERVIDOR
// ============================================================

// Inicia o servidor na porta configurada.
servidor.listen(porta, "0.0.0.0", () => {
  console.log("--------------------------------------");
  console.log("Sistema de Almoxarifado iniciado!");
  console.log(`Servidor: http://localhost:${porta}`);
  console.log(`API: http://localhost:${porta}/api/ferramentas`);
  console.log("--------------------------------------");
});
