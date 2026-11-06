
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
const diretorioProjeto = fileURLToPath(new URL(".", import.meta.url));

// Define o caminho da planilha Excel.
// A planilha está localizada uma pasta acima do projeto.
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

  // Calcula quantas ferramentas possuem e não possuem patrimônio.
  const comPatrimonio = ferramentas.filter(
    ferramenta => ferramenta.patrimonio !== ""
  ).length;

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
servidor.listen(porta, () => {
  console.log("--------------------------------------");
  console.log("Sistema de Almoxarifado iniciado!");
  console.log(`Servidor: http://localhost:${porta}`);
  console.log(`API: http://localhost:${porta}/api/ferramentas`);
  console.log("--------------------------------------");
});
