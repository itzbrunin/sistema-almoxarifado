
// ============================================================
// 1. IMPORTAÇÕES
// ============================================================

// Recursos nativos do Node.js.
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { extname, resolve, sep } from "node:path";

// Biblioteca para leitura de arquivos Excel.
import ExcelJS from "exceljs";


// ============================================================
// 2. CONFIGURAÇÕES DO SERVIDOR
// ============================================================

// Diretório principal do projeto.
const diretorioProjeto = resolve(
  fileURLToPath(new URL(".", import.meta.url))
);

// Porta definida pelo ambiente de hospedagem.
const porta = Number(process.env.PORT || 3000);

// Configurações do Supabase.
const supabaseUrl = process.env.SUPABASE_URL?.replace(/\/+$/, "");
const supabaseSecretKey = process.env.SUPABASE_SECRET_KEY;
const supabaseBucket = process.env.SUPABASE_BUCKET;
const supabaseArquivo = process.env.SUPABASE_FILE;

// Verifica quantas configurações do Supabase foram informadas.
const configuracoesSupabase = [
  supabaseUrl,
  supabaseSecretKey,
  supabaseBucket,
  supabaseArquivo
];

const quantidadeConfiguracoes = configuracoesSupabase.filter(Boolean).length;

// Se alguma variável foi informada, todas devem estar configuradas.
if (
  quantidadeConfiguracoes > 0 &&
  quantidadeConfiguracoes !== configuracoesSupabase.length
) {
  throw new Error(
    "Configuração incompleta do Supabase. Verifique as quatro variáveis de ambiente."
  );
}

// Define a origem dos dados.
const utilizarSupabase = quantidadeConfiguracoes === 4;

// Caminho local utilizado durante o desenvolvimento no Windows.
const caminhoPlanilhaLocal = process.env.PLANILHA_PATH
  ? resolve(process.env.PLANILHA_PATH)
  : resolve(
      diretorioProjeto,
      "..",
      "CONTROLE DE FERRAMENTAS.xlsx"
    );

// Cache dos dados da planilha.
const TEMPO_CACHE = 5 * 60 * 1000;

let dadosCacheados = null;
let cacheExpiraEm = 0;
let leituraEmAndamento = null;

// Ativa diagnósticos detalhados somente quando solicitado.
const debugInventario =
  process.env.DEBUG_INVENTARIO === "true";

// Tipos de conteúdo dos arquivos estáticos.
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

console.log("--------------------------------------");
console.log("CONFIGURAÇÃO DO ALMOXARIFADO");
console.log("Armazenamento:", utilizarSupabase ? "Supabase Storage" : "Arquivo local");
console.log("Cache:", `${TEMPO_CACHE / 1000} segundos`);
console.log("Modo de diagnóstico:", debugInventario ? "Ativado" : "Desativado");

if (!utilizarSupabase) {
  console.log("Caminho local:", caminhoPlanilhaLocal);
}

console.log("--------------------------------------");


// ============================================================
// 3. RESPOSTAS JSON
// ============================================================

// Envia respostas padronizadas em JSON.
function responderJson(resposta, codigo, dados) {
  resposta.writeHead(codigo, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
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
    : String(valor).trim();
}


// ============================================================
// 5. DOWNLOAD DA PLANILHA NO SUPABASE
// ============================================================

// Baixa a planilha de um bucket privado do Supabase.
async function baixarPlanilhaSupabase() {

  console.log("Iniciando download da planilha no Supabase...");

  // Codifica o nome do bucket.
  const bucketCodificado = encodeURIComponent(supabaseBucket);

  // Codifica cada parte do caminho do arquivo.
  const arquivoCodificado = supabaseArquivo
    .split("/")
    .map(parte => encodeURIComponent(parte))
    .join("/");

  // Monta o endereço da API de Storage.
  const urlArquivo =
    `${supabaseUrl}/storage/v1/object/${bucketCodificado}/${arquivoCodificado}`;

  // Faz o download utilizando as credenciais privadas.
  const resposta = await fetch(urlArquivo, {
    method: "GET",
    headers: {
      "apikey": supabaseSecretKey,
      "Authorization": `Bearer ${supabaseSecretKey}`
    },
    signal: AbortSignal.timeout(30000)
  });

  // Verifica se o download foi concluído.
  if (!resposta.ok) {
    const detalheErro = (await resposta.text()).slice(0, 500);

    console.error(
      "Falha no download do Supabase:",
      resposta.status,
      detalheErro
    );

    throw new Error(
      `Não foi possível baixar a planilha do Supabase. HTTP ${resposta.status}.`
    );
  }

  // Converte o arquivo recebido em um Buffer.
  const arquivo = Buffer.from(
    await resposta.arrayBuffer()
  );

  // Verifica se o arquivo não está vazio.
  if (arquivo.length === 0) {
    throw new Error("O arquivo recebido do Supabase está vazio.");
  }

  console.log("Download concluído com sucesso.");
  console.log("Tamanho do arquivo:", arquivo.length, "bytes");

  return arquivo;
}


// ============================================================
// 6. CARREGAMENTO DO EXCEL
// ============================================================

// Carrega a planilha na memória.
async function carregarWorkbook() {

  const workbook = new ExcelJS.Workbook();

  if (utilizarSupabase) {

    // Obtém o arquivo privado do Supabase.
    const arquivo = await baixarPlanilhaSupabase();

    // Interpreta o XLSX diretamente na memória.
    await workbook.xlsx.load(arquivo);

  } else {

    // Durante o desenvolvimento, utiliza o arquivo local.
    console.log("Lendo planilha local:", caminhoPlanilhaLocal);

    await workbook.xlsx.readFile(caminhoPlanilhaLocal);
  }

  return workbook;
}


// ============================================================
// 7. DIAGNÓSTICO OPCIONAL DO INVENTÁRIO
// ============================================================

// Executa análises detalhadas somente quando DEBUG_INVENTARIO=true.
function diagnosticarInventario(worksheet) {

  console.log("\n========== DIAGNÓSTICO DO INVENTÁRIO ==========");

  const ocorrencias = new Map();

  let registrosIdentificados = 0;

  for (
    let numeroLinha = 4;
    numeroLinha <= worksheet.rowCount;
    numeroLinha++
  ) {

    const linha = worksheet.getRow(numeroLinha);

    const item = converterParaTexto(linha, 1);
    const patrimonio = converterParaTexto(linha, 2);
    const descricao = converterParaTexto(linha, 3);

    if (item || patrimonio || descricao) {
      registrosIdentificados++;
    }

    if (!/^\d+$/.test(item)) {
      continue;
    }

    const numero = Number(item);

    if (!ocorrencias.has(numero)) {
      ocorrencias.set(numero, []);
    }

    ocorrencias.get(numero).push({
      linha: numeroLinha,
      patrimonio,
      descricao
    });
  }

  // Identifica itens ausentes.
  const itensAusentes = [];

  for (let numero = 1; numero <= 898; numero++) {
    if (!ocorrencias.has(numero)) {
      itensAusentes.push(numero);
    }
  }

  // Identifica duplicidades.
  const duplicados = [];

  for (const [numero, registros] of ocorrencias) {
    if (registros.length > 1) {
      duplicados.push({
        numero,
        registros
      });
    }
  }

  console.log("Registros identificados:", registrosIdentificados);
  console.log("Itens numéricos distintos:", ocorrencias.size);
  console.log("Itens ausentes:", itensAusentes);
  console.log("Itens duplicados:", duplicados);
  console.log("================================================\n");
}


// ============================================================
// 8. LEITURA E ORGANIZAÇÃO DAS FERRAMENTAS
// ============================================================

// Interpreta a planilha e organiza os registros da API.
async function lerFerramentasDaPlanilha() {

  console.log("Iniciando leitura da planilha com ExcelJS...");

  // Carrega o Excel do Supabase ou do arquivo local.
  const workbook = await carregarWorkbook();

  // Localiza a aba CONTROLE.
  const worksheet = workbook.getWorksheet("CONTROLE");

  if (!worksheet) {
    throw new Error(
      "A aba CONTROLE não foi encontrada na planilha."
    );
  }

  console.log("Total de linhas:", worksheet.rowCount);
  console.log(
    "Linhas efetivamente preenchidas:",
    worksheet.actualRowCount
  );

  // Executa diagnósticos adicionais somente quando necessário.
  if (debugInventario) {
    diagnosticarInventario(worksheet);
  }

  const ferramentas = [];

  // Percorre todas as linhas preenchidas.
  worksheet.eachRow(
    { includeEmpty: false },
    (linha, numeroLinha) => {

      // Os registros começam na linha 4.
      if (numeroLinha < 4) {
        return;
      }

      // Extrai as informações das colunas.
      const item = converterParaTexto(linha, 1);
      const patrimonio = converterParaTexto(linha, 2);
      const descricao = converterParaTexto(linha, 3);

      // Ignora linhas completamente sem identificação.
      if (!item && !patrimonio && !descricao) {
        return;
      }

      // Mantém ferramentas mesmo sem patrimônio.
      ferramentas.push({
        item,
        patrimonio,
        descricao,
        ordemCompra: converterParaTexto(linha, 4),
        situacao: converterParaTexto(linha, 7),
        local: converterParaTexto(linha, 8),
        responsavel: converterParaTexto(linha, 9),
        supervisor: converterParaTexto(linha, 10)
      });
    }
  );

  // Conta ferramentas com patrimônio válido.
  const comPatrimonio = ferramentas.filter(ferramenta => {

    const patrimonio = ferramenta.patrimonio.trim();

    return patrimonio !== "" && patrimonio !== "-";

  }).length;

  // Calcula ferramentas sem patrimônio.
  const semPatrimonio =
    ferramentas.length - comPatrimonio;

  console.log("--------------------------------------");
  console.log("LEITURA DA PLANILHA CONCLUÍDA");
  console.log("Total de ferramentas:", ferramentas.length);
  console.log("Com patrimônio:", comPatrimonio);
  console.log("Sem patrimônio:", semPatrimonio);
  console.log("--------------------------------------");

  // Retorna os dados no formato utilizado pelo frontend.
  return {
    total: ferramentas.length,
    comPatrimonio,
    semPatrimonio,
    ferramentas
  };
}


// ============================================================
// 9. CACHE DA PLANILHA
// ============================================================

// Evita baixar e interpretar a planilha em todas as requisições.
async function obterDadosFerramentas() {

  // Verifica se existe cache válido.
  if (
    dadosCacheados &&
    Date.now() < cacheExpiraEm
  ) {
    console.log("Utilizando dados do cache.");

    return dadosCacheados;
  }

  // Se já existe uma leitura em andamento, reutiliza a mesma operação.
  if (leituraEmAndamento) {
    return leituraEmAndamento;
  }

  // Inicia uma nova leitura.
  leituraEmAndamento = (async () => {

    const dados = await lerFerramentasDaPlanilha();

    // Atualiza o cache somente após uma leitura bem-sucedida.
    dadosCacheados = dados;
    cacheExpiraEm = Date.now() + TEMPO_CACHE;

    return dados;

  })();

  try {

    return await leituraEmAndamento;

  } catch (erro) {

    // Não mantém resultados inválidos no cache.
    dadosCacheados = null;
    cacheExpiraEm = 0;

    throw erro;

  } finally {

    leituraEmAndamento = null;
  }
}


// ============================================================
// 10. ENTREGA DOS ARQUIVOS ESTÁTICOS
// ============================================================

// Entrega HTML, CSS, JavaScript e imagens.
async function enviarArquivo(caminhoRequisitado, resposta) {

  const caminhoArquivo = caminhoRequisitado === "/"
    ? "/index.html"
    : caminhoRequisitado;

  const caminhoCompleto = resolve(
    diretorioProjeto,
    `.${caminhoArquivo}`
  );

  // Impede acesso a arquivos externos ao projeto.
  if (
    caminhoCompleto !== diretorioProjeto &&
    !caminhoCompleto.startsWith(diretorioProjeto + sep)
  ) {
    resposta.writeHead(403);
    resposta.end("Acesso negado.");
    return;
  }

  try {

    const conteudo = await readFile(caminhoCompleto);

    const extensao = extname(caminhoCompleto).toLowerCase();

    const tipoConteudo =
      tiposDeConteudo[extensao] ||
      "application/octet-stream";

    resposta.writeHead(200, {
      "Content-Type": tipoConteudo
    });

    resposta.end(conteudo);

  } catch (erro) {

    if (
      erro.code === "ENOENT" ||
      erro.code === "EISDIR"
    ) {
      resposta.writeHead(404, {
        "Content-Type": "text/plain; charset=utf-8"
      });

      resposta.end("Arquivo não encontrado.");
      return;
    }

    console.error("Erro ao enviar arquivo:", erro);

    resposta.writeHead(500, {
      "Content-Type": "text/plain; charset=utf-8"
    });

    resposta.end("Erro interno ao carregar o arquivo.");
  }
}


// ============================================================
// 11. CRIAÇÃO DO SERVIDOR HTTP
// ============================================================

const servidor = createServer(async (requisicao, resposta) => {

  try {

    const url = new URL(
      requisicao.url,
      `http://${requisicao.headers.host || "localhost"}`
    );


    // --------------------------------------------------------
    // ROTA DE SAÚDE DO SERVIDOR
    // --------------------------------------------------------

    if (
      requisicao.method === "GET" &&
      url.pathname === "/health"
    ) {

      responderJson(resposta, 200, {
        status: "ok",
        armazenamento: utilizarSupabase
          ? "Supabase Storage"
          : "Local"
      });

      return;
    }


    // --------------------------------------------------------
    // API DE FERRAMENTAS
    // --------------------------------------------------------

    if (url.pathname === "/api/ferramentas") {

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

        // Obtém os dados da planilha ou do cache.
        const resultado = await obterDadosFerramentas();

        responderJson(resposta, 200, resultado);

      } catch (erro) {

        console.error("Erro ao ler planilha:", erro);

        responderJson(resposta, 500, {
          erro: "Não foi possível carregar os dados do inventário."
        });
      }

      return;
    }


    // --------------------------------------------------------
    // ARQUIVOS ESTÁTICOS
    // --------------------------------------------------------

    if (requisicao.method !== "GET") {

      resposta.writeHead(405, {
        "Content-Type": "text/plain; charset=utf-8",
        "Allow": "GET"
      });

      resposta.end("Método não permitido.");
      return;
    }

    const caminhoRequisitado = decodeURIComponent(url.pathname);

    await enviarArquivo(caminhoRequisitado, resposta);

  } catch (erro) {

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
// 12. INICIALIZAÇÃO DO SERVIDOR
// ============================================================

servidor.listen(porta, "0.0.0.0", () => {

  console.log("--------------------------------------");
  console.log("Sistema de Almoxarifado iniciado!");
  console.log(`Servidor: http://localhost:${porta}`);
  console.log(`API: http://localhost:${porta}/api/ferramentas`);
  console.log(`Saúde: http://localhost:${porta}/health`);
  console.log("Armazenamento:", utilizarSupabase ? "Supabase" : "Local");
  console.log("--------------------------------------");

});
