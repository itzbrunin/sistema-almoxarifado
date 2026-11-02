import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { extname, resolve } from "node:path";

const diretorioProjeto = fileURLToPath(new URL(".", import.meta.url));
const caminhoPlanilha = resolve(diretorioProjeto, "..", "CONTROLE DE FERRAMENTAS.xlsx");
const caminhoScriptLeitura = resolve(diretorioProjeto, "scripts", "ler-planilha.ps1");
const porta = Number(process.env.PORT || 3000);

const tiposDeConteudo = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8"
};

function responderJson(resposta, codigo, dados) {
  resposta.writeHead(codigo, { "Content-Type": "application/json; charset=utf-8" });
  resposta.end(JSON.stringify(dados));
}

function lerFerramentasDaPlanilha() {
  return new Promise((resolvePromise, rejectPromise) => {
    const processo = spawn(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        caminhoScriptLeitura,
        "-CaminhoPlanilha",
        caminhoPlanilha
      ],
      { windowsHide: true }
    );

    let saida = "";
    let erros = "";

    processo.stdout.on("data", (dados) => {
      saida += dados;
    });

    processo.stderr.on("data", (dados) => {
      erros += dados;
    });

    processo.on("error", rejectPromise);

    processo.on("close", (codigo) => {
      if (codigo !== 0) {
        rejectPromise(new Error(erros || "Não foi possível ler a planilha."));
        return;
      }

      try {
        resolvePromise(JSON.parse(saida));
      } catch {
        rejectPromise(new Error("A planilha retornou dados em um formato inválido."));
      }
    });
  });
}

async function enviarArquivo(resposta, caminhoRequisitado) {
  const caminhosPermitidos = ["/", "/index.html"];
  const ehArquivoDeEstilo = caminhoRequisitado.startsWith("/css/");
  const ehArquivoDeScript = caminhoRequisitado.startsWith("/js/");

  if (!caminhosPermitidos.includes(caminhoRequisitado) && !ehArquivoDeEstilo && !ehArquivoDeScript) {
    resposta.writeHead(404);
    resposta.end("Arquivo não encontrado.");
    return;
  }

  const caminhoRelativo = caminhoRequisitado === "/" ? "index.html" : caminhoRequisitado.slice(1);
  const caminhoArquivo = resolve(diretorioProjeto, caminhoRelativo);

  try {
    const conteudo = await readFile(caminhoArquivo);
    const tipoDeConteudo = tiposDeConteudo[extname(caminhoArquivo)] || "application/octet-stream";

    resposta.writeHead(200, { "Content-Type": tipoDeConteudo });
    resposta.end(conteudo);
  } catch {
    resposta.writeHead(404);
    resposta.end("Arquivo não encontrado.");
  }
}

const servidor = createServer(async (requisicao, resposta) => {
  const url = new URL(requisicao.url, `http://${requisicao.headers.host}`);

  if (requisicao.method === "GET" && url.pathname === "/api/ferramentas") {
    try {
      const dados = await lerFerramentasDaPlanilha();
      responderJson(resposta, 200, dados);
    } catch (erro) {
      console.error(erro);
      responderJson(resposta, 500, { erro: "Não foi possível ler a planilha." });
    }

    return;
  }

  if (requisicao.method === "GET") {
    await enviarArquivo(resposta, url.pathname);
    return;
  }

  responderJson(resposta, 405, { erro: "Método não permitido." });
});

servidor.listen(porta, () => {
  console.log(`Servidor em execução: http://localhost:${porta}`);
});
