const campoBusca = document.querySelector("#campo-busca");
const filtroSituacao = document.querySelector("#filtro-situacao");
const linhasFerramentas = document.querySelectorAll("#corpo-ferramentas tr");
const resultadoBusca = document.querySelector("#resultado-busca");

function atualizarTabela() {
  const termoBuscado = campoBusca.value.trim().toLocaleLowerCase("pt-BR");
  const situacaoSelecionada = filtroSituacao.value;
  let quantidadeEncontrada = 0;

  linhasFerramentas.forEach((linha) => {
    const textoDaLinha = linha.textContent.toLocaleLowerCase("pt-BR");
    const correspondeAoTermo = textoDaLinha.includes(termoBuscado);
    const correspondeASituacao =
      situacaoSelecionada === "" || linha.dataset.situacao === situacaoSelecionada;
    const ferramentaEncontrada = correspondeAoTermo && correspondeASituacao;

    linha.hidden = !ferramentaEncontrada;

    if (ferramentaEncontrada) {
      quantidadeEncontrada += 1;
    }
  });

  const mensagemResultado =
    quantidadeEncontrada === 1
      ? "1 ferramenta encontrada."
      : `${quantidadeEncontrada} ferramentas encontradas.`;

  resultadoBusca.textContent = mensagemResultado;
}

campoBusca.addEventListener("input", atualizarTabela);
filtroSituacao.addEventListener("change", atualizarTabela);
