const campoBusca = document.querySelector("#campo-busca");
const linhasFerramentas = document.querySelectorAll("#corpo-ferramentas tr");
const resultadoBusca = document.querySelector("#resultado-busca");

function atualizarTabela() {
  const termoBuscado = campoBusca.value.trim().toLocaleLowerCase("pt-BR");
  let quantidadeEncontrada = 0;

  linhasFerramentas.forEach((linha) => {
    const textoDaLinha = linha.textContent.toLocaleLowerCase("pt-BR");
    const ferramentaEncontrada = textoDaLinha.includes(termoBuscado);

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
