# Sistema de Almoxarifado

Projeto web para controlar ferramentas e equipamentos do almoxarifado.

> A planilha `CONTROLE DE FERRAMENTAS.xlsx` é a fonte operacional de dados autorizada para o sistema. Antes de qualquer gravação, o sistema criará uma cópia de segurança e registrará a alteração no histórico.

## Objetivo

Construir, passo a passo, um sistema fácil de usar para registrar ferramentas, retiradas, devoluções, manutenções e perdas. Cada colaborador terá um usuário próprio, permitindo saber quem fez cada alteração e quando ela aconteceu.

## O que existe hoje

A primeira tela do sistema exibe o cabeçalho, os indicadores de estoque e uma tabela com dados de exemplo. A busca filtra essa tabela no navegador. A API local para ler a aba `CONTROLE` do Excel também está pronta; a interface ainda será conectada a ela na próxima etapa.

## Tecnologias da primeira etapa

- **HTML:** define os elementos e o conteúdo da página.
- **CSS:** define cores, espaçamentos e o visual.
- **JavaScript:** torna a tela interativa. A busca da tabela usa JavaScript puro.
- **Node.js:** executa o servidor local e disponibiliza a API.
- **PowerShell e Microsoft Excel:** leem o arquivo preservando as estruturas existentes da planilha.

## Executar a aplicação com a planilha

Dentro da pasta do projeto, execute:

```bash
npm run dev
```

Depois, abra [http://localhost:3000](http://localhost:3000). Para conferir somente os dados que a API leu, abra [http://localhost:3000/api/ferramentas](http://localhost:3000/api/ferramentas).

Nesta primeira etapa da integração, a planilha é aberta somente para leitura. Nenhuma alteração é feita no arquivo Excel.

## Estrutura de arquivos

```text
sistema-almoxarifado/
├── index.html       # Estrutura da primeira página
├── css/
│   └── styles.css   # Estilos visuais da página
├── js/
│   └── app.js       # Comportamentos interativos da página
├── scripts/
│   └── ler-planilha.ps1 # Leitura segura do Excel pela automação do Windows
├── server.js        # Servidor web e API local
├── package.json     # Comandos do projeto Node.js
├── README.md        # Documentação do projeto
└── .gitignore       # Arquivos que o Git não deve publicar
```

## Evolução planejada

1. Montar a página inicial e uma tabela de ferramentas com dados de exemplo.
2. Criar interações com JavaScript: busca, filtros e formulários.
3. Conectar a página à API e exibir os dados reais da planilha.
4. Criar gravações seguras no Excel e a aba `HISTORICO`.
5. Implementar login e perfis de acesso.
6. Registrar uma trilha de auditoria: usuário, data, ação, valores anteriores e valores novos.
7. Publicar a aplicação para acesso no computador e no celular.

## Registro de aprendizagem

| Etapa | Conteúdo aprendido | Situação |
| --- | --- | --- |
| 1 | Estrutura HTML e estilos CSS iniciais | Concluída |
| 2 | Painel, cartões de resumo e tabela com dados de exemplo | Concluída |
| 3 | Busca e filtro por situação com JavaScript | Concluída |
| 4 | API local para ler a aba CONTROLE do Excel | Concluída |
| 5 | Atualização do Excel, login e auditoria | Planejada |

## Git e GitHub

O projeto já possui um repositório Git local na branch `main`. Git registra o histórico do código no computador; GitHub guarda esse repositório online.

Os comandos abaixo serão usados a cada melhoria importante:

```bash
git status
git add .
git commit -m "descrição clara da alteração"
git push
```

- `git status` mostra o que mudou.
- `git add .` seleciona os arquivos alterados para o próximo registro.
- `git commit` cria um ponto do histórico no computador.
- `git push` envia os commits para o GitHub, depois que o repositório remoto estiver conectado.

## Princípio de segurança

Senhas, chaves de serviços e dados particulares nunca serão enviados ao GitHub. Quando surgirem, ficarão em arquivos locais ignorados pelo Git.
