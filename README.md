# Sistema de Almoxarifado

Projeto web para controlar ferramentas e equipamentos do almoxarifado.

> A planilha original `CONTROLE DE FERRAMENTAS.xlsx` não será alterada por este projeto. Ela continuará sendo a referência inicial dos dados e poderá ser importada futuramente por meio de uma cópia.

## Objetivo

Construir, passo a passo, um sistema fácil de usar para registrar ferramentas, retiradas, devoluções, manutenções e perdas. Cada colaborador terá um usuário próprio, permitindo saber quem fez cada alteração e quando ela aconteceu.

## O que existe hoje

A primeira tela estática do sistema foi criada com HTML e CSS puros. Ela exibe o cabeçalho, os indicadores de estoque e uma tabela com dados de exemplo. Os dados ainda não são lidos da planilha nem salvos em um banco de dados.

## Tecnologias da primeira etapa

- **HTML:** define os elementos e o conteúdo da página.
- **CSS:** define cores, espaçamentos e o visual.
- **JavaScript:** será incluído na próxima etapa para tornar a tela interativa.

## Estrutura de arquivos

```text
sistema-almoxarifado/
├── index.html       # Estrutura da primeira página
├── css/
│   └── styles.css   # Estilos visuais da página
├── README.md        # Documentação do projeto
└── .gitignore       # Arquivos que o Git não deve publicar
```

## Evolução planejada

1. Montar a página inicial e uma tabela de ferramentas com dados de exemplo.
2. Criar interações com JavaScript: busca, filtros e formulários.
3. Criar uma API e um banco de dados separado da planilha.
4. Implementar login e perfis de acesso.
5. Registrar uma trilha de auditoria: usuário, data, ação, valores anteriores e valores novos.
6. Importar uma cópia dos dados da planilha e criar exportações.
7. Publicar a aplicação para acesso no computador e no celular.

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
