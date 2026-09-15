# Melhorias e gaps do DocGen MCP Server

Este documento reúne os pontos identificados durante a revisão do servidor MCP,
com foco no novo fluxo de geração de apresentações PowerPoint.

## Implementado nesta rodada

- `render_slide` agora gera `.pptx` diretamente com `python-pptx`, usando um
  slide por `<section>`, `div.slide` ou `[data-slide]`.
- O PPTX aceita títulos, parágrafos, listas, tabelas, imagens locais/data URI,
  links clicáveis, cores simples, metadados de título/autor e proporções `16:9`
  e `4:3`.
- Imagens relativas podem ser resolvidas com `assetBasePath`; o rodapé pode usar
  `footerHtml` com `{slideNumber}` e `{slideCount}`, recalculados por chamada.
- O PPTX final é gravado atomicamente, reaberto com `python-pptx` e verificado
  como ZIP antes de substituir o arquivo de destino.
- `read_doc` agora inspeciona texto e tabelas de arquivos `.pptx`.
- `aspectRatio` inválido, formato diferente de `pptx`, opções de PDF (`fonts`,
  `tagged`, `outline`) e imagens remotas agora falham com mensagens explícitas.
- Há detecção de overflow vertical, validação de imagens e testes de contrato
  para o fluxo PPTX.
- Em outras tools, foram corrigidos limites de ZIP, ranges inválidos, entrada
  inválida em `write_sheet`, células inválidas em `patch_sheet` e nomes
  duplicados em `bundle_zip`.
- A conversão PPTX roda em thread para não bloquear o event loop do servidor
  durante a geração.
- O pacote está na versão `0.2.2`, com lockfile regenerado, build validado e
  publicação confirmada no PyPI.

## Prioridade alta

### 1. Melhorar a fidelidade do HTML/CSS no PPTX

O `render_slide` recebe HTML e CSS, mas a conversão atual para PowerPoint
reproduz apenas um subconjunto simples: seletores CSS diretos de tag, classe ou
id e propriedades de cor/fundo. Layout, espaçamentos, fontes, dimensões e
posicionamento definidos por CSS complexo ainda podem não ser preservados.

Existem dois caminhos possíveis:

- gerar elementos editáveis no PowerPoint, implementando um subconjunto bem
  definido de HTML/CSS;
- renderizar cada slide como imagem dentro do PPTX, preservando o visual com
  maior fidelidade, mas deixando o conteúdo não editável.

Uma evolução possível é oferecer os dois modos, por exemplo `editable` e
`visual`.

### 2. Detectar conteúdo que ultrapassa o slide

Textos longos, tabelas grandes e listas extensas usam atualmente dimensões
fixas. O conteúdo pode ficar cortado ou sobreposto sem que a tool informe o
problema.

Já existe uma detecção básica de overflow vertical e a tool retorna erro quando
o conteúdo acumulado passa do limite do slide. Ainda falta medir altura real de
texto, tabela e quebra de linha; portanto, a detecção pode ser conservadora.

O servidor poderia:

- calcular ou estimar o espaço ocupado pelos elementos;
- detectar overflow por slide;
- retornar um aviso ou erro indicando o slide afetado;
- opcionalmente criar slides adicionais quando o conteúdo puder ser dividido
  com segurança.

### 3. Tratar imagens remotas explicitamente

Imagens `http://` ou `https://` agora são rejeitadas explicitamente no PPTX,
evitando uma apresentação incompleta sem explicação.

Opções de correção:

- baixar imagens remotas antes de incorporá-las, com limite de tamanho e
  timeout;
- rejeitar URLs remotas com uma mensagem clara;
- permitir imagens remotas somente quando uma opção explícita estiver ativada.

### 4. Corrigir parâmetros sem efeito real

O contrato de `render_slide` ainda expõe parâmetros que eram usados no fluxo de
PDF, como `fonts`, `tagged` e `outline`, mas eles não têm o mesmo efeito na
geração de PPTX.

Os parâmetros continuam no contrato por compatibilidade, mas são rejeitados
quando usados no PPTX, pois não têm equivalente implementado. O `css` tem
suporte deliberadamente limitado e precisa ser ampliado se for necessário
reproduzir layouts complexos.

As opções futuras são:

- remover esses parâmetros do schema da tool;
- documentar que não são aplicáveis ao PPTX;
- implementar equivalentes reais para PowerPoint.

O parâmetro `css` também precisa ser aplicado ou ter seu escopo claramente
limitado.

## Prioridade média

### 5. Ampliar o suporte a elementos HTML

O conversor atual suporta títulos, parágrafos, listas, tabelas e imagens
simples. Ainda podem ser adicionados:

- links clicáveis (já implementados para `<a href>` simples);
- blockquotes;
- blocos de código;
- listas aninhadas;
- alinhamento de texto;
- cores de fundo;
- bordas e formas;
- posicionamento absoluto;
- notas de apresentação;
- estilos de tabela;
- cabeçalho e rodapé por slide.

Cada novo elemento deve ter um teste de geração e uma validação de abertura do
PPTX.

### 6. Validar `aspectRatio`

Valores diferentes de `16:9` e `4:3` agora geram erro de validação claro.

O contrato deveria aceitar explicitamente apenas:

```text
16:9
4:3
```

Qualquer outro valor deveria produzir um erro de validação claro.

### 7. Evitar bloqueio do event loop

`render_slide` usa `asyncio.to_thread` para que `python-pptx` não bloqueie o
transporte assíncrono durante a geração.

### 8. Adicionar testes de contrato para PPTX

Além dos testes básicos já existentes, cobrir:

- imagem local incorporada;
- imagem relativa via `assetBasePath`;
- imagem em data URI;
- tabela com cabeçalho;
- texto longo;
- lista ordenada e não ordenada;
- proporção `4:3`;
- `outputPath` sem extensão;
- imagem inválida;
- URL remota;
- HTML sem `<section>`;
- múltiplas formas de separar slides;
- metadados de título e autor.
- leitura posterior via `read_doc`;
- diagnóstico contendo o número do slide que falhou;
- rodapé com contagem recalculada.

## Prioridade de manutenção

### 9. Manter a publicação sincronizada com o código

Toda alteração do contrato ou do comportamento do servidor precisa seguir este
fluxo:

1. atualizar a versão no `pyproject.toml`;
2. regenerar o `uv.lock`;
3. executar testes, health check e build;
4. publicar os artefatos no PyPI;
5. validar a instalação com `uvx` usando a versão publicada;
6. atualizar o README e os exemplos de configuração.

Isso evita que a versão pública fique diferente da versão validada localmente.

### 10. Manter a documentação do contrato atualizada

O README deve refletir sempre:

- o formato real de saída de cada tool;
- os parâmetros suportados;
- limitações conhecidas do PPTX;
- exemplos completos de HTML;
- requisitos de instalação;
- versão mínima do Python e das dependências relevantes.

## Sequência recomendada

1. Definir o modelo de fidelidade visual: PPTX editável, visual ou ambos.
2. Evoluir o suporte a CSS e elementos HTML.
3. Melhorar a medição de overflow e permitir quebra segura em slides.
4. Fechar os gaps transversais listados abaixo.

## Outros gaps encontrados nas tools

- `read_doc` pode deixar arquivos temporários quando `extractImages` é usado;
  falta garantir limpeza em `finally`.
- `read_sheet` carrega a planilha inteira antes de aplicar `maxRows`; arquivos
  muito grandes podem consumir memória além do necessário.
- `scan_dir` continua silencioso ao pular arquivos que não consegue ler; seria
  melhor retornar contagem e motivos dos itens ignorados.
- `read_archive` e `read_ods` têm limites de volume, mas ainda vale centralizar
  a política para todos os formatos compactados.
- `write_doc` e `patch_doc` poderiam validar de forma mais uniforme extensão,
  tamanho final e atomicidade das saídas.
- A suíte cobre os contratos principais, mas ainda falta teste de integração
  com o processo MCP via stdio e health check após instalação pelo PyPI.

## Observação sobre o scroll do terminal

O ambiente utiliza WezTerm com Zellij:

```text
TERM_PROGRAM=WezTerm
TMUX=
```

O WezTerm está configurado com `100000` linhas de histórico, portanto o limite
de scrollback não parece ser a causa.

No Zellij, para acessar o histórico do terminal:

- pressione `Ctrl+s` para entrar no modo de scroll;
- use `PageUp`, `PageDown` ou as setas;
- pressione `Esc` para sair do modo de scroll.
