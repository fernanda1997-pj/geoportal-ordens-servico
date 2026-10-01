# Geoportal RTA-MSI — Ordens de Serviço

WebGIS Leaflet de página única com as **O.S.P. (Ordem de Serviço de Pagamento)** —
mostra no mapa os trechos com O.S.P., coloridos por situação, com filtros e um
Painel Executivo pra apresentação/relatório. Dados vêm do banco de dados oficial
do órgão (SISTEMA_AGETO) desde 2026-10 — ver seção "Fonte de dados" abaixo.

Usuária: Fernanda (RTA Engenheiros Consultores). Responder sempre em português.

- **Site**: https://geoportal-ordens-servico.vercel.app
- **Repo**: https://github.com/fernanda1997-pj/geoportal-ordens-servico

Repo próprio desde **2026-08-18** — antes vivia como pasta `ordens-servico/` dentro
do repo `web - fichas` (junto com o app de ficha de inspeção). A usuária pediu
repositório separado ("vou criar só para as OS"). `camadas/` e `logo/` aqui são
**cópia própria** (não referenciam mais `web - fichas`), pra este repo ser
totalmente independente.

## Arquitetura

| Arquivo/pasta | Papel |
|---|---|
| `index.html` | App inteiro (HTML+CSS+JS, sem build). CDN: Leaflet 1.9.4, Chart.js 4, html-to-image 1.11.13 (carregada sob demanda só na hora de exportar imagem) |
| `converter_os.py` | Lê direto do banco oficial do órgão em `G:\...\SISTEMA_AGETO` (ver "Fonte de dados") + `camadas/R<n>_TRECHOS.shp` (campo `Id` = TRECHO_N da O.S.P.), gera uma feature por linha do shapefile que bater com o trecho — geometria do trecho INTEIRO, sem corte por km |
| `fichas/OS/` | **Histórico** — planilhas manuais ("Controle de OSPs LOTE 01/04") usadas até 2026-10, antes da migração pro banco do órgão. Não lidas mais pelo converter; deixadas aqui só de arquivo/backup |
| `dados/os_<REGIAO>.js` | Um GeoJSON (`window.DADOS_OS_REGIAO[regiao]`) por região **geográfica** (R1, R2, R3...) — não por competência, a O.S. não é mensal |
| `dados/manifest_os.js` | Lista de regiões disponíveis (`window.MANIFEST_OS`) |
| `relatorio_qualidade_os.txt` | Gerado a cada rodada (gitignored) — trecho não encontrado no shapefile, pasta/arquivo de região faltando etc. |
| `camadas/` | Cópia de `R<n>_TRECHOS.shp` — R1/R2/R3 (Lote 1/2/3) + R11/R12/R13 (Lote 11/12/13), todas ativas |
| `logo/` | Logos RTA + MSI |

## Fonte de dados — banco oficial do órgão (SISTEMA_AGETO)

Desde **2026-10-02**, `converter_os.py` lê direto de
`G:\.shortcut-targets-by-id\16Cw6zdJvWIidBLYdaIQIh6ITuwcbe6d8\SISTEMA_AGETO\MANUTENÇÃO RODOVIÁRIA`
(Google Drive da usuária, montado como `G:\` — **só funciona na máquina dela**
com o Drive sincronizado; `BASE_SISTEMA` no topo do converter). Substituiu a
planilha manual "Controle de OSPs" (`fichas/OS/`, mantida só de histórico) —
essa planilha não tinha ano de emissão de verdade (só o mês) e tinha pelo
menos 1 contrato errado (Região 24). Cada região tem uma pasta `LOTE XX` com
um arquivo `BD_LOTE_XX.xlsx` (sistema de gestão real do órgão, não uma
planilha feita pra nós) — abas relevantes:

| Aba | Usada pra |
|---|---|
| `CONFIG_CONTRATO` | Contrato (`CodigoContrato` + "." + `AnoContrato`, zero-padded a 3 dígitos — `ler_contrato()`) |
| `BD_OSP` | 1 linha por O.S.P.+trecho: DATA (emissão), TRECHO_N, DESCRIÇÃO, STATUS, VALOR, MEDIDO |
| `ITENS_OSP` + `ORCAMENTO_PADRAO` | Reconstrói "serviço" — matriz item×O.S.P. cruzada com o catálogo de preços pela coluna **Subitem** (não "Código", não Item+Subitem — testado e confirmado). Só linhas `Agrupador="ITEM"` têm descrição de serviço real (GRUPO/SUBGRUPO são rollup) |
| `BD_MED` | 1 linha por O.S.P. por MÊS medido (série real) — vira `medido_mensal` (grade fixa Jan..Dez) e `meses_cronograma` (meses em que teve medição de verdade) |

**Particularidades confirmadas por investigação (não presumir, já foi conferido
nas 12 regiões):**
- **"Região 03" são 2 bancos separados**: `LOTE 03 - ETICA` (contrato
  1452.2026) e `LOTE 03 - LUCENA` (contrato 002.2025) — mesma área física,
  2 contratos. `pastas_do_lote(3)` retorna as duas, ambas alimentam R3.
- **"Região 23" não existe** — sem pasta no Drive ainda. Fica de fora da
  lista `numeros_regiao` em `main()` até aparecer.
- **"Região 16" tem arquivo duplicado** (`BD_LOTE_16 - Copia.xlsx` solto na
  pasta) — `arquivo_bd_da_pasta()` ignora qualquer nome com "COPIA".
- **Tipos inconsistentes DENTRO do próprio banco** (mesma coluna, linhas
  diferentes): número de O.S.P. ora string zero-padded (`"0073"`) ora int
  puro (`73`); `EXT_KM` ora float ora string com vírgula brasileira
  (`"9,25"`); `DATA` ora `datetime` real ora texto `"MÊS-ANO"`/`"MÊS/ANO"`.
  `_float_br()`/`_int_seguro()`/`_parse_mes_ano()` tratam os dois formatos
  em qualquer coluna nova que for ler daqui pra frente.
- **Vocabulário de SITUAÇÃO tem mais valores que a planilha antiga**:
  além de Em elaboração/Em andamento/Concluída/Justificada/Cancelada/
  Correção Fiscal/Análise Gestor, o banco tem **Liberada** e **Para Emissão**
  (novas, com cor/ícone em `CORES_SITUACAO`/`ICONES_SITUACAO`) e
  **Correção Super** (só Região 13 — tratada como sinônimo de Correção
  Fiscal em `STATUS_MAP`, confirmar com a usuária se não for o caso).
  `STATUS_MAP` em `converter_os.py` usa chaves SEM acento (compara contra
  `_norm()`, que sempre tira acento) — **cuidado**: a 1ª versão tinha
  chaves acentuadas e metade dos status vazava em maiúsculo pro frontend.
- **"Serviço" ficou bem mais granular** (O.S.P. pode ter 10+ itens de
  orçamento, cada descrição já longa sozinha) — truncado aos 2 primeiros
  itens distintos + "(+N itens)" (`carregar_servicos_por_osp()`) pra não
  quebrar os gráficos do Painel Executivo (já tropecei nisso: sem limite
  nenhum virou texto ilegível; limitando a 3 itens inteiros ainda estourava
  a largura do ranking "Por serviço").
- **Cronograma agora é a EXECUÇÃO real** (`meses_cronograma`/`cronograma`
  vêm de `BD_MED`, não de um planejamento) — o banco tem uma aba
  `BD_CRONOGRAMA` com cronograma PLANEJADO, mas em formato bem mais
  complexo (percentual por subgrupo de serviço, não por mês simples) e sem
  equivalência direta ao "ago/set" simples de antes — não usada.
- `EMPRESA_POR_CONTRATO` (index.html) foi conferida contrato a contrato
  contra `CONFIG_CONTRATO` de cada região em 2026-10 — todos batiam, EXCETO
  a Região 24 (corrigida de 041.2025/SCR pra 1151.2026/ÉTICA CONSTRUTORA).

## Região de manutenção × restauração — mesma área física, contrato diferente

Confirmado pela usuária em 2026-08-14: a planilha traz duas famílias de código pro
mesmo lugar físico — `1/2/3/11/12/13` = manutenção, `14/15/16/22/23/24` =
restauração da MESMA região geográfica (mesmo trecho, contrato diferente).
`MAPA_REGIAO_GEOGRAFICA` em `converter_os.py` traduz o código de restauração pro
código geográfico (`14→1, 15→2, 16→3, 22→11, 23→12, 24→13`) que é o que existe em
`camadas/` — sem isso a O.S. de restauração nunca acharia shapefile (R14 não existe,
só R1). Cada feature guarda os DOIS códigos:
- `regiao` — código geográfico agrupado (R1, R2, R3...), usado só pra achar a
  geometria/desenhar no mapa certo.
- `regiao_os` — código REAL da planilha (R01, R14, R22...), é o que a usuária pensa
  ("Região 14") e o que aparece em toda exibição pro usuário (`rotuloRegiao()`).
- `tipo_servico` (`"manutencao"`/`"restauracao"`).
- `contrato` — mesmo padrão de exibição: valor cru usado pra filtro/comparação,
  `rotuloContrato()` (backed by `EMPRESA_POR_CONTRATO`) acrescenta " · NOME DA
  EMPRESA" só na exibição, em todo lugar que mostra contrato (select, lista,
  gaveta, resumo de filtros, Painel Executivo). **Exceção deliberada**: o rótulo
  `contrato.osp` do ranking "Por O.S.P." (ex. "034.2025.0009") NÃO leva nome de
  empresa — espelha o formato oficial da planilha, não mexer. Mapeamento
  contrato→empresa é fornecido pela usuária (contratos são únicos no dataset
  todo); se aparecer contrato novo sem empresa no mapa, `rotuloContrato()` cai
  de volta pro número puro, sem quebrar.

**Nunca usar `regiao` (geográfico) pra exibir ou agrupar/somar pro usuário** — já
rendeu bug 2x (filtro Região misturando os dois tipos; gráfico "Previsto×Executado
por região" somando manutenção+restauração juntos). Usar sempre `regiao_os` +
`rotuloRegiao()` pra exibição, e filtrar por `tipo_servico` explicitamente.

## Filtro em cadeia: Tipo de serviço → Região → Contrato

Nessa ordem, cada um dependente do anterior (a usuária pediu explicitamente essa
ordem e esse comportamento, depois de eu tentar um seletor único combinado
"Região — Tipo" que ela não gostou): escolher o Tipo já filtra as opções de Região
pros códigos daquele tipo (`popularSelectRegiao()`); escolher a Região filtra as
opções de Contrato (`popularSelectContrato()`). Ambos preservam a seleção atual se
ainda for válida pro novo filtro, senão voltam a "Todas"/"Todos". Só aparecem
combinações que existem de verdade nos dados.

**Região e Mês são multi-seleção** (`regioesAtivasOS`/`mesesAtivosOS`, Sets —
mesmo padrão de `situacoesAtivasOS`): dropdown fechado por padrão, abre um
popover de checkboxes ao clicar o botão (`.dropdown-check`/`.popover-check`/
`configurarDropdownCheck()`, compartilhado pelos dois). Mês considera tanto
`data_emissao` quanto `meses_cronograma` (`mesesDaEntrada()`) — uma O.S.P.
emitida em Agosto com CRONOGRAMA "ago/set" aparece nos dois meses sem
precisar marcar os dois. Tipo e Contrato continuam seleção única (`<select>`
normal) — só virou multi-seleção o que a usuária pediu explicitamente.

## Ano de emissão — resolvido pela migração pro banco oficial

Histórico (relevante só se for entender commits antigos): a planilha manual
"Controle de OSPs" só guardava o MÊS da emissão (texto, sem ano) — 1ª
tentativa inferiu o ano pelo sufixo ".AAAA" do contrato, a usuária confirmou
em 2026-10 que isso era **falso** (contrato dura mais que 1 ano, sufixo não
reflete emissão real), então virou uma constante única `ANO_EMISSAO_ATUAL =
2026` aplicada a tudo. **Essa constante não existe mais** — o banco oficial
(`BD_OSP.DATA`) tem a data real (dia/mês/ano) de cada O.S.P., lida direto por
`_parse_mes_ano()` (ver "Fonte de dados" acima). Se "ano de emissão" parecer
estranho de novo, suspeitar primeiro de cache do navegador (ver próxima
seção), depois conferir `BD_OSP.DATA` da O.S.P. em questão direto no banco.

## Cache dos dados (`dados/os_*.js`)

Carregados via `<script>` criado em JS (`carregarTodosOsRegiao()`), com
`?v=Date.now()` no `src` — **sempre** força buscar a versão mais nova do
servidor. Sem isso, dados atualizados (planilha nova, correção de ano etc.)
podem ficar presos no cache do navegador sem a usuária perceber, já que o
nome do arquivo não muda quando o conteúdo muda (foi exatamente o que
aconteceu testando a correção do ano acima — `fetch({cache:'no-store'})`
mostrava o valor certo enquanto a página carregada mostrava o errado).
Qualquer novo arquivo de dados carregado dinamicamente neste projeto deve
seguir o mesmo padrão.

## Situação da O.S.P.

Vocabulário vem do campo `STATUS` do banco (`BD_OSP`), mapeado em
`STATUS_MAP` (converter_os.py) pro Título Capitalizado de sempre: `Em
elaboração`, `Em andamento`, `Concluída`, `Justificada`, `Cancelada`,
`Correção Fiscal` (inclui "Correção Super", só Região 13, tratada como
sinônimo), `Análise Gestor`, `Liberada`, `Para Emissão` (as 2 últimas
novas desde a migração pro banco, "Fonte de dados" acima). Cores em
`CORES_SITUACAO`, ícones em `ICONES_SITUACAO`, combinados em
`rotuloComIcone()` pra exibição em pills/badges/legenda.

Valor cru `"0"` ou vazio vira `"Não informada"` — status espúrio visto na
Região 14, provavelmente erro de digitação na fonte, não uma situação real.

## Prazo e atraso (`atrasada`, `prazo_limite`, `prazo_meses`)

`PRAZO` (BD_OSP) é em **meses**, não dias (confirmado contra a distribuição
real: sempre 1-4). `calcular_prazo()` em `converter_os.py`: prazo final =
mês de emissão + PRAZO meses. Compara contra o ÚLTIMO mês com medição real
(`meses_cronograma`) se a O.S.P. já está Concluída/Cancelada/Justificada, ou
contra HOJE (`datetime.date.today()`, roda na máquina da usuária) se ainda
está ativa. Sem emissão ou sem PRAZO na fonte → os 3 campos saem `null`
(não assume nem "no prazo" nem "atrasada" sem dado pra provar). No frontend:
badge "⏰ Atrasada" no item da lista, linha "Prazo" na gaveta, KPI
`data-filtro-atrasada` clicável (painel lateral + Painel Executivo) — é um
**toggle independente** do filtro de Situação (`filtroAtrasadas`, variável
separada de `situacoesAtivasOS`), já que atraso corta através de várias
situações ao mesmo tempo (uma "Em elaboração" parada há meses conta tanto
quanto uma "Em andamento" que passou do prazo).

## O.S.P. sem geometria — ainda contam nos KPIs/lista

Uma O.S.P. pode não ter geometria (trecho ainda "não cadastrado" na planilha, ou
número de trecho que não bate com nenhum shapefile) — ela é um registro
administrativo real (tem valor previsto, orçamento) e **precisa continuar contando
nos totais**, só não desenha no mapa. `converter_os.py` emite a feature mesmo assim
(`geometry: null`); `entradasUnicasOS()` calcula `entrada.semGeometria` (true só se
TODAS as features daquela O.S.P.+trecho não tiverem geometria); a lista mostra tag
"📍⚠️", `abrirOS()` pula o desenho no mapa e a gaveta mostra aviso.

**Isso já foi um bug real**: o converter costumava `continue` (pular) essas linhas,
o que fazia o total do geoportal (155 O.S.P.) não bater com o dashboard nativo da
planilha (163 O.S.P., diferença de R$ 72 mil) — a usuária percebeu comparando print
do Excel. Se "os números não baterem" de novo, suspeitar primeiro desse padrão:
comparar contagem/soma linha a linha contra as abas RESUMO da planilha (script
python direto com openpyxl) antes de desconfiar de outra coisa.

## Mapa geral — todos os trechos sempre visíveis

`camadaGeral`/`renderizarCamadaGeral()`: desenha TODOS os trechos com O.S.P. dentro
do filtro atual, coloridos por situação — sempre visível (não só ao clicar um item
específico da lista). Reconstrói sozinha a cada mudança de filtro. Clicar numa
linha do mapa (não só na lista) abre a gaveta de detalhe da O.S.P. certa (mesmo
lookup `porChave`/`chaveOS()` da lista). Selecionar uma O.S.P. específica na lista
continua dando destaque (casing branco + linha colorida por cima) + zoom
(`fitBounds`) via `abrirOS()`.

## Painel Executivo

Botão "📊 Painel Executivo" no cabeçalho abre um overlay full-screen pensado pra
apresentação/relatório: 6 KPIs (previsto/executado/%/saldo/concluídas/total) + 4
gráficos Chart.js, todos respeitando os MESMOS filtros do painel lateral (recalcula
sozinho a cada mudança, `atualizarDashboard()` chamado no fim de
`renderizarListaOS()`):
- **Situação** (rosca) — exclui as situações `-` e `Correção Fiscal` do cálculo,
  igual o gráfico "Distribuição das OSPs por Situação" nativo do Excel (confirmado
  batendo % contra print da planilha — com elas dentro os % não fecham).
- **Previsto × Executado por região** — agrupado por `regiao_os` (não `regiao`).
- **Evolução mensal do valor medido** — soma `medido_mensal` (vindo da aba
  BD_MED do banco oficial — ver "Fonte de dados") por mês, com acumulado.
- **Ranking Top 10** (toggle Por trecho / Por serviço / Por O.S.P.) — "Por serviço"
  agrupa pela string INTEIRA do campo Serviço (não separa por `/`, mesmo critério
  da tabela oficial "TOP 10 SERVIÇOS" da aba DETALHAMENTO da planilha — muitas
  O.S.P. repetem a mesma combinação exata). "Por O.S.P." não agrega nada, lista
  registros individuais (rótulo `contrato.osp` com 4 dígitos, igual a planilha).
  Todos os 3 modos conferidos batendo valor a valor com as tabelas oficiais da
  planilha (DETALHAMENTO) ou recálculo independente.

Histórico: quando a fonte era a planilha manual, a aba "BASE DASHBOARD" dela
não era confiável (fórmulas quebradas) — não existe mais, só relevante pra
entender commits antigos.

### Exportar imagem (📄 Exportar)

Popover com checkboxes (KPIs/Situação/Região/Evolução/Ranking) + botão "Gerar
imagem" → baixa PNG limpo (sem botões da interface) via **html-to-image** (mesma
lib/versão do geoportal de Levantamento, carregada sob demanda). Captura
`#dashboard-overlay` inteiro.

**2 bugs achados e corrigidos nessa feature:**
1. Imagem saía **cortada na altura da tela** — `#dashboard-overlay` é
   `position:fixed;inset:0` (preso ao viewport); só destravar o scroll interno não
   bastava, a caixa externa continuava com altura fixa. Fix: durante a exportação
   (`.exportando`) o overlay também vira `position:absolute; height:auto`.
2. Falhava com `file://` (usuária abre o HTML direto, clique duplo, sem servidor) —
   a lib tenta re-embutir cada `<img>` via `fetch()`, e `fetch()` de arquivo local é
   bloqueado pelo Chrome. Fix: as 2 logos dentro do Painel Executivo viraram
   `data:` URI (base64) embutido direto no HTML — sem arquivo pra buscar. **Cuidado
   ao regenerar esse base64**: não copiar manualmente via ferramenta de leitura (uma
   vez saiu truncado sem erro nenhum, só a imagem não carregava) — gerar com
   script (`base64.b64encode` em Python) e trocar direto no arquivo, conferindo o
   tamanho da string resultante.

## Testar local

`python -m http.server <porta> --directory .` na raiz do projeto. Ver
`.claude/launch.json` do projeto `web` vizinho (`C:\1. Projetos\RTA\web\.claude\launch.json`,
config `ordens-servico`, porta 8771).

## Publicar

1. Criar um repositório **novo e vazio** no GitHub (usuária faz pela UI — sessões
   de Claude Code não têm `gh` CLI nem token configurado aqui)
2. `git remote add origin <url>` e `git push -u origin main` (branch local é
   `master` — renomear pra `main` antes ou ajustar o push)
3. Importar o repo no Vercel (vercel.com → Add New Project) — deploy automático a
   cada push na `main`, igual aos outros geoportais da família

## Histórico: Regiões 11/12/13/22/24 (antigo "Lote 4")

Shapefiles e `MAPA_REGIAO_GEOGRAFICA` pra essas regiões já estavam prontos
desde o início do projeto; os dados em si vieram primeiro de uma planilha
manual separada ("Controle de OSPs LOTE 04.xlsx", 2026-10-01) e, 1 dia
depois, da migração geral pro banco oficial (ver "Fonte de dados" acima) —
hoje todas as regiões (antigo Lote 1 e Lote 4) vêm da mesma fonte, sem
distinção de "lote" no código.

## Relação com outros projetos

Quarto geoportal da família RTA-MSI (mesmas logos/paleta, apps independentes):
- `C:\1. Projetos\RTA\web` — geoportal principal (Folium), `rta-msi-rodovias.vercel.app`.
  Fonte original dos shapefiles `R*_TRECHOS.shp`.
- `C:\1. Projetos\RTA\web - Mapas` — Levantamento de Trechos, `mapa-levantamento.vercel.app`.
- `C:\1. Projetos\RTA\web - fichas` (pasta `ficha-inspecao/`) — Inspeção do Pavimento.
  Este projeto (`web - OS`) morava dentro desse repo até 2026-08-18, quando virou
  independente a pedido da usuária.
