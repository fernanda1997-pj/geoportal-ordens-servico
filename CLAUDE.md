# Geoportal RTA-MSI — Ordens de Serviço

WebGIS Leaflet de página única com as **O.S.P. (Ordem de Serviço de Pagamento)** da
planilha de Controle — mostra no mapa os trechos com O.S.P., coloridos por situação,
com filtros e um Painel Executivo pra apresentação/relatório.

Usuária: Fernanda (RTA Engenheiros Consultores). Responder sempre em português.

- **Site**: (ainda não publicado — ver seção "Publicar")
- **Repo**: (ainda não publicado)

Repo próprio desde **2026-08-18** — antes vivia como pasta `ordens-servico/` dentro
do repo `web - fichas` (junto com o app de ficha de inspeção). A usuária pediu
repositório separado ("vou criar só para as OS"). `camadas/` e `logo/` aqui são
**cópia própria** (não referenciam mais `web - fichas`), pra este repo ser
totalmente independente.

## Arquitetura

| Arquivo/pasta | Papel |
|---|---|
| `index.html` | App inteiro (HTML+CSS+JS, sem build). CDN: Leaflet 1.9.4, Chart.js 4, html-to-image 1.11.13 (carregada sob demanda só na hora de exportar imagem) |
| `converter_os.py` | Lê `fichas/OS/*.xlsm` + `camadas/R<n>_TRECHOS.shp` (campo `Id` = "N° TRECHO" da planilha), gera uma feature por linha do shapefile que bater com o trecho — geometria do trecho INTEIRO, sem corte por km (O.S.P. não referencia sub-trecho, diferente da ficha de inspeção) |
| `fichas/OS/` | Planilha(s) de Controle de O.S.P., como chegam (ex. `Controle de OSPs LOTE 01.xlsm`) — **não editar**, só adicionar/atualizar arquivo aqui e rodar o converter de novo |
| `dados/os_<REGIAO>.js` | Um GeoJSON (`window.DADOS_OS_REGIAO[regiao]`) por região **geográfica** (R1, R2, R3...) — não por competência, a O.S. não é mensal |
| `dados/manifest_os.js` | Lista de regiões disponíveis (`window.MANIFEST_OS`) |
| `relatorio_qualidade_os.txt` | Gerado a cada rodada (gitignored) — trecho não encontrado no shapefile, erro de fórmula (`#REF!`) na planilha etc. |
| `camadas/` | Cópia de `R<n>_TRECHOS.shp` — só as regiões usadas aqui (R1/R2/R3 = Lote 1 ativo; R11/R12/R13 = Lote 4, shapefiles já prontos aguardando a planilha) |
| `logo/` | Logos RTA + MSI |

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
combinações que existem de verdade nos dados (ex. hoje falta "Região 15" em
Restauração — não tem O.S.P. emitida ali ainda).

## Situação da O.S.P.

Vocabulário vem da aba "Tabelas Auxiliares" da planilha: `Em elaboração`,
`Em andamento`, `Concluída`, `Justificada`, `Cancelada`, `Correção Fiscal`,
`Análise Gestor`. Cores em `CORES_SITUACAO`, ícones em `ICONES_SITUACAO`
(✅🔧✏️📄🚫💲), combinados em `rotuloComIcone()` pra exibição em pills/badges/legenda.

**"-" não é uma situação de verdade** — não está na lista oficial da planilha, é a
célula SITUAÇÃO sem preencher ainda (algumas O.S.P. reais, com valor e trecho
normais, simplesmente não tiveram a situação atualizada). Exibido como
"Situação não informada" (`rotuloSituacao()`), mas o VALOR usado pra filtrar
continua sendo o `"-"` cru vindo da planilha — só a exibição foi trocada.

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
  HISTÓRICO da planilha) por mês, com acumulado; bate 100% com a fonte.
- **Ranking Top 10** (toggle Por trecho / Por serviço / Por O.S.P.) — "Por serviço"
  agrupa pela string INTEIRA do campo Serviço (não separa por `/`, mesmo critério
  da tabela oficial "TOP 10 SERVIÇOS" da aba DETALHAMENTO da planilha — muitas
  O.S.P. repetem a mesma combinação exata). "Por O.S.P." não agrega nada, lista
  registros individuais (rótulo `contrato.osp` com 4 dígitos, igual a planilha).
  Todos os 3 modos conferidos batendo valor a valor com as tabelas oficiais da
  planilha (DETALHAMENTO) ou recálculo independente.

**BASE DASHBOARD (aba da planilha) não é fonte confiável** — as colunas auxiliares
dela (que alimentariam os gráficos nativos do Excel) estão zeradas/quebradas
(fórmula não recalculada). Prefira sempre recalcular a partir de RESUMO+HISTÓRICO
e comparar contra as tabelas que SÃO confiáveis (DETALHAMENTO tem os Top 10 corretos).

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

## Lote 4 (Regiões 11/12/13 manutenção + 22/23/24 restauração)

Shapefiles (`camadas/R11/R12/R13_TRECHOS.shp`) e mapeamento
(`MAPA_REGIAO_GEOGRAFICA`) já prontos — testado com planilha fictícia isolada,
funciona sem mudar nada de código. **Só falta a planilha real** ("Controle de OSPs
LOTE 04.xlsm" ou nome equivalente) em `fichas/OS/` — quando a usuária colocar,
rodar `python converter_os.py` de novo.

## Relação com outros projetos

Quarto geoportal da família RTA-MSI (mesmas logos/paleta, apps independentes):
- `C:\1. Projetos\RTA\web` — geoportal principal (Folium), `rta-msi-rodovias.vercel.app`.
  Fonte original dos shapefiles `R*_TRECHOS.shp`.
- `C:\1. Projetos\RTA\web - Mapas` — Levantamento de Trechos, `mapa-levantamento.vercel.app`.
- `C:\1. Projetos\RTA\web - fichas` (pasta `ficha-inspecao/`) — Inspeção do Pavimento.
  Este projeto (`web - OS`) morava dentro desse repo até 2026-08-18, quando virou
  independente a pedido da usuária.
