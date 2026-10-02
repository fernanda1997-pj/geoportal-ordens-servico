# -*- coding: utf-8 -*-
"""
converter_os.py — Geoportal RTA-MSI / Ordens de Serviço (O.S.P.)

Lê direto do banco de dados oficial do órgão (SISTEMA_AGETO, Google Drive
montado como G:\\), um arquivo BD_LOTE_XX.xlsx por região (pasta "LOTE XX"),
e gera por região GEOGRÁFICA (R1, R2, R3, R11, R12, R13) um GeoJSON com uma
feature por linha do shapefile de trechos (camadas/R<região>_TRECHOS.shp,
campo `Id`) que bater com o TRECHO_N da O.S.P. — geometria do trecho INTEIRO,
sem corte por km (O.S.P. não referencia sub-trecho).

Substituiu em 2026-10 a leitura da planilha manual "Controle de OSPs
LOTE 01/04.xlsx" (fichas/OS/*.xlsm) — essa planilha não tinha ano de emissão
de verdade (só o mês), obrigava atualização manual e tinha pelo menos 1
contrato errado (Região 24). O banco novo tem data completa, é a fonte viva
usada pelo próprio órgão, e os campos adicionais (serviço, cronograma,
medição mensal) vêm de abas cruzadas — ver comentários abaixo de cada leitor.

Região de manutenção x restauração — mesma área geográfica, contrato
diferente (confirmado pela usuária em 2026-08-14). `MAPA_REGIAO_GEOGRAFICA`
traduz o código de restauração pro código geográfico que o shapefile usa.

Particularidades do banco confirmadas por investigação (2026-10-02):
- "Região 03" são DOIS bancos separados na mesma pasta de nível acima
  (LOTE 03 - ETICA, contrato 1452.2026; LOTE 03 - LUCENA, contrato
  002.2025) — mesma área física, dois contratos diferentes. Os dois
  alimentam R3.
- "Região 23" não existe ainda no banco (pendente do órgão criar) — fica
  de fora até aparecer uma pasta LOTE 23.
- "Região 16" tem um arquivo "BD_LOTE_16 - Copia.xlsx" solto na pasta —
  usa sempre o que NÃO tem "Copia" no nome.
- Tipos de dado inconsistentes dentro do próprio banco (confirmado nas 12
  regiões): número de O.S.P. ora string zero-padded ("0073") ora int puro
  (73); EXT_KM ora float ora string com vírgula decimal ("9,25"); datas ora
  datetime real ora texto "MÊS-ANO"/"MÊS/ANO". Todo leitor abaixo trata os
  dois formatos.

Saída (tudo em dados/, consumido pelo index.html sem build, MESMO formato
de antes — o frontend não precisou mudar):
    dados/os_<REGIAO>.js       -- um por região geográfica (R1, R2, ...)
    relatorio_qualidade_os.txt -- trecho não encontrado no shapefile etc.

Rodar:  python converter_os.py
Requer: openpyxl, geopandas, shapely — e acesso de leitura ao Google Drive
        montado em G:\\ (SISTEMA_AGETO), ver BASE_SISTEMA abaixo.
"""
import datetime
import glob
import json
import os
import re
import unicodedata

import openpyxl
import geopandas as gpd
from shapely.geometry import mapping

BASE = os.path.dirname(os.path.abspath(__file__))
CAMADAS_DIR = os.path.join(BASE, 'camadas')
DADOS_DIR = os.path.join(BASE, 'dados')
EPSG_METRICO = 31982  # SIRGAS 2000 / UTM 22S — mesmo do geoportal principal

# Pasta compartilhada do órgão (Google Drive montado como G:\) — se a usuária
# trocar de computador ou o Drive remapear a letra, ajustar aqui.
BASE_SISTEMA = (
    r'G:\.shortcut-targets-by-id\16Cw6zdJvWIidBLYdaIQIh6ITuwcbe6d8'
    r'\SISTEMA_AGETO\MANUTENÇÃO RODOVIÁRIA'
)

# restauração (chave) -> manutenção/geográfico (valor) — mesma área física.
MAPA_REGIAO_GEOGRAFICA = {14: 1, 15: 2, 16: 3, 22: 11, 23: 12, 24: 13}
REGIOES_RESTAURACAO = set(MAPA_REGIAO_GEOGRAFICA.keys())

# Número do lote -> nome(s) de pasta dentro de BASE_SISTEMA. A maioria é só
# "LOTE XX", mas algumas ganharam sufixo com o nome da empresa (confirmado
# variar com o tempo) — por isso usa glob "LOTE XX*" em vez de nome fixo,
# EXCETO a Região 3 que são dois bancos de verdade (não é só sufixo
# cosmético, são dois contratos/empresas diferentes) — lista explícita.
PASTAS_REGIAO_3 = ['LOTE 03 - ETICA', 'LOTE 03 - LUCENA']

MESES_ORDEM = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
MESES_PT = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
            7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
# "ABRIL" (sem acento perdido — Ç/Ã já vêm certos do Excel, só precisa casar
# maiúsculas) -> nome bonito pra exibição.
MESES_COMPLETO_MAP = {nome.upper(): nome for nome in MESES_ORDEM}

# Vocabulário de SITUAÇÃO do banco (tudo maiúsculo) -> rótulo usado no
# geoportal (mesmo Título Capitalizado de sempre). "CORREÇÃO SUPER" = a
# supervisora devolveu a O.S.P. pra correção ANTES de emitir (a usuária
# explicou em 2026-10-02) — situação PRÓPRIA, não é sinônimo de "Correção
# Fiscal" (1ª versão tratava como sinônimo, estava errado; o checklist da
# supervisão, ver carregar_checklists(), mostra o que precisa corrigir).
# "LIBERADA"/"PARA EMISSÃO" são situações novas
# que a planilha antiga não tinha (ver CORES_SITUACAO/ICONES_SITUACAO no
# index.html, também atualizados).
## Chaves SEM acento de propósito — são comparadas contra _norm(), que
## sempre tira acento antes de comparar (ver _norm()). Já tropecei nisso:
## 1ª versão tinha as chaves acentuadas e metade dos status (os com
## acento) caía no fallback "usa o texto original", vazando "CONCLUÍDA"
## etc. em maiúsculo pro frontend em vez de "Concluída".
STATUS_MAP = {
    'EM ELABORACAO': 'Em elaboração',
    'EM ANDAMENTO': 'Em andamento',
    'CONCLUIDA': 'Concluída',
    'JUSTIFICADA': 'Justificada',
    'CANCELADA': 'Cancelada',
    'CORRECAO FISCAL': 'Correção Fiscal',
    'CORRECAO SUPER': 'Correção Super',
    'ANALISE GESTOR': 'Análise Gestor',
    'LIBERADA': 'Liberada',
    'PARA EMISSAO': 'Para Emissão',
}

qa_msgs = []


def qa(msg):
    qa_msgs.append(msg)
    print('  [QA]', msg)


def _norm(s):
    """Maiúsculas, sem acento, sem quebra de linha, espaços colapsados —
    só pra COMPARAR nomes de coluna/aba (nunca usar o resultado como texto
    de exibição, perde acentuação)."""
    if s is None:
        return ''
    s = str(s)
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    s = re.sub(r'\s+', ' ', s).strip().upper()
    return s


def _mapa_colunas(ws, linha_cabecalho=1):
    cols = {}
    for c in range(1, ws.max_column + 1):
        titulo = _norm(ws.cell(row=linha_cabecalho, column=c).value)
        if titulo and titulo not in cols:  # primeira ocorrência vence (colunas duplicadas/vazias no fim)
            cols[titulo] = c
    return cols


def _float_br(v):
    """Número que pode vir float/int de verdade OU string com vírgula
    decimal brasileira ("9,25") — confirmado os dois formatos no mesmo
    banco, até na mesma coluna."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(',', '.'))
    except ValueError:
        return None


def _int_seguro(v):
    if v is None:
        return None
    try:
        return int(_float_br(v))
    except (TypeError, ValueError):
        return None


def _parse_mes_ano(valor):
    """datetime real OU texto "ABRIL-2026"/"ABRIL/2026"/"MAIO/26" (confirmado
    os dois separadores — hífen na BD_OSP e barra na BD_MED — E os dois
    tamanhos de ano, 4 dígitos ou só 2) -> (mês completo bonito, ano) ou
    (None, None) se não reconhecer. **Cuidado, já foi bug real**: a 1ª versão
    só aceitava ano de 4 dígitos e descartava silenciosamente qualquer O.S.P.
    com data tipo "MARÇO/26" (ano com 2 dígitos) — conferir sempre contra
    `relatorio_qualidade_os.txt`/contagem de `data_emissao` nulo se esse
    parser for mexido de novo."""
    if isinstance(valor, datetime.datetime):
        return MESES_ORDEM[valor.month - 1], valor.year
    if isinstance(valor, str) and valor.strip():
        m = re.match(r'^([A-ZÀ-Ü]+)[-/\s]+(\d{2,4})$', valor.strip().upper())
        if m:
            nome = MESES_COMPLETO_MAP.get(m.group(1))
            if nome:
                ano = int(m.group(2))
                if ano < 100:
                    ano += 2000
                return nome, ano
    return None, None


MESES_ABREV_MINUSCULO = {nome: nome[:3].lower() for nome in MESES_ORDEM}

# Estados finais — já não tem mais execução de campo acontecendo, então o
# "atrasou?" olha pro ÚLTIMO mês medido (não pra hoje).
SITUACOES_FINALIZADAS = {'Concluída', 'Cancelada', 'Justificada'}


def calcular_prazo(mes_emissao, ano_emissao, prazo_meses, situacao, meses_cronograma):
    """PRAZO (BD_OSP) é em MESES (sempre visto 1-4, nunca dias — confirmado
    comparando a distribuição real: contrato de manutenção/restauração não
    dá prazo de "2 dias" pra nada). Prazo final = mês de emissão + PRAZO
    meses. O.S.P. já finalizada (Concluída/Cancelada/Justificada) compara
    contra o ÚLTIMO mês com medição real; ainda em andamento compara contra
    HOJE. Sem emissão ou sem prazo (campo vazio na fonte) -> não dá pra
    calcular, retorna tudo None (não assume no prazo nem atrasada)."""
    if not mes_emissao or not prazo_meses:
        return None, None
    idx_emissao = MESES_ORDEM.index(mes_emissao)
    total = idx_emissao + int(prazo_meses)
    ano_limite, idx_limite = ano_emissao + total // 12, total % 12
    mes_limite_str = f'{MESES_ORDEM[idx_limite]}/{ano_limite}'

    if situacao in SITUACOES_FINALIZADAS:
        if not meses_cronograma:
            return mes_limite_str, False  # finalizada sem nenhuma medição registrada — não dá pra provar atraso
        ultimo_mes, ultimo_ano = meses_cronograma[-1].split('/')
        referencia = (int(ultimo_ano), MESES_ORDEM.index(ultimo_mes))
    else:
        hoje = datetime.date.today()
        referencia = (hoje.year, hoje.month - 1)
    return mes_limite_str, referencia > (ano_limite, idx_limite)


def _cronograma_legado(meses_cronograma):
    """Reconstrói o texto curto "abr/mai/jun" (formato que o índice.html já
    sabe exibir no card de detalhe) a partir dos meses reais de medição —
    não é mais um cronograma PLANEJADO (o banco novo guarda isso de um jeito
    bem mais complexo, por subgrupo de serviço e percentual, sem
    equivalência direta ao "mês a mês" simples de antes), é a EXECUÇÃO real."""
    vistos = []
    for mes_ano in meses_cronograma:
        abrev = MESES_ABREV_MINUSCULO.get(mes_ano.split('/')[0])
        if abrev and abrev not in vistos:
            vistos.append(abrev)
    return '/'.join(vistos)


# ---------------------------------------------------------------------
# Geometria: camadas/R<região>_TRECHOS.shp, agrupada pelo campo Id.
# ---------------------------------------------------------------------
_cache_trechos = {}


def carregar_trechos_regiao(regiao_num):
    if regiao_num in _cache_trechos:
        return _cache_trechos[regiao_num]
    caminho = os.path.join(CAMADAS_DIR, f'R{regiao_num}_TRECHOS.shp')
    if not os.path.exists(caminho):
        qa(f'R{regiao_num}: shapefile R{regiao_num}_TRECHOS.shp não encontrado em camadas/')
        _cache_trechos[regiao_num] = {}
        return {}
    gdf = gpd.read_file(caminho)
    if gdf.crs is None:
        gdf = gdf.set_crs(EPSG_METRICO)
    gdf = gdf.to_crs(4326)
    col_id = next((c for c in gdf.columns if _norm(c) == 'ID'), None)
    if col_id is None:
        qa(f'R{regiao_num}: shapefile R{regiao_num}_TRECHOS.shp sem coluna Id reconhecível (colunas: {list(gdf.columns)})')
        _cache_trechos[regiao_num] = {}
        return {}
    por_trecho = {}
    for _, row in gdf.iterrows():
        try:
            n = int(row[col_id])
        except (TypeError, ValueError):
            continue
        if row.geometry is None or row.geometry.is_empty:
            continue
        por_trecho.setdefault(n, []).append(row.geometry)
    _cache_trechos[regiao_num] = por_trecho
    return por_trecho


# ---------------------------------------------------------------------
# CONFIG_CONTRATO — tabela chave/valor (coluna A = campo, coluna B = valor).
# Número do contrato = CodigoContrato (zero-padded a 3 dígitos — o banco
# às vezes guarda isso como int puro, perdendo o zero à esquerda, ex. "49"
# em vez de "049"; :03d restaura sem risco pros códigos de 4 dígitos como
# "1452"/"1151", que não são padded, só maiores mesmo) + "." + AnoContrato.
# ---------------------------------------------------------------------
def ler_contrato(wb):
    ws = wb['CONFIG_CONTRATO']
    config = {}
    for r in range(1, ws.max_row + 1):
        campo = ws.cell(row=r, column=1).value
        if campo:
            config[str(campo).strip()] = ws.cell(row=r, column=2).value
    codigo = config.get('CodigoContrato')
    ano = config.get('AnoContrato')
    if codigo is None or ano is None:
        return None
    return f'{int(_float_br(codigo)):03d}.{int(_float_br(ano))}'


# ---------------------------------------------------------------------
# ORCAMENTO_PADRAO (catálogo de preços) + ITENS_OSP (matriz item x O.S.P.)
# -> "serviço" composto por O.S.P. (ex. "Roçada / Tapa-Buraco CAE").
#
# ITENS_OSP: linha 1 tem números de O.S.P. como cabeçalho de coluna (a
# partir da coluna D); linhas a partir da 1ª com código tipo "X.X" ou
# "X.X.X" na coluna A têm a QUANTIDADE daquele item em cada O.S.P. O código
# da coluna A cruza com ORCAMENTO_PADRAO.Subitem (NÃO com "Código" nem com
# Item+Subitem concatenado — testado e confirmado: só Subitem bate, e só
# as linhas com Agrupador="ITEM" têm descrição de serviço de verdade;
# GRUPO/SUBGRUPO são só os totais de rollup, ignorados aqui).
# ---------------------------------------------------------------------
PADRAO_CODIGO_ITEM = re.compile(r'^\d+(\.\d+)+$')


def carregar_catalogo_itens(wb):
    if 'ORCAMENTO_PADRAO' not in wb.sheetnames:
        return {}
    ws = wb['ORCAMENTO_PADRAO']
    cols = _mapa_colunas(ws)
    col_agrupador, col_subitem, col_descricao = cols.get('AGRUPADOR'), cols.get('SUBITEM'), cols.get('DESCRICAO')
    if not (col_agrupador and col_subitem and col_descricao):
        return {}
    catalogo = {}
    for r in range(2, ws.max_row + 1):
        if _norm(ws.cell(row=r, column=col_agrupador).value) != 'ITEM':
            continue
        subitem = ws.cell(row=r, column=col_subitem).value
        descricao = ws.cell(row=r, column=col_descricao).value
        if subitem is None or not descricao:
            continue
        catalogo[str(subitem).strip()] = str(descricao).strip()
    return catalogo


def carregar_servicos_por_osp(wb, catalogo):
    if 'ITENS_OSP' not in wb.sheetnames or not catalogo:
        return {}
    ws = wb['ITENS_OSP']
    osp_por_coluna = {}
    for c in range(4, ws.max_column + 1):
        osp = _int_seguro(ws.cell(row=1, column=c).value)
        if osp is not None:
            osp_por_coluna[c] = osp

    primeira_linha_item = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is not None and PADRAO_CODIGO_ITEM.match(str(v).strip()):
            primeira_linha_item = r
            break
    if primeira_linha_item is None or not osp_por_coluna:
        return {}

    servicos = {}
    for r in range(primeira_linha_item, ws.max_row + 1):
        codigo_raw = ws.cell(row=r, column=1).value
        if codigo_raw is None:
            continue
        descricao = catalogo.get(str(codigo_raw).strip())
        if not descricao:
            continue  # rollup de GRUPO/SUBGRUPO ou código não catalogado
        for c, osp in osp_por_coluna.items():
            qtd = _float_br(ws.cell(row=r, column=c).value)
            if qtd and qtd > 0:
                servicos.setdefault(osp, []).append(descricao)

    # O banco novo é bem mais granular que a planilha antiga (O.S.P. com 10+
    # itens de orçamento, cada descrição já comprida sozinha — algumas
    # passam de 100 caracteres) — juntar tudo vira um texto ilegível
    # (quebrou o gráfico de ranking "Por serviço" mesmo limitando a 3
    # itens inteiros). Mostra só os 2 primeiros itens distintos, cada um
    # cortado em ~45 caracteres, + "(+N itens)" quando sobrar mais.
    resultado = {}
    for osp, descs in servicos.items():
        unicos = list(dict.fromkeys(descs))  # remove duplicata preservando ordem
        curtos = [d if len(d) <= 45 else d[:42].rstrip() + '...' for d in unicos[:2]]
        texto = ' / '.join(curtos)
        if len(unicos) > 2:
            texto += f' (+{len(unicos) - 2} itens)'
        resultado[osp] = texto
    return resultado


# ---------------------------------------------------------------------
# BD_MED — uma linha por O.S.P. por MÊS medido (série real de execução,
# confirmado com exemplo de 3 meses seguidos pra mesma O.S.P.). Vira:
#   - meses_cronograma: todo (mês/ano) em que essa O.S.P. teve medição —
#     usado no filtro de Mês (mesesDaEntrada() no index.html).
#   - medido_mensal: grade fixa Jan..Dez (mesmo formato de antes, pro KPI
#     "executado no mês"/"acumulado" do Painel Executivo continuar
#     funcionando sem mudar o index.html) — soma por NOME do mês,
#     ignorando ano (só relevante quando há 1 ano de medição ativo, que é
#     o caso de hoje; se algum dia cruzar virada de ano no meio da mesma
#     O.S.P., juntaria os dois anos no mesmo slot — não visto ainda).
# ---------------------------------------------------------------------
def carregar_medicoes(wb):
    if 'BD_MED' not in wb.sheetnames:
        return {}
    ws = wb['BD_MED']
    cols = _mapa_colunas(ws)
    col_osp, col_data, col_valor = cols.get('BD_OSP'), cols.get('DATA'), cols.get('MEDICAO 1')
    col_num, col_periodo, col_status = cols.get('MEDICAO'), cols.get('PERIODO'), cols.get('STATUS')
    if not (col_osp and col_data):
        return {}
    por_osp = {}
    for r in range(2, ws.max_row + 1):
        osp = _int_seguro(ws.cell(row=r, column=col_osp).value)
        if osp is None:
            continue
        mes_nome, ano = _parse_mes_ano(ws.cell(row=r, column=col_data).value)
        if mes_nome is None:
            continue
        valor = _float_br(ws.cell(row=r, column=col_valor).value) if col_valor else None
        status_raw = ws.cell(row=r, column=col_status).value if col_status else None
        status_chave = _norm(status_raw)
        periodo = ws.cell(row=r, column=col_periodo).value if col_periodo else None
        por_osp.setdefault(osp, []).append({
            'mes': mes_nome, 'ano': ano, 'valor': valor or 0,
            'n': _int_seguro(ws.cell(row=r, column=col_num).value) if col_num else None,
            'periodo': re.sub(r'\s+', ' ', str(periodo)).strip() if periodo else None,
            'status': STATUS_MAP.get(status_chave) or (str(status_raw).strip() if status_chave not in ('', '0') else None),
        })
    for registros in por_osp.values():
        registros.sort(key=lambda m: (m['ano'], MESES_ORDEM.index(m['mes'])))
    return por_osp


def montar_medido_mensal(registros):
    """Grade fixa Jan..Dez — mesmo formato consumido por atualizarKpisMes()
    no index.html desde a época da aba HISTÓRICO."""
    soma_por_mes = {nome: 0 for nome in MESES_ORDEM}
    for reg in registros:
        soma_por_mes[reg['mes']] += reg['valor']
    return [{'mes': MESES_PT[i + 1], 'valor': soma_por_mes[nome]} for i, nome in enumerate(MESES_ORDEM)]


def montar_meses_cronograma(registros):
    vistos = {}
    for reg in registros:
        chave = f"{reg['mes']}/{reg['ano']}"
        vistos[chave] = True
    return sorted(vistos.keys(), key=lambda s: (int(s.split('/')[1]), MESES_ORDEM.index(s.split('/')[0])))


# ---------------------------------------------------------------------
# MEDIÇÕES por O.S.P. (mostradas na gaveta de detalhe): BD_MED (1 linha por
# O.S.P. por medição: nº, período, valor medido, situação — a soma bate com
# MEDIDO de BD_OSP, conferido em 83 de 83 O.S.P. da Região 13) + BD_BOLETIM
# (nome dos PDFs do boletim de cada medição) + JUSTIFICATIVA_MED (texto da
# justificativa de cada medição). Os PDFs em si ficam no Drive — o geoportal
# só mostra quais documentos existem.
# ---------------------------------------------------------------------
# Regras em ordem: a primeira que casar (no nome SEM acento) define o rótulo.
# Os nomes dos arquivos vêm bagunçados ("MEMORIA", "MEMORIAL DE CALCULO",
# "RELATORIO OSP 091"...), então classifica por palavra-chave.
REGRAS_DOC_BOLETIM = [
    ('JUSTIFICATIVA', 'Justificativa técnica'),
    ('FOTOGRAF', 'Relatório fotográfico'),
    ('MEMORI', 'Memória de cálculo'),
    ('DIARIO', 'Diário de obras'),
    ('TECNOLOGICO', 'Controle tecnológico'),
    ('INSPECAO', 'Relatório de inspeção'),
    ('BOLETIM', 'Boletim da O.S.P.'),
    ('RELATORIO OSP', 'Relatório da O.S.P.'),
    ('MEDICAO', 'Relatório de medição'),
    ('RELATORIO', 'Relatório de medição'),
    ('ANEXO', 'Anexos'),
    ('COMPLEMENTAR', 'Anexos'),
]


def rotulo_doc_boletim(tipo_bruto):
    t = _norm(tipo_bruto.replace('.pdf', ''))
    if not t.strip():
        return 'Boletim de medição'
    if sum(k in t for k in ('MEMORI', 'FOTOGRAF', 'JUSTIFICATIVA')) >= 2:
        return 'Memória, fotos e justificativa'
    for chave, rotulo in REGRAS_DOC_BOLETIM:
        if chave in t:
            return rotulo
    return tipo_bruto.strip().capitalize()


def tipo_do_boletim(nome):
    sufixo = re.search(r'MEDI[ÇC][ÃA]O[^_]*_(.+?)(?:\.pdf)?$', str(nome).strip(), flags=re.I)
    return rotulo_doc_boletim(sufixo.group(1) if sufixo else '')


def carregar_boletins(wb, pasta_pdf=None):
    """{osp: {nº da medição: [{tipo, arquivo}, ...]}} — tipo = trecho final
    do nome do arquivo ("BOLETIM_OSP_0078_MED_11ª MEDIÇÃO_MEMORIA DE CALCULO.pdf"
    -> "Memória de cálculo")."""
    if 'BD_BOLETIM' not in wb.sheetnames:
        return {}
    ws = wb['BD_BOLETIM']
    # PDFs de fato presentes em LOTE XX/BOLETIM_PDF (a planilha lista nomes que
    # às vezes ainda não foram subidos, ou sem os acentos do arquivo real).
    # Chave sem acento/caixa -> nome real do arquivo no Drive.
    no_drive = {}
    if pasta_pdf and os.path.isdir(pasta_pdf):
        for f in os.listdir(pasta_pdf):
            no_drive[_norm(f)] = f
    por_osp = {}
    for r in range(1, ws.max_row + 1):
        osp = _int_seguro(ws.cell(row=r, column=1).value)
        nome = ws.cell(row=r, column=2).value
        if osp is None or not nome:
            continue
        m = re.search(r'MED_(\d+)', str(nome))
        if not m:
            continue
        # Tipo = o que vem depois de "MEDIÇÃO_" ("..._11ª MEDIÇÃO_MEMORIA DE
        # CALCULO.pdf"); boletim sem sufixo (só "..._13ª MEDIÇÃO.pdf") vira
        # "Boletim de medição". Os nomes dos arquivos vêm SEM acento, por
        # isso o rótulo passa por rotulo_doc_boletim().
        tipo = tipo_do_boletim(nome)
        arquivo = str(nome).strip()
        real = no_drive.get(_norm(arquivo))
        docs = por_osp.setdefault(osp, {}).setdefault(int(m.group(1)), [])
        if all(d['arquivo'] != (real or arquivo) for d in docs):
            # 'no_drive' = o PDF existe mesmo no Drive (nome real, com acentos);
            # senão fica o nome da planilha e o geoportal não oferece abrir.
            docs.append({'tipo': tipo, 'arquivo': real or arquivo, 'no_drive': bool(real)})
    # BD_BOLETIM nem sempre está preenchida (ex.: LOTE 16 tem os PDFs na pasta e
    # a aba vazia) — então também lê os nomes direto da pasta do Drive
    # ("BOLETIM_OSP_0005_MED_6ª MEDIÇÃO PARCIAL_....pdf"). INSP_CONSOL_* não tem
    # nº de O.S.P. no nome, só entra se a planilha listar.
    for real in no_drive.values():
        if not real.lower().endswith('.pdf'):
            continue
        m = re.search(r'OSP_(\d+)_MED_(\d+)', real)
        if not m:
            continue
        docs = por_osp.setdefault(int(m.group(1)), {}).setdefault(int(m.group(2)), [])
        if all(d['arquivo'] != real for d in docs):
            docs.append({'tipo': tipo_do_boletim(real), 'arquivo': real, 'no_drive': True})
    # A planilha às vezes lista o mesmo documento com o nome digitado errado
    # ("CÁCULO") e o arquivo real existe com outro nome: se já há um PDF
    # existente desse tipo na medição, descarta o item "sem arquivo".
    for medicoes in por_osp.values():
        for docs in medicoes.values():
            tipos_ok = {d['tipo'] for d in docs if d['no_drive']}
            docs[:] = [d for d in docs if d['no_drive'] or d['tipo'] not in tipos_ok]
    return por_osp


def carregar_justificativas_med(wb):
    """{(osp, nº da medição): texto da justificativa}."""
    if 'JUSTIFICATIVA_MED' not in wb.sheetnames:
        return {}
    ws = wb['JUSTIFICATIVA_MED']
    cols = _mapa_colunas(ws)
    c_osp, c_num, c_txt = cols.get('NUM_OSP'), cols.get('NUMERO MED'), cols.get('TEXTO_JUSTIFICATIVA')
    if not (c_osp and c_num and c_txt):
        return {}
    textos = {}
    for r in range(2, ws.max_row + 1):
        osp, n = _int_seguro(ws.cell(row=r, column=c_osp).value), _int_seguro(ws.cell(row=r, column=c_num).value)
        txt = ws.cell(row=r, column=c_txt).value
        if osp is not None and n is not None and txt and str(txt).strip():
            textos[(osp, n)] = re.sub(r'[ \t]+', ' ', str(txt)).strip()
    return textos


def montar_medicoes(osp, registros, boletins, justificativas):
    medicoes = []
    for reg in registros:
        n = reg.get('n')
        medicoes.append({
            'n': n, 'mes': f"{reg['mes']}/{reg['ano']}", 'periodo': reg.get('periodo'),
            'valor': reg['valor'], 'status': reg.get('status'),
            'docs': boletins.get(osp, {}).get(n, []),
            'justificativa': justificativas.get((osp, n)),
        })
    return medicoes


# ---------------------------------------------------------------------
# CHECKLIST DA SUPERVISÃO (aba CHECKLIST_OSP do sistema principal, que a
# equipe demorava a abrir) — dá a lista do que precisa ser corrigido quando
# a O.S.P. está em "Correção Super".
#
# As RESPOSTAS ficam em BD_CHECK_OSP de cada banco: cabeçalho da linha 1 =
# nº da O.S.P. (uma coluna por O.S.P.), cada linha a partir da 10 = 1 item
# do checklist, célula no formato "SIM|NÃO|NA|observação" (ex. "X|||" = SIM;
# "|X||FOTOS INAPLICÁVEIS" = NÃO com observação; "||X|" = NA; "|||" = não
# respondido). O TEXTO dos itens (1.1, 1.2...) não está no banco — vem da
# aba CHECKLIST_OSP de "Sistema de Gestão OSP – AGETO_V*.xlsm": a linha N
# dessa aba corresponde à linha N-3 de BD_CHECK_OSP (conferido contra o
# checklist real da O.S.P. 051.2024.0163: itens 1.8 em NÃO, 1.3 e 2.2 em NA).
# ---------------------------------------------------------------------
OFFSET_LINHA_CHECKLIST = 3
CACHE_CHECKLIST = os.path.join(DADOS_DIR, 'checklist_itens.json')
_catalogo_checklist = None


def carregar_catalogo_checklist():
    """{linha_BD_CHECK_OSP: (codigo, descricao, eh_titulo_de_grupo)}. Lê do
    sistema principal (arquivo grande, ~8 MB — só 1 vez por rodada) e guarda
    cópia em dados/checklist_itens.json; se o sistema não estiver acessível,
    usa a cópia da última rodada."""
    global _catalogo_checklist
    if _catalogo_checklist is not None:
        return _catalogo_checklist
    catalogo = {}
    mestres = sorted(
        f for f in glob.glob(os.path.join(BASE_SISTEMA, 'Sistema de Gest*.xlsm'))
        if not os.path.basename(f).startswith('~$')
    )
    if mestres:
        try:
            wm = openpyxl.load_workbook(mestres[-1], data_only=True, keep_vba=False, read_only=True)
            ws = wm['CHECKLIST_OSP']
            for n, (codigo, descricao) in enumerate(
                    ws.iter_rows(min_row=1, max_row=400, min_col=2, max_col=3, values_only=True), start=1):
                if n < 13 or codigo in (None, '') or not descricao:
                    continue
                codigo = str(codigo).strip()
                catalogo[n - OFFSET_LINHA_CHECKLIST] = (codigo, str(descricao).strip(), codigo.endswith('.'))
            wm.close()
            with open(CACHE_CHECKLIST, 'w', encoding='utf-8') as f:
                json.dump({str(k): list(v) for k, v in catalogo.items()}, f, ensure_ascii=False)
        except Exception as e:
            qa(f'{mestres[-1]}: falha ao ler CHECKLIST_OSP ({e}) — tentando cópia da última rodada')
            catalogo = {}
    if not catalogo and os.path.exists(CACHE_CHECKLIST):
        with open(CACHE_CHECKLIST, encoding='utf-8') as f:
            catalogo = {int(k): tuple(v) for k, v in json.load(f).items()}
    if not catalogo:
        qa('Texto dos itens do checklist indisponível (sistema principal não acessível e sem cópia) — checklist não gerado')
    _catalogo_checklist = catalogo
    return catalogo


def carregar_checklists(wb):
    """{osp: {'n_sim','n_nao','n_na','pendencias':[{item,descricao,obs}],
    'observacoes':[{item,resposta,descricao,obs}]}} — só O.S.P. que tiveram
    pelo menos 1 item respondido. 'pendencias' = itens marcados NÃO (o que a
    supervisão precisa corrigir); 'observacoes' = SIM/NA com comentário."""
    catalogo = carregar_catalogo_checklist()
    if not catalogo or 'BD_CHECK_OSP' not in wb.sheetnames:
        return {}
    ws = wb['BD_CHECK_OSP']
    resultado = {}
    for c in range(2, ws.max_column + 1):
        osp = _int_seguro(ws.cell(row=1, column=c).value)
        if osp is None:
            continue
        info = {'n_sim': 0, 'n_nao': 0, 'n_na': 0, 'pendencias': [], 'observacoes': []}
        for linha in range(10, ws.max_row + 1):
            item = catalogo.get(linha)
            if not item or item[2]:  # sem texto no catálogo, ou título de grupo
                continue
            bruto = ws.cell(row=linha, column=c).value
            if not isinstance(bruto, str) or '|' not in bruto:
                continue
            partes = bruto.split('|')
            partes += [''] * (4 - len(partes))
            sim, nao, na = (partes[i].strip() != '' for i in range(3))
            obs = '|'.join(partes[3:]).strip()
            if not (sim or nao or na):
                continue
            codigo, descricao, _ = item
            if nao:
                info['n_nao'] += 1
                info['pendencias'].append({'item': codigo, 'descricao': descricao, 'obs': obs})
            else:
                info['n_sim' if sim else 'n_na'] += 1
                if obs:
                    info['observacoes'].append({'item': codigo, 'resposta': 'SIM' if sim else 'NA',
                                                'descricao': descricao, 'obs': obs})
        if info['n_sim'] + info['n_nao'] + info['n_na']:
            resultado[osp] = info
    return resultado


# ---------------------------------------------------------------------
# O.S.P. CRIADAS MAS AINDA SEM EMISSÃO — têm levantamento/inventário
# (BD_INVENTARIO) mas não foram lançadas em BD_OSP. Só os arquivos existem:
# sem valor, data nem situação. Regras (investigadas em 2026-10-02):
#  - O nº real da O.S.P. vem do NOME do arquivo quando tem o padrão
#    "042.2025.0004" — o índice às vezes registra o nº errado (Região 22
#    listava a O.S.P. 0004 como "1"; como a 4 já está em BD_OSP, não é nova).
#  - Só entra número DEPOIS do primeiro cadastrado do contrato: a Região 02
#    começa na 106, e o inventário cita 2, 3, 88, 102, 104 (numeração de
#    antes, de contrato anterior) — esses ficam de fora e vão pro relatório.
# ---------------------------------------------------------------------
def carregar_inventario(wb):
    if 'BD_INVENTARIO' not in wb.sheetnames:
        return {}
    ws = wb['BD_INVENTARIO']
    por_osp = {}
    for r in range(1, ws.max_row + 1):
        chave = _int_seguro(ws.cell(row=r, column=1).value)
        nome = ws.cell(row=r, column=2).value
        if chave is None or not nome:
            continue
        nome = str(nome).strip()
        m = re.search(r'\d{3,4}\.\d{4}\.(\d{4})', nome)
        arquivos = por_osp.setdefault(int(m.group(1)) if m else chave, [])
        if nome not in arquivos:
            arquivos.append(nome)
    return por_osp


def carregar_nomes_trechos(wb):
    """{nº do trecho: nome} da aba TRECHOS (colunas D e E, uma linha por S.R.E.)."""
    nomes = {}
    if 'TRECHOS' in wb.sheetnames:
        ws = wb['TRECHOS']
        for r in range(2, ws.max_row + 1):
            n = _int_seguro(ws.cell(row=r, column=4).value)
            if n is not None and n not in nomes:
                nomes[n] = str(ws.cell(row=r, column=5).value or '').strip()
    return nomes


# ---------------------------------------------------------------------
# BD_OSP — uma linha por O.S.P.+trecho. Fonte principal: data de emissão,
# trecho, situação, valor previsto/executado.
# ---------------------------------------------------------------------
def processar_banco(caminho, regiao_num_planilha, resultados_por_regiao):
    print(f'Lendo {os.path.basename(os.path.dirname(caminho))}/{os.path.basename(caminho)} ...')
    wb = openpyxl.load_workbook(caminho, data_only=True, keep_vba=False)

    contrato = ler_contrato(wb)
    if not contrato:
        qa(f'{caminho}: CONFIG_CONTRATO sem CodigoContrato/AnoContrato — arquivo ignorado')
        return

    if 'BD_OSP' not in wb.sheetnames:
        qa(f'{caminho}: aba BD_OSP não encontrada — arquivo ignorado')
        return
    ws = wb['BD_OSP']
    cols = _mapa_colunas(ws)
    obrigatorias = ['BD_OSP', 'DATA', 'TRECHO_N', 'DESCRICAO', 'STATUS', 'VALOR', 'MEDIDO']
    if not all(c in cols for c in obrigatorias):
        qa(f'{caminho}: aba BD_OSP sem as colunas esperadas (achadas: {sorted(cols.keys())}) — arquivo ignorado')
        return

    catalogo_itens = carregar_catalogo_itens(wb)
    servicos_por_osp = carregar_servicos_por_osp(wb, catalogo_itens)
    medicoes_por_osp = carregar_medicoes(wb)
    checklists_por_osp = carregar_checklists(wb)
    boletins_por_osp = carregar_boletins(wb, os.path.join(os.path.dirname(caminho), 'BOLETIM_PDF'))
    justificativas_med = carregar_justificativas_med(wb)

    tipo_servico = 'restauracao' if regiao_num_planilha in REGIOES_RESTAURACAO else 'manutencao'
    regiao_geo = MAPA_REGIAO_GEOGRAFICA.get(regiao_num_planilha, regiao_num_planilha)
    regiao_geo_label = f'R{regiao_geo}'

    n_lidos = n_pulados_sem_trecho = n_sem_geometria = 0
    osps_cadastradas = set()

    for r in range(2, ws.max_row + 1):
        osp = _int_seguro(ws.cell(row=r, column=cols['BD_OSP']).value)
        if osp is None:
            continue
        osps_cadastradas.add(osp)

        trecho_num = _int_seguro(ws.cell(row=r, column=cols['TRECHO_N']).value)
        trecho_nome = ws.cell(row=r, column=cols['DESCRICAO']).value
        sem_trecho_cadastrado = trecho_num is None or trecho_num == 0
        if sem_trecho_cadastrado:
            n_pulados_sem_trecho += 1
            trecho_num = None

        geoms = None
        if not sem_trecho_cadastrado:
            geoms = carregar_trechos_regiao(regiao_geo).get(trecho_num)
            if not geoms:
                qa(f'{regiao_geo_label} (via {os.path.basename(caminho)}): Trecho {trecho_num} '
                   f'({trecho_nome}) não encontrado no shapefile R{regiao_geo}_TRECHOS.shp — '
                   f'mantido na lista/KPIs, sem geometria no mapa')
                n_sem_geometria += 1

        mes_nome, ano = _parse_mes_ano(ws.cell(row=r, column=cols['DATA']).value)
        data_emissao = f'{mes_nome}/{ano}' if mes_nome else None

        registros_med = medicoes_por_osp.get(osp, [])
        meses_cronograma = montar_meses_cronograma(registros_med)
        medido_mensal = montar_medido_mensal(registros_med)

        situacao_raw = ws.cell(row=r, column=cols['STATUS']).value
        situacao_chave = _norm(situacao_raw)
        situacao = STATUS_MAP.get(situacao_chave) or (
            'Não informada' if situacao_chave in ('', '0') else str(situacao_raw).strip()
        )

        valor_previsto = _float_br(ws.cell(row=r, column=cols['VALOR']).value) or 0
        valor_executado = _float_br(ws.cell(row=r, column=cols['MEDIDO']).value) or 0

        prazo_meses = _int_seguro(ws.cell(row=r, column=cols['PRAZO']).value) if 'PRAZO' in cols else None
        prazo_limite, atrasada = calcular_prazo(mes_nome, ano, prazo_meses, situacao, meses_cronograma)

        teve_med_justificada = cols.get('TEVE MED JUSTIFICADA')
        observacao = None
        if teve_med_justificada and _norm(ws.cell(row=r, column=teve_med_justificada).value) == 'SIM':
            observacao = 'Teve medição justificada'

        props = {
            'regiao': regiao_geo_label,
            'regiao_os': f'R{regiao_num_planilha:02d}',
            'tipo_servico': tipo_servico,
            'contrato': contrato,
            'osp': osp,
            'data_emissao': data_emissao,
            'cronograma': _cronograma_legado(meses_cronograma),
            'meses_cronograma': meses_cronograma,
            'trecho_num': trecho_num,
            'trecho_nome': trecho_nome,
            'servico': servicos_por_osp.get(osp),
            'valor_previsto': valor_previsto,
            'situacao': situacao,
            'observacao': observacao,
            'valor_executado': valor_executado,
            'saldo': valor_previsto - valor_executado,
            'pct_executado': (valor_executado / valor_previsto * 100) if valor_previsto > 0 else None,
            'situacao_final': situacao,
            'medido_mensal': medido_mensal,
            'prazo_meses': prazo_meses,
            'prazo_limite': prazo_limite,
            'atrasada': atrasada,
            'checklist': checklists_por_osp.get(osp),
            'medicoes': montar_medicoes(osp, registros_med, boletins_por_osp, justificativas_med),
        }

        alvo = (resultados_por_regiao.setdefault(regiao_geo_label, []))
        if geoms:
            for geom in geoms:
                alvo.append({'type': 'Feature', 'geometry': mapping(geom), 'properties': props})
        else:
            alvo.append({'type': 'Feature', 'geometry': None, 'properties': props})
        n_lidos += 1

    # O.S.P. criadas mas ainda sem emissão (só têm levantamento/inventário).
    n_sem_emissao = 0
    inventario = carregar_inventario(wb)
    nomes_trechos = carregar_nomes_trechos(wb)
    primeira = min(osps_cadastradas) if osps_cadastradas else None
    for osp in sorted(inventario):
        if osp in osps_cadastradas or primeira is None:
            continue
        arquivos = inventario[osp]
        if osp < primeira:
            qa(f'{contrato} (R{regiao_num_planilha:02d}): O.S.P. {osp} tem levantamento/inventário mas o número é anterior '
               f'ao 1º cadastrado do contrato ({primeira}) — ignorada (provável numeração de contrato anterior)')
            continue
        trecho_num = next((int(m.group(1)) for a in arquivos
                           for m in [re.search(r'TRECHO\s*-?\s*(\d+)', a, re.I)] if m), None)
        geoms = carregar_trechos_regiao(regiao_geo).get(trecho_num) if trecho_num else None
        props = {
            'regiao': regiao_geo_label, 'regiao_os': f'R{regiao_num_planilha:02d}', 'tipo_servico': tipo_servico,
            'contrato': contrato, 'osp': osp, 'data_emissao': None, 'cronograma': '', 'meses_cronograma': [],
            'trecho_num': trecho_num, 'trecho_nome': nomes_trechos.get(trecho_num) if trecho_num else None,
            'servico': servicos_por_osp.get(osp), 'valor_previsto': 0, 'situacao': 'Sem emissão',
            'observacao': None, 'valor_executado': 0, 'saldo': 0, 'pct_executado': None,
            'situacao_final': 'Sem emissão', 'medido_mensal': montar_medido_mensal([]),
            'prazo_meses': None, 'prazo_limite': None, 'atrasada': None,
            'checklist': checklists_por_osp.get(osp),
            'sem_emissao': True, 'inventario': arquivos,
        }
        alvo = resultados_por_regiao.setdefault(regiao_geo_label, [])
        if geoms:
            for geom in geoms:
                alvo.append({'type': 'Feature', 'geometry': mapping(geom), 'properties': props})
        else:
            alvo.append({'type': 'Feature', 'geometry': None, 'properties': props})
        n_sem_emissao += 1

    print(f'  {os.path.basename(caminho)}: {n_lidos} O.S.P. lida(s), {n_pulados_sem_trecho} sem trecho cadastrado ainda, '
          f'{n_sem_geometria} sem geometria no shapefile, {n_sem_emissao} criada(s) sem emissão')


def pastas_do_lote(numero):
    if numero == 3:
        return [os.path.join(BASE_SISTEMA, nome) for nome in PASTAS_REGIAO_3]
    candidatas = sorted(glob.glob(os.path.join(BASE_SISTEMA, f'LOTE {numero:02d}*')))
    # Evita "LOTE 2X" casar com "LOTE 2" (sem zero) por engano — compara o
    # número logo após "LOTE " caractere a caractere.
    return [c for c in candidatas if re.match(rf'^LOTE 0*{numero}(\D|$)', os.path.basename(c))]


def arquivo_bd_da_pasta(pasta):
    candidatos = [
        f for f in glob.glob(os.path.join(pasta, 'BD_LOTE_*.xlsx'))
        if not os.path.basename(f).startswith('~$') and 'COPIA' not in _norm(os.path.basename(f))
    ]
    if not candidatos:
        qa(f'{pasta}: nenhum BD_LOTE_*.xlsx encontrado (ignorando "- Copia")')
        return None
    if len(candidatos) > 1:
        qa(f'{pasta}: mais de um BD_LOTE_*.xlsx encontrado ({candidatos}) — usando o primeiro')
    return candidatos[0]


def main():
    if not os.path.isdir(BASE_SISTEMA):
        print(f'Pasta do SISTEMA_AGETO não encontrada: {BASE_SISTEMA}')
        print('Confirme se o Google Drive está montado em G:\\ e sincronizado.')
        return

    # Regiões de manutenção (1,2,3,11,12,13) + restauração (14,15,16,22,24)
    # — 23 fica de fora, não existe banco ainda (ver topo do arquivo).
    numeros_regiao = [1, 2, 3, 11, 12, 13, 14, 15, 16, 22, 24]

    geral = {}
    for numero in numeros_regiao:
        pastas = pastas_do_lote(numero)
        if not pastas:
            qa(f'Região {numero}: nenhuma pasta "LOTE {numero:02d}*" encontrada — pulando')
            continue
        for pasta in pastas:
            caminho = arquivo_bd_da_pasta(pasta)
            if not caminho:
                continue
            try:
                processar_banco(caminho, numero, geral)
            except Exception as e:
                qa(f'{caminho}: falha ao ler ({e}) — arquivo ignorado')

    os.makedirs(DADOS_DIR, exist_ok=True)
    total_osp = 0
    for regiao, feats in sorted(geral.items()):
        pacote = {'type': 'FeatureCollection', 'features': feats}
        nome_arquivo = f'os_{regiao}.js'
        with open(os.path.join(DADOS_DIR, nome_arquivo), 'w', encoding='utf-8') as f:
            f.write('// Gerado por converter_os.py — não editar à mão\n')
            f.write('window.DADOS_OS_REGIAO = window.DADOS_OS_REGIAO || {};\n')
            f.write(f'window.DADOS_OS_REGIAO["{regiao}"] = {json.dumps(pacote, ensure_ascii=False)};\n')
        n_osp = len(set((f['properties']['contrato'], f['properties']['osp'], f['properties']['trecho_num'])
                        for f in feats))
        total_osp += n_osp
        print(f'  -> dados/{nome_arquivo} ({len(feats)} feature(s), {n_osp} O.S.P.+trecho)')

    with open(os.path.join(DADOS_DIR, 'manifest_os.js'), 'w', encoding='utf-8') as f:
        f.write('// Gerado por converter_os.py — não editar à mão\n')
        f.write('window.MANIFEST_OS = ' + json.dumps(sorted(geral.keys()), ensure_ascii=False) + ';\n')

    with open(os.path.join(BASE, 'relatorio_qualidade_os.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(qa_msgs) + '\n' if qa_msgs else 'Nenhum problema encontrado.\n')

    print(f'\n{total_osp} O.S.P.+trecho no total, {len(qa_msgs)} observação(ões) — ver relatorio_qualidade_os.txt')


if __name__ == '__main__':
    main()
