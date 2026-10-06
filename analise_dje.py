"""Leitura e análise do texto do DJE: divisão em atos, portarias de cargo/função e apostilas."""
from __future__ import annotations

import os
import re
import bisect
import unicodedata
from dataclasses import dataclass
from datetime import date
from functools import lru_cache

import fitz  # PyMuPDF

# ==========================================
# TERMOS DE BUSCA
# ==========================================

TERMOS_CARGOS = re.compile(r'(CAI-|DAI-|DAS-)', re.IGNORECASE)
TERMOS_ACAO = re.compile(r'(nomeação|nomear|designação|designar|dispensa|dispensar|exoneração|exonerar)', re.IGNORECASE)
# Matrícula iniciada por 5 com no mínimo 4 dígitos, aceitando "." e "-" entre eles (ex.: 5012345, 5.012.345, 501234-5)
TERMO_MATRICULA = re.compile(r'matr[a-zí\.]*[^\d]{0,20}\b(5(?:[.\-]?\d){3,})\b', re.IGNORECASE)

# Início de cada ato publicado. A linha precisa conter só o título (ex.: "PORTARIA CGJ nº 2331/2026",
# "APOSTILA (SEI ...)", "DECISÃO"); citações como "PORTARIA Nº 2046/2026, publicada no DJERJ..." não abrem ato.
# A linha "id: 17273256", que o DJERJ põe antes de cada publicação, também separa os atos.
TITULO_ATO = re.compile(
    r'^[ \t]*(?:'
    r'(?P<portaria>PORTARIA\b(?![^\n]*[Pp]ublicad)[^\n]*)'
    r'|(?P<apostila>(?:APOSTILA(?:MENTO)?|Apostila(?:mento)?)\b[^\n]{0,60})'
    r'|(?P<outro>(?:DECIS[ÃA]O|ATO\s+(?:EXECUTIVO|NORMATIVO)|AVISO|ORDEM\s+DE\s+SERVI[ÇC]O|RESOLU[ÇC][ÃA]O|EDITAL'
    r'|DESPACHO|PROVIMENTO|COMUNICADO)\b[^\n]*|Decis[ãa]o)'
    r'|(?P<id>id:\s*\d+)'
    r')[ \t]*$',
    re.MULTILINE
)

TORNAR_SEM_EFEITO = re.compile(r'\b(?:torn\w*|fica[mr]?)\s+(?:\w+\s+)?sem\s+efeito', re.IGNORECASE)
REF_PORTARIA = re.compile(r'Portaria\b[^\d]{0,25}?n[º°o.]*\s*(\d[\d.]*\d|\d)(?:\s*/\s*(\d{4}))?', re.IGNORECASE)
NUMERO_TITULO = re.compile(r'PORTARIA[^\d\n]*?(\d[\d.]*\d|\d)(?:\s*/\s*(\d{4}))?')
SIMBOLO = re.compile(r'\b(CAI|DAI|DAS)\s*-\s*(\d+)', re.IGNORECASE)
# Indica que o ato citado é de cargo em comissão ou função gratificada ("cargo efetivo" não conta)
INDICIO_CARGO_FUNCAO = re.compile(r'\b(?:CAI|DAI|DAS)\s*-|cargo\s+em\s+comiss[ãa]o|fun[çc][ãa]o\s+gratificada', re.IGNORECASE)
REMOCAO_LOTACAO = re.compile(r'referente\s+(?:à|a|ao)s?\s+(?:remo[çc][ãa]o|lota[çc][ãa]o)', re.IGNORECASE)

# Onde começa e termina o trecho que a apostila efetivamente altera
INICIO_ALTERACAO = re.compile(
    r'fica(?:m|ndo)?\s+(?:declarad[oa]|esclarecid[oa]|retificad[oa])|devendo\s+constar|para\s+constar|onde\s+se\s+l[êe]',
    re.IGNORECASE
)
FIM_ALTERACAO = re.compile(r'(?:e\s+)?n[ãa]o\s+como\s+constou|mantid[oa]s|mantendo-se', re.IGNORECASE)
ONDE_SE_LE = re.compile(
    r'onde\s+se\s+l[êe]\b(?P<antes>.*?)leia-se(?P<depois>.*?)'
    r'(?=' + FIM_ALTERACAO.pattern + r'|⟦|Publique|$)',
    re.IGNORECASE | re.DOTALL
)

MESES = {"janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
         "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}
DATA = re.compile(
    r'(\d{1,2})º?\s*[/.]\s*(\d{1,2})\s*[/.]\s*(\d{4})'
    r'|(\d{1,2})º?\s+de\s+(' + '|'.join(MESES) + r')\s+de\s+(\d{4})',
    re.IGNORECASE
)
A_CONTAR = re.compile(r'a\s+contar\s+(?:de|do\s+dia|da\s+data\s+de)?\s*', re.IGNORECASE)

# Trechos que o PDF traz desenhados (contorno das letras) em vez de texto. Acontece em páginas geradas pelo
# "Imprimir em PDF" do Windows; o DJE original traz texto normal. São lidos por OCR quando ele está instalado.
MARCA_ILEGIVEL = "⟦trecho ilegível⟧"
MARCA_OCR_INICIO, MARCA_OCR_FIM = "⟪", "⟫"
LARGURA_MINIMA_ILEGIVEL = 25  # pontos; abaixo disso costuma ser só um travessão ou aspas

CONTA, NAO_CONTA, VERIFICAR = "CONTA", "NÃO CONTA", "VERIFICAR"

# ==========================================
# LEITURA DO PDF
# ==========================================

def listar_pdfs(pasta):
    return [f for f in sorted(os.listdir(pasta)) if f.lower().endswith(".pdf")]

def _trechos_desenhados(pagina):
    """Retângulos das linhas de texto que o PDF desenhou como contorno, sem texto extraível."""
    # Letra desenhada: forma preenchida, curva e pequena (exclui as linhas que separam as seções)
    glifos = [d["rect"] for d in pagina.get_drawings()
              if d.get("fill") is not None and len(d["items"]) > 3 and d["rect"].height < 20 and d["rect"].width < 40]
    linhas = []  # glifos agrupados pela altura (mesma linha do texto)
    for r in sorted(glifos, key=lambda r: (r.y0 + r.y1) / 2):
        centro = (r.y0 + r.y1) / 2
        if linhas and abs(centro - linhas[-1][0]) < 5:
            linhas[-1][1].append(r)
        else:
            linhas.append([centro, [r]])

    trechos = []
    for _, rets in linhas:
        rets.sort(key=lambda r: r.x0)
        atual = fitz.Rect(rets[0])
        for r in rets[1:]:
            if r.x0 - atual.x1 < 15:
                atual |= r
            else:
                trechos.append(atual)
                atual = fitz.Rect(r)
        trechos.append(atual)
    return [t for t in trechos if t.width >= LARGURA_MINIMA_ILEGIVEL]

_motor_ocr = None

def _ocr():
    """Motor de OCR (RapidOCR), carregado só quando aparece o primeiro trecho desenhado.

    O OCR é opcional: se a biblioteca não estiver instalada, retorna None e o trecho fica como MARCA_ILEGIVEL.
    """
    global _motor_ocr
    if _motor_ocr is None:
        try:
            from rapidocr import RapidOCR
            _motor_ocr = RapidOCR()
        except Exception:
            _motor_ocr = False
    return _motor_ocr or None

def ocr_disponivel():
    return _ocr() is not None

def _ler_trecho_desenhado(pagina, trecho):
    """Texto do trecho desenhado lido por OCR, entre MARCA_OCR_INICIO e MARCA_OCR_FIM; MARCA_ILEGIVEL se não der."""
    motor = _ocr()
    if motor:
        try:
            area = fitz.Rect(trecho.x0 - 3, trecho.y0 - 3, trecho.x1 + 3, trecho.y1 + 3)
            resultado = motor(pagina.get_pixmap(dpi=300, clip=area).tobytes("png"))
            partes = [t for t, nota in zip(resultado.txts or (), resultado.scores or ()) if nota >= 0.5]
            if partes:
                return MARCA_OCR_INICIO + " ".join(partes) + MARCA_OCR_FIM
        except Exception:
            pass
    return MARCA_ILEGIVEL

def _texto_da_pagina(pagina):
    """Mesmo texto de pagina.get_text(), com o conteúdo dos trechos desenhados lido por OCR (ou MARCA_ILEGIVEL)."""
    linhas = [(fitz.Rect(linha["bbox"]), "".join(s["text"] for s in linha["spans"]))
              for bloco in pagina.get_text("dict")["blocks"] for linha in bloco.get("lines", [])]
    inseridos = {}  # índice da linha -> textos a inserir depois dela (-1: antes da primeira linha)
    for trecho in _trechos_desenhados(pagina):
        sobrepostas = [i for i, (r, _) in enumerate(linhas) if r.y0 < trecho.y1 and trecho.y0 < r.y1]
        if sobrepostas:
            i = sobrepostas[0]
        else:
            seguintes = [i for i, (r, _) in enumerate(linhas) if r.y0 >= trecho.y1]
            i = (seguintes[0] - 1) if seguintes else len(linhas) - 1
        inseridos.setdefault(i, []).append(_ler_trecho_desenhado(pagina, trecho))

    saida = list(inseridos.get(-1, []))
    for i, (_, texto) in enumerate(linhas):
        saida.append(texto)
        saida.extend(inseridos.get(i, []))
    return "\n".join(saida) + "\n" if saida else ""

def ler_paginas_pdf(caminho_pdf):
    # O texto lido fica em memória: rodar a Aba 1 e depois a Aba 5 não relê os mesmos PDFs.
    # A data de modificação e o tamanho entram na chave para que um arquivo substituído seja relido.
    info = os.stat(caminho_pdf)
    return list(_ler_paginas_em_cache(os.path.abspath(caminho_pdf), info.st_mtime, info.st_size))

@lru_cache(maxsize=64)
def _ler_paginas_em_cache(caminho_pdf, _mtime, _tamanho):
    with fitz.open(caminho_pdf) as doc:
        return tuple(_texto_da_pagina(pagina) for pagina in doc)

# ==========================================
# DIVISÃO EM ATOS
# ==========================================

@dataclass
class Ato:
    tipo: str      # "portaria", "apostila" ou "outro"
    texto: str
    inicio: int    # posição no texto completo do documento

def dividir_em_atos(textos_paginas):
    """Junta as páginas de um PDF e divide o texto em atos (portarias, apostilas, decisões etc.).

    Retorna a lista de atos e a posição em que cada página começa no texto completo,
    para que um trecho do ato possa ser associado à página de origem.
    """
    texto = ""
    inicios_pagina = []
    for texto_pagina in textos_paginas:
        inicios_pagina.append(len(texto))
        texto += texto_pagina + "\n"

    cortes = [(0, "outro")] + [(m.start(), m.lastgroup) for m in TITULO_ATO.finditer(texto)] + [(len(texto), None)]
    atos = []
    for (ini, tipo), (fim, _) in zip(cortes, cortes[1:]):
        if fim > ini:
            atos.append(Ato("outro" if tipo == "id" else tipo, texto[ini:fim], ini))
    return atos, inicios_pagina

def numero_da_pagina(inicios_pagina, posicao):
    return bisect.bisect_right(inicios_pagina, posicao)  # numeração começando em 1

def paginas_do_ato(ato, inicios_pagina):
    fim = ato.inicio + len(ato.texto.rstrip()) - 1
    return list(range(numero_da_pagina(inicios_pagina, ato.inicio), numero_da_pagina(inicios_pagina, fim) + 1))

def eh_portaria_de_cargo(texto):
    return bool(TERMOS_CARGOS.search(texto) and TERMOS_ACAO.search(texto))

def resumir(texto, limite=300):
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto if len(texto) <= limite else texto[:limite].rstrip() + "…"

# ==========================================
# DATAS E SÍMBOLOS
# ==========================================

def _converter_data(m):
    try:
        if m.group(1):
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        return date(int(m.group(6)), MESES[m.group(5).lower()], int(m.group(4)))
    except ValueError:
        return None

def primeira_data(texto, limite=None):
    m = DATA.search(texto, 0, limite if limite is not None else len(texto))
    return _converter_data(m) if m else None

def datas_a_contar(texto):
    datas = set()
    for m in A_CONTAR.finditer(texto):
        d = DATA.match(texto, m.end())
        if d and _converter_data(d):
            datas.add(_converter_data(d))
    return datas

def simbolos(texto):
    return {f"{m.group(1).upper()}-{m.group(2)}" for m in SIMBOLO.finditer(texto)}

def formatar(d):
    return d.strftime("%d/%m/%Y")

# ==========================================
# BUSCA DA PORTARIA REFERENCIADA
# ==========================================

@dataclass
class Referencia:
    numero: str
    ano: str | None
    data: date | None

    def __str__(self):
        numero = f"{self.numero}/{self.ano}" if self.ano else self.numero
        return f"Portaria nº {numero}" + (f", publicada em {formatar(self.data)}" if self.data else ", data de publicação não identificada")

def _referencia(m, data):
    return Referencia(m.group(1).replace(".", ""), m.group(2), data)

class BuscaPortarias:
    """Procura portarias publicadas em dias anteriores nas pastas informadas (DO e acervo)."""

    def __init__(self, pastas):
        self.pastas = []
        for pasta in pastas:
            if pasta and os.path.isdir(pasta) and os.path.abspath(pasta) not in [os.path.abspath(p) for p in self.pastas]:
                self.pastas.append(pasta)
        self._portarias_por_dia = {}

    def _portarias_do_dia(self, dia):
        if dia not in self._portarias_por_dia:
            prefixo = dia.strftime("%Y%m%d")
            encontradas = []
            for pasta in self.pastas:
                for arquivo in listar_pdfs(pasta):
                    if not arquivo.startswith(prefixo):
                        continue
                    try:
                        atos, inicios = dividir_em_atos(ler_paginas_pdf(os.path.join(pasta, arquivo)))
                    except Exception:
                        continue
                    for ato in atos:
                        if ato.tipo == "portaria":
                            encontradas.append((arquivo, numero_da_pagina(inicios, ato.inicio), ato))
            self._portarias_por_dia[dia] = encontradas
        return self._portarias_por_dia[dia]

    def buscar(self, ref):
        """Retorna (ato, descrição de onde foi achado) ou (None, motivo)."""
        if not ref.data:
            return None, "data de publicação da portaria original não identificada no texto"
        portarias = self._portarias_do_dia(ref.data)
        if not portarias:
            return None, f"DJE de {formatar(ref.data)} não está na pasta DO nem no acervo"
        for arquivo, pagina, ato in portarias:
            m = NUMERO_TITULO.search(ato.texto.split("\n", 1)[0])
            if m and m.group(1).replace(".", "") == ref.numero and not (ref.ano and m.group(2) and m.group(2) != ref.ano):
                return ato, f"{arquivo}, p. {pagina}"
        return None, f"portaria não localizada no DJE de {formatar(ref.data)}"

# ==========================================
# AVALIAÇÃO DE APOSTILAS E "TORNAR SEM EFEITO"
# ==========================================

def _referencias(ato, cancelamento):
    """Portarias citadas pela apostila (a primeira) ou pelo ato que as torna sem efeito (todas após a expressão)."""
    if cancelamento:
        inicio = TORNAR_SEM_EFEITO.search(ato.texto).end()
        trecho = ato.texto[inicio:]
        publicada = re.search(r'publicad', trecho, re.IGNORECASE)
        data_comum = primeira_data(trecho[publicada.start():]) if publicada else None
        refs = []
        for m in REF_PORTARIA.finditer(trecho):
            data = primeira_data(trecho[m.end():], 200) if not data_comum else data_comum
            refs.append(_referencia(m, data or data_comum))
        return refs

    m = REF_PORTARIA.search(ato.texto)
    if not m:
        return []
    return [_referencia(m, primeira_data(ato.texto[m.end():], 300))]

def _separar_alteracao(texto):
    """Divide a apostila em descrição do ato original e trecho alterado.

    Retorna (descrição, trecho alterado, se o trecho alterado está ilegível no PDF).
    """
    # Formato "onde se lê X, leia-se Y": só X e Y importam; um trecho ilegível depois de Y não atrapalha
    onde = ONDE_SE_LE.search(texto)
    if onde:
        antes, depois = onde.group("antes"), onde.group("depois")
        ilegivel = MARCA_ILEGIVEL in antes or not re.search(r'\w', depois)
        return texto[:onde.start()], antes + depois, ilegivel

    inicio = INICIO_ALTERACAO.search(texto)
    if not inicio:
        return texto, "", MARCA_ILEGIVEL in texto
    fim = FIM_ALTERACAO.search(texto, inicio.end())
    alteracao = texto[inicio.start():fim.start() if fim else len(texto)]
    return texto[:inicio.start()], alteracao, MARCA_ILEGIVEL in alteracao

def avaliar_apostila(ato, busca):
    """Decide se a apostila (ou o ato que torna portaria sem efeito) tem efeito financeiro.

    Retorna (resultado, motivo, referência consultada).
    """
    cancelamento = bool(TORNAR_SEM_EFEITO.search(ato.texto))
    refs = _referencias(ato, cancelamento)
    if not refs:
        return NAO_CONTA, "não altera portaria", ""

    descricao, alteracao, ilegivel = _separar_alteracao(ato.texto)
    if cancelamento:
        descricao, alteracao, ilegivel = ato.texto, "", False

    if not cancelamento and REMOCAO_LOTACAO.search(descricao):
        return NAO_CONTA, "refere-se a remoção/lotação", str(refs[0])

    datas_alteracao = datas_a_contar(alteracao)
    simbolos_alteracao = simbolos(alteracao)
    if not cancelamento and not datas_alteracao and not simbolos_alteracao and not ilegivel:
        return NAO_CONTA, "a alteração não envolve data \"a contar\" nem símbolo (nome, matrícula, lotação etc.)", str(refs[0])

    cargo_na_descricao = bool(INDICIO_CARGO_FUNCAO.search(descricao))

    # Busca as portarias originais (uma só na apostila; uma ou mais no "tornar sem efeito")
    originais, onde = [], []
    for ref in refs:
        original, local = busca.buscar(ref)
        originais.append(original)
        onde.append(f"{ref} → {'encontrada em ' + local if original else local}")
    referencia = "; ".join(onde)

    if cancelamento:
        if cargo_na_descricao or any(o and INDICIO_CARGO_FUNCAO.search(o.texto) for o in originais):
            return CONTA, "torna sem efeito portaria de cargo/função", referencia
        if all(originais):
            return NAO_CONTA, "torna sem efeito portaria que não é de cargo/função", referencia
        return VERIFICAR, "torna portaria sem efeito, mas a portaria original não foi encontrada", referencia

    original = originais[0]
    if original and not INDICIO_CARGO_FUNCAO.search(original.texto):
        return NAO_CONTA, "a portaria original não é de cargo/função", referencia
    if not original and not cargo_na_descricao:
        return VERIFICAR, "não foi possível confirmar se a portaria original é de cargo/função", referencia

    if ilegivel:
        return VERIFICAR, "portaria de cargo/função, mas o trecho alterado está ilegível no PDF (desenhado como imagem)", referencia

    if original:
        novas_datas = datas_alteracao - datas_a_contar(original.texto)
        novos_simbolos = simbolos_alteracao - simbolos(original.texto)
        if novas_datas or novos_simbolos:
            mudancas = []
            if novas_datas:
                antes = ", ".join(sorted(formatar(d) for d in datas_a_contar(original.texto))) or "?"
                mudancas.append(f"data \"a contar\" de {antes} para {', '.join(sorted(formatar(d) for d in novas_datas))}")
            if novos_simbolos:
                antes = ", ".join(sorted(simbolos(original.texto))) or "?"
                mudancas.append(f"símbolo de {antes} para {', '.join(sorted(novos_simbolos))}")
            return CONTA, "altera " + " e ".join(mudancas), referencia
        return NAO_CONTA, "data \"a contar\" e símbolo iguais aos da portaria original", referencia

    # Original não encontrado, mas a apostila já mostra que é de cargo/função
    if simbolos_alteracao:
        return VERIFICAR, "reescreve a portaria de cargo/função, mas sem a original não dá para saber o que mudou", referencia
    return CONTA, f"altera a data \"a contar\" para {', '.join(sorted(formatar(d) for d in datas_alteracao))}", referencia

# ==========================================
# ANÁLISE DE UM DJE
# ==========================================

def analisar_documento(caminho_pdf, busca):
    """Retorna (páginas que contam, páginas a verificar, lista de detalhes por ato)."""
    atos, inicios_pagina = dividir_em_atos(ler_paginas_pdf(caminho_pdf))
    contadas, a_verificar = set(), set()
    detalhes = []

    for ato in atos:
        if ato.tipo == "apostila" or (ato.tipo == "portaria" and TORNAR_SEM_EFEITO.search(ato.texto)):
            resultado, motivo, referencia = avaliar_apostila(ato, busca)
            if MARCA_OCR_INICIO in ato.texto:
                motivo += " (parte do texto foi lida por OCR: conferir)"
            paginas = paginas_do_ato(ato, inicios_pagina)
            tipo = "Tornar sem efeito" if TORNAR_SEM_EFEITO.search(ato.texto) else "Apostila"
        elif ato.tipo == "portaria" and eh_portaria_de_cargo(ato.texto):
            # A página entra quando o cargo e a ação aparecem na MESMA portaria.
            # Se a portaria passa de uma página para outra, todas as páginas com os termos entram.
            resultado, motivo, referencia = CONTA, "portaria de cargo/função", ""
            ocorrencias = list(TERMOS_CARGOS.finditer(ato.texto)) + list(TERMOS_ACAO.finditer(ato.texto))
            paginas = sorted({numero_da_pagina(inicios_pagina, ato.inicio + m.start()) for m in ocorrencias})
            tipo = "Portaria"
        else:
            continue

        if resultado == CONTA:
            contadas.update(paginas)
        elif resultado == VERIFICAR:
            a_verificar.update(paginas)
        detalhes.append({"paginas": paginas, "tipo": tipo, "resultado": resultado, "motivo": motivo,
                         "referencia": referencia, "trecho": resumir(ato.texto)})

    # Página já contabilizada dispensa a análise das demais apostilas nela
    for d in detalhes:
        if d["resultado"] == VERIFICAR and set(d["paginas"]) <= contadas:
            d["resultado"] = "VERIFICAR (dispensável)"
            d["motivo"] += "; a página já é contabilizada por outro ato"
    return contadas, a_verificar - contadas, detalhes

# ==========================================
# SAÍDAS DA FOLHA (EXONERAÇÃO DE COMISSIONADOS E REQUISITADOS)
# ==========================================

SAIU, NAO_SAIU = "SAIU DA FOLHA", "NÃO SAIU"

# Servidores efetivos continuam na folha mesmo perdendo o cargo em comissão ou a função
CARGOS_ESTAVEIS = re.compile(
    r't[ée]cnico\s+de\s+atividade\s+judici[áa]ria|analista\s+judici[áa]rio|t[ée]cnico\s+judici[áa]rio|oficial\s+de\s+justi[çc]a',
    re.IGNORECASE
)
NOME = r"[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý'\-]*(?:\s+[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý'\-]*)+"
A_PEDIDO = r"(?:\s*,\s*a\s+pedido\s*,)?(?:\s+(?i:o|a)\s+(?i:servidora?))?"
EXONERACAO = re.compile(r'\b(?i:exonerar|dispensar)\b' + A_PEDIDO + r'\s+(?P<nome>' + NOME + r')')
# "servidores" no plural; o artigo é opcional porque o DJE às vezes erra ("Exonerar a servidores...")
EXONERACAO_COLETIVA = re.compile(r'\b(?:exonerar|dispensar)\s+(?:\w{1,3}\s+)?(?:seguintes\s+)?servidores\b', re.IGNORECASE)
NOMEACAO = re.compile(r'\b(?i:nomear|designar)\b' + A_PEDIDO + r'\s+(?P<nome>' + NOME + r')')
NOMEACAO_COLETIVA = re.compile(r'nome[áa]-l[oa]s|\b(?:nomear|designar)\s+(?:\w{1,3}\s+)?(?:seguintes\s+)?servidores\b', re.IGNORECASE)
RENOMEADO_NO_ATO = re.compile(r'\be\s+nome[áa]-l[oa]s?\b|nomeando-[oa]', re.IGNORECASE)
VINCULO_E_MATRICULA = re.compile(
    r'\s*,(?P<vinculo>.{0,150}?),?\s*matr[íi]cula\s*(?:funcional\s*)?n?[º°o.]*\s*(?P<matricula>\d[\d./\-]*\d)',
    re.DOTALL
)
FIM_DO_ARTIGO = re.compile(r'Art\.\s*2', re.IGNORECASE)

def normalizar_nome(nome):
    sem_acento = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode()
    return re.sub(r'\s+', ' ', sem_acento).strip().upper()

def _data_do_arquivo(arquivo):
    try:
        return date(int(arquivo[:4]), int(arquivo[4:6]), int(arquivo[6:8]))
    except ValueError:
        return None

def _primeira_data_a_contar(texto):
    for m in A_CONTAR.finditer(texto):
        d = DATA.match(texto, m.end())
        if d and _converter_data(d):
            return _converter_data(d)
    return None

def _classificar_vinculo(vinculo, matricula, texto_ato):
    """Retorna 'Estável', 'Comissionado', 'Requisitado' ou None (não identificado)."""
    if CARGOS_ESTAVEIS.search(vinculo):
        return "Estável"
    if re.search(r'comissionad', vinculo, re.IGNORECASE) or re.search(r'exclusivamente\s+comissionad', texto_ato, re.IGNORECASE):
        return "Comissionado"
    numero = re.sub(r'[.\-]', '', matricula.split("/")[-1]) if matricula else ""
    if numero.startswith("4000") and len(numero) >= 7:
        return "Comissionado"
    if numero.startswith("5") and len(numero) > 5:
        return "Requisitado"
    return None

@dataclass
class Nomeacao:
    data: date | None
    arquivo: str
    pagina: int
    portaria: str
    nomes: set      # nomes normalizados logo após "Nomear"/"Designar"
    coletiva: str   # texto normalizado quando o ato nomeia uma relação de servidores (tabela)

def _titulo(ato):
    return next((linha.strip() for linha in ato.texto.split("\n") if linha.strip()), "")

def analisar_saidas_da_folha(pasta, progresso=None):
    """Procura exonerações/dispensas de cargo ou função (CAI, DAI, DAS) de servidores sem vínculo efetivo.

    A renomeação é procurada nas portarias do mesmo DJE e dos DJEs posteriores da pasta.
    Retorna (lista de registros, lista de erros por arquivo).
    """
    arquivos = listar_pdfs(pasta)
    documentos, erros = [], []
    for i, arquivo in enumerate(arquivos):
        try:
            atos, inicios = dividir_em_atos(ler_paginas_pdf(os.path.join(pasta, arquivo)))
            documentos.append((arquivo, _data_do_arquivo(arquivo), atos, inicios))
        except Exception as e:
            erros.append(f"Erro ao processar {arquivo}: {e}")
        if progresso:
            progresso((i + 1) / len(arquivos))

    # Todas as nomeações/designações do período, para descobrir quem foi renomeado
    nomeacoes = []
    for arquivo, data, atos, inicios in documentos:
        for ato in atos:
            if ato.tipo != "portaria":
                continue
            nomes = {normalizar_nome(m.group("nome")) for m in NOMEACAO.finditer(ato.texto)}
            coletiva = normalizar_nome(ato.texto) if NOMEACAO_COLETIVA.search(ato.texto) else ""
            if nomes or coletiva:
                nomeacoes.append(Nomeacao(data, arquivo, numero_da_pagina(inicios, ato.inicio), _titulo(ato), nomes, coletiva))

    def buscar_renomeacao(nome, data, arquivo_exoneracao, titulo_exoneracao):
        alvo = normalizar_nome(nome)
        for n in nomeacoes:
            if data and n.data and n.data < data:
                continue
            if n.arquivo == arquivo_exoneracao and n.portaria == titulo_exoneracao:
                continue  # o próprio ato de exoneração
            if alvo in n.nomes or (n.coletiva and re.search(r'\b' + re.escape(alvo) + r'\b', n.coletiva)):
                return n
        return None

    registros = []
    for arquivo, data, atos, inicios in documentos:
        for ato in atos:
            if ato.tipo != "portaria":
                continue
            base = {"data": data, "arquivo": arquivo, "portaria": _titulo(ato), "trecho": resumir(ato.texto), "texto": ato.texto}

            # Exoneração de uma relação de servidores (tabela): os nomes não podem ser lidos um a um
            if EXONERACAO_COLETIVA.search(ato.texto) and SIMBOLO.search(ato.texto):
                renomeia = NOMEACAO_COLETIVA.search(ato.texto)
                registros.append({**base,
                    "pagina": numero_da_pagina(inicios, ato.inicio), "servidor": "(relação de servidores)", "matricula": "",
                    "vinculo": "Comissionado" if re.search(r'comissionad', ato.texto, re.IGNORECASE) else "",
                    "simbolo": ", ".join(sorted(simbolos(ato.texto))), "a_contar": _primeira_data_a_contar(ato.texto),
                    "situacao": NAO_SAIU if renomeia else VERIFICAR,
                    "motivo": "exoneração e renomeação no mesmo ato" if renomeia else "exoneração coletiva: conferir a relação de servidores"})
                continue

            for m in EXONERACAO.finditer(ato.texto):
                depois = ato.texto[m.end():m.end() + 700]
                fim = FIM_DO_ARTIGO.search(depois)
                depois = depois[:fim.start()] if fim else depois
                simbolo = SIMBOLO.search(depois)
                if not simbolo:
                    continue  # não é cargo/função CAI, DAI ou DAS

                vm = VINCULO_E_MATRICULA.match(depois)
                vinculo_texto = re.sub(r'\s+', ' ', vm.group("vinculo")).strip() if vm else ""
                matricula = vm.group("matricula") if vm else ""
                vinculo = _classificar_vinculo(vinculo_texto, matricula, ato.texto)
                if vinculo == "Estável":
                    continue

                nome = re.sub(r'\s+', ' ', m.group("nome")).strip()
                registro = {**base,
                    "pagina": numero_da_pagina(inicios, ato.inicio + m.start()), "servidor": nome,
                    "matricula": matricula, "vinculo": vinculo or vinculo_texto or "Não identificado",
                    "simbolo": f"{simbolo.group(1).upper()}-{simbolo.group(2)}", "a_contar": _primeira_data_a_contar(depois)}

                renomeacao = buscar_renomeacao(nome, data, arquivo, _titulo(ato))
                if RENOMEADO_NO_ATO.search(depois):
                    registro.update(situacao=NAO_SAIU, motivo="renomeado no mesmo ato")
                elif renomeacao:
                    quando = f"{formatar(renomeacao.data)}, " if renomeacao.data else ""
                    registro.update(situacao=NAO_SAIU,
                                    motivo=f"renomeado em {quando}{renomeacao.arquivo}, p. {renomeacao.pagina} ({renomeacao.portaria})")
                elif not vinculo:
                    registro.update(situacao=VERIFICAR, motivo="vínculo não identificado (sem 'comissionado' nem matrícula 4000… ou 5…)")
                else:
                    registro.update(situacao=SAIU, motivo="nenhuma renomeação encontrada no período")
                registros.append(registro)

    return registros, erros
