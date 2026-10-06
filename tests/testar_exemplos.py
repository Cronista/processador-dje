"""Confere a análise contra o gabarito de exemplos_apostilas/gabarito.txt.

Uso (na pasta do projeto, com o venv ativado):
    python tests/testar_exemplos.py

Sem o OCR instalado, os casos que dependem de trechos desenhados no PDF são aceitos como "VERIFICAR".
"""
import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import analise_dje as a  # noqa: E402

EXEMPLOS = os.path.join(RAIZ, "exemplos_apostilas")
DO = os.path.join(EXEMPLOS, "DO")
ACERVO = os.path.join(EXEMPLOS, "acervo")

# (pasta, arquivo, trecho que identifica a apostila, resultado esperado, resultado aceito sem OCR)
APOSTILAS = [
    (DO, "20260810ADMDJETJRJ_pg61.pdf", "FLAVIA SAMPAIO MUSSE", a.CONTA, a.VERIFICAR),
    (DO, "20260810ADMDJETJRJ_pg61.pdf", "ALESSANDRA PRATAGY CABRAL", a.NAO_CONTA, None),
    (DO, "20260812ADMDJETJRJ_pg3.pdf", "HELLEN SANTOS REGO", a.CONTA, None),
    (DO, "20260818ADMDJETJRJ_pg2.pdf", "BEATRIZ DO NASCIMENTO LIMA", a.CONTA, None),
    (DO, "20260818ADMDJETJRJ_pg2.pdf", "ESTER FREIRE DA SILVA", a.CONTA, None),
    (DO, "20260818ADMDJETJRJ_pg2.pdf", "IGOR AZEVEDO MARTINET", a.CONTA, None),
    (DO, "20260819ADMDJETJRJ_pg11.pdf", "JULIANA ROCHA ALVES", a.NAO_CONTA, None),
    (DO, "20260828ADMDJETJRJ_pg65.pdf", "2499/2026", a.NAO_CONTA, a.VERIFICAR),
    (DO, "20260902ADMDJETJRJ_pg2.pdf", "ANA CRISTINA SATURNINO VAZ", a.NAO_CONTA, None),
    (DO, "20260903ADMDJETJRJ_pg69.pdf", "Thaíssa Fernanda", a.NAO_CONTA, None),
    (DO, "20260903ADMDJETJRJ_pg69.pdf", "Renata Lucia Lima", a.NAO_CONTA, None),
    (DO, "20260904ADMDJETJRJ_pg65.pdf", "Ana Lucia Parada Macedo", a.NAO_CONTA, None),
    (DO, "20260908ADMDJETJRJ_pg50.pdf", "CAROLAINE DA SILVA", a.NAO_CONTA, None),
    (ACERVO, "20260902ADMDJETJRJ_pg197.pdf", "Referente à lotação", a.NAO_CONTA, None),
]

# Páginas que não podem entrar no índice (ex.: DECISÃO que cita cargo e exoneração)
PAGINAS_FORA = [(DO, "20260824ADMDJETJRJ_pg17.pdf", 1)]

SAIRAM_DA_FOLHA = {"BEATRIZ CAVALCANTI SILVA", "RODRIGO PORTUGAL DE CARVALHO", "MARIANE CRISTINA COSTA GOMES"}


def main():
    ocr = a.ocr_disponivel()
    print(f"OCR {'disponível' if ocr else 'NÃO instalado (casos dependentes de OCR aceitam VERIFICAR)'}\n")
    busca = a.BuscaPortarias([DO, ACERVO])
    falhas = 0

    for pasta, arquivo, chave, esperado, sem_ocr in APOSTILAS:
        _, _, detalhes = a.analisar_documento(os.path.join(pasta, arquivo), busca)
        achados = [d for d in detalhes if d["tipo"] != "Portaria" and chave.upper() in d["trecho"].upper()]
        obtido = achados[0]["resultado"].split(" (")[0] if achados else "(não encontrada)"
        aceitos = {esperado} | ({sem_ocr} if sem_ocr and not ocr else set())
        ok = obtido in aceitos
        falhas += not ok
        print(f"{'OK  ' if ok else 'FALHA'} {arquivo} | {chave}: esperado {esperado}, obtido {obtido}")

    for pasta, arquivo, pagina in PAGINAS_FORA:
        contadas, a_verificar, _ = a.analisar_documento(os.path.join(pasta, arquivo), busca)
        ok = pagina not in contadas | a_verificar
        falhas += not ok
        print(f"{'OK  ' if ok else 'FALHA'} {arquivo} | página {pagina} fora do índice")

    registros, _ = a.analisar_saidas_da_folha(ACERVO)
    sairam = {r["servidor"] for r in registros if r["situacao"] == a.SAIU}
    ok = sairam == SAIRAM_DA_FOLHA
    falhas += not ok
    print(f"{'OK  ' if ok else 'FALHA'} Saídas da folha (acervo): {sorted(sairam)}")

    print(f"\n{'Tudo certo.' if not falhas else f'{falhas} falha(s).'}")
    return 1 if falhas else 0


if __name__ == "__main__":
    sys.exit(main())
