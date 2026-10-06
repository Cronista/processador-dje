# Retomar — Processador de DJE

> Ponto de partida para continuar o desenvolvimento a partir do GitHub/nuvem, em qualquer computador.
> Atualizado em 06/10/2026. Manual do usuário em [MANUAL.docx](MANUAL.docx).

---

## 1. O que é

App **Streamlit** que automatiza a triagem do **Diário da Justiça Eletrônico do TJRJ (DJERJ)**,
Caderno Administrativo. A partir dos PDFs completos do DJE de um período:

| Aba | O que faz |
|---|---|
| 1. Indexação | Planilha com as páginas que têm portaria de cargo/função (CAI-, DAI-, DAS-) ou apostila/"tornar sem efeito" com efeito financeiro. Aba "Detalhes" com o motivo de cada ato. |
| 2. Separação | Recorta essas páginas em `DJE_Separados/<DD MÊS>/<NN>.pdf`. |
| 3. Distribuição | Divide as páginas entre a equipe de revisão (diferença máxima de 1 página). |
| 4. Requisitados | TXT com as portarias de cargo/função de servidores com matrícula iniciada por 5. |
| 5. Saídas da folha | Planilha + TXT com exonerações de comissionados/requisitados que **não** foram renomeados. |

Usuários não técnicos abrem pelo `Iniciar.bat` (cria o `venv`, instala dependências, abre o navegador).

---

## 2. Onde as coisas estão

| O quê | Onde |
|---|---|
| Código-fonte | GitHub `Cronista/processador-dje` (privado), branch `main` |
| Interface (Streamlit) e geração de Excel/TXT | [app.py](app.py) |
| Regras de análise do texto do DJE | [analise_dje.py](analise_dje.py) |
| Teste contra o gabarito | [tests/testar_exemplos.py](tests/testar_exemplos.py) |
| Exemplos reais + gabarito do usuário | [exemplos_apostilas/](exemplos_apostilas/) (`DO/`, `acervo/`, `gabarito.txt`) |
| Dependências | `requirements.txt` (obrigatórias), `requirements-ocr.txt` (OCR opcional) |
| PDFs de trabalho (`DO/`, `Acervo/`, `DJE_Separados/`) e arquivos gerados | Só no computador do usuário — **fora do Git** |

---

## 3. Como rodar e testar

```bash
python -m venv venv
venv\Scripts\activate            # Windows
pip install -r requirements.txt
pip install -r requirements-ocr.txt   # opcional
streamlit run app.py
python tests/testar_exemplos.py       # deve terminar com "Tudo certo."
```

Os PDFs do DJE vão em `DO/` (nome começando por `AAAAMMDD`). DJEs de datas anteriores, usados para
consultar a portaria original citada por apostilas, vão em `Acervo/`.

Para testar a interface sem navegador há o `streamlit.testing.v1.AppTest` (foi assim que as abas foram
validadas durante o desenvolvimento).

---

## 4. Regras de negócio (definidas com o usuário)

### 4.1 Divisão do texto em atos
O texto de cada PDF é dividido em atos pelo **título na própria linha** (`PORTARIA ... nº`, `APOSTILA`,
`APOSTILAMENTO`, `DECISÃO`, `ATO EXECUTIVO`, `AVISO`...) e pela linha `id: NNNNNNNN` do DJERJ.
Uma linha `PORTARIA Nº X, publicada...` é **citação**, não título. Só atos do tipo portaria entram na
busca de cargo/função — decisões não contam, mesmo citando CAI-6 e exoneração.

### 4.2 Portaria de cargo/função (Aba 1)
Conta a página quando a **mesma portaria** tem `CAI-`/`DAI-`/`DAS-` e nomeação/designação/dispensa/exoneração.

### 4.3 Apostilas e "tornar sem efeito" (Aba 1)
Conta quando altera portaria de **cargo em comissão/função gratificada** e:
- muda a data **"a contar"** (de um dia para outro), ou
- muda o **símbolo** (ex.: DAI → DAS), ou
- torna a portaria sem efeito.

Não contam: correção de nome, matrícula, lotação, remoção, horário, "cargo efetivo", substituto eventual.
Quando a apostila só reescreve o ato ("fica declarado que a referida portaria..."), a **portaria original**
é buscada no DJE da data citada (pastas DO e Acervo) e as datas "a contar" e os símbolos são comparados.
Não confirmado → a página entra no índice e na coluna `PÁGINAS A VERIFICAR`.
Página já contabilizada por outro ato → `VERIFICAR (dispensável)`.

### 4.4 Saídas da folha (Aba 5)
Exoneração/dispensa de CAI/DAI/DAS de quem **não tem vínculo efetivo**:
"Comissionado"/"exclusivamente comissionado", matrícula `4000…` ou requisitado (`5…`, mais de 5 dígitos).
Estáveis (ficam de fora): Técnico de Atividade Judiciária, Analista Judiciário, Técnico Judiciário,
Oficial de Justiça. **Renomeação** (nomear/designar a mesma pessoa no mesmo DJE ou em DJE posterior da
pasta DO) → "NÃO SAIU". A busca é **pelo nome** (a matrícula muda); só vale o nome logo após
"Nomear"/"Designar", para não confundir com "vaga decorrente da exoneração de FULANO".

### 4.5 Trechos desenhados e OCR
Páginas salvas pelo **"Imprimir em PDF" do Windows** (é o caso dos exemplos) convertem parte do texto
em desenho. O **DJE original** (Word + iTextSharp) traz texto normal — 0 trechos desenhados em 198
páginas verificadas. O sistema detecta esses trechos e, se o RapidOCR estiver instalado, lê o conteúdo
e o insere entre `⟪ ⟫`; senão, marca `⟦trecho ilegível⟧`.

---

## 5. Estado atual (06/10/2026)

- `tests/testar_exemplos.py`: 16/16 casos do gabarito corretos (com OCR).
- Divisão em atos validada nos DJEs completos de 03/08 e 10/08: nenhuma portaria de cargo perdida;
  6 falsos positivos antigos (decisões) eliminados.
- Histórico do Git: estado inicial → limpeza → apostilas → Aba 5 → OCR.

## 6. Pendências e limitações conhecidas

- "Tornar sem efeito" da PORTARIA 2725 (02/09 p.197) fica "a verificar": as portarias canceladas
  (01/09) não estão no acervo. Não estava no gabarito — confirmar com o usuário se conta.
- Aba 5: renomeação publicada depois do último DJE da pasta DO não é vista; exoneração em **tabela**
  sem renomeação aparece como uma linha "(relação de servidores)" para conferência manual.
- Não há exemplos reais de **requisitado exonerado** nem de **renomeação em outro dia** (testados só
  com PDFs sintéticos durante o desenvolvimento).
- Aba 4 não inclui apostilas de requisitados (decisão do usuário).
- Para calibrar novos casos: acrescentar PDFs em `exemplos_apostilas/`, uma linha no `gabarito.txt` e
  um item em `tests/testar_exemplos.py`.
