# Processador de DJE — instruções para o Claude Code

App **Streamlit** que faz a triagem do Diário da Justiça Eletrônico do TJRJ (Caderno Administrativo):
portarias de cargo/função (CAI-, DAI-, DAS-), apostilas com efeito financeiro e saídas da folha.

**Antes de qualquer mudança, leia o [RETOMAR.md](RETOMAR.md)**: regras de negócio (definidas com o usuário, caso a
caso), estado atual e pendências. Manual do usuário: [MANUAL.docx](MANUAL.docx).

## Onde as coisas estão

- `app.py`: interface (5 abas) e geração de Excel/TXT.
- `analise_dje.py`: regras de análise do texto (atos, apostilas, "tornar sem efeito", saídas da folha, OCR).
- `tests/testar_exemplos.py`: confere a análise contra `exemplos_apostilas/gabarito.txt`.
- `exemplos_apostilas/`: PDFs de calibração (`DO/` = apostilas avaliadas, `acervo/` = DJEs/portarias citadas)
  e o gabarito escrito pelo usuário. **Versionados no Git.**
- Fora do Git (`.gitignore`), só no computador de trabalho: `DO/`, `Acervo/`, `DJE_Separados/` na raiz e os
  arquivos gerados (`DO-exp_Processado_Agrupado*.xlsx`, `Portarias_requisitados*.txt`, `Saidas_da_folha*`).
  São dados reais do usuário: ler à vontade, mas não apagar nem sobrescrever.

## Ambiente e teste

O `Iniciar.bat` (Windows) cria o `venv`, instala as dependências e abre o app. Para rodar o teste:

```bash
venv\Scripts\python tests\testar_exemplos.py     # Windows
venv/bin/python tests/testar_exemplos.py         # Linux/Mac
```

Precisa terminar com `Tudo certo.`. Com OCR (`requirements-ocr.txt`) todos os casos dão o resultado esperado;
sem OCR, os casos que dependem de trechos desenhados no PDF aceitam `VERIFICAR`.
Para testar a interface sem navegador: `streamlit.testing.v1.AppTest`.

## Calibrar um caso novo (ex.: página que saiu em "PÁGINAS A VERIFICAR")

1. Copiar a página para `exemplos_apostilas/DO/` com o nome `AAAAMMDDADMDJETJRJ_pgNN.pdf` (a data do DJE
   no início do nome é obrigatória: é por ela que a portaria original é procurada).
2. Se a apostila cita portaria de outra data, copiar o DJE (ou a página) daquela data para `exemplos_apostilas/acervo/`.
3. Acrescentar a decisão do usuário ao `gabarito.txt`, no formato das linhas existentes.
4. Acrescentar o item em `APOSTILAS` (ou `PAGINAS_FORA` / `SAIRAM_DA_FOLHA`) em `tests/testar_exemplos.py`.
5. Ajustar `analise_dje.py` até o teste fechar, sem quebrar os casos anteriores.
6. Atualizar o `RETOMAR.md` (regra nova na seção 4; seções 5 e 6 com estado e pendências).

## Convenções

- Tudo em **português**: textos, nomes de variáveis, mensagens da interface e **mensagens de commit**.
- Histórico linear no `main`. O projeto é usado em mais de um lugar (computador do usuário e VPS):
  `git pull` antes de começar, `git push` ao terminar.
- Regra de negócio nova ou alterada só com a decisão do usuário; na dúvida, perguntar e registrar no `RETOMAR.md`.
- Instruções específicas de uma máquina vão em `CLAUDE.local.md` (fora do Git), não neste arquivo.
