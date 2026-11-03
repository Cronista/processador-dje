# Contexto — Processador de DJE na VPS

> Ponto de partida para o Claude que vai trabalhar na VPS (sessão via SSH).
> Escrito em 03/11/2026, na migração do ambiente de nuvem do Claude Code para a VPS do usuário.
> Regras de negócio, estado e pendências do projeto: [RETOMAR.md](RETOMAR.md). Manual do usuário: [MANUAL.docx](MANUAL.docx).

---

## 1. O projeto em uma frase

App **Streamlit** (Python) que faz a triagem do **Diário da Justiça Eletrônico do TJRJ**, Caderno
Administrativo: localiza portarias de cargo/função (CAI-, DAI-, DAS-) e apostilas com efeito financeiro,
recorta as páginas, distribui entre a equipe e gera listas de requisitados e de saídas da folha.

| Arquivo | Papel |
|---|---|
| `app.py` | Interface (5 abas) e geração de Excel/TXT |
| `analise_dje.py` | Regras de análise do texto (divisão em atos, apostilas, saídas da folha, OCR) |
| `tests/testar_exemplos.py` | Confere a análise contra `exemplos_apostilas/gabarito.txt` |
| `exemplos_apostilas/` | PDFs reais de calibração (`DO/`, `acervo/`) e o gabarito do usuário |
| `requirements.txt` / `requirements-ocr.txt` | Dependências obrigatórias / OCR opcional (RapidOCR) |
| `Iniciar.bat` | Atalho do **Windows** — não serve na VPS |

**Leia o `RETOMAR.md` antes de mexer nas regras de análise.** As regras foram definidas com o usuário, caso a caso,
e o teste do gabarito precisa continuar terminando com `Tudo certo.` (16/16 com OCR).

---

## 2. De onde veio e o que muda

- Até agora o projeto rodava no **Windows do usuário** (via `Iniciar.bat`) e era desenvolvido em sessões de nuvem
  do Claude Code. O código está todo no GitHub: `Cronista/processador-dje` (privado), branch `main`.
- A sessão de nuvem não tinha alterações pendentes: **clonar o `main` é suficiente**.
- **Não estão no Git** (ver `.gitignore`) e precisam ser copiados do computador do usuário se forem necessários:
  `DO/` (PDFs do período), `Acervo/` (DJEs anteriores), `DJE_Separados/` (recortes) e os arquivos gerados
  (`DO-exp_Processado_Agrupado*.xlsx`, `Portarias_requisitados*.txt`, `Saidas_da_folha*`).

---

## 3. Instalação na VPS

```bash
sudo apt install -y python3 python3-venv git
git clone https://github.com/Cronista/processador-dje.git ~/processador-dje   # repo privado: gh auth login ou chave SSH
cd ~/processador-dje
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
pip install -r requirements-ocr.txt      # opcional; se falhar, o sistema funciona sem OCR
python tests/testar_exemplos.py          # deve terminar com "Tudo certo."
```

- Testado com Python 3.13. O RapidOCR baixa/usa modelos ONNX; sem ele, os casos com trechos desenhados
  ficam como `VERIFICAR` (o teste aceita isso).
- `.gitattributes` força `eol=lf` (exceto `.bat`): não há problema de fim de linha entre Windows e Linux.

### Copiar os dados de trabalho (do computador do usuário para a VPS)

```bash
rsync -av DO/ Acervo/ usuario@vps:~/processador-dje/   # ou scp -r; ajuste conforme as pastas que existirem
```

Os nomes dos PDFs precisam começar com `AAAAMMDD` (ex.: `20260810ADMDJETJRJ.pdf`).

---

## 4. Rodar o app na VPS

```bash
streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless true
```

**Segurança:** o app **não tem login** e lê/escreve pastas do servidor a partir de caminhos digitados na barra
lateral (`resolver_caminho` aceita caminho absoluto). Não expor a porta 8501 na internet. Opções:

1. **Túnel SSH** (mais simples): `ssh -L 8501:localhost:8501 usuario@vps` e abrir `http://localhost:8501`.
2. **nginx** na frente, com HTTPS e autenticação (basic auth ou similar), fazendo proxy para `127.0.0.1:8501`
   (o Streamlit usa WebSocket: repassar `Upgrade`/`Connection`).

### Serviço systemd (opcional, sobe sozinho após reiniciar)

`/etc/systemd/system/processador-dje.service`:

```ini
[Unit]
Description=Processador de DJE (Streamlit)
After=network.target

[Service]
User=USUARIO
WorkingDirectory=/home/USUARIO/processador-dje
ExecStart=/home/USUARIO/processador-dje/venv/bin/streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless true
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload && sudo systemctl enable --now processador-dje
```

---

## 5. Limitações do app para uso remoto (candidatas a adaptação — combinar com o usuário antes)

O app foi feito para rodar **no mesmo computador** onde estão os PDFs. Na VPS, isso gera lacunas:

| Aba | Hoje | Problema na VPS |
|---|---|---|
| Barra lateral | Pastas `./DO`, `./DJE_Separados`, `./Acervo` relativas à pasta do `app.py` | Não há envio de PDFs pelo navegador (`st.file_uploader` não é usado) — os PDFs precisam chegar por `scp`/`rsync` |
| 1. Indexação | Grava a planilha `.xlsx` na pasta do sistema | **Sem botão de download** |
| 2. Separação | Grava os recortes em `DJE_Separados/<DD MÊS>/<NN>.pdf` | **Sem download** (um `.zip` resolveria) |
| 3. Distribuição | Atualiza a planilha ativa | **Sem download** |
| 4. Requisitados | Gera o TXT | Já tem `st.download_button` |
| 5. Saídas da folha | Gera `.xlsx` + `.txt` | Já tem `st.download_button` |

Também: `Iniciar.bat` só funciona no Windows; o `MANUAL.docx` descreve o uso local (pode precisar de uma seção
sobre o acesso pela VPS).

Nada disso foi implementado ainda. Se o usuário pedir, a ordem natural é: upload dos PDFs → download nas abas 1–3
→ ajuste do manual.

---

## 6. Convenções do repositório

- Textos, nomes de variáveis, mensagens de interface e **mensagens de commit em português**
  (ex.: "Nova Aba 5: saídas da folha de comissionados e requisitados").
- Histórico linear e pequeno; `main` é o branch principal.
- Para calibrar um caso novo: PDF em `exemplos_apostilas/`, linha no `gabarito.txt` e item em
  `tests/testar_exemplos.py`.
- Para testar a interface sem navegador: `streamlit.testing.v1.AppTest`.
- Atualizar o `RETOMAR.md` (seções 5 e 6) quando o estado ou as pendências mudarem.

---

## 7. Checklist da primeira sessão na VPS

- [ ] Repositório clonado e `venv` criado
- [ ] `python tests/testar_exemplos.py` → `Tudo certo.`
- [ ] OCR instalado (ou decidido ficar sem)
- [ ] `DO/` e `Acervo/` copiados, se o usuário for processar DJEs na VPS
- [ ] App acessível só por túnel SSH ou atrás de nginx com senha
- [ ] (Opcional) serviço systemd ativo
- [ ] Perguntar ao usuário se quer as adaptações da seção 5
