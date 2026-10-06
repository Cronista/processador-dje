import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import re
import math
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E VARIÁVEIS GERAIS
# ==========================================
st.set_page_config(page_title="Processador de DJE", page_icon="📄", layout="wide")

MESES_EXTENSO = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

# ==========================================
# FUNÇÕES DE PROCESSAMENTO
# ==========================================

def gerar_indice_excel(pasta_diarios, arquivo_saida):
    termos_cargos = re.compile(r'(CAI-|DAI-|DAS-)', re.IGNORECASE)
    termos_acao = re.compile(r'(nomeação|nomear|designação|designar|dispensa|dispensar|exoneração|exonerar)', re.IGNORECASE)
    resultados = []

    if not os.path.exists(pasta_diarios):
        return None, f"A pasta '{pasta_diarios}' não existe."

    arquivos_pdf = [f for f in sorted(os.listdir(pasta_diarios)) if f.lower().endswith(".pdf")]
    
    if not arquivos_pdf:
        return None, "Nenhum PDF encontrado na pasta de íntegras."

    barra_progresso = st.progress(0)
    
    for i, arquivo in enumerate(arquivos_pdf):
        caminho_pdf = os.path.join(pasta_diarios, arquivo)
        data_str = arquivo[:8]
        
        if len(data_str) >= 8 and data_str.isdigit():
            data_formatada = f"{data_str[6:8]}/{data_str[4:6]}/{data_str[:4]}"
        else:
            data_formatada = "Data Inválida"
            
        paginas_do_dia = []
        
        try:
            doc = fitz.open(caminho_pdf)
            for num_pagina in range(len(doc)):
                texto_pagina = doc[num_pagina].get_text()
                if termos_cargos.search(texto_pagina) and termos_acao.search(texto_pagina):
                    paginas_do_dia.append(str(num_pagina + 1))
            doc.close()
            
            if paginas_do_dia:
                resultados.append({
                    "DATA": data_formatada,
                    "PORTARIA": "CAI, DAI, DAS",
                    "PÁGINAS": ", ".join(paginas_do_dia),
                    "FOLHA": "PROCESSADO"
                })
        except Exception as e:
            st.error(f"Erro ao processar {arquivo}: {e}")
            
        barra_progresso.progress((i + 1) / len(arquivos_pdf))

    if not resultados:
        return None, "Nenhuma ocorrência encontrada nos PDFs."

    df = pd.DataFrame(resultados)
    df.to_excel(arquivo_saida, index=False)

    # Aplicação de Estilo Visual no Excel
    wb = load_workbook(arquivo_saida)
    ws = wb.active
    cor_cabecalho = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fonte_cabecalho = Font(bold=True, name='Calibri')
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[1]:
        cell.fill = cor_cabecalho
        cell.font = fonte_cabecalho
        cell.alignment = alinhamento_centro
        cell.border = borda_fina

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = alinhamento_centro
            cell.border = borda_fina

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 4

    wb.save(arquivo_saida)
    return df, "Sucesso"

def obter_nome_pasta(data_campo):
    if isinstance(data_campo, str):
        partes = data_campo.split('/')
        if len(partes) < 2: return None
        dia = partes[0].zfill(2)
        mes_num = int(partes[1])
    else:
        dia = str(data_campo.day).zfill(2)
        mes_num = data_campo.month
    mes_nome = MESES_EXTENSO.get(mes_num)
    if dia and mes_nome:
        return f"{dia} {mes_nome}"
    return None

def separar_pdfs(caminho_planilha, diretorio_integra, diretorio_destino):
    if not os.path.exists(caminho_planilha):
        return False, [f"Planilha de índice não encontrada: {caminho_planilha}"]
    
    if not os.path.exists(diretorio_integra):
        return False, [f"Diretório de PDFs na íntegra não encontrado: {diretorio_integra}"]
        
    df = pd.read_excel(caminho_planilha)
    logs = []
    
    for idx, row in df.iterrows():
        data_bruta = row['DATA']
        paginas_brutas = row['PÁGINAS']

        if pd.isna(paginas_brutas) or str(paginas_brutas).strip().upper() == "SEM":
            continue

        nome_pasta_dia = obter_nome_pasta(data_bruta)
        if not nome_pasta_dia: continue
            
        if isinstance(data_bruta, str):
            partes = data_bruta.split('/')
            if len(partes) < 3: continue
            prefixo_data = f"{partes[2]}{partes[1].zfill(2)}{partes[0].zfill(2)}"
        else:
            prefixo_data = f"{data_bruta.year}{str(data_bruta.month).zfill(2)}{str(data_bruta.day).zfill(2)}"

        pdf_principal = None
        for arquivo in os.listdir(diretorio_integra):
            if arquivo.lower().endswith('.pdf') and arquivo.startswith(prefixo_data):
                pdf_principal = arquivo
                break

        if not pdf_principal:
            logs.append(f"Aviso: PDF na íntegra com prefixo {prefixo_data} não foi encontrado em {diretorio_integra}.")
            continue

        caminho_pdf_completo = os.path.join(diretorio_integra, pdf_principal)
        caminho_pasta_dia = os.path.join(diretorio_destino, nome_pasta_dia)
        os.makedirs(caminho_pasta_dia, exist_ok=True)

        lista_paginas = [p.strip() for p in re.split(r'[,;\s]+', str(paginas_brutas)) if p.strip()]

        try:
            doc = fitz.open(caminho_pdf_completo)
            for pag_str in lista_paginas:
                if not pag_str.isdigit(): continue
                num_pagina = int(pag_str)
                idx_pagina = num_pagina - 1 

                if 0 <= idx_pagina < len(doc):
                    nome_saida_pdf = f"{num_pagina:02d}.pdf"
                    caminho_saida = os.path.join(caminho_pasta_dia, nome_saida_pdf)
                    
                    novo_doc = fitz.open()
                    novo_doc.insert_pdf(doc, from_page=idx_pagina, to_page=idx_pagina)
                    novo_doc.save(caminho_saida)
                    novo_doc.close()
                    logs.append(f"Extraída com sucesso: {nome_pasta_dia} -> {nome_saida_pdf}")
                else:
                    logs.append(f"Erro: Página {num_pagina} está fora dos limites do documento {pdf_principal}.")
            doc.close()
        except Exception as e:
            logs.append(f"Erro ao processar o arquivo {pdf_principal}: {e}")
            
    return True, logs

def distribuir_paginas(caminho_planilha, nomes_equipe):
    if not os.path.exists(caminho_planilha):
        return None, f"Planilha não encontrada: {caminho_planilha}"
        
    df = pd.read_excel(caminho_planilha)
    todas_paginas = []
    
    for idx, row in df.iterrows():
        data = row['DATA']
        paginas_str = str(row['PÁGINAS'])
        if pd.isna(row['PÁGINAS']) or paginas_str.strip().upper() == "SEM":
            continue
        paginas = [p.strip() for p in paginas_str.split(',') if p.strip()]
        for p in paginas:
            todas_paginas.append((data, p))
            
    total_paginas = len(todas_paginas)
    if total_paginas == 0:
        return None, "Nenhuma página para distribuir na planilha informada."

    qtd_membros = len(nomes_equipe)
    base_qtd = total_paginas // qtd_membros
    
    distribuicao = {}
    inicio = 0
    
    for i, nome in enumerate(nomes_equipe):
        if i == qtd_membros - 1:
            distribuicao[nome] = todas_paginas[inicio:]
        else:
            distribuicao[nome] = todas_paginas[inicio : inicio + base_qtd]
            inicio += base_qtd

    texto_saida = ""
    for nome in nomes_equipe:
        texto_saida += f"**{nome}**\n"
        paginas_por_data = {}
        for data, pag in distribuicao.get(nome, []):
            if data not in paginas_por_data:
                paginas_por_data[data] = []
            paginas_por_data[data].append(pag)
            
        for data, pags in paginas_por_data.items():
            texto_saida += f"- {data}: {', '.join(pags)}\n"
        texto_saida += "\n"
        
    return texto_saida, total_paginas

def extrair_texto_portarias(pasta_diarios, arquivo_saida_txt):
    termos_cargos = re.compile(r'(CAI-|DAI-|DAS-)', re.IGNORECASE)
    termos_acao = re.compile(r'(nomeação|nomear|designação|designar|dispensa|dispensar|exoneração|exonerar)', re.IGNORECASE)
    termo_matricula = re.compile(r'matr[a-zí\.]*[^\d]{0,20}\b(5\d*)\b', re.IGNORECASE)

    if not os.path.exists(pasta_diarios):
        return 0, f"A pasta '{pasta_diarios}' não existe."

    arquivos_pdf = [f for f in sorted(os.listdir(pasta_diarios)) if f.lower().endswith(".pdf")]
    contador_total = 0

    if not arquivos_pdf:
        return 0, "Nenhum PDF na íntegra encontrado."

    barra_progresso = st.progress(0)

    with open(arquivo_saida_txt, 'w', encoding='utf-8') as arquivo_txt:
        for i, arquivo in enumerate(arquivos_pdf):
            caminho_pdf = os.path.join(pasta_diarios, arquivo)
            try:
                doc = fitz.open(caminho_pdf)
                texto_documento = ""
                for num_pagina in range(len(doc)):
                    texto_documento += doc[num_pagina].get_text() + "\n"
                doc.close()
                
                blocos = re.split(r'(?i)(?=PORTARIA)', texto_documento)
                for bloco in blocos:
                    if (termos_cargos.search(bloco) and termos_acao.search(bloco) and termo_matricula.search(bloco)):
                        bloco_limpo = re.sub(r'\s+', ' ', bloco).strip()
                        arquivo_txt.write(bloco_limpo + "\n\n\n")
                        contador_total += 1
            except Exception as e:
                pass
                
            barra_progresso.progress((i + 1) / len(arquivos_pdf))
            
    return contador_total, "Sucesso"

# ==========================================
# INTERFACE DO USUÁRIO (UI)
# ==========================================

st.title("📄 Sistema de Processamento - DJE")
st.markdown("Automação de extração de portarias de requisitados, separação de páginas e distribuição de metas da equipe.")

# Inicializa variável na memória para que os menus se comuniquem
if 'planilha_ativa' not in st.session_state:
    st.session_state.planilha_ativa = "DO-exp_Processado_Agrupado.xlsx"

# --- MENU LATERAL (CONFIGURAÇÕES) ---
with st.sidebar:
    st.header("⚙️ Configurações de Diretórios")
    diretorio_integra = st.text_input("Pasta dos PDFs na Íntegra", value="./DO")
    diretorio_separados = st.text_input("Pasta dos PDFs Separados", value="./DJE_Separados")
    
    st.divider()
    st.header("🗂️ Padrão de Nomenclatura")
    planilha_base = st.text_input("Nome Base da Planilha", value="DO-exp_Processado_Agrupado.xlsx")
    txt_base = st.text_input("Nome Base do TXT", value="Portarias_requisitados.txt")
    
    # Aviso flutuante para mostrar qual planilha o sistema está considerando para leitura
    st.info(f"📌 **Planilha Atual em Memória:**\n\nO sistema usará este arquivo nas Abas 2 e 3:\n\n`{st.session_state.planilha_ativa}`")

    st.divider()
    st.header("👥 Equipe de Revisão")
    equipe_input = st.text_area("Membros (um por linha)", value="Álvaro\nSidnei\nCatiane\nFred")
    nomes_equipe = [nome.strip() for nome in equipe_input.split('\n') if nome.strip()]

# --- ABAS DE NAVEGAÇÃO ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Indexação (Excel)", 
    "✂️ 2. Separação de Páginas", 
    "👥 3. Distribuição", 
    "📝 4. Extração de Texto (Requisitados)"
])

with tab1:
    st.subheader("Gerar Planilha de Ocorrências no DJE")
    st.write("Varre os arquivos na íntegra buscando páginas com incidência de cargos (CAI, DAI, DAS) e ações administrativas.")
    
    if st.button("▶️ Mapear e Gerar Excel", type="primary"):
        with st.spinner("Analisando páginas do DJE..."):
            # Gera o nome do arquivo com a data e hora
            nome, ext = os.path.splitext(planilha_base)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            arquivo_saida = f"{nome}_{timestamp}{ext}"
            
            df, msg = gerar_indice_excel(diretorio_integra, arquivo_saida)
            if df is not None:
                # Atualiza a memória com o nome exato do novo arquivo gerado
                st.session_state.planilha_ativa = arquivo_saida
                
                st.success(f"Arquivo '{arquivo_saida}' gerado com sucesso! As próximas abas já estão configuradas para utilizá-lo.")
                st.dataframe(df, use_container_width=True)
            else:
                st.error(msg)

with tab2:
    st.subheader("Separação de Páginas Específicas")
    st.write(f"Extrai as páginas identificadas no índice atual (`{st.session_state.planilha_ativa}`) e organiza na pasta de destino.")
    
    if st.button("▶️ Separar PDFs"):
        with st.spinner("Extraindo e organizando páginas em novos arquivos..."):
            sucesso, logs = separar_pdfs(st.session_state.planilha_ativa, diretorio_integra, diretorio_separados)
            if sucesso:
                st.success(f"Processo de separação finalizado! Os arquivos foram salvos em '{diretorio_separados}'.")
                with st.expander("Ver Detalhes do Log de Separação"):
                    for log in logs:
                        st.text(log)
            else:
                st.error("\n".join(logs))

with tab3:
    st.subheader("Distribuição da Carga de Trabalho")
    st.write(f"Divide de forma matemática as páginas mapeadas no índice atual (`{st.session_state.planilha_ativa}`) entre a equipe.")
    
    if st.button("▶️ Calcular Distribuição"):
        texto_dist, total = distribuir_paginas(st.session_state.planilha_ativa, nomes_equipe)
        if texto_dist:
            st.success(f"Divisão concluída! {total} páginas distribuídas entre {len(nomes_equipe)} servidores.")
            st.markdown(texto_dist)
        else:
            st.error(total)

with tab4:
    st.subheader("Consolidação de Portarias de Requisitados")
    st.write("Filtra e extrai o texto literal exclusivamente das portarias associadas a servidores requisitados (matrículas iniciadas com 5).")
    
    if st.button("▶️ Extrair Portarias (TXT)"):
        with st.spinner("Varrendo documentos e isolando atos de requisitados..."):
            # Gera o nome do TXT com a data e hora
            nome_txt, ext_txt = os.path.splitext(txt_base)
            timestamp_txt = datetime.now().strftime("%Y%m%d_%H%M%S")
            arquivo_txt_saida = f"{nome_txt}_{timestamp_txt}{ext_txt}"
            
            total_portarias, msg = extrair_texto_portarias(diretorio_integra, arquivo_txt_saida)
            if total_portarias > 0:
                st.success(f"Processo concluído! {total_portarias} portarias foram compiladas no arquivo '{arquivo_txt_saida}'.")
                
                with open(arquivo_txt_saida, "r", encoding="utf-8") as f:
                    st.download_button(
                        label="📥 Baixar Portarias Consolidadas (TXT)",
                        data=f,
                        file_name=arquivo_txt_saida,
                        mime="text/plain"
                    )
            else:
                st.warning(f"Nenhuma portaria de requisitados foi extraída. Detalhes: {msg}")