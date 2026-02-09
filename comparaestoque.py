import streamlit as st
import pandas as pd
import zipfile
import io
import unicodedata
import re
from io import StringIO
from datetime import datetime

# =====================================================
# CONFIGURAÇÃO DA PÁGINA (GLOBAL)
# =====================================================
st.set_page_config(
    page_title="Ferramentas de Pedido e Estoque",
    layout="wide",
    page_icon="📦"
)

st.title("📦 Ferramentas de Pedido e Estoque")

# =====================================================
# CRIAR ABAS
# =====================================================
tab1, tab2 = st.tabs(["📦 CONVERTER PEDIDO WHATSAPP", "📊 COMPARAR ESTOQUES"])

import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
from fpdf import FPDF

# =====================================================
# ABA 1: CONVERTER PEDIDO WHATSAPP
# =====================================================
with tab1:
    st.markdown("### 📦 CONVERTER PEDIDO WHATSAPP")
    st.markdown("Cole abaixo o texto do pedido exatamente como recebido:")

    if "texto_pedido" not in st.session_state:
        st.session_state.texto_pedido = ""

    # ---------------- BOTÕES ----------------
    col1, col2, col3 = st.columns(3)

    with col1:
        converter = st.button("🔄 Converter para CSV", key="btn_csv")

    with col2:
        limpar = st.button("🧹 Limpar texto", key="btn_limpar")
        if limpar:
            st.session_state["texto_pedido"] = ""


    with col3:
        gerar_pdf_btn = st.button("📄 Gerar PDF", key="btn_pdf")


    # ---------------- TEXT AREA ----------------
    texto = st.text_area(
        "📋 Texto do Pedido",
        value=st.session_state["texto_pedido"],
        key="texto_pedido",
        height=420
    )


    # ---------------- FUNÇÕES AUXILIARES ----------------
    def so_numeros(v):
        return re.sub(r"\D", "", v or "")

    def parse_valor(v):
        if not v:
            return 0.0
        v = v.strip()
        if "," in v and "." in v:
            if v.rfind(",") > v.rfind("."):
                v = v.replace(".", "").replace(",", ".")
            else:
                v = v.replace(",", "")
        elif "," in v:
            v = v.replace(".", "").replace(",", ".")
        else:
            v = v.replace(",", "")
        return float(v)

    def formatar_br(v):
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def formatar_codigo(c):
        return c.zfill(7)

    def formatar_cnpj(cnpj):
        return f'="{so_numeros(cnpj)}"'

    # ---------------- EXTRAÇÃO DE ITENS ----------------
    def extrair_itens(texto):
        itens = []

        bloco = re.search(
            r"📦\s*\*ITENS DO PEDIDO\*(.*?)💰\s*\*TOTAL DO PEDIDO",
            texto,
            re.DOTALL
        )

        if not bloco:
            return pd.DataFrame()

        padrao = re.findall(
            r"\*\s*(.*?)\s*\*\s*\n"
            r"Cód:\s*(\d+)\s*\n"
            r"(\d+)\s*x\s*R\$\s*([\d,.]+)\s*=\s*\*R\$\s*([\d,.]+)\*",
            bloco.group(1)
        )

        for prod, cod, qtd, unit, total in padrao:
            itens.append({
                "Codigo": cod,
                "Produto": prod,
                "Quantidade": int(qtd),
                "Valor_Unitario": parse_valor(unit),
                "Total": parse_valor(total)
            })

        return pd.DataFrame(itens)

    # ---------------- EXTRAÇÃO CLIENTE ----------------
    def extrair_dados_cliente(texto):
        campos = {
            "Razao_Social": r"Razão Social:\s*(.*)",
            "CNPJ": r"CNPJ:\s*([\d./-]+)",
            "IE": r"IE:\s*(.*)",
            "Telefone": r"Telefone:\s*([\d()+\s-]+)",
            "Email": r"E-mail:\s*(.*)"
        }

        dados = {}
        for campo, regex in campos.items():
            m = re.search(regex, texto)
            dados[campo] = m.group(1).strip() if m else ""

        return dados

    # ---------------- ENDEREÇO (2 LINHAS) ----------------
    def extrair_endereco(texto):
        padrao = r"📍 Endereço:\s*\n(.+)\n(.+)"
        m = re.search(padrao, texto)
        if m:
            return m.group(1).strip(), m.group(2).strip()
        return "", ""

    # ---------------- TOTAL ----------------
    def extrair_total(texto):
        m = re.search(r"TOTAL DO PEDIDO:\s*R\$\s*([\d,.]+)", texto)
        return parse_valor(m.group(1)) if m else 0.0

    # ---------------- MONTAGEM PDF ----------------
    def montar_dados_para_pdf(dados_cliente, df_itens, texto):
        end1, end2 = extrair_endereco(texto)

        dados_cliente_pdf = {
            "razao": dados_cliente.get("Razao_Social", ""),
            "endereco_linha1": end1,
            "endereco_linha2": end2
        }

        carrinho = []
        for _, r in df_itens.iterrows():
            carrinho.append({
                "codigo": r["Codigo"],
                "descricao": r["Produto"],
                "qtd": r["Quantidade"],
                "preco": r["Valor_Unitario"],
                "total": r["Total"]
            })

        return dados_cliente_pdf, carrinho

    # =====================================================
    # PDF OFICIAL ZIONNE
    # =====================================================
    class PedidoPDF(FPDF):

        def header(self):
            self.set_font("Arial", "B", 14)
            self.cell(0, 8, "PEDIDO DE VENDA - FEIRA ABUP SHOW HOME", 0, 1, "C")
            self.set_font("Arial", "", 9)
            self.cell(
                0, 5,
                f"Emissão: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                0, 1, "R"
            )
            self.ln(3)

        def footer(self):
            self.set_y(-22)
            self.set_font("Arial", "", 8)
            self.cell(0, 4, "Instagram: @zionne.oficial", 0, 1, "C")
            self.cell(0, 4, "Telefone / WhatsApp: (41) 3043-0595", 0, 1, "C")
            self.cell(0, 4, "Site: zionne.com.br | E-mail: comercial@zionne.com", 0, 1, "C")
            self.cell(0, 4, "R. Gen. Mário Tourinho, 2465 - Curitiba - PR", 0, 0, "C")

    def gerar_pdf(
        dados_cliente,
        carrinho,
        total,
        cond_pag,
        frete,
        obs,
        cnpj,
        telefone,
        email,
        ie
    ):
        pdf = PedidoPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=25)

        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 6, "DADOS DO CLIENTE", 0, 1)
        pdf.set_font("Arial", "", 9)

        pdf.cell(0, 5, f"Cliente: {dados_cliente.get('razao','')}", 0, 1)
        pdf.cell(0, 5, f"CNPJ: {cnpj}    IE: {ie}", 0, 1)
        pdf.cell(0, 5, f"Telefone: {telefone}    E-mail: {email}", 0, 1)

        pdf.multi_cell(
            0, 5,
            f"Endereço:\n{dados_cliente.get('endereco_linha1','')}\n{dados_cliente.get('endereco_linha2','')}"
        )

        pdf.ln(3)

        pdf.set_font("Arial", "B", 9)
        pdf.cell(10, 6, "Item", 1)
        pdf.cell(25, 6, "Código", 1)
        pdf.cell(90, 6, "Descrição", 1)
        pdf.cell(15, 6, "Qtde", 1)
        pdf.cell(25, 6, "Vlr Unit", 1)
        pdf.cell(25, 6, "Vlr Total", 1, 1)

        pdf.set_font("Arial", "", 9)

        for i, item in enumerate(carrinho, start=1):
            pdf.cell(10, 6, str(i), 1)
            pdf.cell(25, 6, item["codigo"], 1)
            pdf.cell(90, 6, item["descricao"], 1)
            pdf.cell(15, 6, str(item["qtd"]), 1, 0, "C")
            pdf.cell(25, 6, f"{item['preco']:.2f}", 1, 0, "R")
            pdf.cell(25, 6, f"{item['total']:.2f}", 1, 1, "R")

        pdf.ln(4)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(130, 6, "")
        pdf.cell(30, 6, "TOTAL:", 1)
        pdf.cell(30, 6, f"{total:.2f}", 1, 1, "R")

        pdf.ln(4)
        pdf.set_font("Arial", "", 9)
        pdf.multi_cell(0, 5, f"Pagamento: {cond_pag}")
        pdf.multi_cell(0, 5, f"Frete: {frete}")
        pdf.multi_cell(0, 5, f"Observações: {obs}")

        return pdf.output(dest="S").encode("latin1")

    # ---------------- AÇÃO PRINCIPAL ----------------
    if converter or gerar_pdf_btn:
        if not texto.strip():
            st.warning("Cole o texto do pedido antes de continuar.")
        else:
            df_itens = extrair_itens(texto)
            dados_cliente = extrair_dados_cliente(texto)
            total_pedido = extrair_total(texto)

            cnpj = so_numeros(dados_cliente.get("CNPJ"))
            telefone = so_numeros(dados_cliente.get("Telefone"))
            nome_arquivo = f"{cnpj}_{telefone}"

            st.subheader("📦 Itens do Pedido")
            st.dataframe(df_itens, use_container_width=True)

            st.metric("💰 Total", f"R$ {formatar_br(total_pedido)}")

            if converter:
                df_csv = df_itens.copy()
                df_csv["Cnpj"] = formatar_cnpj(dados_cliente.get("CNPJ"))
                df_csv["Codigo"] = df_csv["Codigo"].apply(formatar_codigo)
                df_csv["Valor_Unitario"] = df_csv["Valor_Unitario"].apply(formatar_br)
                df_csv["Total"] = df_csv["Total"].apply(formatar_br)

                csv = StringIO()
                df_csv.to_csv(csv, index=False, sep=";")

                st.download_button(
                    "⬇️ Baixar CSV",
                    csv.getvalue(),
                    file_name=f"{nome_arquivo}.csv",
                    mime="text/csv"
                )

            if gerar_pdf_btn:
                dados_pdf, carrinho = montar_dados_para_pdf(
                    dados_cliente, df_itens, texto
                )

                pdf_bytes = gerar_pdf(
                    dados_cliente=dados_pdf,
                    carrinho=carrinho,
                    total=total_pedido,
                    cond_pag="Conforme combinado",
                    frete="A combinar",
                    obs="Pedido gerado via WhatsApp",
                    cnpj=dados_cliente.get("CNPJ"),
                    telefone=dados_cliente.get("Telefone"),
                    email=dados_cliente.get("Email"),
                    ie=dados_cliente.get("IE")
                )

                st.download_button(
                    "⬇️ Baixar PDF",
                    pdf_bytes,
                    file_name=f"{nome_arquivo}.pdf",
                    mime="application/pdf"
                )


# =====================================================
# ABA 2: COMPARAR ESTOQUES
# =====================================================
with tab2:
    st.markdown("### 📊 COMPARAR ESTOQUES")
    
    # =====================================================
    # LIMPAR CACHE / RESET
    # =====================================================
    if st.button("🧹 Limpar cache e recarregar"):
        st.cache_data.clear()
        st.session_state.clear()
        st.rerun()

    st.title("📦 Alocação de Pedido de Venda (QM x MF)")
    st.markdown("Ferramenta para alocar pedidos de venda com base nos estoques de QM e MF, gerando relatórios de faturamento e downloads organizados.")

    # =====================================================
    # UPLOAD DOS ESTOQUES
    # =====================================================
    st.subheader("📂 Upload dos Estoques")
    estoque_qm = st.file_uploader("Estoque QM (CSV)", type="csv", help="Arquivo CSV com separador ';', encoding latin1.")
    estoque_mf = st.file_uploader("Estoque MF (CSV)", type="csv", help="Arquivo CSV com separador ';', encoding latin1.")

    # =====================================================
    # UPLOAD DOS PEDIDOS
    # =====================================================
    st.subheader("🛒 Upload dos Pedidos de Venda")
    pedidos_files = st.file_uploader(
        "Pedidos (um ou vários CSVs)",
        type="csv",
        accept_multiple_files=True,
        help="Arquivos CSV com separador ';', encoding latin1, decimal ','."
    )

    # =====================================================
    # FUNÇÃO AUXILIAR PARA LOCALIZAR COLUNAS
    # =====================================================
    def localizar_coluna(df, possiveis_nomes):
        """
        Localiza uma coluna no DataFrame baseada em possíveis nomes (ignorando acentos e maiúsculas).
        """
        for col in df.columns:
            col_norm = unicodedata.normalize('NFD', col).encode('ascii', 'ignore').decode('ascii').upper()
            if any(nome in col_norm for nome in possiveis_nomes):
                return col
        return None

    # =====================================================
    # FUNÇÃO PARA CONVERTER VALORES PARA FLOAT (FORMATO BRASILEIRO)
    # =====================================================
    def to_float_br(valor):
        if pd.isna(valor):
            return 0.0

        # Se já for número, retorna direto
        if isinstance(valor, (int, float)):
            return float(valor)

        # Se for string, converte formato BR
        valor = str(valor).strip()
        valor = valor.replace(".", "").replace(",", ".")
        return float(valor)


    # =====================================================
    # PROCESSAMENTO PRINCIPAL
    # =====================================================
    if estoque_qm and estoque_mf and pedidos_files:
        try:
            # -------------------------------------------------
            # LER E VALIDAR ESTOQUES
            # -------------------------------------------------
            with st.spinner("Carregando estoques..."):
                df_qm = pd.read_csv(estoque_qm, sep=";", encoding="latin1")
                df_mf = pd.read_csv(estoque_mf, sep=";", encoding="latin1")
            
            df_estoque = pd.concat([df_qm, df_mf], ignore_index=True)
            df_estoque.columns = df_estoque.columns.str.strip()
            
            # Detectar colunas obrigatórias do estoque
            col_produto = localizar_coluna(df_estoque, {"PRODUTO", "CODIGO", "COD_PROD", "CODPROD"})
            col_qtde = localizar_coluna(df_estoque, {"QTDE", "QUANTIDADE", "SALDO", "QTD"})
            col_grupo = localizar_coluna(df_estoque, {"LK-GRUPO", "GRUPO", "EMPRESA", "LKGRUPO"})
            
            if not all([col_produto, col_qtde, col_grupo]):
                st.error("❌ Os arquivos de estoque não contêm todas as colunas necessárias (Produto, Qtde, LK-GRUPO).")
                st.write("Colunas encontradas no estoque QM:", df_qm.columns.tolist())
                st.write("Colunas encontradas no estoque MF:", df_mf.columns.tolist())
                st.stop()
            
            # Normalizar dados do estoque (usar strings para evitar NaN em códigos mistos)
            df_estoque["Produto"] = (
                df_estoque[col_produto]
                .astype(str)
                .str.strip()
                .str.upper()
                .str.lstrip('0')  # Remove zeros à esquerda se numérico
            )
            df_estoque["Qtde"] = pd.to_numeric(df_estoque[col_qtde], errors="coerce").fillna(0)
            df_estoque["LK-GRUPO"] = df_estoque[col_grupo].astype(str).str.strip().str.upper()
            
            # Filtrar estoques válidos
            df_estoque = df_estoque.dropna(subset=["Produto"])
            
            # Agrupar estoque por Produto e LK-GRUPO
            df_estoque_agrupado = (
                df_estoque
                .groupby(["Produto", "LK-GRUPO"], as_index=False)["Qtde"]
                .sum()
            )
            
            with st.expander("🔍 Ver estoque agrupado (primeiras 100 linhas)"):
                st.dataframe(df_estoque_agrupado.head(100), use_container_width=True)
            
            # -------------------------------------------------
            # PROCESSAR PEDIDOS
            # -------------------------------------------------
            resultados = []
            progresso = st.progress(0)
            total_arquivos = len(pedidos_files)
            
            for i, file in enumerate(pedidos_files):
                with st.spinner(f"Processando pedido {i+1}/{total_arquivos}: {file.name}"):
                    df_pedido = pd.read_csv(file, sep=";", encoding="latin1", decimal=",")
                    df_pedido.columns = df_pedido.columns.str.strip()
                    
                    # Detectar colunas do pedido
                    col_codigo = localizar_coluna(df_pedido, {"CODIGO", "PRODUTO", "COD_PROD", "CODPROD"})
                    col_quant = localizar_coluna(df_pedido, {"QUANTIDADE", "QTDE", "QTD"})
                    col_cnpj = localizar_coluna(df_pedido, {"CNPJ", "CNPJ_CLIENTE", "CPF_CNPJ", "CLIENTE"})
                    col_valor_unit = localizar_coluna(df_pedido, {"VALOR_UNITARIO", "VALOR", "PRECO", "VL_UNIT"})
                    col_total = localizar_coluna(df_pedido, {"TOTAL", "VALOR_TOTAL", "VL_TOTAL"})
                    
                    if not col_codigo or not col_quant:
                        st.warning(f"⚠️ Pedido {file.name} ignorado: colunas obrigatórias não encontradas.")
                        continue
                    
                    # Normalizar dados do pedido
                    df_pedido["Codigo"] = (
                        df_pedido[col_codigo]
                        .astype(str)
                        .str.strip()
                        .str.upper()
                        .str.lstrip('0')  # Remove zeros à esquerda
                    )
                    df_pedido["Quantidade"] = pd.to_numeric(df_pedido[col_quant], errors="coerce").fillna(0)
                    
                    # Calcular valor do item (usando to_float_br para garantir formato brasileiro)
                    if col_total:
                        df_pedido["Total_Item"] = df_pedido[col_total].apply(to_float_br)
                        df_pedido["Valor_Unitario"] = (
                            df_pedido["Total_Item"] /
                            df_pedido["Quantidade"].replace(0, 1)
                        )

                    elif col_valor_unit:
                        df_pedido["Valor_Unitario"] = df_pedido[col_valor_unit].apply(to_float_br)
                        df_pedido["Total_Item"] = (
                            df_pedido["Quantidade"] *
                            df_pedido["Valor_Unitario"]
                        )

                    else:
                        df_pedido["Valor_Unitario"] = 0
                        df_pedido["Total_Item"] = 0
                    
                    # Normalizar CNPJ
                    if col_cnpj:
                        df_pedido["CNPJ"] = df_pedido[col_cnpj].astype(str).str.replace(r"\D", "", regex=True)
                    else:
                        df_pedido["CNPJ"] = file.name
                    
                    # Filtrar pedidos válidos
                    df_pedido = df_pedido.dropna(subset=["Codigo"])
                    
                    # -------------------------------------------------
                    # COMPARAÇÃO COM ESTOQUE
                    # -------------------------------------------------
                    for _, row in df_pedido.iterrows():
                        if row["Quantidade"] <= 0:
                            continue
                        
                        codigo = row["Codigo"]
                        qtd_pedida = row["Quantidade"]
                        
                        empresa = "SEM ESTOQUE"
                        qtde_disp = 0
                        
                        estoque_prod = df_estoque_agrupado[df_estoque_agrupado["Produto"] == codigo]
                        
                        if not estoque_prod.empty:
                            qm = estoque_prod[estoque_prod["LK-GRUPO"] == "QM"]
                            mf = estoque_prod[estoque_prod["LK-GRUPO"] == "MF"]
                            
                            if not qm.empty and qm.iloc[0]["Qtde"] >= qtd_pedida:
                                empresa = "QM"
                                qtde_disp = qm.iloc[0]["Qtde"]
                            elif not mf.empty and mf.iloc[0]["Qtde"] >= qtd_pedida:
                                empresa = "MF"
                                qtde_disp = mf.iloc[0]["Qtde"]
                        
                        resultados.append({
                            "CNPJ": row["CNPJ"],
                            "Codigo": codigo,
                            "Quantidade_Pedido": qtd_pedida,
                            "Valor_Unitario": row["Valor_Unitario"],
                            "Empresa_Atendimento": empresa,
                            "Qtde_Disponivel": qtde_disp,
                            "Valor_Item": row["Total_Item"],
                            "Arquivo_Pedido": file.name,
                        })
                
                progresso.progress((i + 1) / total_arquivos)
            
            progresso.empty()
            
            # =====================================================
            # RESULTADO FINAL
            # =====================================================
            df_resultado = pd.DataFrame(resultados)
            
            st.subheader("📊 Resultado da Alocação")
            st.dataframe(df_resultado, use_container_width=True)
            
            # =====================================================
            # FATURAMENTO POR EMPRESA (QM E MF)
            # =====================================================
            df_faturamento = (
                df_resultado
                .query("Empresa_Atendimento != 'SEM ESTOQUE'")
                .groupby("Empresa_Atendimento", as_index=False)["Valor_Item"]
                .sum()
            )
            
            # Garantir que QM e MF sempre apareçam (mesmo com 0)
            qm_valor = df_faturamento[df_faturamento["Empresa_Atendimento"] == "QM"]["Valor_Item"].sum() if not df_faturamento[df_faturamento["Empresa_Atendimento"] == "QM"].empty else 0.0
            mf_valor = df_faturamento[df_faturamento["Empresa_Atendimento"] == "MF"]["Valor_Item"].sum() if not df_faturamento[df_faturamento["Empresa_Atendimento"] == "MF"].empty else 0.0
            
            st.subheader("💰 Faturamento Previsto por Empresa (QM e MF)")
            col1, col2 = st.columns(2)
            
            valor_qm = f"R$ {qm_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            valor_mf = f"R$ {mf_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            col1.metric("QM", valor_qm)
            col2.metric("MF", valor_mf)
            
            # =====================================================
            # FATURAMENTO POR EMPRESA (INCLUINDO SEM ESTOQUE)
            # =====================================================
            df_faturamento_total = (
                df_resultado
                .groupby("Empresa_Atendimento", as_index=False)["Valor_Item"]
                .sum()
            )
            
            # Garantir que QM, MF e SEM ESTOQUE sempre apareçam (mesmo com 0)
            qm_valor_total = df_faturamento_total[df_faturamento_total["Empresa_Atendimento"] == "QM"]["Valor_Item"].sum() if not df_faturamento_total[df_faturamento_total["Empresa_Atendimento"] == "QM"].empty else 0.0
            mf_valor_total = df_faturamento_total[df_faturamento_total["Empresa_Atendimento"] == "MF"]["Valor_Item"].sum() if not df_faturamento_total[df_faturamento_total["Empresa_Atendimento"] == "MF"].empty else 0.0
            sem_estoque_valor = df_faturamento_total[df_faturamento_total["Empresa_Atendimento"] == "SEM ESTOQUE"]["Valor_Item"].sum() if not df_faturamento_total[df_faturamento_total["Empresa_Atendimento"] == "SEM ESTOQUE"].empty else 0.0
            
            st.subheader("💰 Faturamento Previsto por Empresa (Incluindo Sem Estoque)")
            col1, col2, col3 = st.columns(3)
            
            valor_qm_total = f"R$ {qm_valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            valor_mf_total = f"R$ {mf_valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            valor_sem_estoque = f"R$ {sem_estoque_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            col1.metric("QM", valor_qm_total)
            col2.metric("MF", valor_mf_total)
            col3.metric("SEM ESTOQUE", valor_sem_estoque)
            
            # =====================================================
            # GERAR ZIP COM CSVs SEPARADOS
            # =====================================================
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for (arquivo, empresa), df_grupo in df_resultado.groupby(["Arquivo_Pedido", "Empresa_Atendimento"]):
                    nome_csv = f"{arquivo.replace('.csv', '')}_{empresa}.csv"
                    zipf.writestr(nome_csv, df_grupo.to_csv(index=False, sep=";", decimal=",").encode("utf-8"))
            
            zip_buffer.seek(0)
            
            st.download_button(
                "⬇️ Baixar Resultados (ZIP)",
                zip_buffer,
                "alocacao_pedidos.zip",
                "application/zip"
            )
        
        except Exception as e:
            st.error(f"❌ Erro durante o processamento: {str(e)}")
            st.info("Verifique os arquivos enviados (formato, encoding, separadores) e tente novamente.")

    else:
        st.info("⬆️ Envie os arquivos de estoque (QM e MF) e os pedidos para iniciar o processamento.")

