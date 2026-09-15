import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Calculadora de Markup & Descontos AGRU", layout="wide")

st.title("📊 Análise Comercial: Markup, Impostos e Descontos")
st.markdown("Automação para apuração de preços líquidos, markups estimados/reais e recálculo de descontos sucessivos.")

# --- BARRA LATERAL: PARÂMETROS E UPLOAD DAS BASES ---
with st.sidebar:
    st.header("⚙️ 1. Arquivos de Custos")
    
    file_lp = st.file_uploader("Planilha: Lista de Preços", type=["xlsx", "xls"])
    col_lp_cod = st.text_input("Coluna Código (Lista de Preços)", value="C").strip().upper()
    col_lp = st.text_input("Coluna Landed (Custo Estimado)", value="Y").strip().upper()
    
    file_ev = st.file_uploader("Planilha: Evolução de Custos", type=["xlsx", "xls"])
    col_ev_protheus = st.text_input("Coluna Custo Protheus", value="C").strip().upper()
    col_ev_rot = st.text_input("Coluna Custo Rotação", value="AN").strip().upper()
    
    st.divider()
    st.header("📄 2. Proposta Comercial")
    pdf_file = st.file_uploader("Upload da Proposta Comercial (PDF)", type=["pdf"])

def col2idx(col_letter):
    col_letter = col_letter.upper()
    idx = 0
    for c in col_letter:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1

# --- CARREGAMENTO DAS BASES DE CUSTOS ---
@st.cache_data
def carregar_bases(file_lp, col_lp_cod_idx, col_lp_idx, file_ev, col_c_idx, col_an_idx):
    custos = {}
    
    # 1. Lista de Preços (Custo Estimado Landed)
    if file_lp is not None:
        try:
            df_lp = pd.read_excel(file_lp, sheet_name=0)
            for _, row in df_lp.iterrows():
                if col_lp_cod_idx < len(row):
                    code_raw = str(row.iloc[col_lp_cod_idx]).split('.')[0].strip()
                    if code_raw and code_raw != 'nan':
                        val = pd.to_numeric(row.iloc[col_lp_idx], errors='coerce') if col_lp_idx < len(row) else 0.0
                        custos[code_raw] = {'custo_landed': float(val) if pd.notnull(val) else 0.0, 'custo_real': 0.0}
        except Exception as e:
            st.error(f"Erro ao ler Lista de Preços: {e}")

    # 2. Evolução de Custos (Maior valor entre Protheus e Rotação)
    if file_ev is not None:
        try:
            df_ev = pd.read_excel(file_ev, sheet_name=0)
            for _, row in df_ev.iterrows():
                code_raw = str(row.iloc[0]).split('.')[0].strip()
                if code_raw and code_raw != 'nan':
                    val_c = pd.to_numeric(row.iloc[col_c_idx], errors='coerce') if col_c_idx < len(row) else 0.0
                    val_an = pd.to_numeric(row.iloc[col_an_idx], errors='coerce') if col_an_idx < len(row) else 0.0
                    val_c = float(val_c) if pd.notnull(val_c) else 0.0
                    val_an = float(val_an) if pd.notnull(val_an) else 0.0
                    maior_real = max(val_c, val_an)
                    
                    if code_raw in custos:
                        custos[code_raw]['custo_real'] = maior_real
                    else:
                        custos[code_raw] = {'custo_landed': 0.0, 'custo_real': maior_real}
        except Exception as e:
            st.error(f"Erro ao ler Evolução de Custos: {e}")
            
    return custos

# --- PARSER DEFINITIVO DO PDF DA PROPOSTA ---
def parse_proposta_pdf(uploaded_pdf):
    reader = pypdf.PdfReader(uploaded_pdf)
    full_text = "\n".join([page.extract_text() or "" for page in reader.pages])
    
    # 1. Impostos destacados nas observações (ICMS, PIS e COFINS)
    imp_match = re.search(r'Impostos\s+inclu[íi]dos:\s*([^.\n\r]+)', full_text, re.IGNORECASE)
    aliquotas = []
    detalhe_impostos = ""
    if imp_match:
        detalhe_impostos = imp_match.group(1).strip()
        aliquotas_matches = re.findall(r'([0-9.,]+)\s*%', detalhe_impostos)
        aliquotas = [float(a.replace(',', '.')) for a in aliquotas_matches]
    aliquota_total = sum(aliquotas) / 100.0

    # 2. Representante
    rep_match = re.search(r'Representante\s*([^\n\r]+)', full_text)
    representante = "Sem representante"
    if rep_match:
        rep_raw = rep_match.group(1).strip()
        rep_clean = re.split(r'Telefones|Cidade|Contato|e-mail', rep_raw)[0].strip()
        if rep_clean and "Cidade" not in rep_clean:
            representante = rep_clean
            
    # 3. Frete destacado
    frete_match = re.search(r'\bFRETE\s+([0-9.,]+)', full_text)
    frete_destacado = float(frete_match.group(1).replace('.', '').replace(',', '.')) if frete_match else 0.0

    # 4. Itens da Proposta
    item_matches = list(re.finditer(r'(\d+\.\d+)\s*(\d{11})', full_text))
    fim_tabela = re.search(r'\n\s*VENDEDOR\b', full_text)
    fim_idx = fim_tabela.start() if fim_tabela else len(full_text)
    
    itens = []
    for i, m in enumerate(item_matches):
        pos = m.group(1)
        cod = m.group(2)
        start = m.end()
        end = item_matches[i+1].start() if (i + 1 < len(item_matches)) else fim_idx
        bloco = full_text[start:end].strip()
        
        # Padrão numérico final: NCM(8) + ST% + IPI% + QTD + UN + [UNITÁRIO + SUBTOTAL]
        vals_match = re.search(r'(\d{8})\s*([0-9.,]+)%\s*([0-9.,]+)%\s*([0-9.,]+)\s*([A-Z]{2})\s*(.*)', bloco, re.DOTALL)
        if vals_match:
            ncm = vals_match.group(1)
            qtd_str = vals_match.group(4)
            un = vals_match.group(5)
            resto_precos = vals_match.group(6).strip()
            
            desc_raw = bloco[:vals_match.start()].strip()
            desc_raw = re.sub(r'(Imediata|\d+\s*dias|\d{2}/\d{2}/\d{4})\s*$', '', desc_raw, flags=re.IGNORECASE).strip()
            desc = " ".join(desc_raw.split())
            
            qtd = float(qtd_str.replace('.', '').replace(',', '.'))
            
            # Tratamento para separar valores unitários e subtotais (mesmo se vierem colados)
            resto_clean = " ".join(resto_precos.split())
            parts = resto_clean.split()
            if len(parts) >= 2:
                unit = float(parts[0].replace('.', '').replace(',', '.'))
                subtotal = float(parts[1].replace('.', '').replace(',', '.'))
            else:
                m_colados = re.match(r'^([0-9.,]+?,\d{2})([0-9.,]+?,\d{2})$', resto_clean)
                if m_colados:
                    unit = float(m_colados.group(1).replace('.', '').replace(',', '.'))
                    subtotal = float(m_colados.group(2).replace('.', '').replace(',', '.'))
                else:
                    unit = 0.0
                    subtotal = 0.0
                    
            itens.append({
                'pos': pos,
                'codigo': cod,
                'descricao': desc,
                'ncm': ncm,
                'quantidade': qtd,
                'un': un,
                'unit_original': unit,
                'subtotal_original': subtotal
            })
            
    return representante, frete_destacado, aliquota_total, detalhe_impostos, itens

# --- FLUXO PRINCIPAL ---
if pdf_file is not None:
    rep, frete_destacado, aliquota_total, detalhe_imp, itens_extraidos = parse_proposta_pdf(pdf_file)
    
    if not itens_extraidos:
        st.error("Não foi possível identificar itens no formato da AGRU neste arquivo PDF.")
    else:
        custos_db = carregar_bases(file_lp, col2idx(col_lp_cod), col2idx(col_lp), file_ev, col2idx(col_ev_protheus), col2idx(col_ev_rot))
        
        st.subheader("📋 Dados da Proposta Identificados")
        c1, c2, c3 = st.columns(3)
        c1.metric("Representante", rep)
        c2.metric("Frete Destacado em Nota", f"R$ {frete_destacado:,.2f}")
        c3.metric("Impostos Descontados", f"{aliquota_total * 100:.2f}%")
        if detalhe_imp:
            st.caption(f"ℹ️ **Tributos identificados nas observações:** {detalhe_imp}")
        
        # Frete Embutido
        tem_frete_embutido = False
        valor_frete_embutido = 0.0
        if frete_destacado == 0.0:
            c_frete1, c_frete2 = st.columns(2)
            tem_frete_embutido = c_frete1.checkbox("A proposta possui frete embutido nos produtos?", value=False)
            if tem_frete_embutido:
                valor_frete_embutido = c_frete2.number_input("Valor Total do Frete Embutido a expurgar (R$):", min_value=0.0, value=0.0, step=50.0)
                
        # Simulação de Desconto
        st.divider()
        st.subheader("🎯 Simulação de Desconto e Margem")
        
        tipo_desconto = st.radio("Onde deseja aplicar o desconto adicional?", ["Em Toda a Proposta", "Apenas em um Item Específico"], horizontal=True)
        
        lista_pos_itens = [f"{it['pos']} - {it['descricao']} (Cód: {it['codigo']})" for it in itens_extraidos]
        pos_selecionada = None
        
        if tipo_desconto == "Apenas em um Item Específico":
            item_sel = st.selectbox("Selecione o Item:", lista_pos_itens)
            pos_selecionada = item_sel.split(' - ')[0]

        d_col1, d_col2 = st.columns(2)
        desconto_atual_pct = d_col1.number_input("Desconto atual na proposta (%):", min_value=0.0, max_value=99.0, value=71.0, step=0.5)
        adicional_pct = d_col2.number_input("Conceder desconto adicional (%):", min_value=0.0, max_value=99.0, value=0.0, step=0.5)
        
        fator_atual = 1.0 - (desconto_atual_pct / 100.0)
        fator_novo = fator_atual * (1.0 - (adicional_pct / 100.0))
        novo_desconto_composto_pct = (1.0 - fator_novo) * 100.0
        
        if adicional_pct > 0.0:
            alvo_txt = "na proposta" if tipo_desconto == "Em Toda a Proposta" else f"no item {pos_selecionada}"
            st.info(f"💡 **Novo Desconto Total {alvo_txt}:** **{novo_desconto_composto_pct:.2f}%** (calculado via {desconto_atual_pct:.2f}% + {adicional_pct:.2f}%)")

        dados = []
        subtotal_faturado_total = sum(i['subtotal_original'] for i in itens_extraidos)
        fator_liquido = (1.0 - aliquota_total)

        for it in itens_extraidos:
            cod = str(it['codigo'])
            qtd = it['quantidade']
            sub_orig = it['subtotal_original']
            
            aplica_desc = (tipo_desconto == "Em Toda a Proposta") or (tipo_desconto == "Apenas em um Item Específico" and it['pos'] == pos_selecionada)
            fator_mult_preco = (1.0 - adicional_pct / 100.0) if aplica_desc else 1.0
            
            sub_novo = sub_orig * fator_mult_preco
            unit_novo = sub_novo / qtd
            
            # Preço sem impostos = Faturado * (1 - alíquotas)
            unit_sem_imp = unit_novo * fator_liquido
            
            # Rateio do frete embutido
            rateio_frete_item = ((sub_orig / subtotal_faturado_total) * valor_frete_embutido) if (tem_frete_embutido and subtotal_faturado_total > 0) else 0.0
            unit_frete = rateio_frete_item / qtd
            unit_liquido_efetivo = unit_sem_imp - (unit_frete * fator_liquido)
            
            info_c = custos_db.get(cod, {'custo_landed': 0.0, 'custo_real': 0.0})
            c_est = info_c['custo_landed']
            c_real = info_c['custo_real']
            
            dados.append({
                'Item': it['pos'],
                'Código': cod,
                'Descrição': it['descricao'],
                'Qtd': qtd,
                'Unit. Faturado (R$)': unit_novo,
                'Unit. sem Imposto (R$)': unit_sem_imp,
                'Frete Unit. Rateado (R$)': unit_frete,
                'Unit. Líq. Efetivo (R$)': unit_liquido_efetivo,
                'Custo Estimado (R$)': c_est,
                'Custo Real (R$)': c_real,
                'Markup Unit. Estimado': (unit_liquido_efetivo / c_est) if c_est > 0 else 0.0,
                'Markup Unit. Real': (unit_liquido_efetivo / c_real) if c_real > 0 else 0.0,
                'Subtotal Líquido': unit_liquido_efetivo * qtd,
                'Sub Custo Estimado': c_est * qtd,
                'Sub Custo Real': c_real * qtd,
                'Alerta Custo': (c_est == 0.0 or c_real == 0.0)
            })
            
        df_resultado = pd.DataFrame(dados)
        
        tot_liq = df_resultado['Subtotal Líquido'].sum()
        tot_custo_est = df_resultado['Sub Custo Estimado'].sum()
        tot_custo_real = df_resultado['Sub Custo Real'].sum()
        
        mk_geral_est = (tot_liq / tot_custo_est) if tot_custo_est > 0 else 0.0
        mk_geral_real = (tot_liq / tot_custo_real) if tot_custo_real > 0 else 0.0
        
        st.divider()
        st.subheader("📈 Resultado Consolidado da Proposta")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Faturamento Líquido Total", f"R$ {tot_liq:,.2f}")
        m2.metric("Markup Geral Estimado", f"{mk_geral_est:.2f}x" if mk_geral_est > 0 else "N/A")
        m3.metric("Markup Geral Real", f"{mk_geral_real:.2f}x" if mk_geral_real > 0 else "N/A")
        m4.metric("Diferença Margem (Real vs Est.)", f"{(mk_geral_real - mk_geral_est):+.2f}x" if mk_geral_est > 0 and mk_geral_real > 0 else "-")
        
        itens_alerta = df_resultado[df_resultado['Alerta Custo']]
        if not itens_alerta.empty:
            st.error(f"⚠️ Atenção: {len(itens_alerta)} item(ns) estão com Custo Estimado ou Real zerado/não localizado na base (destacados em vermelho abaixo).")

        def highlight_missing_cost(row):
            return ['color: #d90429; font-weight: bold; background-color: #ffebee;'] * len(row) if row['Alerta Custo'] else [''] * len(row)

        df_view = df_resultado.drop(columns=['Subtotal Líquido', 'Sub Custo Estimado', 'Sub Custo Real'])
        styled_df = df_view.style.apply(highlight_missing_cost, axis=1).format({
            'Unit. Faturado (R$)': '{:,.2f}',
            'Unit. sem Imposto (R$)': '{:,.2f}',
            'Frete Unit. Rateado (R$)': '{:,.2f}',
            'Unit. Líq. Efetivo (R$)': '{:,.2f}',
            'Custo Estimado (R$)': '{:,.2f}',
            'Custo Real (R$)': '{:,.2f}',
            'Markup Unit. Estimado': '{:.2f}x',
            'Markup Unit. Real': '{:.2f}x',
            'Qtd': '{:,.2f}'
        })
        
        st.write("### Detalhamento por Item")
        st.dataframe(styled_df, use_container_width=True, height=520)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_view.drop(columns=['Alerta Custo']).to_excel(writer, index=False, sheet_name='Markup Proposta')
            
        st.download_button(
            label="📥 Baixar Análise em Excel",
            data=buffer.getvalue(),
            file_name="analise_markup_proposta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Por favor, faça o upload do PDF da proposta na barra lateral para iniciar a análise.")