import streamlit as st
import pandas as pd
from datetime import datetime
import unidecode
import io

# Configuração da página
st.set_page_config(page_title="Processador Palanilha", layout="wide")

st.title("📊 Processador de planilha de licenças")
st.markdown("Suba sua planilha original e defina os filtros para gerar o relatório consolidado.")

# --- BARRA LATERAL (INPUTS) ---
with st.sidebar:
    st.header("Configurações")
    data_ref = st.date_input(
    "Data de Referência", 
    value=datetime(2025, 7, 31),
    format="DD/MM/YYYY"
)
    hoje = pd.Timestamp(data_ref)
    
    uploaded_file = st.file_uploader("Upload do arquivo BDU.xlsx", type=["xlsx"])
    
    st.divider()
    st.header("Filtros de Saída")
    # Lista de documentos que você deseja permitir filtrar
    # Você pode expandir esta lista conforme necessário
    docs_opcoes = ["ALVARA", "AVCB", "SANITARIA", "PUBLICIDADE", "L.AMBIENTAL"]
    docs_selecionados = st.multiselect(
        "Selecione os documentos para o relatório final:",
        options=docs_opcoes,
        default=["ALVARA", "AVCB"] # Deixa estes dois como padrão
    )

# --- FUNÇÕES DE PROCESSAMENTO ---
def processar_dados(file, data_referencia, lista_docs):
    dados = pd.read_excel(file, sheet_name='ANÁLITICO LICENÇAS', header=1)
    df = pd.DataFrame(dados)
    
    df.columns = [unidecode.unidecode(col).strip().upper() for col in df.columns]
    
    colunas_para_remover = [
        "ID. IMOVEL BENNER", "EMPRESA LEGAL", "BCO", "TIPO",
        "UNIDADE", "ENDERECO", "CIDADE", "UF", "CEP", "BAIRRO", "SUBPREFEITURA",
        "REDE", "REGIONAL", "REGIAO", "LOCALIDADE", "CNPJ", "AREA CONSTRUIDA",
        "PROPRIEDADE", "N. MATRICULA", "TIPO DO IMOVEL", "IBM",
        "SANTANDER", "TOMBAMENTO", "CONFERENCIA2", "DETALHE", "SIGLA ORGAO",
        "CONFERENCIA", "INAUGURACAO", "PROPRIEDADE2", "ENTREGAVEIS SIM OU NAO",
        "DATA ENTREGA OU PREVISAO DE ENTREGA"
    ]
    df = df.drop(columns=[col for col in colunas_para_remover if col in df.columns])
    
    df_melted = df.melt(id_vars=["AGENCIA","CARTEIRA"], var_name="ATRIBUTO", value_name="VALOR")
    df_melted["ATRIBUTO"] = df_melted["ATRIBUTO"].str.replace("DETALHE DA PENDENCIA", "AVCB - DETALHE DA PENDENCIA", regex=False)
    
    # Split com limpeza para evitar deslocamento de colunas
    df_temp = df_melted["ATRIBUTO"].str.split(" - ", n=1, expand=True)
    df_melted["DOCUMENTO"] = df_temp[0].str.strip()
    df_melted["CAMPO"] = df_temp[1].str.strip()
    
    df_melted["CAMPO"] = df_melted["CAMPO"].str.replace("PENDENCIASQue mais afeta PROC ADM", "PENDENCIAS Que mais afeta", regex=False)
    df_melted["CAMPO"] = df_melted["CAMPO"].str.replace("STATUS PROC ADM", "STATUS", regex=False)
    
    df_pivot = df_melted.pivot_table(
        index=["AGENCIA", "CARTEIRA","DOCUMENTO"],
        columns="CAMPO",
        values="VALOR",
        aggfunc="first"
    ).reset_index()
    
    df_pivot.columns = [col.strip().upper() for col in df_pivot.columns]
    
    # --- FILTRO DE DOCUMENTOS ---
    # Filtramos aqui para que os cálculos e o arquivo final contenham apenas o desejado
    if lista_docs:
        df_pivot = df_pivot[df_pivot["DOCUMENTO"].isin(lista_docs)]
    
    # Datas e Cálculos
    for col_dt in ["EMISSAO", "VALIDADE", "PROTOCOLO"]:
        if col_dt in df_pivot.columns:
            df_pivot[col_dt] = pd.to_datetime(df_pivot[col_dt].astype(str).str.replace("-", "").replace("PERMANENTE", "").str.strip(), errors='coerce')
            # Trava para evitar o erro de 1969/Datas bizzaras
            df_pivot.loc[df_pivot[col_dt] < pd.Timestamp("1980-01-01"), col_dt] = pd.NaT

    df_pivot["TEMPO EXECUÇÃO"] = (data_referencia - df_pivot["PROTOCOLO"]).dt.days
    df_pivot["PERM_RENOV"] = df_pivot["VALIDADE"].apply(lambda x: "PERMANENTE" if pd.isna(x) else "RENOVAVEL")

    def classificar_tempo(dias):
        if pd.isna(dias): return None
        elif dias <= 183: return "0 a 6 meses"
        elif dias <= 365: return "6 meses a 1 ano"
        else: return "Acima de 1 ano"

    df_pivot["FAIXA_TEMPO"] = df_pivot["TEMPO EXECUÇÃO"].apply(classificar_tempo)

    if "RESPONSABILIDADE" in df_pivot.columns:
        df_pivot["RESPONSABILIDADE"] = df_pivot["RESPONSABILIDADE"].replace({"JURIDICO": "JURÍDICO", "CONDOMINIO": "CONDOMÍNIO", "MANUTENCAO": "MANUTENÇÃO"})
    else:
        df_pivot["RESPONSABILIDADE"] = pd.NA

    df_pivot["DATA_PROCESSAMENTO"] = data_referencia
    df_pivot = df_pivot[df_pivot["AGENCIA"].notna()]
    
    colunas_finais = ["AGENCIA","CARTEIRA" ,"DOCUMENTO", "RESPONSABILIDADE", "EMISSAO", "VALIDADE","PROTOCOLO","TEMPO EXECUÇÃO","FAIXA_TEMPO" ,'PENDENCIAS"QUE MAIS AFETA"',"HISTORICO DETALHADO","STATUS", "PERM_RENOV", "DATA_PROCESSAMENTO"]
    colunas_existentes = [c for c in colunas_finais if c in df_pivot.columns]
    
    return df_pivot[colunas_existentes]

# --- LÓGICA PRINCIPAL ---
if uploaded_file is not None:
    try:
        with st.spinner('Processando dados...'):
            # Passamos a lista de documentos selecionados para a função
            df_final = processar_dados(uploaded_file, hoje, docs_selecionados)
            
            # --- CÁLCULOS DE RESUMO ---
            ativos = df_final[df_final["CARTEIRA"] != "FUSIONADA"]
            
            # Filtro de responsabilidade para os indicadores do Dashboard
            docs_alvos = ativos[ativos["RESPONSABILIDADE"].isin(["EMPRESA LEGAL", "-", ""])]
            
            total_docs_alvos = len(docs_alvos)
            docs_validos = len(docs_alvos[docs_alvos["STATUS"].isin(["VÁLIDO", "EM EXECUÇÃO"])])
            percentual_regular = (docs_validos/ total_docs_alvos)*100 if total_docs_alvos > 0 else 0
            
            # --- DASHBOARD ---
            st.subheader(f"Indicadores: {', '.join(docs_selecionados)}")
            col1, col2, col3 = st.columns(3)
            col1.metric("Unidades com estes docs", len(df_final["AGENCIA"].unique()))
            col2.metric("Percentual Regular", f"{percentual_regular:.2f}%")
            col3.metric("Total de Docs Processados", len(df_final))

            # --- EXCEL PARA DOWNLOAD ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Base_Filtrada')
                
                resumo_data = {
                    'DATA_REFERENCIA': [hoje.strftime('%d/%m/%Y')],
                    'DOCS_FILTRADOS': [", ".join(docs_selecionados)],
                    'PERCENTUAL_REGULAR': [percentual_regular],
                    'TOTAL_LINHAS': [len(df_final)]
                }
                pd.DataFrame(resumo_data).to_excel(writer, sheet_name='Resumo', index=False)
            
            processed_data = output.getvalue()

            st.success("✅ Processamento concluído com filtros aplicados!")
            
            st.download_button(
                label="📥 Baixar Excel Filtrado",
                data=processed_data,
                file_name=f"bdu_filtrado_{hoje.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.subheader("Visualização dos Dados (Filtrados)")
            st.dataframe(df_final)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("Aguardando upload do arquivo Excel.")