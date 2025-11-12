import streamlit as st
import pandas as pd
import io
import zipfile

# Configurar página para modo wide
st.set_page_config(page_title="Arezzo Bolsa Valor", layout="wide")

def process_excel(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # 1. Criar novo DataFrame
        new_df = pd.DataFrame()

        # 2. Mapear e transformar colunas
        new_df['PEDIDO'] = df['PEDIDO'].astype(str).str.replace(',', '', regex=False)
        new_df['STYLE NAME'] = df['STYLE_NAME']
        new_df['MATERIAL-COLOR'] = df['MATERIAL_COLOR']
        new_df['SKU'] = df['SKU']
        new_df['MODELO'] = df['REFERENCIA_LOJA']
        new_df['EAN'] = df['EAN'].astype(str).str.replace(',', '', regex=False)
        new_df['TAM'] = df['TAM']
        new_df['QUANT'] = df['QUANT']
        new_df['QUANT EXTRA'] = df['QUANT'] + 5
        new_df['VALOR'] = df['MSRP'].apply(lambda x: f"{x:.2f}")
        new_df['NUM DA ETQ'] = ''
        new_df['VALOR DO FILTRO'] = 1
        df_ean_str = df['EAN'].astype(str).str.zfill(13)
        new_df['PREFIXO DA EMP'] = df_ean_str.str[:7]
        new_df['ITEM DE REF'] = '0' + df_ean_str.str[7:12]
        new_df['SERIAL'] = ''

        # 3. Duplicar linhas
        expanded_rows = []
        for _, row in new_df.iterrows():
            for _ in range(row['QUANT EXTRA']):
                expanded_rows.append(row.copy())
        
        expanded_df = pd.DataFrame(expanded_rows)

        # 4. Adicionar NUM DA ETQ
        # expanded_df['NUM DA ETQ'] = [f"{i:06d}" for i in range(1, len(expanded_df) + 1)]

        # 5. Reordenar colunas para a ordem final desejada
        final_columns = [
            'PEDIDO', 'STYLE NAME', 'MATERIAL-COLOR', 'SKU', 'MODELO', 'EAN', 
            'TAM', 'QUANT', 'QUANT EXTRA', 'VALOR', 'NUM DA ETQ', 'VALOR DO FILTRO',
            'PREFIXO DA EMP', 'ITEM DE REF', 'SERIAL'
        ]
        expanded_df = expanded_df[final_columns]

        return expanded_df

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
        return None

st.title("Processador de Planilhas Excel")
st.write("Faça o upload do arquivo Excel para processamento.")

uploaded_file = st.file_uploader("Escolha um arquivo Excel (.xlsx, .xlsm)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    st.write("Arquivo recebido! Processando...")
    processed_df = process_excel(uploaded_file)

    if processed_df is not None:
        st.success("Arquivo processado com sucesso!")
        st.write("Amostra dos dados processados (primeiras 50 linhas):")
        st.dataframe(processed_df.head(50))

        pedidos = processed_df['PEDIDO'].unique()
        st.write(f"Encontrados {len(pedidos)} pedidos únicos. Gerando planilhas individuais...")

        # Criar um arquivo ZIP em memória contendo todos os arquivos .xls
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for pedido in pedidos:
                pedido_df = processed_df[processed_df['PEDIDO'] == pedido].copy()
                
                # Converter todas as colunas para texto
                for col in pedido_df.columns:
                    pedido_df[col] = pedido_df[col].astype(str)

                # Criar um buffer para cada arquivo Excel individual
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    pedido_df.to_excel(writer, sheet_name='Dados', index=False)
                    
                    # Formatar todas as células como texto
                    worksheet = writer.sheets['Dados']
                    for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, 
                                                   min_col=1, max_col=worksheet.max_column):
                        for cell in row:
                            cell.number_format = '@'  # @ é o formato de texto no Excel
                
                # Adicionar o arquivo Excel ao ZIP
                excel_buffer.seek(0)
                zip_file.writestr(f"pedido_{pedido}.xlsx", excel_buffer.getvalue())

        # Voltar ao início do buffer para que o st.download_button possa lê-lo
        zip_buffer.seek(0)

        st.download_button(
            label="Baixar todas as planilhas (.zip com arquivos .xlsx)",
            data=zip_buffer,
            file_name="planilhas_por_pedido.zip",
            mime="application/zip"
        )

# Footer
st.divider()
st.markdown(
    "<div style='text-align: center; color: gray;'>Sistema de Geração de Planilhas | Desenvolvido por Fabio</div>",
    unsafe_allow_html=True
)
