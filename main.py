import streamlit as st
import pandas as pd
import plotly.express as px

# Definindo um tema personalizado e configurações da página
custom_theme = {
    "primaryColor": "#3498db",
    "backgroundColor": "#f4f4f4",
    "secondaryBackgroundColor": "#e0f7fa",
    "textColor": "#2c3e50",
    "font": "sans serif"
}
st.set_page_config(page_title="Análise de Dados", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    f"""
    <style>
        .reportview-container .main .block-container{{
            max-width: 1200px;
            padding-top: 2rem;
            padding-right: 2rem;
            padding-left: 2rem;
            padding-bottom: 3rem;
        }}
    </style>
    """,
    unsafe_allow_html=True
)
st.markdown(
    f"""
    <style>
        .sidebar .sidebar-content {{
            background-image: linear-gradient(
                0deg,
                #2c3e50,
                #2c3e50 50%,
                #3498db 50%
            );
            color: white;
        }}
        .btn-outline-secondary {{
            color: #2c3e50;
            background-color: white;
            border-color: #2c3e50;
        }}
        .btn-outline-secondary:hover {{
            color: white;
            background-color: #2c3e50;
            border-color: #2c3e50;
        }}
    </style>
    """,
    unsafe_allow_html=True
)


def highlight_vencido(row):
    if 'status' in row.index and 'Vencido' in row['status']:
        return ['background-color: #e74c3c; color: white;'] * len(row)
    return [''] * len(row)


def create_pie_chart(df, sheet_name):
    # Verificar se a coluna 'status' está presente no DataFrame
    if 'status' not in df.columns:
        if sheet_name not in ["CIV", "Opacidade"]:
            st.warning(f"A coluna 'status' não foi encontrada na aba '{sheet_name}'. Ignorando a aba.")
        return None

    fig = px.pie(df, names='status', title=f'Distribuição de Status - {sheet_name}', hole=0.5)

    try:
        fig.update_traces(marker=dict(
            colors=[
                '#118DFF' if 'Atenção' in status else
                '#164888' if 'Conforme' in status else
                '#EF4824' if 'Vencido' in status else
                '#3498db'  # Cor padrão para outras categorias
                for status in df['status']
            ]
        ))
        fig.update_traces(textposition='inside', textinfo='percent+label',
                          pull=[0.1 if 'Vencido' in status else 0 for status in df['status']])
    except TypeError:
        st.warning("Ocorreu um erro ao tentar criar o gráfico. Verifique se os dados estão corretos.")
        return None

    labels = df['status'].unique()

    if not all(isinstance(label, str) for label in labels):
        st.warning("Ocorreu um erro ao tentar obter os rótulos. Verifique se os dados estão corretos.")
        return None

    values = [df[df['status'] == label].shape[0] for label in labels]

    # Adicionando anotações com o número total de cada categoria abaixo do gráfico
    annotations = [
        f"{label}: {value}" for label, value in zip(labels, values)
    ]

    # Adicionando a contagem de 'Vencido em operação' abaixo do gráfico
    vencido_count = df[(df['status'] == 'Vencido') & (df['observação\n'] == 'EM OPERAÇÃO')].shape[0]
    annotations.append(f"Vencido em operação: {vencido_count}")

    fig.update_layout(
        annotations=[
            {
                "x": 0.5,
                "y": -0.2,  # Ajuste para a posição abaixo dos rótulos
                "text": "<br>".join(annotations),
                "showarrow": False,
                "font": {"size": 12},
                "xref": "paper",
                "yref": "paper",
                "xanchor": "center",
                "yanchor": "bottom"
            }
        ]
    )

    return fig


def main():
    st.title("Análise de Dados")

    # Especificando o nome do arquivo Excel
    excel_filename = "PythonStreamlit.xlsx"

    # Lista de abas no arquivo Excel
    sheet_names = pd.ExcelFile(excel_filename).sheet_names + ["Todos os Gráficos"]

    # Adicionando a opção de visualizar todos os gráficos
    selected_sheet = st.sidebar.selectbox("Escolha a aba:", sheet_names)

    try:
        # Carregando os dados do Excel usando Pandas, especificando a aba
        if selected_sheet == "Todos os Gráficos":
            # Visualização de todos os gráficos
            cols = st.columns(2)
            for i, sheet_name in enumerate(sheet_names[:-1]):  # Excluindo a última entrada ("Perfil Securitário")
                df_sheet = pd.read_excel(excel_filename, sheet_name=sheet_name)
                if 'status' in df_sheet.columns:
                    with cols[i % 2]:
                        st.plotly_chart(create_pie_chart(df_sheet, sheet_name), use_container_width=True)

        elif selected_sheet in ["Veículos Bloqueados", "Data"]:
            # Visualização da tabela para as abas específicas "Veículos Bloqueados" e "Data"
            df = pd.read_excel(excel_filename, sheet_name=selected_sheet)
            st.write(f"Dados carregados da aba '{selected_sheet}':")
            st.dataframe(df.style.apply(highlight_vencido, axis=1))

        else:
            # Visualização individual para uma aba específica
            df = pd.read_excel(excel_filename, sheet_name=selected_sheet)
            pie_chart = create_pie_chart(df, selected_sheet)

            # Verificar se o gráfico foi criado com sucesso antes de exibi-lo
            if pie_chart is not None:
                st.plotly_chart(pie_chart, use_container_width=True)

                # Adicionando mensagens de depuração
                st.write(f"Dados carregados da aba '{selected_sheet}':")
                st.dataframe(df.style.apply(highlight_vencido, axis=1))

    except FileNotFoundError:
        st.error(
            f"Arquivo {excel_filename} não encontrado. Certifique-se de que o arquivo está no mesmo diretório do script.")
    except pd.errors.EmptyDataError:
        st.error(f"A aba selecionada ({selected_sheet}) não contém dados.")
    except Exception as e:
        st.error(f"Erro durante a leitura do arquivo Excel: {e}")


if __name__ == "__main__":
    main()

