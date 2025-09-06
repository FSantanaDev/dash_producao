# -------------------------- IMPORTAÇÃO DE BIBLIOTECAS --------------------------
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import os
import io
from datetime import datetime

# -------------------------- CONFIGURAÇÃO DA PÁGINA --------------------------
st.set_page_config(page_title="Dashboard de Análise - Agosto", layout="wide")

# -------------------------- FUNÇÃO PARA LER ARQUIVOS --------------------------
def read_excel_file(file_path):
    """
    Lê arquivo Excel específico e retorna um DataFrame do pandas.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None

# -------------------------- INTERFACE STREAMLIT --------------------------
st.title("📊 Dashboard de Análise - Agosto 2025")
st.markdown(
    "Este dashboard apresenta os dados de atendimentos do mês de agosto de 2025, com filtros interativos, KPIs e visualizações.")

# Caminho do arquivo Excel
file_path = "Analise_Agosto.xlsx"

# Verifica se o arquivo existe
if os.path.exists(file_path):
    # Lê a planilha
    df = read_excel_file(file_path)

    if df is not None and not df.empty:
        # Exibe as primeiras linhas
        with st.expander("👀 Amostra dos dados (primeiras linhas)"):
            st.dataframe(df.head())
        
        # Adiciona coluna de receita
        df['Receita'] = df['Quantidade'] * df['ValorUnitario']
        
        # Converte a coluna de data para o formato correto
        df['dataRealizado'] = pd.to_datetime(df['dataRealizado'])
        df['Dia'] = df['dataRealizado'].dt.day
        
        # -------------------- FILTROS --------------------
        st.sidebar.header("🔍 Filtros")
        
        # Filtro de unidade
        unidades = ["Todas"] + sorted(df['Unidade'].unique().tolist())
        unidade_selecionada = st.sidebar.selectbox("Unidade", unidades)
        
        # Filtro de categoria
        categorias = ["Todas"] + sorted(df['Categoria'].unique().tolist())
        categoria_selecionada = st.sidebar.selectbox("Categoria", categorias)
        
        # Filtro de subárea
        subareas = ["Todas"] + sorted(df['Subarea'].unique().tolist())
        subarea_selecionada = st.sidebar.selectbox("Subárea", subareas)
        
        # Filtro de tipo de atendimento
        tipos_atendimento = ["Todos"] + sorted(df['TipoAtendimento'].unique().tolist())
        tipo_atendimento_selecionado = st.sidebar.selectbox("Tipo de Atendimento", tipos_atendimento)
        
        # Filtro de tipo de serviço
        tipos_servico = ["Todos"] + sorted(df['TipoServico'].unique().tolist())
        tipo_servico_selecionado = st.sidebar.selectbox("Tipo de Serviço", tipos_servico)
        
        # Aplicar filtros
        df_filtrado = df.copy()
        
        if unidade_selecionada != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Unidade'] == unidade_selecionada]
            
        if categoria_selecionada != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Categoria'] == categoria_selecionada]
            
        if subarea_selecionada != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Subarea'] == subarea_selecionada]
            
        if tipo_atendimento_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['TipoAtendimento'] == tipo_atendimento_selecionado]
            
        if tipo_servico_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado['TipoServico'] == tipo_servico_selecionado]
        
        # -------------------- CÁLCULO DE KPIs --------------------
        # KPIs gerais
        qtd_total = df_filtrado['Quantidade'].sum()
        rec_total = df_filtrado['Receita'].sum()
        valor_medio = rec_total / qtd_total if qtd_total > 0 else 0
        num_atendimentos = len(df_filtrado)
        
        # -------------------- INDICADORES (KPIs) --------------------
        st.subheader("🔢 Indicadores (KPIs) do Filtro Atual")
        
        # KPIs gerais
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Quantidade Total", f"{int(qtd_total):,}".replace(",", "."))
        col2.metric("Receita Total (R$)", f"{rec_total:,.2f}".replace(",", "."))
        col3.metric("Valor Médio (R$)", f"{valor_medio:,.2f}".replace(",", "."))
        col4.metric("Número de Atendimentos", f"{num_atendimentos:,}".replace(",", "."))
        
        # -------------------- VISUALIZAÇÕES --------------------
        st.subheader("📈 Visualizações")
        
        # Layout de duas colunas para os gráficos
        col1, col2 = st.columns(2)
        
        # 1. Gráfico de barras: Quantidade por Unidade
        with col1:
            df_unidade = df_filtrado.groupby('Unidade')['Quantidade'].sum().reset_index().sort_values('Quantidade', ascending=False)
            fig1 = px.bar(
                df_unidade, 
                x='Unidade', 
                y='Quantidade', 
                title="Quantidade por Unidade", 
                color='Unidade'
            )
            fig1.update_layout(xaxis_title="Unidade", yaxis_title="Quantidade")
            st.plotly_chart(fig1, use_container_width=True)
        
        # 2. Gráfico de barras: Receita por Unidade
        with col2:
            df_unidade_receita = df_filtrado.groupby('Unidade')['Receita'].sum().reset_index().sort_values('Receita', ascending=False)
            fig2 = px.bar(
                df_unidade_receita, 
                x='Unidade', 
                y='Receita', 
                title="Receita por Unidade", 
                color='Unidade'
            )
            fig2.update_layout(xaxis_title="Unidade", yaxis_title="Receita (R$)")
            st.plotly_chart(fig2, use_container_width=True)
        
        # 3. Gráfico de pizza: Distribuição por Categoria
        with col1:
            df_categoria = df_filtrado.groupby('Categoria')['Quantidade'].sum().reset_index()
            fig3 = px.pie(
                df_categoria, 
                values='Quantidade', 
                names='Categoria', 
                title="Distribuição por Categoria"
            )
            st.plotly_chart(fig3, use_container_width=True)
        
        # 4. Gráfico de pizza: Distribuição por Tipo de Atendimento
        with col2:
            df_tipo_atendimento = df_filtrado.groupby('TipoAtendimento')['Quantidade'].sum().reset_index()
            fig4 = px.pie(
                df_tipo_atendimento, 
                values='Quantidade', 
                names='TipoAtendimento', 
                title="Distribuição por Tipo de Atendimento"
            )
            st.plotly_chart(fig4, use_container_width=True)
        
        # 5. Gráfico de barras: Top 10 Serviços mais realizados
        df_servicos = df_filtrado.groupby('NMServico')['Quantidade'].sum().reset_index().sort_values('Quantidade', ascending=False).head(10)
        fig5 = px.bar(
            df_servicos, 
            x='Quantidade', 
            y='NMServico', 
            orientation='h',
            title="Top 10 Serviços mais realizados",
            color='Quantidade',
            color_continuous_scale='Viridis'
        )
        fig5.update_layout(
            yaxis={'categoryorder':'total ascending'},
            xaxis_title="Quantidade",
            yaxis_title="Serviço"
        )
        st.plotly_chart(fig5, use_container_width=True)
        
        # 5.1 Gráfico de barras: Top 10 Serviços que mais trouxeram faturamento
        df_servicos_faturamento = df_filtrado.groupby('NMServico')['Receita'].sum().reset_index().sort_values('Receita', ascending=False).head(10)
        fig5_1 = px.bar(
            df_servicos_faturamento, 
            x='Receita', 
            y='NMServico', 
            orientation='h',
            title="Top 10 Serviços que mais trouxeram faturamento",
            color='Receita',
            color_continuous_scale='Viridis'
        )
        fig5_1.update_layout(
            yaxis={'categoryorder':'total ascending'},
            xaxis_title="Receita (R$)",
            yaxis_title="Serviço"
        )
        st.plotly_chart(fig5_1, use_container_width=True)
        
        # 6. Gráfico de linha: Evolução diária de atendimentos
        df_diario = df_filtrado.groupby('Dia')['Quantidade'].sum().reset_index()
        fig6 = px.line(
            df_diario, 
            x='Dia', 
            y='Quantidade', 
            title="Evolução Diária de Atendimentos",
            markers=True
        )
        fig6.update_layout(
            xaxis_title="Dia do Mês",
            yaxis_title="Quantidade"
        )
        st.plotly_chart(fig6, use_container_width=True)
        
        # 7. Mapa de calor: Subárea vs Tipo de Atendimento
        df_heatmap = df_filtrado.pivot_table(
            index='Subarea', 
            columns='TipoAtendimento', 
            values='Quantidade', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # Seleciona as top 10 subáreas para o mapa de calor
        top_subareas = df_filtrado.groupby('Subarea')['Quantidade'].sum().nlargest(10).index.tolist()
        df_heatmap_filtered = df_filtrado[df_filtrado['Subarea'].isin(top_subareas)]
        
        if not df_heatmap_filtered.empty:
            df_heatmap = df_heatmap_filtered.pivot_table(
                index='Subarea', 
                columns='TipoAtendimento', 
                values='Quantidade', 
                aggfunc='sum',
                fill_value=0
            )
            
            fig7 = px.imshow(
                df_heatmap,
                labels=dict(x="Tipo de Atendimento", y="Subárea", color="Quantidade"),
                title="Mapa de Calor: Subárea vs Tipo de Atendimento",
                color_continuous_scale='Viridis'
            )
            st.plotly_chart(fig7, use_container_width=True)
        
        # -------------------- TABELA DETALHADA --------------------
        st.subheader("📋 Tabela Detalhada")
        
        # Agrupa os dados por Unidade, Subárea e Tipo de Atendimento
        df_agrupado = df_filtrado.groupby(['Unidade', 'Subarea', 'TipoAtendimento']).agg({
            'Quantidade': 'sum',
            'Receita': 'sum'
        }).reset_index()
        
        # Adiciona coluna de valor médio
        df_agrupado['Valor Médio'] = df_agrupado['Receita'] / df_agrupado['Quantidade']
        
        # Ordena por quantidade
        df_agrupado = df_agrupado.sort_values('Quantidade', ascending=False)
        
        # Exibe a tabela
        st.dataframe(df_agrupado, use_container_width=True)
        
        # -------------------- DOWNLOAD DOS DADOS FILTRADOS --------------------
        st.subheader("⬇️ Baixar Dados Filtrados")
        
        # Opções de formato para download
        col1, col2 = st.columns(2)
        
        # Download CSV
        with col1:
            csv = df_filtrado.to_csv(index=False).encode("utf-8-sig")
            st.download_button("Download CSV", csv, "dados_filtrados.csv", "text/csv")
        
        # Download Excel
        with col2:
            # Cria um buffer para o arquivo Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Dados')
                # Ajusta automaticamente a largura das colunas
                for column in df_filtrado.columns:
                    column_width = max(df_filtrado[column].astype(str).map(len).max(), len(column)) + 2
                    col_idx = df_filtrado.columns.get_loc(column)
                    writer.sheets['Dados'].column_dimensions[chr(65 + col_idx)].width = column_width
            
            buffer.seek(0)
            st.download_button("Download Excel", buffer, "dados_filtrados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.warning("O arquivo está vazio ou não pôde ser lido.")
else:
    st.error(f"Arquivo não encontrado: {file_path}")
    st.info("Verifique se o arquivo Excel está no diretório correto.")