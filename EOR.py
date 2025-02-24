import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Cargar el archivo Excel principal
file_path = "Screening EOR  para programar.xlsm"
xls = pd.ExcelFile(file_path)

def load_data():
    """Carga los datos del archivo Excel y extrae los puntajes de la hoja Ui."""
    df_ui = xls.parse("Ui")
    df_ui = df_ui.iloc[54:63, [19, 20, 21]]
    df_ui.columns = ["Método EOR", "Validación", "Puntaje Final"]
    df_ui.dropna(inplace=True)
    df_ui.reset_index(drop=True, inplace=True)
    df_ui["Puntaje Final"] = pd.to_numeric(df_ui["Puntaje Final"], errors='coerce').round(1)
    df_ui["Validación"] = pd.to_numeric(df_ui["Validación"], errors='coerce')
    return df_ui

def load_criteria_table():
    """Carga los datos de la hoja Tabla_Puntaje y filtra criterios que no cumplen."""
    df_puntaje = xls.parse("Tabla_Puntaje")
    df_puntaje = df_puntaje.iloc[:, [1, 2, 9, 11]]
    df_puntaje.columns = ["Método EOR", "Criterio", "Valor", "Cumple"]
    df_puntaje["Método EOR"].fillna(method='ffill', inplace=True)
    df_puntaje = df_puntaje.dropna(subset=["Criterio"])
    df_puntaje = df_puntaje[df_puntaje["Cumple"] == 0]
    df_puntaje["Cumple"] = "NO CUMPLE"
    return df_puntaje

def highlight_rows(row):
    """Resalta en rojo las filas donde la columna 'Validación' sea 0."""
    return ['background-color: red' if row["Validación"] == 0 else '' for _ in row]

# Mostrar el nombre del autor en la parte superior izquierda
st.markdown("<h6 style='text-align: left; color: gray;'>Autor: MSc. Wladimir Chávez</h6>", unsafe_allow_html=True)

st.title("Evaluación de Métodos EOR")

df_ui = load_data()
df_criteria = load_criteria_table()

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["🏆 Puntajes", "⚠️ Criterios Incumplidos", "📊 Gráfico de Barras", "📈 Gráfico Radar", "ℹ️ Información del Autor"]
)

with tab1:
    st.subheader("Puntajes de Métodos EOR")
    st.dataframe(df_ui.style.apply(highlight_rows, axis=1).format({"Puntaje Final": "{:.1f}"}))

with tab2:
    st.subheader("Criterios que No Cumplen")
    st.dataframe(df_criteria)

with tab3:
    st.subheader("Gráfico de Barras - Puntajes por Método EOR")
    fig = px.bar(df_ui, x="Método EOR", y="Puntaje Final", color="Método EOR",
                 title="Puntajes Finales por Método EOR",
                 text=df_ui["Puntaje Final"].astype(str),
                 labels={"Puntaje Final": "Puntaje"},
                 template="plotly_white")
    fig.update_layout(xaxis_title="Método EOR", yaxis_title="Puntaje Final",
                      font=dict(family="Arial", size=12),
                      plot_bgcolor="rgba(0,0,0,0)", 
                      margin=dict(l=40, r=40, t=40, b=40))
    st.plotly_chart(fig)

with tab4:
    st.subheader("Gráfico Radar - Comparación de Métodos")
    categories = df_ui["Método EOR"].tolist()
    values = df_ui["Puntaje Final"].tolist()
    fig_radar = go.Figure()
    fig_radar.add_trace(go.Scatterpolar(
        r=values + [values[0]],
        theta=categories + [categories[0]],
        fill='toself',
        name='Puntajes EOR',
        marker=dict(color='blue', opacity=0.6)
    ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 10])),
        showlegend=False,
        template="plotly_white",
        margin=dict(l=40, r=40, t=40, b=40)
    )
    st.plotly_chart(fig_radar)

with tab5:
    st.subheader("Información del Autor")
    st.write("👨‍🏫 **MSc. Wladimir Chávez**")
    st.write("📧 **Correo Electrónico:**")
    st.write("- [wladimir.chavez@epn.edu.ec](mailto:wladimir.chavez@epn.edu.ec)")
    st.write("- [wladpato@hotmail.com](mailto:wladpato@hotmail.com)")
    st.write("📱 **Celular:** 0984458153")

    # Cargar y descargar automáticamente el archivo "DATOS EOR.xlsx"
    datos_eor_path = "DATOS EOR.xlsx"
    if os.path.exists(datos_eor_path):
        st.success("📄 Archivo 'DATOS EOR.xlsx' cargado automáticamente.")
        with open(datos_eor_path, "rb") as file:
            st.download_button(
                label="📥 Descargar Archivo DATOS EOR.xlsx",
                data=file,
                file_name="DATOS EOR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("⚠️ No se encontró el archivo 'DATOS EOR.xlsx' en la carpeta del script.")
