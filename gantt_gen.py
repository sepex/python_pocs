import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import load_workbook

def alocar_tarefas(tarefas, data_inicio):
    tarefas_agendadas = []
    data_atual = data_inicio
    tarefas_em_execucao = []

    for _, row in tarefas.iterrows():
        nome = row["Tarefa"]
        duracao = int(row["Prazo"])
        plano_acao = row["Plano de Ação"]

        tarefas_em_execucao = [t for t in tarefas_em_execucao if t["fim"] > data_atual]
        while len(tarefas_em_execucao) >= 3:
            data_atual += timedelta(days=1)
            tarefas_em_execucao = [t for t in tarefas_em_execucao if t["fim"] > data_atual]

        inicio = data_atual
        fim = inicio + timedelta(days=duracao - 1)

        tarefas_agendadas.append({
            "Nome da ação": nome.strip(),
            "Prazo": duracao,
            "Data de início": inicio,
            "Data de término": fim,
            "Plano de Ação": plano_acao.strip()
        })

        tarefas_em_execucao.append({"fim": fim})

    return tarefas_agendadas, fim

def processar_checklist(file):
    trilha1_df = pd.read_excel(file, sheet_name="trilha 1", skiprows=1)
    trilha2_df = pd.read_excel(file, sheet_name="trilha 2", skiprows=1)
    trilha1_df.columns = ["Tarefa", "Prazo", "Descricao", "Plano de Ação", "Atuar", "Responsável"]
    trilha2_df.columns = ["Tarefa", "Prazo", "Descricao", "Plano de Ação", "Atuar", "Responsável"]

    trilha1_df = trilha1_df[trilha1_df["Atuar"].str.upper() == "SIM"]
    trilha2_df = trilha2_df[trilha2_df["Atuar"].str.upper() == "SIM"]

    trilha1_df["Prazo"] = pd.to_numeric(trilha1_df["Prazo"], errors="coerce").fillna(1).astype(int)
    trilha2_df["Prazo"] = pd.to_numeric(trilha2_df["Prazo"], errors="coerce").fillna(1).astype(int)

    return trilha1_df, trilha2_df

def preencher_template(modelo_bytes, dados):
    wb = load_workbook(filename=BytesIO(modelo_bytes))
    planner = wb.worksheets[0]
    detalhes = wb.worksheets[1]

    for i, row in enumerate(dados.itertuples(index=False), start=5):
        planner[f'B{i}'] = row[0]
        planner[f'C{i}'] = row[2]
        planner[f'D{i}'] = row[1]

    for i, row in enumerate(dados.itertuples(index=False), start=2):
        detalhes[f'A{i}'] = row[0]
        detalhes[f'B{i}'] = row[1]
        detalhes[f'C{i}'] = row[2]
        detalhes[f'D{i}'] = row[3]
        detalhes[f'E{i}'] = row[4]

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

st.title("Gerador de Gantt por Trilha")

checklist_file = st.file_uploader("Upload do Checklist", type="xlsx")
gantt_template = st.file_uploader("Upload do Gantt Planner (modelo)", type="xlsx")

if checklist_file and gantt_template:
    trilha1_df, trilha2_df = processar_checklist(checklist_file)
    inicio_trilha1 = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

    ag_trilha1, fim1 = alocar_tarefas(trilha1_df, inicio_trilha1)
    ag_trilha2, _ = alocar_tarefas(trilha2_df, fim1 + timedelta(days=1))

    df1 = pd.DataFrame(ag_trilha1)
    df2 = pd.DataFrame(ag_trilha2)

    gantt_bytes = gantt_template.read()
    out1 = preencher_template(gantt_bytes, df1)
    out2 = preencher_template(gantt_bytes, df2)

    st.download_button("Download Trilha 1", out1, file_name="Gantt_Trilha_1.xlsx")
    st.download_button("Download Trilha 2", out2, file_name="Gantt_Trilha_2.xlsx")
