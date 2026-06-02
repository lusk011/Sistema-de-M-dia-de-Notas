from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Sistema de Avaliação Escolar",
    page_icon="📚",
    layout="wide",
)


def iniciar_estado():
    valores_iniciais = {
        "etapa": "configuracao",
        "sala": "",
        "total_alunos": 1,
        "alunos": [],
        "indice_atual": 0,
    }

    for chave, valor in valores_iniciais.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def calcular_situacao(media):
    if media >= 7:
        return "Aprovado"
    if media >= 5:
        return "Recuperação"
    return "Reprovado"


def calcular_observacao(situacao):
    if situacao == "Aprovado":
        return "Parabens!"
    if situacao == "Recuperação":
        return "Precisa fazer recuperação."
    return "Média insuficiente."


def montar_dataframe(alunos):
    dados = []

    for numero, aluno in enumerate(alunos, start=1):
        dados.append(
            {
                "N": numero,
                "Nome": aluno["nome"],
                "Nota 1": aluno["nota1"],
                "Nota 2": aluno["nota2"],
                "Nota 3": aluno["nota3"],
                "Média": aluno["media"],
                "Situação": aluno["situacao"],
                "Observação": calcular_observacao(aluno["situacao"]),
            }
        )

    return pd.DataFrame(dados)


def montar_dataframe_edicao(alunos):
    dados = []

    for numero, aluno in enumerate(alunos, start=1):
        dados.append(
            {
                "N": numero,
                "Nome": aluno["nome"],
                "Nota 1": aluno["nota1"],
                "Nota 2": aluno["nota2"],
                "Nota 3": aluno["nota3"],
            }
        )

    return pd.DataFrame(dados)


def atualizar_alunos_por_dataframe(df):
    alunos_corrigidos = []

    for _, linha in df.iterrows():
        nome = str(linha["Nome"]).strip()
        nota1 = float(linha["Nota 1"])
        nota2 = float(linha["Nota 2"])
        nota3 = float(linha["Nota 3"])
        media = round((nota1 + nota2 + nota3) / 3, 2)

        alunos_corrigidos.append(
            {
                "nome": nome,
                "nota1": nota1,
                "nota2": nota2,
                "nota3": nota3,
                "media": media,
                "situacao": calcular_situacao(media),
            }
        )

    return alunos_corrigidos


def gerar_planilha_excel(sala, alunos):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"

    azul = "1F4E78"
    azul_claro = "DDEBF7"
    verde = "E2F0D9"
    amarelo = "FFF2CC"
    vermelho = "FCE4D6"
    cinza = "F2F2F2"

    borda = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    centro = Alignment(horizontal="center", vertical="center")
    esquerda = Alignment(horizontal="left", vertical="center")

    def formatar_intervalo(linha_inicio, coluna_inicio, linha_fim, coluna_fim, preenchimento, alinhamento):
        for linha in range(linha_inicio, linha_fim + 1):
            for coluna in range(coluna_inicio, coluna_fim + 1):
                celula = ws.cell(row=linha, column=coluna)
                celula.fill = PatternFill("solid", fgColor=preenchimento)
                celula.border = borda
                celula.alignment = alinhamento

    ws.merge_cells("A1:H1")
    ws["A1"] = f"Relatório de Avaliações - Sala: {sala}"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = PatternFill("solid", fgColor=azul)
    ws["A1"].alignment = centro
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(italic=True, color="666666")
    ws["A2"].fill = PatternFill("solid", fgColor=cinza)
    ws["A2"].alignment = centro

    cabecalhos = [
        "N",
        "Nome",
        "Nota 1",
        "Nota 2",
        "Nota 3",
        "Média",
        "Situação",
        "Observação",
    ]

    for coluna, titulo in enumerate(cabecalhos, start=1):
        celula = ws.cell(row=3, column=coluna, value=titulo)
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor=azul)
        celula.alignment = centro
        celula.border = borda

    cores_situacao = {
        "Aprovado": verde,
        "Recuperação": amarelo,
        "Reprovado": vermelho,
    }

    for linha, aluno in enumerate(alunos, start=4):
        cor = cores_situacao.get(aluno["situacao"], "FFFFFF")
        valores = [
            linha - 3,
            aluno["nome"],
            aluno["nota1"],
            aluno["nota2"],
            aluno["nota3"],
            aluno["media"],
            aluno["situacao"],
            calcular_observacao(aluno["situacao"]),
        ]

        for coluna, valor in enumerate(valores, start=1):
            celula = ws.cell(row=linha, column=coluna, value=valor)
            celula.fill = PatternFill("solid", fgColor=cor)
            celula.border = borda
            celula.alignment = esquerda if coluna in [2, 8] else centro

            if coluna in [3, 4, 5, 6]:
                celula.number_format = "0.00"

    ultima_linha = len(alunos) + 3
    resumo_linha = ultima_linha + 2

    ws.merge_cells(start_row=resumo_linha, start_column=1, end_row=resumo_linha, end_column=8)
    ws.cell(row=resumo_linha, column=1, value="Resumo da Turma")
    ws.cell(row=resumo_linha, column=1).font = Font(bold=True, color="FFFFFF")
    formatar_intervalo(resumo_linha, 1, resumo_linha, 8, azul, centro)

    aprovados = sum(1 for aluno in alunos if aluno["situacao"] == "Aprovado")
    recuperacao = sum(1 for aluno in alunos if aluno["situacao"] == "Recuperação")
    reprovados = sum(1 for aluno in alunos if aluno["situacao"] == "Reprovado")

    resumo = [
        ("Total de alunos", len(alunos)),
        ("Aprovados", aprovados),
        ("Em recuperação", recuperacao),
        ("Reprovados", reprovados),
        ("Média geral", f"=AVERAGE(F4:F{ultima_linha})"),
    ]

    for deslocamento, (rotulo, valor) in enumerate(resumo, start=1):
        linha = resumo_linha + deslocamento
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
        ws.merge_cells(start_row=linha, start_column=5, end_row=linha, end_column=8)

        celula_rotulo = ws.cell(row=linha, column=1, value=rotulo)
        celula_valor = ws.cell(row=linha, column=5, value=valor)

        celula_rotulo.font = Font(bold=True)
        formatar_intervalo(linha, 1, linha, 4, cinza, esquerda)
        formatar_intervalo(linha, 5, linha, 8, azul_claro, centro)
        celula_valor.number_format = "0.00" if rotulo == "Média geral" else "General"

    larguras = [8, 32, 12, 12, 12, 12, 16, 28]
    for coluna, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(coluna)].width = largura

    arquivo = BytesIO()
    wb.save(arquivo)
    arquivo.seek(0)
    return arquivo


def reiniciar():
    for chave in ["etapa", "sala", "total_alunos", "alunos", "indice_atual"]:
        if chave in st.session_state:
            del st.session_state[chave]
    iniciar_estado()


iniciar_estado()

st.title("Sistema de Avaliação Escolar")
st.caption("Cadastre alunos, registre notas e gere uma planilha Excel automaticamente.")


if st.session_state.etapa == "configuracao":
    st.subheader("Configuração da turma")

    col_sala, col_total = st.columns([2, 1])
    with col_sala:
        sala = st.text_input("Nome ou código da sala", placeholder="Exemplo: 3º Ano A")
    with col_total:
        total_alunos = st.number_input(
            "Quantidade de alunos",
            min_value=1,
            max_value=100,
            value=5,
            step=1,
        )

    if st.button("Iniciar cadastro", type="primary", use_container_width=True):
        if not sala.strip():
            st.error("Informe o nome ou código da sala.")
        else:
            st.session_state.sala = sala.strip()
            st.session_state.total_alunos = int(total_alunos)
            st.session_state.alunos = []
            st.session_state.indice_atual = 0
            st.session_state.etapa = "cadastro"
            st.rerun()


elif st.session_state.etapa == "cadastro":
    total = st.session_state.total_alunos
    indice = st.session_state.indice_atual

    st.subheader(f"Aluno {indice + 1} de {total}")
    st.progress(indice / total)

    with st.form(f"formulario_aluno_{indice}"):
        nome = st.text_input("Nome completo do aluno")

        col1, col2, col3 = st.columns(3)
        with col1:
            nota1 = st.number_input("Nota 1", min_value=0.0, max_value=10.0, step=0.1)
        with col2:
            nota2 = st.number_input("Nota 2", min_value=0.0, max_value=10.0, step=0.1)
        with col3:
            nota3 = st.number_input("Nota 3", min_value=0.0, max_value=10.0, step=0.1)

        salvar = st.form_submit_button("Salvar aluno", type="primary", use_container_width=True)

    if salvar:
        if not nome.strip():
            st.error("Informe o nome do aluno.")
        else:
            media = round((nota1 + nota2 + nota3) / 3, 2)
            st.session_state.alunos.append(
                {
                    "nome": nome.strip(),
                    "nota1": nota1,
                    "nota2": nota2,
                    "nota3": nota3,
                    "media": media,
                    "situacao": calcular_situacao(media),
                }
            )

            st.session_state.indice_atual += 1
            if st.session_state.indice_atual >= total:
                st.session_state.etapa = "resultado"
            st.rerun()

    if st.session_state.alunos:
        st.divider()
        st.subheader("Alunos já cadastrados")
        st.dataframe(montar_dataframe(st.session_state.alunos), use_container_width=True)

    st.button("Recomeçar", on_click=reiniciar, use_container_width=True)


elif st.session_state.etapa == "resultado":
    sala = st.session_state.sala
    alunos = st.session_state.alunos
    df = montar_dataframe(alunos)

    aprovados = int((df["Situação"] == "Aprovado").sum())
    recuperacao = int((df["Situação"] == "Recuperação").sum())
    reprovados = int((df["Situação"] == "Reprovado").sum())
    media_geral = float(df["Média"].mean())

    st.subheader(f"Resultado final - {sala}")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total", len(alunos))
    col2.metric("Aprovados", aprovados)
    col3.metric("Recuperação", recuperacao)
    col4.metric("Reprovados", reprovados)

    st.metric("Média geral da turma", f"{media_geral:.2f}")

    st.divider()
    st.subheader("Corrigir dados antes de gerar a planilha")
    st.caption("Edite nomes ou notas na tabela abaixo e clique em Salvar correções.")

    df_edicao = montar_dataframe_edicao(alunos)
    df_corrigido = st.data_editor(
        df_edicao,
        use_container_width=True,
        hide_index=True,
        disabled=["N"],
        column_config={
            "Nota 1": st.column_config.NumberColumn("Nota 1", min_value=0.0, max_value=10.0, step=0.1),
            "Nota 2": st.column_config.NumberColumn("Nota 2", min_value=0.0, max_value=10.0, step=0.1),
            "Nota 3": st.column_config.NumberColumn("Nota 3", min_value=0.0, max_value=10.0, step=0.1),
        },
    )

    if st.button("Salvar correções", type="secondary", use_container_width=True):
        if df_corrigido["Nome"].astype(str).str.strip().eq("").any():
            st.error("Todos os alunos precisam ter nome.")
        else:
            st.session_state.alunos = atualizar_alunos_por_dataframe(df_corrigido)
            st.success("Correções salvas. A planilha será gerada com os dados atualizados.")
            st.rerun()

    st.divider()
    st.subheader("Tabela de resultados")
    df = montar_dataframe(st.session_state.alunos)
    st.dataframe(df, use_container_width=True)

    arquivo_excel = gerar_planilha_excel(sala, st.session_state.alunos)
    nome_arquivo = f"avaliacao_{sala.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    st.download_button(
        "Gerar e baixar planilha Excel",
        data=arquivo_excel,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.button("Nova avaliação", on_click=reiniciar, use_container_width=True)
