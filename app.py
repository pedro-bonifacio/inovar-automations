import streamlit as st
import pandas as pd
import numpy as np

# --- Configuração da Página ---
st.set_page_config(
    page_title="Análise de Notas Escolares",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- Funções de Processamento ---

def process_student_grades(file_obj):
    """
    Lê um ficheiro Excel, deteta automaticamente o número de alunos
    e converte-o num DataFrame estruturado.
    """
    # Carregar o ficheiro sem cabeçalho
    df = pd.read_excel(file_obj, header=None)

    # Definir índices das linhas
    SUBJECT_ROW_IDX = 11
    CF_ROW_IDX = 12
    DATA_START_IDX = 13

    # Extrair linhas de cabeçalho
    subject_row = df.iloc[SUBJECT_ROW_IDX]
    cf_row = df.iloc[CF_ROW_IDX]

    # 1. Mapear colunas para disciplinas
    col_mapping = {}
    current_subject = None

    for col_idx in range(df.shape[1]):
        val = subject_row.iloc[col_idx]
        if pd.notna(val):
            current_subject = str(val).strip()
        
        sub_header = cf_row.iloc[col_idx]
        if sub_header == 'CF' and current_subject:
            col_mapping[current_subject] = col_idx

    # 2. Definir o mapeamento de notas
    grade_map = {
        "Insuficiente": 2,
        "Suficiente": 3,
        "Bom": 4,
        "Muito Bom": 5,
        "--": np.nan
    }

    # 3. Extrair Dados dos Alunos (Automático)
    students_data = []
    
    # Selecionar todas as linhas a partir do início dos dados
    raw_data_rows = df.iloc[DATA_START_IDX:]

    for _, row in raw_data_rows.iterrows():
        # Verificação de Paragem: Se a coluna do ID (índice 0) estiver vazia, assumimos o fim da lista
        student_id = row.iloc[0]
        student_name = row.iloc[2]

        if pd.isna(student_id):
            break
            
        # Opcional: Se o ID for texto (ex: "Média"), também paramos
        # Isto previne ler rodapés de estatísticas que possam existir no Excel
        if isinstance(student_id, str) and not student_id.isdigit():
             # Verifica se é apenas um ID alfanumérico válido ou lixo
             # Para segurança, se não tiver nome associado, paramos
             if pd.isna(student_name):
                 break

        student_dict = {}
        student_dict['ID'] = student_id
        student_dict['Name'] = student_name
        
        # Extrair notas
        for subject, col_idx in col_mapping.items():
            raw_grade = row.iloc[col_idx]
            
            if isinstance(raw_grade, str):
                raw_grade = raw_grade.strip()
            
            if raw_grade in grade_map:
                mapped_grade = grade_map[raw_grade]
            else:
                try:
                    mapped_grade = float(raw_grade)
                except (ValueError, TypeError):
                    mapped_grade = np.nan
            
            student_dict[subject] = mapped_grade
        
        students_data.append(student_dict)

    # 4. Criar DataFrame Final
    result_df = pd.DataFrame(students_data)
    
    if 'ID' in result_df.columns:
        result_df.set_index('ID', inplace=True)
    
    if not result_df.empty:
        cols = ['Name'] + [c for c in result_df.columns if c != 'Name']
        result_df = result_df[cols]
    
    return result_df

def calculate_class_statistics(df):
    """
    Calcula estatísticas específicas da turma.
    """
    if df.empty:
        return {"Erro": "Não foram encontrados dados de alunos."}

    subject_cols = [col for col in df.columns if col != 'Name']
    grades_df = df[subject_cols]

    total_students = len(df)
    
    not_evaluated_mask = grades_df.isna().all(axis=1)
    num_not_evaluated = not_evaluated_mask.sum()
    
    evaluated_df = grades_df[~not_evaluated_mask]
    
    negatives_per_student = (evaluated_df < 3).sum(axis=1)
    num_no_negatives = (negatives_per_student == 0).sum()
    num_3_plus_negatives = (negatives_per_student >= 3).sum()
    num_any_negative = (negatives_per_student >= 1).sum()
    
    port_col = 'PORT.'
    mat_col = 'Mat'
    
    if port_col in evaluated_df.columns and mat_col in evaluated_df.columns:
        port_negative = evaluated_df[port_col] < 3
        mat_negative = evaluated_df[mat_col] < 3
        num_port_mat_negative = (port_negative & mat_negative).sum()
    else:
        num_port_mat_negative = "N/A (Colunas não encontradas)"

    mb_per_student = (evaluated_df == 5).sum(axis=1)
    num_3_plus_mb = (mb_per_student >= 3).sum()
    num_any_mb = (mb_per_student >= 1).sum()

    insuficiente_counts = (evaluated_df == 2).sum().sort_values(ascending=False)
    top_3_insuficiente = insuficiente_counts.head(3).index.tolist()
    
    amplitude = evaluated_df.max() - evaluated_df.min()
    if not amplitude.empty:
        max_amp_value = amplitude.max()
        max_dispersion_subjects = amplitude[amplitude == max_amp_value].index.tolist()
        dispersion_str = ", ".join(max_dispersion_subjects) + f" (Amplitude: {max_amp_value})"
    else:
        dispersion_str = "N/A"

    stats = {
        "N.º de alunos da Turma": total_students,
        "N.º de alunos não avaliados": num_not_evaluated,
        "N.º de alunos SEM menções inferiores a Suficiente": num_no_negatives,
        "N.º de alunos com TRÊS OU MAIS menções inferiores a Suficiente": num_3_plus_negatives,
        "N.º TOTAL de alunos com menções inferiores a Suficiente": num_any_negative,
        "N.º de alunos com menções inferiores a Suficiente cumulativamente a PORTUGUÊS e MATEMÁTICA": num_port_mat_negative,
        "N.º de alunos com TRÊS OU MAIS menções de MUITO BOM": num_3_plus_mb,
        "N.º TOTAL de alunos com menções de MUITO BOM": num_any_mb,
        "As 3 disciplinas com maior número de menções de INSUFICIENTE": ", ".join(top_3_insuficiente),
        "As disciplinas com MAIOR DISPERSÃO/AMPLITUDE DE RESULTADOS": dispersion_str,
    }
    
    return stats

# --- Interface Streamlit ---

st.title("Gerador de Estatísticas Escolares")
st.markdown("**Made by Pedro Bonifácio**")
st.markdown("---")

col_instrucoes, col_upload = st.columns([2, 1], gap="large")

with col_instrucoes:
    st.subheader("Instruções Passo-a-Passo")
    st.write("Siga as imagens abaixo para preparar e carregar o seu ficheiro corretamente.")
    
    # Placeholders para as capturas de ecrã
    st.info('Passo 1: Entrar no inovar, selecionar a direção de turma no canto superior direito e depois selecionar "Área Docente" -> "Intercalares" -> "Sínteses Globais" -> "EB031a" (IMPORTANTE)')
    st.image("static/image1.webp")
    
    st.info("Passo 2: Selecionar exatamente estas opções de configuração.")
    st.image("static/image2.webp")
    
    st.info("Passo 3: IMPORTANTE - Abrir o ficheiro excel e remover manualmente os alunos transferidos. Carregar com o botão do lado direito no número da linha -> 'Eliminar'. Depois é só gravar e fazer upload do ficheiro neste website.")
    st.image("static/image3.webp")

with col_upload:
    st.subheader("Carregar Ficheiro e Processar")
    
    with st.container(border=True):
        st.write("Carregue o ficheiro Excel (.xls ou .xlsx). O número de alunos será calculado automaticamente.")
        
        # Botão de upload (Input numérico removido)
        uploaded_file = st.file_uploader("Escolha o ficheiro Excel", type=['xls', 'xlsx'])

        if uploaded_file is not None:
            st.divider()
            with st.spinner('A processar o ficheiro...'):
                try:
                    # Processar o ficheiro carregado (sem argumento num_students)
                    df_result = process_student_grades(uploaded_file)
                    
                    # Calcular estatísticas
                    stats = calculate_class_statistics(df_result)
                    
                    st.success(f"Ficheiro processado com sucesso! {stats.get('N.º de alunos da Turma', 0)} alunos detetados.")
                    
                    st.subheader("Resultados Estatísticos")
                    
                    for key, value in stats.items():
                        st.markdown(f"**{key}:**")
                        st.markdown(f"> {value}")
                    
                    st.divider()
                    
                    with st.expander("Ver Tabela de Notas Processada"):
                        st.dataframe(df_result)
                        
                except Exception as e:
                    st.error(f"Ocorreu um erro ao processar o ficheiro: {e}")
                    st.warning("Verifique se o ficheiro Excel segue a estrutura correta (linhas 11, 12 e 13).")