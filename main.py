import pandas as pd
from fpdf import FPDF
from docx import Document
import matplotlib.pyplot as plt
from sqlalchemy import create_engine

# Configura a conexão usando SQLAlchemy
def conectar_bd():
    conexao_str = 'mysql+mysqlconnector://root@localhost/iAxxMES'
    engine = create_engine(conexao_str)
    return engine

# Consulta de dados do RPM de uma máquina específica ao longo do tempo
def obter_dados_rpm(maquina_id):
    engine = conectar_bd()
    query = f"""
    SELECT data_hora, rpm
    FROM maquina_dados
    WHERE maquina_id = {maquina_id}
    ORDER BY data_hora;
    """
    dados = pd.read_sql(query, engine)
    return dados

# Função para gerar relatório em PDF
def gerar_relatorio_pdf(dados, maquina_id):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f'Relatório de RPM - Máquina {maquina_id}', 0, 1, 'C')

    plt.figure(figsize=(10, 5))
    plt.plot(dados['data_hora'], dados['rpm'], marker='o', linestyle='-', color='b')
    plt.title(f'RPM da Máquina {maquina_id} ao longo do tempo')
    plt.xlabel('Data e Hora')
    plt.ylabel('RPM')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("rpm_chart.png")
    plt.close()

    pdf.image("rpm_chart.png", x=10, y=30, w=180)
    pdf.output(f"Relatorio_RPM_Maquina_{maquina_id}.pdf")

# Função para gerar relatório em Excel
def gerar_relatorio_excel(dados, maquina_id):
    dados.to_excel(f"Relatorio_RPM_Maquina_{maquina_id}.xlsx", index=False)

# Função para gerar relatório em Word
def gerar_relatorio_word(dados, maquina_id):
    doc = Document()
    doc.add_heading(f'Relatório de RPM - Máquina {maquina_id}', 0)
    doc.add_paragraph(f'Relatório de RPM ao longo do tempo para a máquina {maquina_id}.')

    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Data e Hora'
    hdr_cells[1].text = 'RPM'

    for index, row in dados.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['data_hora'])
        row_cells[1].text = str(row['rpm'])

    doc.save(f"Relatorio_RPM_Maquina_{maquina_id}.docx")

# Exemplo de execução
maquina_id = 1
dados_rpm = obter_dados_rpm(maquina_id)
gerar_relatorio_pdf(dados_rpm, maquina_id)
gerar_relatorio_excel(dados_rpm, maquina_id)
gerar_relatorio_word(dados_rpm, maquina_id)
