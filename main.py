import pandas as pd
from fpdf import FPDF
from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import matplotlib.pyplot as plt
from sqlalchemy import create_engine
from abc import ABC, abstractmethod
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# Configura a conexão usando SQLAlchemy
def conectar_bd():
    conexao_str = 'mysql+mysqlconnector://root@localhost/iAxxMES'
    engine = create_engine(conexao_str)
    return engine

# Dicionário de cores para o status
STATUS_CORES = {
    'Rodando': '008000',  # Verde
    'Parada': 'FF0000',   # Vermelho
    'Setup': 'ADD8E6',    # Azul claro
    'Carga de fio': 'FFA500',  # Laranja
    'Sem programação': '808080'  # Cinza
}

# Classe base para relatórios
class Relatorio(ABC):
    def __init__(self, maquina_id=None, data_inicio=None, data_fim=None):
        self.maquina_id = maquina_id
        self.data_inicio = data_inicio
        self.data_fim = data_fim
        self.dados = None
        self.engine = conectar_bd()

    @abstractmethod
    def obter_dados(self):
        pass

    @abstractmethod
    def gerar_grafico(self, dados, titulo):
        pass

    def gerar_pdf(self, titulo):
        if self.maquina_id is not None:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, titulo, 0, 1, 'C')
            self.gerar_grafico(self.dados, f"Máquina {self.maquina_id}")
            pdf.image(f"{self.__class__.__name__}_chart_maquina_{self.maquina_id}.png", x=10, y=30, w=180)
            pdf.output(f"{self.__class__.__name__}_Relatorio.pdf")

    def gerar_excel(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Relatório Todas Máquinas" if self.maquina_id is None else f"Máquina {self.maquina_id}"
        colunas = list(self.dados.columns)
        sheet.append(colunas)

        for idx, row in self.dados.iterrows():
            linha = [row[col] for col in colunas]
            sheet.append(linha)
            status = row['status'] if 'status' in row else None
            if status in STATUS_CORES:
                cor = STATUS_CORES[status]
                status_cell = sheet.cell(row=sheet.max_row, column=colunas.index('status') + 1)
                status_cell.fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")

            if self.maquina_id is None and (idx < len(self.dados) - 1) and row['maquina_id'] != self.dados.loc[idx + 1, 'maquina_id']:
                sheet.append([""] * len(colunas))

        nome_arquivo = f"{self.__class__.__name__}_Relatorio.xlsx"
        workbook.save(nome_arquivo)

    def gerar_word(self):
        doc = Document()
        titulo = f'Relatório de {self.__class__.__name__}' + (
            f' - Máquina {self.maquina_id}' if self.maquina_id else ' - Todas as Máquinas')
        doc.add_heading(titulo, 0)

        if self.maquina_id is None:
            maquinas = self.dados['maquina_id'].unique()
            for maquina in maquinas:
                dados_maquina = self.dados[self.dados['maquina_id'] == maquina]
                doc.add_paragraph(f"Máquina {maquina}", style="Heading 1")
                table = doc.add_table(rows=1, cols=len(dados_maquina.columns))
                table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                for idx, col_name in enumerate(dados_maquina.columns):
                    hdr_cells[idx].text = col_name

                for _, row in dados_maquina.iterrows():
                    row_cells = table.add_row().cells
                    for idx, value in enumerate(row):
                        cell = row_cells[idx]
                        cell.text = str(value)
                        if dados_maquina.columns[idx] == 'status' and value in STATUS_CORES:
                            cor_hex = STATUS_CORES[value]
                            cell._element.get_or_add_tcPr().append(
                                parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), cor_hex)))

                doc.add_paragraph()
        else:
            table = doc.add_table(rows=1, cols=len(self.dados.columns))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for idx, col_name in enumerate(self.dados.columns):
                hdr_cells[idx].text = col_name

            for _, row in self.dados.iterrows():
                row_cells = table.add_row().cells
                for idx, value in enumerate(row):
                    cell = row_cells[idx]
                    cell.text = str(value)
                    if self.dados.columns[idx] == 'status' and value in STATUS_CORES:
                        cor_hex = STATUS_CORES[value]
                        cell._element.get_or_add_tcPr().append(
                            parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), cor_hex)))

        nome_arquivo = f"{self.__class__.__name__}_Relatorio.docx"
        doc.save(nome_arquivo)

    def gerar_relatorios(self, tipos_relatorio):
        self.obter_dados()
        titulo = f"Relatório de {self.__class__.__name__}" + (
            f" - Máquina {self.maquina_id}" if self.maquina_id else " - Todas as Máquinas")

        if 'pdf' in tipos_relatorio and self.maquina_id is not None:
            self.gerar_pdf(titulo)

        if 'excel' in tipos_relatorio:
            self.gerar_excel()

        if 'word' in tipos_relatorio:
            self.gerar_word()

# Relatório de RPM
class RelatorioRPM(Relatorio):
    def obter_dados(self):
        base_query = """
            SELECT maquina_id, data_hora, rpm FROM maquina_dados
            WHERE data_hora BETWEEN %(data_inicio)s AND %(data_fim)s
        """
        params = {'data_inicio': self.data_inicio, 'data_fim': self.data_fim}
        if self.maquina_id is not None:
            base_query += " AND maquina_id = %(maquina_id)s"
            params['maquina_id'] = self.maquina_id
        base_query += " ORDER BY maquina_id, data_hora;"
        self.dados = pd.read_sql(base_query, self.engine, params=params)

    def gerar_grafico(self, dados, titulo):
        plt.figure(figsize=(10, 5))
        plt.plot(dados['data_hora'], dados['rpm'], marker='o', linestyle='-', color='b')
        plt.title(titulo)
        plt.xlabel('Data e Hora')
        plt.ylabel('RPM')
        plt.xticks(rotation=45)
        plt.tight_layout()
        nome_arquivo = f"{self.__class__.__name__}_chart_maquina_{self.maquina_id}.png"
        plt.savefig(nome_arquivo)
        plt.close()

# Relatório de Status
class RelatorioStatus(Relatorio):
    def obter_dados(self):
        base_query = """
            SELECT md.maquina_id, md.data_hora, ms.descricao AS status
            FROM maquina_dados md
            JOIN maquina_status ms ON md.status = ms.id
            WHERE md.data_hora BETWEEN %(data_inicio)s AND %(data_fim)s
        """
        params = {'data_inicio': self.data_inicio, 'data_fim': self.data_fim}
        if self.maquina_id is not None:
            base_query += " AND md.maquina_id = %(maquina_id)s"
            params['maquina_id'] = self.maquina_id
        base_query += " ORDER BY md.maquina_id, md.data_hora;"
        self.dados = pd.read_sql(base_query, self.engine, params=params)

    def gerar_grafico(self, dados, titulo):
        plt.figure(figsize=(10, 5))
        status_values, status_labels = pd.factorize(dados['status'])
        plt.plot(dados['data_hora'], status_values, marker='o', linestyle='-', color='g')
        plt.title(titulo)
        plt.xlabel('Data e Hora')
        plt.ylabel('Status')
        plt.xticks(rotation=45)
        plt.yticks(ticks=range(len(status_labels)), labels=status_labels)
        plt.tight_layout()
        nome_arquivo = f"{self.__class__.__name__}_chart_maquina_{self.maquina_id}.png"
        plt.savefig(nome_arquivo)
        plt.close()

# Menu de seleção
def main():
    print("Selecione o tipo de relatório:")
    print("1. Relatório de RPM")
    print("2. Relatório de Status")

    opcao_relatorio = input("Digite o número do tipo de relatório: ")

    print("\nSelecione uma opção para o relatório:")
    print("1. Relatório para uma máquina específica")
    print("2. Relatório para todas as máquinas")

    opcao_maquina = input("Digite o número da opção desejada: ")
    maquina_id = None

    if opcao_maquina == "1":
        maquina_id = int(input("\nDigite o ID da máquina: "))

    data_inicio = input("\nDigite a data e hora de início (AAAA-MM-DD HH:MM:SS): ")
    data_fim = input("Digite a data e hora de término (AAAA-MM-DD HH:MM:SS): ")

    print("\nSelecione os tipos de relatório que deseja gerar (separados por vírgula):")
    print("Opções: pdf, excel, word")
    tipos_relatorio = input("Digite as opções: ").replace(" ", "").split(",")

    if opcao_relatorio == "1":
        relatorio = RelatorioRPM(maquina_id, data_inicio, data_fim)
    elif opcao_relatorio == "2":
        relatorio = RelatorioStatus(maquina_id, data_inicio, data_fim)
    else:
        print("\nOpção de relatório inválida.")
        return

    relatorio.gerar_relatorios(tipos_relatorio)
    print("\nRelatórios gerados com sucesso.")

if __name__ == "__main__":
    main()