"""
Script Python para processar dados financeiros brutos e gerar Tabela Dinâmica no Excel.

Este script:
1. Carrega dados brutos de um arquivo CSV
2. Realiza tratamento de valores nulos
3. Normaliza os dados
4. Corrige tipos de dados
5. Cria uma Tabela Dinâmica formatada no Excel
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import warnings
import locale

warnings.filterwarnings('ignore')

# Tenta configurar locale para português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR')
    except locale.Error:
        # Se não conseguir, mantém o locale padrão (nomes em inglês)
        pass


class ProcessadorFinanceiro:
    """Classe para processar dados financeiros e gerar relatórios."""
    
    def __init__(self, arquivo_entrada):
        """
        Inicializa o processador com o arquivo de entrada.
        
        Args:
            arquivo_entrada (str): Caminho para o arquivo CSV com dados brutos
        """
        self.arquivo_entrada = arquivo_entrada
        self.df = None
        self.df_processado = None
        
    def carregar_dados(self):
        """Carrega os dados brutos do arquivo CSV."""
        print("📂 Carregando dados brutos...")
        self.df = pd.read_csv(self.arquivo_entrada)
        print(f"✅ {len(self.df)} registros carregados")
        print(f"📊 Colunas: {', '.join(self.df.columns.tolist())}")
        return self
    
    def tratar_nulos(self):
        """Trata valores nulos nos dados."""
        print("\n🔧 Tratando valores nulos...")
        
        # Contabiliza nulos antes do tratamento
        nulos_antes = self.df.isnull().sum()
        print("Valores nulos encontrados:")
        for col, count in nulos_antes[nulos_antes > 0].items():
            print(f"  - {col}: {count}")
        
        # Cria cópia para processamento
        self.df_processado = self.df.copy()
        
        # Tratamento específico por coluna
        if 'Categoria' in self.df_processado.columns:
            # Preenche categorias vazias com "Não Classificado"
            self.df_processado['Categoria'] = self.df_processado['Categoria'].fillna('Não Classificado')
        
        if 'Valor' in self.df_processado.columns:
            # Remove linhas onde Valor é nulo (transações inválidas)
            linhas_antes = len(self.df_processado)
            self.df_processado = self.df_processado.dropna(subset=['Valor'])
            linhas_removidas = linhas_antes - len(self.df_processado)
            if linhas_removidas > 0:
                print(f"  ⚠️  {linhas_removidas} linhas removidas por Valor nulo")
        
        if 'Status' in self.df_processado.columns:
            # Preenche status vazios com "Não Informado"
            self.df_processado['Status'] = self.df_processado['Status'].fillna('Não Informado')
        
        if 'Moeda' in self.df_processado.columns:
            # Preenche moeda vazia com a mais comum
            moeda_padrao = self.df_processado['Moeda'].mode()[0] if not self.df_processado['Moeda'].mode().empty else 'BRL'
            self.df_processado['Moeda'] = self.df_processado['Moeda'].fillna(moeda_padrao)
        
        print("✅ Valores nulos tratados")
        return self
    
    def normalizar_dados(self):
        """Normaliza os dados (padronização de strings, formatos, etc)."""
        print("\n🔄 Normalizando dados...")
        
        # Normaliza strings: remove espaços extras e padroniza capitalização
        colunas_texto = self.df_processado.select_dtypes(include=['object']).columns
        for col in colunas_texto:
            if col != 'Data':  # Não normalizar a coluna de data
                # Garante que a coluna é string antes de aplicar métodos de string
                self.df_processado[col] = self.df_processado[col].astype(str)
                self.df_processado[col] = self.df_processado[col].str.strip()
                self.df_processado[col] = self.df_processado[col].str.title()
        
        # Normaliza valores numéricos (garante formato consistente)
        if 'Valor' in self.df_processado.columns:
            self.df_processado['Valor'] = self.df_processado['Valor'].round(2)
        
        print("✅ Dados normalizados")
        return self
    
    def corrigir_tipos(self):
        """Corrige os tipos de dados das colunas."""
        print("\n🔢 Corrigindo tipos de dados...")
        
        # Converte coluna Data para datetime
        if 'Data' in self.df_processado.columns:
            self.df_processado['Data'] = pd.to_datetime(self.df_processado['Data'])
            print(f"  - Data: convertido para datetime")
        
        # Garante que Valor é float
        if 'Valor' in self.df_processado.columns:
            self.df_processado['Valor'] = pd.to_numeric(self.df_processado['Valor'], errors='coerce')
            print(f"  - Valor: convertido para numérico")
        
        # Converte colunas de texto para categoria (economiza memória)
        colunas_categoricas = ['Categoria', 'Moeda', 'Status']
        for col in colunas_categoricas:
            if col in self.df_processado.columns:
                self.df_processado[col] = self.df_processado[col].astype('category')
                print(f"  - {col}: convertido para categoria")
        
        print("✅ Tipos de dados corrigidos")
        return self
    
    def adicionar_colunas_derivadas(self):
        """
        Adiciona colunas derivadas úteis para análise.
        
        Nota: Os nomes dos meses e dias da semana serão em português se o 
        sistema tiver locale pt_BR configurado, caso contrário serão em inglês.
        """
        print("\n➕ Adicionando colunas derivadas...")
        
        if 'Data' in self.df_processado.columns:
            self.df_processado['Ano'] = self.df_processado['Data'].dt.year
            self.df_processado['Mês'] = self.df_processado['Data'].dt.month
            self.df_processado['Mês_Nome'] = self.df_processado['Data'].dt.strftime('%B')
            self.df_processado['Dia_Semana'] = self.df_processado['Data'].dt.day_name()
            print("  - Colunas de tempo adicionadas: Ano, Mês, Mês_Nome, Dia_Semana")
        
        print("✅ Colunas derivadas adicionadas")
        return self
    
    def criar_pivot_table(self, arquivo_saida='relatorio_financeiro.xlsx'):
        """
        Cria uma Tabela Dinâmica formatada no Excel.
        
        Args:
            arquivo_saida (str): Nome do arquivo Excel de saída
        """
        print(f"\n📊 Criando Tabela Dinâmica...")
        
        # Cria tabela dinâmica: Categoria x Mês com soma de valores
        pivot = pd.pivot_table(
            self.df_processado,
            values='Valor',
            index='Categoria',
            columns='Mês',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Total'
        )
        
        # Cria segunda pivot: Categoria x Status
        pivot_status = pd.pivot_table(
            self.df_processado,
            values='Valor',
            index='Categoria',
            columns='Status',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Total'
        )
        
        # Salva no Excel
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            # Escreve dados processados
            self.df_processado.to_excel(writer, sheet_name='Dados Processados', index=False)
            
            # Escreve pivot por mês
            pivot.to_excel(writer, sheet_name='Pivot por Mês')
            
            # Escreve pivot por status
            pivot_status.to_excel(writer, sheet_name='Pivot por Status')
            
            # Escreve resumo estatístico
            resumo = self.df_processado.groupby('Categoria')['Valor'].agg([
                ('Total', 'sum'),
                ('Média', 'mean'),
                ('Mínimo', 'min'),
                ('Máximo', 'max'),
                ('Contagem', 'count')
            ]).round(2)
            resumo.to_excel(writer, sheet_name='Resumo Estatístico')
        
        print(f"✅ Arquivo Excel criado: {arquivo_saida}")
        
        # Formata o arquivo Excel
        self._formatar_excel(arquivo_saida)
        
        return arquivo_saida
    
    def _formatar_excel(self, arquivo):
        """Aplica formatação ao arquivo Excel."""
        print("🎨 Aplicando formatação...")
        
        wb = load_workbook(arquivo)
        
        # Cores para formatação
        cor_header = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cor_total = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        fonte_header = Font(bold=True, color="FFFFFF", size=11)
        fonte_total = Font(bold=True, size=11)
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        borda = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Formata cada planilha
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Formata cabeçalho
            for cell in ws[1]:
                cell.fill = cor_header
                cell.font = fonte_header
                cell.alignment = alinhamento_centro
                cell.border = borda
            
            # Ajusta largura das colunas
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except (TypeError, AttributeError):
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # Formata linha de totais (se existir)
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = borda
                    
                    # Destaca linha "Total"
                    if cell.value == 'Total':
                        for c in row:
                            c.fill = cor_total
                            c.font = fonte_total
            
            # Formata números como moeda nas planilhas de pivot
            if 'Pivot' in sheet_name or 'Resumo' in sheet_name:
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = 'R$ #,##0.00'
            
            # Congela painéis
            ws.freeze_panes = 'B2'
        
        wb.save(arquivo)
        print("✅ Formatação aplicada")
    
    def gerar_relatorio(self, arquivo_saida='relatorio_financeiro.xlsx'):
        """
        Executa todo o pipeline de processamento e gera o relatório.
        
        Args:
            arquivo_saida (str): Nome do arquivo Excel de saída
            
        Returns:
            str: Caminho do arquivo gerado
        """
        print("=" * 60)
        print("🚀 INICIANDO PROCESSAMENTO DE DADOS FINANCEIROS")
        print("=" * 60)
        
        self.carregar_dados()
        self.tratar_nulos()
        self.normalizar_dados()
        self.corrigir_tipos()
        self.adicionar_colunas_derivadas()
        arquivo = self.criar_pivot_table(arquivo_saida)
        
        print("\n" + "=" * 60)
        print("✨ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
        print("=" * 60)
        print(f"\n📄 Relatório gerado: {arquivo}")
        print(f"📊 Total de registros processados: {len(self.df_processado)}")
        print(f"💰 Valor total: R$ {self.df_processado['Valor'].sum():,.2f}")
        
        return arquivo


def main():
    """Função principal."""
    # Arquivo de entrada com dados brutos
    arquivo_entrada = 'dados_financeiros_brutos.csv'
    
    # Arquivo de saída
    arquivo_saida = 'relatorio_financeiro.xlsx'
    
    try:
        # Cria processador e gera relatório
        processador = ProcessadorFinanceiro(arquivo_entrada)
        processador.gerar_relatorio(arquivo_saida)
        
    except FileNotFoundError:
        print(f"❌ Erro: Arquivo '{arquivo_entrada}' não encontrado!")
        print("Por favor, certifique-se de que o arquivo existe no diretório atual.")
    except Exception as e:
        print(f"❌ Erro durante o processamento: {str(e)}")
        raise


if __name__ == '__main__':
    main()
