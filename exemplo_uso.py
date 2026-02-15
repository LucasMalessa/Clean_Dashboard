"""
Exemplo de uso customizado do ProcessadorFinanceiro.

Este script demonstra como usar a classe ProcessadorFinanceiro
de forma programática para ter mais controle sobre o processamento.
"""

from processar_financeiro import ProcessadorFinanceiro

# Exemplo 1: Uso básico
print("=" * 60)
print("EXEMPLO 1: Uso Básico")
print("=" * 60)

processador = ProcessadorFinanceiro('dados_financeiros_brutos.csv')
arquivo = processador.gerar_relatorio('exemplo_basico.xlsx')
print(f"\n✅ Relatório gerado: {arquivo}\n")


# Exemplo 2: Processamento passo a passo
print("=" * 60)
print("EXEMPLO 2: Processamento Passo a Passo")
print("=" * 60)

processador2 = ProcessadorFinanceiro('dados_financeiros_brutos.csv')

# Executar cada etapa individualmente
processador2.carregar_dados()
print(f"\nDados carregados: {len(processador2.df)} registros")

processador2.tratar_nulos()
print(f"Após tratamento: {len(processador2.df_processado)} registros")

processador2.normalizar_dados()
print("Dados normalizados")

processador2.corrigir_tipos()
print("Tipos corrigidos")

processador2.adicionar_colunas_derivadas()
print("Colunas derivadas adicionadas")

# Ver estatísticas antes de gerar o relatório
print("\n📊 Estatísticas dos dados processados:")
print(f"  - Total de registros: {len(processador2.df_processado)}")
print(f"  - Categorias únicas: {processador2.df_processado['Categoria'].nunique()}")
print(f"  - Valor total: R$ {processador2.df_processado['Valor'].sum():,.2f}")
print(f"  - Valor médio: R$ {processador2.df_processado['Valor'].mean():,.2f}")

arquivo2 = processador2.criar_pivot_table('exemplo_passo_a_passo.xlsx')
print(f"\n✅ Relatório gerado: {arquivo2}\n")


# Exemplo 3: Análise customizada dos dados processados
print("=" * 60)
print("EXEMPLO 3: Análise Customizada")
print("=" * 60)

processador3 = ProcessadorFinanceiro('dados_financeiros_brutos.csv')
processador3.carregar_dados()
processador3.tratar_nulos()
processador3.normalizar_dados()
processador3.corrigir_tipos()
processador3.adicionar_colunas_derivadas()

# Acessar os dados processados diretamente
df = processador3.df_processado

print("\n📈 Análise por Categoria:")
analise_categoria = df.groupby('Categoria')['Valor'].agg(['sum', 'mean', 'count'])
print(analise_categoria)

print("\n📅 Análise por Mês:")
analise_mes = df.groupby('Mês')['Valor'].sum()
print(analise_mes)

print("\n✅ Análise concluída!")
print("\nVocê pode usar processador.df_processado para fazer análises customizadas")
print("antes de gerar o relatório Excel.\n")
