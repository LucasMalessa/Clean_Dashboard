# Clean_Dashboard

Script Python para processar dados financeiros brutos e gerar relatórios em Excel com Tabelas Dinâmicas formatadas.

## 📋 Descrição

Este projeto contém um script Python completo que:

1. **Carrega dados brutos** de arquivos CSV com informações financeiras
2. **Trata valores nulos** de forma inteligente:
   - Remove transações sem valor
   - Preenche categorias vazias com "Não Classificado"
   - Preenche status vazios com "Não Informado"
   - Preenche moeda vazia com a mais comum
3. **Normaliza os dados**:
   - Padroniza strings (capitalização, espaços extras)
   - Formata valores numéricos
4. **Corrige tipos de dados**:
   - Converte datas para formato datetime
   - Garante valores numéricos em campos monetários
   - Otimiza memória usando tipos categóricos
5. **Adiciona colunas derivadas**:
   - Ano, Mês, Nome do Mês, Dia da Semana
6. **Gera Tabelas Dinâmicas formatadas** no Excel:
   - Pivot por Mês e Categoria
   - Pivot por Status e Categoria
   - Resumo estatístico
   - Formatação profissional com cores e bordas

## 🚀 Instalação

### Pré-requisitos
- Python 3.7 ou superior
- pip (gerenciador de pacotes Python)

### Instalar dependências

```bash
pip install -r requirements.txt
```

As dependências incluem:
- `pandas` - Manipulação de dados
- `openpyxl` - Leitura/escrita de arquivos Excel
- `numpy` - Operações numéricas

## 💻 Uso

### Uso Básico

1. Coloque seu arquivo CSV com dados brutos no mesmo diretório do script
2. O arquivo deve ter as colunas: Data, Categoria, Descrição, Valor, Moeda, Status
3. Execute o script:

```bash
python processar_financeiro.py
```

### Uso Programático

```python
from processar_financeiro import ProcessadorFinanceiro

# Criar processador com arquivo de entrada
processador = ProcessadorFinanceiro('dados_financeiros_brutos.csv')

# Gerar relatório completo
processador.gerar_relatorio('meu_relatorio.xlsx')
```

### Uso Avançado

```python
from processar_financeiro import ProcessadorFinanceiro

# Processar passo a passo
processador = ProcessadorFinanceiro('dados_financeiros_brutos.csv')
processador.carregar_dados()
processador.tratar_nulos()
processador.normalizar_dados()
processador.corrigir_tipos()
processador.adicionar_colunas_derivadas()
processador.criar_pivot_table('relatorio_customizado.xlsx')
```

## 📊 Estrutura dos Dados

### Formato de Entrada (CSV)

O arquivo CSV deve conter as seguintes colunas:

| Coluna | Tipo | Descrição | Obrigatório |
|--------|------|-----------|-------------|
| Data | String/Data | Data da transação (formato: YYYY-MM-DD) | ✅ |
| Categoria | String | Categoria da transação (Receita/Despesa) | ❌ |
| Descrição | String | Descrição detalhada | ❌ |
| Valor | Numérico | Valor monetário | ✅ |
| Moeda | String | Código da moeda (ex: BRL, USD) | ❌ |
| Status | String | Status da transação | ❌ |

### Exemplo de Dados

```csv
Data,Categoria,Descrição,Valor,Moeda,Status
2024-01-15,Receita,Venda de produtos,1500.50,BRL,Confirmado
2024-01-16,Despesa,Compra de estoque,850.75,BRL,Confirmado
```

### Formato de Saída (Excel)

O arquivo Excel gerado contém 4 planilhas:

1. **Dados Processados**: Dados limpos e normalizados com colunas derivadas
2. **Pivot por Mês**: Tabela dinâmica agregando valores por Categoria e Mês
3. **Pivot por Status**: Tabela dinâmica agregando valores por Categoria e Status
4. **Resumo Estatístico**: Estatísticas descritivas por Categoria (Total, Média, Min, Max, Contagem)

## 📁 Estrutura do Projeto

```
Clean_Dashboard/
├── processar_financeiro.py          # Script principal
├── dados_financeiros_brutos.csv     # Dados de exemplo
├── requirements.txt                 # Dependências Python
├── .gitignore                      # Arquivos ignorados pelo Git
└── README.md                       # Este arquivo
```

## ✨ Funcionalidades

### Tratamento de Dados
- ✅ Remoção automática de registros inválidos
- ✅ Preenchimento inteligente de valores nulos
- ✅ Normalização de strings
- ✅ Conversão de tipos de dados
- ✅ Validação de dados

### Análise e Relatórios
- ✅ Tabelas dinâmicas por múltiplas dimensões
- ✅ Estatísticas descritivas
- ✅ Agregações por período temporal
- ✅ Formatação profissional do Excel

### Formatação Excel
- ✅ Cabeçalhos com cores e fonte em negrito
- ✅ Bordas em todas as células
- ✅ Largura de colunas ajustada automaticamente
- ✅ Formatação monetária (R$)
- ✅ Destaque visual para linhas de total
- ✅ Painéis congelados para melhor navegação

## 🎯 Exemplo de Saída

Ao executar o script, você verá uma saída como:

```
============================================================
🚀 INICIANDO PROCESSAMENTO DE DADOS FINANCEIROS
============================================================
📂 Carregando dados brutos...
✅ 26 registros carregados
📊 Colunas: Data, Categoria, Descrição, Valor, Moeda, Status

🔧 Tratando valores nulos...
✅ Valores nulos tratados

🔄 Normalizando dados...
✅ Dados normalizados

🔢 Corrigindo tipos de dados...
✅ Tipos de dados corrigidos

➕ Adicionando colunas derivadas...
✅ Colunas derivadas adicionadas

📊 Criando Tabela Dinâmica...
✅ Arquivo Excel criado: relatorio_financeiro.xlsx
🎨 Aplicando formatação...
✅ Formatação aplicada

============================================================
✨ PROCESSAMENTO CONCLUÍDO COM SUCESSO!
============================================================

📄 Relatório gerado: relatorio_financeiro.xlsx
📊 Total de registros processados: 24
💰 Valor total: R$ 54,551.10
```

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## 📝 Licença

Este projeto é de código aberto e está disponível para uso livre.

## 👨‍💻 Autor

Lucas Malessa
