
import pandas as pd

# Carregar os arquivos
clientes_df = pd.read_excel('Base de Clientes.xlsx')
estoque_df = pd.read_excel('Relatório Estoque.xlsx')
pedidos_df = pd.read_excel('Base de dados - Detalhes dos pedidos.xlsx')

# Renomear colunas para padronizar
clientes_df.rename(columns={'Código do Pedido': 'Pedido'}, inplace=True)
pedidos_df.rename(columns={'Detalhe do pedido': 'Produto'}, inplace=True)
estoque_df.rename(columns={'Produtos': 'Produto'}, inplace=True)

# Juntar clientes com pedidos
clientes_pedidos = pd.merge(clientes_df, pedidos_df, on='Pedido', how='left')

# Juntar com estoque
clientes_pedidos_estoque = pd.merge(clientes_pedidos, estoque_df, on='Produto', how='left')

# Filtrar produtos com estoque 0
pedidos_travados = clientes_pedidos_estoque[clientes_pedidos_estoque['Estoque'] == 0].copy()

# Selecionar colunas para exportação
pedidos_travados_final = pedidos_travados[[
    'Nome do Cliente', 'Telefone', 'E-mail', 'Pedido', 'Produto', 'Data do Pedido'
]]

# Exportar para Excel
pedidos_travados_final.to_excel('pedidos_travados.xlsx', index=False)

print("Arquivo 'pedidos_travados.xlsx' gerado com sucesso.")
