import pandas as pd
import unicodedata
import re

# Carregar a base de NPS
nps_df = pd.read_excel(
    r'\Gog\Part I\1. Base de NPS.xlsx'
)

# Categorias expandidas com palavras-chave
categorias = {
    'Entrega': [
        'atraso', 'demora', 'entrega', 'demorou', 'transportadora', 'frete', 'atrasada',
        'prazo', 'chegou tarde', 'chegou depois', 'muito tempo', 'não chegou', 'não recebi',
        'entregue errado', 'entregaram errado', 'esperando entrega', 'demorando muito',
        'saiu para entrega e não chegou', 'entrega incompleta'
    ],
    'Produto': [
        'quebrado', 'defeito', 'arranhado', 'descascado', 'danificado', 'manchado',
        'estragado', 'veio errado', 'produto errado', 'produto ruim', 'qualidade ruim',
        'não funciona', 'falta peça', 'incompleto', 'mal feito', 'fragil', 'rachado',
        'imperfeito', 'decepcionado com produto', 'diferença de cor', 'nada a ver com a foto',
        'esperava mais do produto'
    ],
    'Atendimento': [
        'atendimento', 'suporte', 'demora resposta', 'resposta', 'reclamei',
        'ninguém me respondeu', 'chat não funciona', 'atendente grosso',
        'mal atendido', 'demora no suporte', 'demora para responder', 'péssimo atendimento',
        'fui ignorado', 'cansei de esperar', 'não resolveram', 'ninguém ajuda',
        'me deixaram no vácuo'
    ],
    'Pagamento': [
        'boleto', 'cartão', 'pagamento', 'cobrança', 'não confirmou pagamento',
        'paguei e não recebi', 'problema com cartão', 'pagamento em duplicidade',
        'desconto não aplicado', 'valor errado', 'cobrança indevida',
        'problema no checkout', 'falha na cobrança', 'pix não foi reconhecido'
    ],
    'Outros': []
}

# Função para normalizar texto
def normalizar_texto(texto):
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join([c for c in texto if not unicodedata.combining(c)])
    texto = re.sub(r'[^\w\s]', '', texto)
    return texto

# Função para classificar todos os registros
def classificar_problema(row):
    nota = row['Rating']
    texto = row['Body']

    if nota <= 6:
        texto_normalizado = normalizar_texto(texto)
        for categoria, palavras in categorias.items():
            for palavra in palavras:
                if palavra in texto_normalizado:
                    return categoria
        return 'Outros'
    elif nota <= 8:
        return 'Neutro (sem problema)'
    else:
        return 'Promotor (sem problema)'

# Aplicar classificação em toda a base
nps_df['categoria_problema'] = nps_df.apply(classificar_problema, axis=1)

# Criar relatório geral
relatorio = nps_df.groupby('categoria_problema').agg(
    qtd_respostas=('Order ID', 'count')
).reset_index()

# Exportar os arquivos
nps_df.to_excel(
    r'\Gog\project\doc\result\nps_classificado_todos.xlsx',
    index=False
)

relatorio.to_excel(
    r'\Gog\project\doc\result\relatorio_geral_nps.xlsx',
    index=False
)

print("Classificação de todos os registros concluída com sucesso.")
