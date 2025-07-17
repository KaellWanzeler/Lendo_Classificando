
import pandas as pd
import unicodedata
import re

# Carregar a base de NPS
nps_df = pd.read_excel(r'C:\Users\Kosmos Soluções\OneDrive - Kosmos Construtora Ltda\Documentos\ProjectsUser\Gog\Part I\1. Base de NPS.xlsx')

# Filtrar os detratores (notas de 0 a 6)
detratores_df = nps_df[nps_df['Rating'] <= 6].copy()

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

# Função para classificar comentário com base no texto normalizado
def classificar_problema(texto):
    texto_normalizado = normalizar_texto(texto)
    for categoria, palavras in categorias.items():
        for palavra in palavras:
            if palavra in texto_normalizado:
                return categoria
    return 'Outros'

# Aplicar classificação
detratores_df['categoria_problema'] = detratores_df['Body'].apply(classificar_problema)

# Criar relatório semanal
relatorio = detratores_df.groupby('categoria_problema').agg(
    qtd_detratores=('Order ID', 'count')
).reset_index()

# Exportar os arquivos
detratores_df.to_excel(r'C:\Users\Kosmos Soluções\OneDrive - Kosmos Construtora Ltda\Documentos\ProjectsUser\Gog\project\doc\result\detratores_classificados.xlsx', index=False)
relatorio.to_excel(r'C:\Users\Kosmos Soluções\OneDrive - Kosmos Construtora Ltda\Documentos\ProjectsUser\Gog\project\doc\result\relatorio_impacto_semanal.xlsx', index=False)

print("Relatórios gerados com sucesso!")
