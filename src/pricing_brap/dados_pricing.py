'''
Camada de dados - Pricing BRAP

Função:
- Localizar os arquivos de origem (controle, custo, tabela padrão)
de forma eficiente, independente da máquina de utilização.
- Garantir que o controle tenha as colunas de status necessárias
para o pipeline.
- Ler, tratar e enriquecer as pendências de precificação, preservando
a referência de linha original do controle.

O módulo em questão só LÊ dados e prepara o "arquivo de análise".
A escrita de volta no controle (precificação em si) fica em outro
módulo.
'''

import os
from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook

# ======================================================================
# 0. Configuração central

CONFIG = {
    # Biblioteca sincronizada (Teams/SharePoint) onde armazena o controle
    # Ajuste se os arquivos estiverem em outro local
    'pasta_biblioteca_controle': 'afip.com.br/Pricing - Geral',
    'arquivo_controle': 'CONSOLIDADO_DEMANDAS_BRAP_v1.xlsx',
    'aba_controle': 'WhatsApp',

    # Arquivos dentro da pasta de rede
    'arquivo_custo': 'Nova planilha de PROC V2 - 27.04.2026.xlsx',
    'aba_custo': 'PROC',
    'skiprows_custo': 3,

    'arquivo_tabela_padrao': r'V:\Gestao de Excelencia Comercial\AreaComum\BRAP\ROGER'
                              r'\TABELA_PADRÃO\TABELAS_BRAP_05.01.2026.xlsx',
    'aba_tabela_padrao': 'TABELA_PADRÃO',
    'skiprows_tabela_padrao': 2,
}

# Colunas de controle
COLUNAS_STATUS = ['STATUS', 'CHAMADO', 'RETORNO']
STATUS_PENDENTE = 'PENDENTE'
STATUS_CONCLUIDO = 'FINALIZADO'

# =======================================================================
# 1. Localização de arquivos

def localizar_arquivo(nome_arquivo: str, pastas_candidatas: list[str] | None = None) -> Path:
    '''
    Localiza um arquivo dentro das pastas sincronizadas do usuário,
    sem depender do nome do usuário do Windows nem nome exato da biblioteca/canal.
    '''
    home = Path.home()
    ignorar = {'AppData', 'Application Data', '.git', 'node_modules'}

    if pastas_candidatas:
        bases = [home / p for p in pastas_candidatas]
    else:
        bases = [p for p in home.iterdir() if p.is_dir() and p.name not in ignorar]

    encontrados: list[Path] = []
    for base in bases:
        if base.exists():
            encontrados += list(base.rglob(nome_arquivo))

    if not encontrados:
        raise FileNotFoundError(
            f'Arquivo "{nome_arquivo}" não encontrado nas pastas sincronizadas de {home}. '
            f'Confirme se a biblioteca foi sincronizada (não apenas "apatalho adicionado").'
        )

    if len(encontrados) > 1:
        raise RuntimeError(
            f'Mais de um arquivo "{nome_arquivo}" encontrado: {encontrados}. '
            f'Especifique a pasta em "pasta_candidatas".'
        )

    return encontrados[0]

# =======================================================================
# 2. Migração de colunas de STATUS

def garantir_colunas_status(caminho_arquivo: Path, aba: str) -> bool:
    wb = load_workbook(caminho_arquivo)
    ws = wb[aba]

    cabecalho = {cell.value: cell.column for cell in ws[1] if cell.value}
    faltantes = [c for c in COLUNAS_STATUS if c not in cabecalho]

    if not faltantes:
        return False

    proxima_coluna = ws.max_column + 1
    col_responsavel = cabecalho.get('RESPONSÁVEL')

    for i, nome_coluna in enumerate(faltantes):
        col = proxima_coluna + i
        ws.cell(row=1, column=col, value=nome_coluna)
        cabecalho[nome_coluna] = col

    if 'STATUS' in faltantes:
        col_status = cabecalho['STATUS']
        for row in range(2, ws.max_row + 1):
            responsavel_vazio = (
                col_responsavel is None
                or ws.cell(row=row, column=col_responsavel).value in (None, '', ' ')
            )
            status_inicial = STATUS_PENDENTE if responsavel_vazio else STATUS_CONCLUIDO
            ws.cell(row=row, column=col_status, value=status_inicial)

    wb.save(caminho_arquivo)
    print(f'Migração aplicada: colunas {faltantes} adicionadas em "{aba}". '
          f'Revise o STATUS inicial das linhas antigas antes de confiar nelas.')
    return True

# =======================================================================
# 3. Leitura das bases

def ler_controle() -> tuple[pd.DataFrame, Path]:
    caminho = localizar_arquivo(
        CONFIG['arquivo_controle'],
        pastas_candidatas=[CONFIG['pasta_biblioteca_controle']],
    )
    garantir_colunas_status(caminho, CONFIG['aba_controle'])
    df = pd.read_excel(caminho, sheet_name=CONFIG['aba_controle'])
    return df, caminho

def ler_custo() -> pd.DataFrame:
    return pd.read_excel(
        localizar_arquivo(
            CONFIG['arquivo_custo'],
            pastas_candidatas=[CONFIG['pasta_biblioteca_controle']],
        ),
        sheet_name=CONFIG['aba_custo'],
        skiprows=CONFIG['skiprows_custo'],
    )

def ler_tabela_padrao() -> pd.DataFrame:
    return pd.read_excel(
        CONFIG['arquivo_tabela_padrao'],
        sheet_name=CONFIG['aba_tabela_padrao'],
        skiprows=CONFIG['skiprows_tabela_padrao']
    )

# =======================================================================
# 4. Tratamento das bases

def tratar_custo(df_custo: pd.DataFrame) -> pd.DataFrame:
    colunas = ['MN SHIFT', 'Insumos de Coleta', 'Mão de Obra de Coleta',
               'Custo Fixo de Processamento', 'Custo Variável de Processamento']

    colunas_numericas = colunas[1:]

    df = df_custo[colunas].copy().rename(columns={'MN SHIFT': 'MN_SHIFT'})
    for col in colunas_numericas:
        df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

    df['CUSTO_TOTAL'] = df[colunas_numericas].sum(axis=1)
    return df

def tratar_tabela_padrao(df_padrao: pd.DataFrame) -> pd.DataFrame:
    colunas = ['MN_SHIFT', 'DESCRIÇÃO_EXAME', 'MN_PAI', 'STATUS', 'MATRIZ',
               'TABELA_PADRÃO', 'TM_BRAP', 'MEDIA_MERCADO', 'VALOR_CONCORRENTE',
               'FAIXA_1', 'FAIXA_2', 'FAIXA_3', 'FAIXA_4', 'FAIXA_5', 'FAIXA_6',
               'FAIXA_7', 'FAIXA_8']
    colunas_numericas = ['TABELA_PADRÃO', 'TM_BRAP', 'MEDIA_MERCADO', 'VALOR_CONCORRENTE',
                          'FAIXA_1', 'FAIXA_2', 'FAIXA_3', 'FAIXA_4', 'FAIXA_5',
                          'FAIXA_6', 'FAIXA_7', 'FAIXA_8']

    df = df_padrao[colunas].copy()
    for col in colunas_numericas:
        df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

    df['QTD_REP_MN'] = df['MN_PAI'].map(df['MN_PAI'].value_counts())
    df['LOCAL'] = np.where(
        df['MATRIZ'].str.contains('Apoi', case=False, na=False), 'Apoio', 'Interno'
    )
    return df

def verificar_chave_unica(df: pd.DataFrame, coluna: str, nome_base: str) -> None:
    '''
    Alerta (não interrompe) se a coluna usada como chave de merge tiver
    duplicatas - isso é o que causa multiplicação silenciosa de linhas
    depois do merge.
    '''
    duplicados = df[df[coluna].duplicated(keep=False)][coluna].unique()
    if len(duplicados) > 0:
        print(
            f'[ATENÇÃO] "{nome_base}" tem {len(duplicados)} valor(es) de '
            f'"{coluna}" duplicado(s): {list(duplicados)[:10]}'
            f'{ " ..." if len(duplicados) > 10 else ""}'
            f'O merge vai multiplicar as linhas correspondentes no controle.'
        )

# =======================================================================
# 5. Orquestrador - ponto de entrada deste módulo

def coletar_pendencias() -> tuple[pd.DataFrame, pd.DataFrame, Path]:
    '''
    Lê e enriquece as demandas pendentes de precificação

    Retorno:
        arquivo_analise: DataFrame pronto para etapa de precificação
        df_controle: controle mestre completo, sem filtro
        caminho_controle: onde o arquivo de controle está localizado
    '''

    df_controle, caminho_controle = ler_controle()

    pendentes = df_controle.loc[df_controle['STATUS'] == STATUS_PENDENTE].copy()
    pendentes = pendentes.reset_index().rename(columns={'index': '_indice_controle'})

    colunas_pendentes = [
        '_indice_controle', 'RECEBIMENTO', 'REPRESENTANTE', 'LABORATÓRIO', 'FP',
        'MN_SHIFT', 'PREÇO_CONCORRENTE', 'PREÇO_SUGERIDO', 'VOLUME', 'NEGOCIADO',
        'APROVADO', 'OBS.',
    ]

    pendentes = pendentes[colunas_pendentes]

    df_custo = tratar_custo(ler_custo())
    df_padrao = tratar_tabela_padrao(ler_tabela_padrao())

    verificar_chave_unica(df_custo, 'MN_SHIFT', 'planilha de custo')
    verificar_chave_unica(df_padrao, 'MN_SHIFT', 'tabela padrão')

    arquivo_analise = (
        pendentes
        .merge(df_custo, how='left', on='MN_SHIFT')
        .merge(df_padrao, how='left', on='MN_SHIFT')
    )

    arquivo_analise['MARGEM'] = (
        (arquivo_analise['APROVADO'] - arquivo_analise['CUSTO_TOTAL'])
        / arquivo_analise['APROVADO']
    )

    colunas_finais = [
        '_indice_controle', 'RECEBIMENTO', 'REPRESENTANTE', 'LABORATÓRIO', 'FP',
        'MN_SHIFT', 'QTD_REP_MN', 'DESCRIÇÃO_EXAME', 'MATRIZ', 'LOCAL', 'STATUS',
        'PREÇO_CONCORRENTE', 'PREÇO_SUGERIDO', 'VOLUME', 'NEGOCIADO', 'APROVADO',
        'MARGEM', 'OBS.', 'CUSTO_TOTAL', 'TABELA_PADRÃO', 'TM_BRAP',
        'MEDIA_MERCADO', 'VALOR_CONCORRENTE', 'FAIXA_1',
    ]
    arquivo_analise = arquivo_analise[colunas_finais]

    print(f'{len(arquivo_analise)} pendência(s) coletada(s) e enriquecida(s).')
    return arquivo_analise, df_controle, caminho_controle

if __name__ == '__main__':
    analise, controle, caminho = coletar_pendencias()
    print(analise.head())
