import pandas as pd
import numpy as np

# --- CONFIGURAÇÕES INICIAIS ---
nome_arquivo = "00_Controle do encaminhamento dos alertas - 21 de agosto, 13_53.xlsx"

mapeamento = {
    'ALERTAS-SEI 2023': '2023',
    'ALERTAS-SEI 2024': '2024',
    'ALERTAS-SEI 2025': '2025',
    'ALERTAS-SEI-2026': '2026'
}

# --- FUNÇÕES DE PROCESSAMENTO ---

def expandir_alertas(df):
    """Padroniza separadores e explode alertas em linhas únicas."""
    if df is None: return None
    df = df.copy()
    df['Alertas'] = df['Alertas'].astype(str).str.replace(' e ', ',')
    df['Alertas'] = df['Alertas'].astype(str).str.replace(';', ',')
    df['Alertas'] = df['Alertas'].apply(lambda x: [i.strip() for i in x.split(',') if i.strip()])
    return df.explode('Alertas').reset_index(drop=True)

def processar_e_verificar(df_exp, ano):
    """Realiza a limpeza seletiva e reporta a integridade dos dados."""
    if df_exp is None or df_exp.empty: return None
    
    df_limpo = df_exp.dropna(subset=['Processo SEI', 'Alertas']).copy()
    df_limpo = df_limpo[df_limpo['Alertas'].astype(str).str.lower() != 'nan']
    
    total_inicial = len(df_limpo)

    # REMOVER DUPLICATAS INTERNAS (Mesmo Alerta + Mesmo Processo na mesma aba)
    df_sem_internas = df_limpo.drop_duplicates(subset=['Processo SEI', 'Alertas'])
    duplicatas_internas = total_inicial - len(df_sem_internas)

    # IDENTIFICAR DUPLICATAS REAIS (Mesmo Alerta em Processos Diferentes dentro do mesmo ano)
    alertas_compartilhados = df_sem_internas.duplicated(subset=['Alertas'], keep=False).sum()
    
    print(f"\n--- RELATÓRIO DE VERIFICAÇÃO [{ano}] ---")
    print(f"✅ Alertas processados: {len(df_sem_internas)}")
    print(f"🧹 Erros de digitação removidos (mesmo processo): {duplicatas_internas}")
    print(f"🔗 Alertas em múltiplos processos (mantidos): {alertas_compartilhados}")
    
    return df_sem_internas

# --- EXECUÇÃO DO FLUXO ---

print("🚀 Iniciando processamento das guias...")

try:
    todas_guias = pd.read_excel(nome_arquivo, sheet_name=None, header=1)
except Exception as e:
    print(f"❌ Erro ao abrir o arquivo: {e}")
    exit()

dfs_processados = []

for guia_excel, ano_label in mapeamento.items():
    if guia_excel in todas_guias:
        df = todas_guias[guia_excel].copy()
        df.columns = df.columns.astype(str).str.strip()
        
        try:
            df_resumo = df[["Processo SEI", "Alertas"]]
            df_exp = expandir_alertas(df_resumo)
            df_final = processar_e_verificar(df_exp, ano_label)
            
            if df_final is not None:
                dfs_processados.append(df_final)
        except KeyError:
            print(f"❌ Colunas não encontradas na guia {guia_excel}.")
    else:
        print(f"⚠ Guia {guia_excel} não encontrada.")

# --- CONSOLIDAÇÃO E EXPORTAÇÃO PARA EXCEL ---

if dfs_processados:
    df_master = pd.concat(dfs_processados, ignore_index=True)
    
    # Identifica duplicatas reais na base toda (entre anos inclusive)
    df_duplicados_reais = df_master[df_master.duplicated(subset=['Alertas'], keep=False)].sort_values(by='Alertas')

    # Nome do arquivo de saída
    arquivo_saida = "Relatorio_Alertas_SEI_Final.xlsx"

    # Criando o arquivo Excel com múltiplas abas
    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        # Aba 1: Base completa e limpa
        df_master.to_excel(writer, sheet_name='Alertas_SEI', index=False)
        
        # Aba 2: Auditoria (se houver duplicatas)
        if not df_duplicados_reais.empty:
            df_duplicados_reais.to_excel(writer, sheet_name='Alertas_multiprocessos', index=False)
            status_auditoria = f"🔍 {len(df_duplicados_reais)} linhas de duplicatas reais encontradas."
        else:
            status_auditoria = "✅ Nenhuma duplicata real encontrada."

    print("\n" + "="*40)
    print("📊 RESULTADO FINAL")
    print(f"Arquivo gerado: {arquivo_saida}")
    print(f"Total de alertas na base: {len(df_master)}")
    print(status_auditoria)
    print("="*40)
else:
    print("\nNenhum dado processado.")
