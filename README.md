# Limpeza de alertas-SEI

## Descrição
Este script realiza a expansão dos alertas para cada SEI correspondente, gerando um arquivo final com únicos valores por células. Também faz a verificação de dados, removendo alertas iguais no mesmo processo e verificando a existência de um mesmo alerta em processos SEI diferentes. O arquivo final `Relatorio_Alertas_SEI_Final.xlsx` contém umas guias: `Alertas_SEI` e `Alertas_multiprocessos`, sendo esta última apenas para a verificação dos alertas em mais de um processo. 

## Uso
Na pasta desse repositório, coloque o arquivo `00_Controle do encaminhamento dos alertas - 21 de agosto, 13_53.xlsx`. No seu compilador de preferência, execute o arquivo `alertas_SEI.py`.  

