# Guia de reprodução da solução analítica

Este repositório contém os componentes utilizados na Prova de Conceito de uma solução analítica baseada em Mineração de Processos para análise de conformidade e desempenho de processos tramitados no SEI.

## Fluxo geral

1. Extrair os registros de tramitação do SEI.
2. Armazenar os registros extraídos.
3. Executar as consultas SQL de tratamento dos dados.
4. Parametrizar as unidades e macroatividades.
5. Executar a análise de conformidade com o modelo BPMN de referência.
6. Gerar os arquivos de resultados da análise.
7. Carregar os resultados no template do Power BI.
8. Analisar indicadores de conformidade estrutural, temporal e desempenho.

## Componentes do repositório

- `01_extracao`: rotina de extração dos registros do SEI.
- `02_tratamento_sql`: consultas SQL para reconstrução das etapas e relações de fluxo.
- `03_conformidade`: rotina de conformance checking com PM4Py.
- `04_powerbi`: template do painel e documentação das medidas DAX.
- `05_bpmn`: modelo BPMN computável utilizado como referência.
- `06_parametrizacao`: arquivos de parametrização das unidades e critérios normativos.
- `documentacao`: orientações para instalação, execução e interpretação da solução.

## Observações metodológicas

O comportamento observado é reconstruído a partir dos registros disponíveis no SEI e representa o comportamento observável no sistema.

A divergência entre o fluxo observado e o modelo BPMN de referência indica desvio estrutural em relação ao modelo adotado para comparação, não implicando, isoladamente, declaração de ilegalidade.

Os indicadores temporais devem ser interpretados conforme os parâmetros normativos aplicáveis a cada macroetapa. Não foi adotado prazo legal global para a duração total do processo de cessão.

O lead time representa a duração cronológica total da instância do processo. O cycle time representa a duração cronológica das macroetapas analisadas. Os registros do SEI não permitem distinguir diretamente tempo de trabalho ativo e tempo de espera.
