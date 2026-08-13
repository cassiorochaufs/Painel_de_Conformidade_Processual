# Solução Analítica para Conformidade e Desempenho no SEI

Repositório da Prova de Conceito desenvolvida para análise de conformidade estrutural, conformidade temporal e desempenho de processos tramitados no Sistema Eletrônico de Informações (SEI), com aplicação ao processo de cessão de servidores da Universidade Federal de Sergipe (UFS).

A solução articula Gestão de Processos de Negócio, BPMN e Mineração de Processos para comparar o comportamento observado nos registros de tramitação do SEI com um modelo de referência computável.

## Finalidade

O repositório disponibiliza os principais componentes necessários para reprodução e transferência da solução analítica:

- extração de registros de tramitação do SEI;
- tratamento e reconstrução do fluxo observado;
- parametrização das unidades e critérios normativos;
- verificação de conformidade com PM4Py;
- cálculo de indicadores de desempenho;
- visualização dos resultados em Power BI.

## Estrutura do repositório

```text
01_extracao/
02_tratamento_sql/
03_conformidade/
04_powerbi/
05_bpmn/
06_parametrizacao/
documentacao/
README.md
