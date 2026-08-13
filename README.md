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
```

### `01_extracao`

Contém a rotina de extração dos registros necessários à análise.

Principais arquivos:

- `sei_scraper.py`
- `.env.example`
- `requirements.txt`

Credenciais e parâmetros locais devem ser fornecidos por variáveis de ambiente. Nenhuma credencial real é disponibilizada neste repositório.

### `02_tratamento_sql`

Contém as consultas utilizadas para tratamento dos registros e reconstrução do comportamento observado.

- `vw_cessao.sql`: reconstrói as etapas e respectivos intervalos temporais.
- `vw_cessao_fluxo.sql`: reconstrói as relações diretamente sucessivas entre unidades ou macroatividades.

### `03_conformidade`

Contém a implementação da análise de conformidade baseada em PM4Py.

Principais arquivos:

- `conformance_check.py`
- `.env.example`
- `requirements.txt`
- `config/mapeamento_unidades.exemplo.csv`

A análise produz métricas de fitness, precision e registros de desvios estruturais identificados entre o comportamento observado e o modelo BPMN de referência.

### `04_powerbi`

Contém a camada de visualização da solução.

- `painel_conformidade_desempenho_sei.pbit`: template parametrizado do painel.
- `medidas_powerbi_documentadas.dax`: documentação das principais medidas utilizadas no painel.

O arquivo `.pbit` não contém os dados da aplicação original. As fontes devem ser configuradas pelo usuário.

### `05_bpmn`

Contém o modelo BPMN computável adotado como referência para a análise de conformidade.

- `modelo_referencia_computavel.bpmn`

### `06_parametrizacao`

Contém os parâmetros institucionais e normativos utilizados na Prova de Conceito.

- `mapeamento_unidades.xlsx`: padronização e classificação das unidades observadas no SEI.
- `matriz_conformidade_normativa.xlsx`: fundamentos e critérios utilizados na análise de conformidade temporal.

### `documentacao`

Contém orientações complementares para reprodução e interpretação da solução.

- `guia_reproducao.md`

## Fluxo analítico

```text
Registros do SEI
      ↓
Extração
      ↓
Tratamento SQL
      ↓
Reconstrução do comportamento observado
      ↓
Mapeamento das unidades e macroatividades
      ↓
Comparação com o BPMN de referência
      ↓
Conformance checking e análise de desempenho
      ↓
Arquivos de resultados
      ↓
Power BI
```

## Indicadores de conformidade

A conformidade estrutural é analisada pela comparação entre o comportamento observado e o modelo BPMN de referência.

A implementação utiliza alinhamentos do PM4Py para obtenção de:

- `fitness`: grau de aderência de cada instância ao modelo de referência;
- `precision`: medida global relacionada ao comportamento permitido pelo modelo em relação ao observado;
- movimentos sincronizados;
- movimentos somente no log;
- movimentos somente no modelo.

No painel, os movimentos são apresentados de forma gerencial como:

- `Sync Move`: Sincronizado;
- `Log Move`: Atividade Extra / Retrabalho;
- `Model Move`: Atividade Omitida / Pulada.

Essas denominações devem ser interpretadas como apoio analítico. Um desvio estrutural não implica, isoladamente, irregularidade ou ilegalidade do processo.

## Indicadores temporais

Na solução:

- **Lead time** corresponde à duração cronológica total da instância do processo.
- **Cycle time** corresponde à duração cronológica da macroetapa analisada.

Os registros de tramitação do SEI permitem observar tempos decorridos entre eventos, mas não permitem distinguir diretamente tempo de trabalho ativo e tempo de espera.

Os parâmetros normativos são aplicados somente às situações previstas na matriz de conformidade normativa. Não foi adotado prazo legal global para conclusão do processo de cessão.

## Reprodutibilidade

Para reproduzir a solução é necessário:

1. configurar as variáveis de ambiente da extração;
2. disponibilizar uma base PostgreSQL para tratamento dos registros;
3. executar as consultas SQL;
4. configurar o mapeamento das unidades;
5. disponibilizar o modelo BPMN de referência;
6. executar a rotina de conformidade;
7. apontar o template do Power BI para as fontes geradas.

Instruções complementares estão disponíveis em [`documentacao/guia_reproducao.md`](documentacao/guia_reproducao.md).

## Segurança e dados

Este repositório não disponibiliza:

- credenciais de acesso ao SEI;
- senhas de banco de dados;
- arquivos `.env` utilizados na aplicação;
- base de dados original extraída do SEI;
- arquivos `.pbix` com dados carregados.

Os arquivos publicados destinam-se à demonstração, reprodução metodológica e transferência da solução.

## Licença

O código-fonte disponibilizado neste repositório é licenciado sob a [MIT License](LICENSE), permitindo seu uso, cópia, modificação e redistribuição nos termos da licença.

Essa autorização aplica-se aos componentes de código da solução, incluindo scripts Python, consultas SQL e medidas DAX.

Os documentos, modelos, planilhas de parametrização e demais materiais acadêmicos disponibilizados no repositório não são abrangidos pela licença MIT, salvo indicação expressa em contrário, devendo ser utilizados com a devida atribuição de autoria e referência à pesquisa.

## Como citar

Se esta solução, seus códigos ou sua estrutura metodológica forem utilizados em trabalhos acadêmicos ou aplicações institucionais, recomenda-se citar o autor e o repositório.

As informações estruturadas para citação estão disponíveis no arquivo [`CITATION.cff`](CITATION.cff).

**Autor:** Cassio Santos Araujo Rocha
