/*
===============================================================================
VIEW: vw_cessao_fluxo
Finalidade:
    Reconstruir as relações de origem e destino observadas nas tramitações
    dos processos de cessão, produzindo a estrutura utilizada na representação
    do comportamento observado (modelo descoberto).

Saída principal:
    - Processo
    - Unidade_Origem
    - Unidade_Destino
    - Data_Transicao
    - Horas_de_Transicao

Observações:
    1. A view considera apenas processos identificados pelo marcador "CESSÃO".
    2. Os eventos são ordenados cronologicamente dentro de cada processo.
    3. A primeira ocorrência observada é representada como transição
       INÍCIO -> unidade inicial.
    4. Eventos de remessa produzem relações entre unidade de origem e unidade
       de destino.
    5. Conclusões e anexações produzem relações unidade -> FIM.
    6. A lógica abaixo foi mantida fiel à rotina utilizada na Prova de Conceito.
===============================================================================
*/

CREATE OR REPLACE VIEW vw_cessao_fluxo AS

WITH HistoricoBase AS (
    /*
    ---------------------------------------------------------------------------
    1. Seleção dos campos necessários e conversão da data/hora do histórico.
    ---------------------------------------------------------------------------
    */
    SELECT
        "Processo",
        "Unidade",
        "Obs",
        "Marcador",
        TO_TIMESTAMP("Data", 'DD/MM/YYYY HH24:MI') AS DataConvertida
    FROM historico
),

ProcessosComFlag AS (
    /*
    ---------------------------------------------------------------------------
    2. Identificação dos processos de cessão.

    A função bool_or() verifica, para cada processo, se pelo menos um registro
    possui o marcador correspondente à cessão.
    ---------------------------------------------------------------------------
    */
    SELECT
        *,
        bool_or(
            TRIM("Marcador") ~* '^Marcador\s+CESSÃO$'
        ) OVER (
            PARTITION BY "Processo"
        ) AS ProcessoDeCessao
    FROM HistoricoBase
),

EventosOrdenados AS (
    /*
    ---------------------------------------------------------------------------
    3. Ordenação cronológica dos eventos de cada processo.

    O ROW_NUMBER() identifica a primeira ocorrência registrada de cada
    instância, utilizada posteriormente para construir a relação
    INÍCIO -> unidade inicial.
    ---------------------------------------------------------------------------
    */
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY "Processo"
            ORDER BY DataConvertida ASC
        ) AS num_linha
    FROM ProcessosComFlag
    WHERE ProcessoDeCessao = TRUE
),

EventosTransicao AS (
    /*
    ---------------------------------------------------------------------------
    4. Reconstrução das transições observadas.

    Unidade_Origem:
        - primeira ocorrência: INÍCIO;
        - remessa: unidade de origem extraída da descrição do evento;
        - conclusão/anexação: unidade registrada no histórico.

    Unidade_Destino:
        - primeira ocorrência: unidade que inicia o processo;
        - remessa: unidade de destino registrada no campo "Unidade";
        - conclusão/anexação: FIM.

    Quando a extração da unidade de origem de uma remessa não encontra valor,
    mantém-se o rótulo "DESCONHECIDA", conforme a lógica da rotina original.
    ---------------------------------------------------------------------------
    */
    SELECT
        "Processo",
        DataConvertida AS "Data_Transicao",

        CASE
            WHEN num_linha = 1
                THEN 'INÍCIO'

            WHEN "Obs" ILIKE '%Processo remetido pela unidade%'
                THEN TRIM(
                    COALESCE(
                        SUBSTRING(
                            "Obs"
                            FROM '(?i)remetido pela unidade[[:space:]]+([A-Za-z0-9/\-]+)'
                        ),
                        'DESCONHECIDA'
                    )
                )

            WHEN "Obs" ILIKE '%Conclusão do processo na unidade%'
                OR "Obs" ILIKE '%Anexado ao processo%'
                THEN TRIM("Unidade")
        END AS "Unidade_Origem",

        CASE
            WHEN num_linha = 1
                THEN TRIM("Unidade")

            WHEN "Obs" ILIKE '%Processo remetido pela unidade%'
                THEN TRIM("Unidade")

            WHEN "Obs" ILIKE '%Conclusão do processo na unidade%'
                OR "Obs" ILIKE '%Anexado ao processo%'
                THEN 'FIM'
        END AS "Unidade_Destino"

    FROM EventosOrdenados

    WHERE
        num_linha = 1
        OR "Obs" ILIKE '%Processo remetido pela unidade%'
        OR "Obs" ILIKE '%Conclusão do processo na unidade%'
        OR "Obs" ILIKE '%Anexado ao processo%'
)

/*
-------------------------------------------------------------------------------
5. Consolidação das relações de transição.

Horas_de_Transicao:
    Diferença, em horas decimais, entre o instante da transição atual e o
    instante da transição anterior do mesmo processo.

A primeira transição de cada processo não possui valor anterior e, portanto,
apresentará valor nulo para Horas_de_Transicao.
-------------------------------------------------------------------------------
*/
SELECT
    "Processo",
    "Unidade_Origem",
    "Unidade_Destino",
    "Data_Transicao",

    EXTRACT(
        EPOCH FROM (
            "Data_Transicao"
            -
            LAG("Data_Transicao") OVER (
                PARTITION BY "Processo"
                ORDER BY "Data_Transicao"
            )
        )
    ) / 3600.0 AS "Horas_de_Transicao"

FROM EventosTransicao

WHERE
    "Unidade_Origem" IS NOT NULL
    AND "Unidade_Destino" IS NOT NULL

ORDER BY
    "Processo",
    "Data_Transicao";
