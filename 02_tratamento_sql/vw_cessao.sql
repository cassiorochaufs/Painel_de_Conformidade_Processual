/*
===============================================================================
VIEW: vw_cessao
Finalidade:
    Reconstruir episódios de permanência/tramitação do processo de cessão
    a partir do histórico do SEI.

Saída principal:
    - Processo
    - Unidade
    - StayID
    - Data_Inicio
    - Data_Fim
    - Status_Etapa
    - Duracao_Decimal_Dias

Observações:
    1. A view considera apenas processos identificados pelo marcador "CESSÃO".
    2. Entradas e saídas são inferidas a partir das descrições registradas
       no histórico do SEI.
    3. O StayID diferencia retornos sucessivos de um mesmo processo à mesma
       unidade.
    4. Para episódios ainda em andamento, a duração é calculada até o instante
       corrente da consulta.
    5. A lógica abaixo foi mantida fiel à rotina utilizada na Prova de Conceito.
===============================================================================
*/

CREATE VIEW vw_cessao AS

WITH HistoricoBase AS (
    /*
    ---------------------------------------------------------------------------
    1. Conversão e seleção dos campos necessários do histórico.
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

    A função bool_or() verifica, dentro de cada processo, se pelo menos um
    registro possui o marcador correspondente à cessão.
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

ProcessosFiltrados AS (
    /*
    ---------------------------------------------------------------------------
    3. Restrição da análise aos processos identificados como cessão.
    ---------------------------------------------------------------------------
    */
    SELECT
        "Processo",
        "Unidade",
        "Obs",
        DataConvertida
    FROM ProcessosComFlag
    WHERE ProcessoDeCessao = TRUE
),

EventosChave AS (
    /*
    ---------------------------------------------------------------------------
    4. Classificação dos registros relevantes como eventos de INÍCIO ou FIM.

    Eventos de início:
        - Processo público gerado
        - Processo recebido na unidade
        - Reabertura do processo na unidade

    Eventos de fim:
        - Conclusão do processo na unidade
        - Processo remetido pela unidade
        - Anexado ao processo

    Para eventos de remessa, a unidade de origem é recuperada, quando possível,
    da própria descrição do andamento. Caso a extração não encontre valor, é
    utilizada a unidade registrada no campo "Unidade".
    ---------------------------------------------------------------------------
    */
    SELECT
        "Processo",
        DataConvertida,
        "Obs",

        CASE
            WHEN "Obs" ILIKE '%Processo público gerado%'
                THEN 'INICIO'
            WHEN "Obs" ILIKE '%Processo recebido na unidade%'
                THEN 'INICIO'
            WHEN "Obs" ILIKE '%Reabertura do processo na unidade%'
                THEN 'INICIO'

            WHEN "Obs" ILIKE '%Conclusão do processo na unidade%'
                THEN 'FIM'
            WHEN "Obs" ILIKE '%Processo remetido pela unidade%'
                THEN 'FIM'
            WHEN "Obs" ILIKE '%Anexado ao processo%'
                THEN 'FIM'
        END AS "TipoEvento",

        TRIM(
            CASE
                WHEN "Obs" ILIKE '%Processo remetido pela unidade%'
                    THEN COALESCE(
                        SUBSTRING(
                            "Obs"
                            FROM '(?i)remetido pela unidade[[:space:]]+([A-Za-z0-9/\-]+)'
                        ),
                        "Unidade"
                    )
                ELSE "Unidade"
            END
        ) AS "UnidadeCorreta"

    FROM ProcessosFiltrados

    WHERE
        "Obs" ILIKE '%Processo público gerado%'
        OR "Obs" ILIKE '%Processo recebido na unidade%'
        OR "Obs" ILIKE '%Reabertura do processo na unidade%'
        OR "Obs" ILIKE '%Conclusão do processo na unidade%'
        OR "Obs" ILIKE '%Processo remetido pela unidade%'
        OR "Obs" ILIKE '%Anexado ao processo%'
),

IdentificaEstadias AS (
    /*
    ---------------------------------------------------------------------------
    5. Identificação dos episódios de permanência por processo e unidade.

    Cada evento de FIM encerra uma estadia. O contador acumulado dos eventos de
    fim anteriores, acrescido de 1, gera o StayID. Dessa forma, retornos do
    mesmo processo à mesma unidade são tratados como episódios distintos.
    ---------------------------------------------------------------------------
    */
    SELECT
        *,
        COALESCE(
            SUM(
                CASE
                    WHEN "TipoEvento" = 'FIM' THEN 1
                    ELSE 0
                END
            ) OVER (
                PARTITION BY
                    "Processo",
                    "UnidadeCorreta"
                ORDER BY
                    DataConvertida,
                    CASE
                        WHEN "TipoEvento" = 'INICIO' THEN 1
                        ELSE 2
                    END
                ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING
            ),
            0
        ) + 1 AS "StayID"

    FROM EventosChave
)

/*
-------------------------------------------------------------------------------
6. Consolidação de cada episódio de tramitação.

Data_Inicio:
    Primeiro evento de INÍCIO do episódio.

Data_Fim:
    Último evento de FIM do episódio.

Status_Etapa:
    "Em Andamento" quando não existe Data_Fim.
    "Concluída" quando existe Data_Fim.

Duracao_Decimal_Dias:
    Diferença cronológica entre início e fim, expressa em dias decimais.
    Para episódios em andamento, utiliza CURRENT_TIMESTAMP como referência.
-------------------------------------------------------------------------------
*/
SELECT
    "Processo",
    "UnidadeCorreta" AS "Unidade",
    "StayID",

    MIN(DataConvertida)
        FILTER (WHERE "TipoEvento" = 'INICIO')
        AS "Data_Inicio",

    MAX(DataConvertida)
        FILTER (WHERE "TipoEvento" = 'FIM')
        AS "Data_Fim",

    CASE
        WHEN MAX(DataConvertida)
            FILTER (WHERE "TipoEvento" = 'FIM') IS NULL
            THEN 'Em Andamento'
        ELSE 'Concluída'
    END AS "Status_Etapa",

    (
        EXTRACT(
            EPOCH FROM (
                COALESCE(
                    MAX(DataConvertida)
                        FILTER (WHERE "TipoEvento" = 'FIM'),
                    CURRENT_TIMESTAMP
                )
                -
                MIN(DataConvertida)
                    FILTER (WHERE "TipoEvento" = 'INICIO')
            )
        ) / 86400.0
    ) AS "Duracao_Decimal_Dias"

FROM IdentificaEstadias

GROUP BY
    "Processo",
    "UnidadeCorreta",
    "StayID"

HAVING
    MIN(DataConvertida)
        FILTER (WHERE "TipoEvento" = 'INICIO') IS NOT NULL

ORDER BY
    "Processo",
    "Data_Inicio";
