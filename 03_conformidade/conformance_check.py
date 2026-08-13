"""
Verificação de conformidade estrutural com PM4Py.

A rotina:
1. carrega os episódios tratados do processo;
2. aplica a parametrização de unidades/macroatividades;
3. constrói o event log;
4. carrega o modelo BPMN de referência e o converte para rede de Petri;
5. executa alignments e calcula fitness por trace e precision global;
6. detalha os movimentos dos alinhamentos;
7. exporta arquivos CSV estáveis para consumo pela camada de visualização.

Segurança
---------
Nenhuma credencial, caminho pessoal ou dado de acesso deve ser escrito no
código-fonte. As configurações sensíveis são fornecidas por variáveis de
ambiente ou por um arquivo .env local, que não deve ser versionado.

Observação metodológica
-----------------------
Os rótulos "Log Move (Atividade Extra / Retrabalho)" e
"Model Move (Atividade Omitida / Pulada)" são mantidos para compatibilidade
com a interface analítica. Eles representam tipos de movimento de alinhamento
e não comprovam, isoladamente, retrabalho, omissão, ilegalidade ou
irregularidade administrativa.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import pm4py
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Config:
    """Parâmetros necessários à execução da análise."""

    source_mode: str
    database_url: str | None
    database_view: str
    source_csv: Path | None

    mapping_file: Path
    bpmn_file: Path
    output_dir: Path

    case_column: str
    unit_column: str
    timestamp_column: str

    fitness_alert_threshold: float
    unmapped_activity: str | None
    strip_model_prepositions: bool


def configurar_logging(verbose: bool = False) -> None:
    """Configura logs sem exibir credenciais ou strings de conexão."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def parse_bool(valor: str | None, default: bool = False) -> bool:
    """Converte textos comuns de configuração em booleano."""
    if valor is None:
        return default

    valor_normalizado = valor.strip().lower()
    if valor_normalizado in {"1", "true", "yes", "sim", "s"}:
        return True
    if valor_normalizado in {"0", "false", "no", "nao", "não", "n"}:
        return False

    raise ValueError(f"Valor booleano inválido: {valor!r}")


def caminho_opcional(valor: str | None) -> Path | None:
    """Converte um caminho textual opcional em Path absoluto."""
    if not valor or not valor.strip():
        return None
    return Path(valor).expanduser().resolve()


def validar_identificador_sql(nome: str) -> str:
    """
    Valida nome simples de view/tabela ou schema.view.

    A validação evita que a configuração de DATABASE_VIEW seja utilizada como
    texto SQL arbitrário.
    """
    padrao = r"^[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?$"
    if not re.fullmatch(padrao, nome):
        raise ValueError(
            "DATABASE_VIEW deve conter apenas um identificador SQL simples "
            "ou schema.identificador."
        )
    return nome


def carregar_configuracao() -> Config:
    """Carrega parâmetros de linha de comando e do arquivo .env local."""
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Executa conformance checking do event log contra um BPMN."
    )
    parser.add_argument(
        "--source-mode",
        choices=("database", "csv"),
        default=os.getenv("SOURCE_MODE", "database").strip().lower(),
        help="Origem dos dados tratados: database ou csv.",
    )
    parser.add_argument(
        "--source-csv",
        default=os.getenv("SOURCE_CSV", ""),
        help="CSV tratado quando --source-mode=csv.",
    )
    parser.add_argument(
        "--mapping-file",
        default=os.getenv("MAPPING_FILE", "./config/mapeamento_unidades.xlsx"),
        help="Arquivo CSV/XLSX com UNIDADE_ORIGINAL e AJUSTE.",
    )
    parser.add_argument(
        "--bpmn-file",
        default=os.getenv("BPMN_FILE", "./config/modelo_referencia.bpmn"),
        help="Modelo BPMN computável de referência.",
    )
    parser.add_argument(
        "--output-dir",
        default=os.getenv("OUTPUT_DIR", "./output"),
        help="Diretório dos CSVs consumidos pela camada de visualização.",
    )
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    configurar_logging(args.verbose)

    database_url = os.getenv("DATABASE_URL") or None
    database_view = validar_identificador_sql(
        os.getenv("DATABASE_VIEW", "vw_cessao").strip()
    )

    source_csv = caminho_opcional(args.source_csv)
    mapping_file = Path(args.mapping_file).expanduser().resolve()
    bpmn_file = Path(args.bpmn_file).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    case_column = os.getenv("CASE_COLUMN", "Processo").strip()
    unit_column = os.getenv("UNIT_COLUMN", "Unidade").strip()
    timestamp_column = os.getenv("TIMESTAMP_COLUMN", "Data_Inicio").strip()

    threshold = float(os.getenv("FITNESS_ALERT_THRESHOLD", "0.80"))
    if not 0 <= threshold <= 1:
        raise ValueError("FITNESS_ALERT_THRESHOLD deve estar entre 0 e 1.")

    fallback = os.getenv("UNMAPPED_ACTIVITY", "UNIDADE DE ORIGEM").strip()
    unmapped_activity = fallback if fallback else None

    strip_model_prepositions = parse_bool(
        os.getenv("STRIP_MODEL_PREPOSITIONS"),
        default=True,
    )

    if args.source_mode == "database" and not database_url:
        raise ValueError(
            "DATABASE_URL é obrigatória quando SOURCE_MODE=database."
        )
    if args.source_mode == "csv" and source_csv is None:
        raise ValueError(
            "SOURCE_CSV é obrigatório quando SOURCE_MODE=csv."
        )

    output_dir.mkdir(parents=True, exist_ok=True)

    return Config(
        source_mode=args.source_mode,
        database_url=database_url,
        database_view=database_view,
        source_csv=source_csv,
        mapping_file=mapping_file,
        bpmn_file=bpmn_file,
        output_dir=output_dir,
        case_column=case_column,
        unit_column=unit_column,
        timestamp_column=timestamp_column,
        fitness_alert_threshold=threshold,
        unmapped_activity=unmapped_activity,
        strip_model_prepositions=strip_model_prepositions,
    )


# ---------------------------------------------------------------------------
# Utilidades de validação e normalização
# ---------------------------------------------------------------------------

def normalizar_texto(valor: Any) -> str:
    """Normaliza espaços e converte o valor para texto."""
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def normalizar_atividade(valor: Any) -> str:
    """Padroniza o rótulo de uma atividade para comparação."""
    return normalizar_texto(valor).upper()


def validar_arquivo(caminho: Path, descricao: str) -> None:
    """Interrompe a execução quando um arquivo obrigatório não existe."""
    if not caminho.is_file():
        raise FileNotFoundError(f"{descricao} não encontrado: {caminho}")


def validar_colunas(df: pd.DataFrame, colunas: set[str], origem: str) -> None:
    """Verifica a existência das colunas mínimas esperadas."""
    ausentes = sorted(colunas.difference(df.columns))
    if ausentes:
        raise ValueError(
            f"{origem} não contém as colunas obrigatórias: {', '.join(ausentes)}"
        )


# ---------------------------------------------------------------------------
# Aquisição dos dados tratados
# ---------------------------------------------------------------------------

def criar_engine(database_url: str) -> Engine:
    """Cria uma engine SQLAlchemy sem registrar a URL nos logs."""
    return create_engine(database_url, pool_pre_ping=True)


def carregar_dados_database(cfg: Config) -> pd.DataFrame:
    """Carrega os dados da view configurada."""
    assert cfg.database_url is not None

    logging.info("Carregando dados tratados da view %s...", cfg.database_view)
    engine = criar_engine(cfg.database_url)

    try:
        consulta = text(f"SELECT * FROM {cfg.database_view}")
        with engine.connect() as conn:
            return pd.read_sql_query(consulta, conn)
    finally:
        engine.dispose()


def carregar_dados_csv(cfg: Config) -> pd.DataFrame:
    """Carrega dados tratados previamente exportados em CSV."""
    assert cfg.source_csv is not None
    validar_arquivo(cfg.source_csv, "Arquivo CSV de origem")

    logging.info("Carregando dados tratados do CSV...")
    return pd.read_csv(
        cfg.source_csv,
        sep=None,
        engine="python",
        encoding="utf-8-sig",
    )


def carregar_dados(cfg: Config) -> pd.DataFrame:
    """Carrega a origem configurada e retém os campos necessários."""
    if cfg.source_mode == "database":
        df = carregar_dados_database(cfg)
    else:
        df = carregar_dados_csv(cfg)

    colunas = {cfg.case_column, cfg.unit_column, cfg.timestamp_column}
    validar_colunas(df, colunas, "Fonte de dados")

    return df[
        [cfg.case_column, cfg.unit_column, cfg.timestamp_column]
    ].copy()


# ---------------------------------------------------------------------------
# Parametrização das unidades e construção do event log
# ---------------------------------------------------------------------------

def carregar_mapeamento(caminho: Path) -> dict[str, str]:
    """Carrega a tabela UNIDADE_ORIGINAL -> AJUSTE de CSV ou Excel."""
    validar_arquivo(caminho, "Arquivo de parametrização")

    extensao = caminho.suffix.lower()
    if extensao in {".xlsx", ".xls"}:
        df = pd.read_excel(caminho)
    elif extensao == ".csv":
        df = pd.read_csv(caminho, sep=None, engine="python", encoding="utf-8-sig")
    else:
        raise ValueError(
            "MAPPING_FILE deve ser um arquivo .csv, .xlsx ou .xls."
        )

    validar_colunas(
        df,
        {"UNIDADE_ORIGINAL", "AJUSTE"},
        "Arquivo de parametrização",
    )

    df = df[["UNIDADE_ORIGINAL", "AJUSTE"]].copy()
    df["UNIDADE_ORIGINAL"] = df["UNIDADE_ORIGINAL"].map(normalizar_atividade)
    df["AJUSTE"] = df["AJUSTE"].map(normalizar_atividade)

    df = df[
        (df["UNIDADE_ORIGINAL"] != "")
        & (df["AJUSTE"] != "")
    ].drop_duplicates(subset=["UNIDADE_ORIGINAL"], keep="last")

    return dict(zip(df["UNIDADE_ORIGINAL"], df["AJUSTE"]))


def preparar_event_log_dataframe(cfg: Config, df: pd.DataFrame) -> pd.DataFrame:
    """
    Parametriza as unidades e cria as colunas reconhecidas pelo PM4Py.

    A coluna original de unidade é preservada como atributo do evento para
    rastreabilidade e apresentação dos diagnósticos.
    """
    logging.info("Aplicando parametrização das unidades/macroatividades...")
    mapping = carregar_mapeamento(cfg.mapping_file)

    df = df.rename(
        columns={
            cfg.case_column: "Processo",
            cfg.unit_column: "Unidade",
            cfg.timestamp_column: "Data_Inicio",
        }
    )

    df["Processo"] = df["Processo"].map(normalizar_texto)
    df["Unidade"] = df["Unidade"].map(normalizar_texto)
    df["_unidade_chave"] = df["Unidade"].map(normalizar_atividade)

    df["Unidade_Padronizada"] = df["_unidade_chave"].map(mapping)

    nao_mapeadas = sorted(
        df.loc[df["Unidade_Padronizada"].isna(), "Unidade"]
        .dropna()
        .unique()
        .tolist()
    )

    if nao_mapeadas:
        if cfg.unmapped_activity is None:
            amostra = ", ".join(nao_mapeadas[:10])
            raise ValueError(
                "Há unidades sem parametrização e UNMAPPED_ACTIVITY não foi "
                f"definida. Exemplos: {amostra}"
            )

        logging.warning(
            "%d unidade(s) sem correspondência explícita serão classificadas "
            "como %s. Revise esta regra ao adaptar a solução a outro processo.",
            len(nao_mapeadas),
            cfg.unmapped_activity,
        )
        df["Unidade_Padronizada"] = df["Unidade_Padronizada"].fillna(
            normalizar_atividade(cfg.unmapped_activity)
        )

    df["Data_Inicio"] = pd.to_datetime(df["Data_Inicio"], errors="coerce")

    invalidos = int(df["Data_Inicio"].isna().sum())
    if invalidos:
        logging.warning(
            "%d registro(s) sem timestamp válido serão removidos.",
            invalidos,
        )

    df = df.dropna(subset=["Data_Inicio"])
    df = df[
        (df["Processo"] != "")
        & (df["Unidade_Padronizada"] != "")
    ].copy()

    df = df.sort_values(
        by=["Processo", "Data_Inicio"],
        kind="stable",
    ).reset_index(drop=True)

    df["case:concept:name"] = df["Processo"]
    df["concept:name"] = df["Unidade_Padronizada"].map(normalizar_atividade)
    df["time:timestamp"] = df["Data_Inicio"]

    return df[
        [
            "case:concept:name",
            "concept:name",
            "time:timestamp",
            "Unidade",
        ]
    ].copy()


# ---------------------------------------------------------------------------
# Modelo de referência e conformance checking
# ---------------------------------------------------------------------------

_PREPOSITION_SUFFIX_PATTERN = re.compile(
    r"\b(?:na|no|da|do|das|dos|em|para a|para o|para|pela|pelo)\s+(.+)",
    flags=re.IGNORECASE,
)


def normalizar_rotulos_modelo(net: Any, strip_prepositions: bool) -> None:
    """
    Normaliza os rótulos visíveis da rede de Petri.

    Quando STRIP_MODEL_PREPOSITIONS=true, preserva a regra usada na Prova de
    Conceito para extrair o trecho posterior a determinadas preposições. Em
    novas aplicações, esta regra deve ser revista e pode ser desativada.
    """
    for transition in net.transitions:
        if transition.label is None:
            continue

        rotulo = normalizar_texto(transition.label)

        if strip_prepositions:
            match = _PREPOSITION_SUFFIX_PATTERN.search(rotulo)
            if match:
                rotulo = match.group(1)

        transition.label = normalizar_atividade(rotulo)


def carregar_modelo_referencia(cfg: Config):
    """Lê o BPMN, converte-o para rede de Petri e normaliza seus rótulos."""
    validar_arquivo(cfg.bpmn_file, "Modelo BPMN de referência")

    logging.info("Carregando modelo BPMN de referência...")
    bpmn_graph = pm4py.read_bpmn(str(cfg.bpmn_file))
    net, initial_marking, final_marking = pm4py.convert_to_petri_net(bpmn_graph)

    normalizar_rotulos_modelo(
        net,
        strip_prepositions=cfg.strip_model_prepositions,
    )

    return net, initial_marking, final_marking


def executar_conformance(
    event_df: pd.DataFrame,
    net: Any,
    initial_marking: Any,
    final_marking: Any,
):
    """Executa alignments e calcula precision global baseada em alignments."""
    logging.info("Convertendo dados para event log...")
    event_log = pm4py.convert_to_event_log(event_df)

    logging.info("Executando alignments...")
    alignments = pm4py.conformance_diagnostics_alignments(
        event_log,
        net,
        initial_marking,
        final_marking,
    )

    logging.info("Calculando precision global...")
    precision = pm4py.precision_alignments(
        event_log,
        net,
        initial_marking,
        final_marking,
    )

    return event_log, alignments, float(precision)


# ---------------------------------------------------------------------------
# Interpretação dos alignments
# ---------------------------------------------------------------------------

def limpar_move(move: Any) -> Any | None:
    """Converte o marcador >> em ausência de movimento."""
    if move is None or move == ">>":
        return None
    return move


def nomes_transicoes_silenciosas(net: Any) -> set[str]:
    """Obtém nomes internos de transições sem rótulo visível."""
    return {
        str(t.name)
        for t in net.transitions
        if getattr(t, "label", None) is None
    }


def eh_movimento_silencioso(
    log_move: Any | None,
    model_move: Any | None,
    silent_names: set[str],
) -> bool:
    """Identifica movimentos de modelo associados a transições invisíveis."""
    if log_move is not None:
        return False

    if model_move is None:
        return True

    texto_modelo = normalizar_texto(model_move)
    texto_lower = texto_modelo.lower()

    return (
        texto_modelo in silent_names
        or texto_lower.startswith("tau")
        or texto_lower.startswith("skip")
    )


def classificar_movimento(
    log_move: Any | None,
    model_move: Any | None,
) -> str:
    """Classifica o par de alinhamento para a camada de visualização."""
    if log_move == model_move and log_move is not None:
        return "Sync Move (Sincronizado)"

    if log_move is not None and model_move is None:
        return "Log Move (Atividade Extra / Retrabalho)"

    if log_move is None and model_move is not None:
        return "Model Move (Atividade Omitida / Pulada)"

    # Em alignments usuais, uma substituição tende a ser representada como
    # movimentos separados. A categoria abaixo preserva o diagnóstico caso a
    # biblioteca devolva um par não sincronizado com ambos os lados preenchidos.
    return "Movimento não sincronizado (Substituição)"


def extrair_metricas_e_movimentos(
    event_log: Any,
    alignments: list[dict[str, Any]],
    precision: float,
    net: Any,
    fitness_alert_threshold: float,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Produz uma linha de métricas por trace e uma linha por movimento relevante.

    Fitness igual a 1 representa aderência completa do trace ao modelo.
    O limiar de alerta gerencial é mantido em campo separado e não redefine a
    fronteira de conformidade estrutural.
    """
    logging.info("Estruturando métricas e diagnósticos dos alignments...")

    metricas: list[dict[str, Any]] = []
    movimentos: list[dict[str, Any]] = []
    silent_names = nomes_transicoes_silenciosas(net)

    if len(event_log) != len(alignments):
        raise RuntimeError(
            "Quantidade de traces e de resultados de alignment não coincide."
        )

    for index, trace_alignment in enumerate(alignments):
        trace = event_log[index]
        case_id = normalizar_texto(
            trace.attributes.get("concept:name", f"trace_{index + 1}")
        )

        fitness = float(trace_alignment.get("fitness", 0.0))

        metricas.append(
            {
                "Processo": case_id,
                "fitness_escore": round(fitness, 4),
                "processo_conforme": 1 if fitness >= 1.0 else 0,
                "alerta_fitness": 1
                if fitness < fitness_alert_threshold
                else 0,
                "limiar_alerta_fitness": fitness_alert_threshold,
                "precision_global": round(precision, 4),
            }
        )

        alignment_steps = trace_alignment.get("alignment", [])
        event_idx = 0
        ordem = 1

        for step in alignment_steps:
            if not isinstance(step, (tuple, list)) or len(step) != 2:
                logging.warning(
                    "Passo de alignment não reconhecido no processo %s: %r",
                    case_id,
                    step,
                )
                continue

            raw_log_move, raw_model_move = step
            log_move = limpar_move(raw_log_move)
            model_move = limpar_move(raw_model_move)

            if eh_movimento_silencioso(
                log_move,
                model_move,
                silent_names,
            ):
                continue

            if log_move is None and model_move is None:
                continue

            unidade_original = ""

            # Sync moves e log moves consomem um evento real do trace.
            if log_move is not None:
                if event_idx < len(trace):
                    unidade_original = normalizar_texto(
                        trace[event_idx].get("Unidade", "")
                    )
                else:
                    logging.warning(
                        "Índice de evento excedeu o tamanho do trace %s.",
                        case_id,
                    )
                event_idx += 1

            tipo = classificar_movimento(log_move, model_move)

            movimentos.append(
                {
                    "Processo": case_id,
                    "Ordem": ordem,
                    "Tipo_Desvio": tipo,
                    "Atividade_Modelo": normalizar_texto(model_move),
                    "Atividade_Log": normalizar_texto(log_move),
                    "Unidade_Original": unidade_original,
                }
            )
            ordem += 1

    return pd.DataFrame(metricas), pd.DataFrame(movimentos)


# ---------------------------------------------------------------------------
# Exportação para a camada de visualização
# ---------------------------------------------------------------------------

def salvar_csv_atomico(df: pd.DataFrame, destino: Path) -> None:
    """
    Grava o CSV em arquivo temporário e o substitui ao final.

    Isso reduz o risco de o Power BI consumir um arquivo parcialmente escrito
    durante uma atualização automática.
    """
    destino.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".csv",
        prefix=f".{destino.stem}_",
        dir=destino.parent,
        delete=False,
        encoding="utf-8-sig",
        newline="",
    ) as tmp:
        temporario = Path(tmp.name)
        df.to_csv(tmp, index=False)

    temporario.replace(destino)


def gerar_estatisticas_desvios(df_movimentos: pd.DataFrame) -> pd.DataFrame:
    """Agrupa somente os movimentos não sincronizados para apoio visual."""
    colunas = ["Atividade_Alvo", "Tipo_Desvio", "Quantidade"]

    if df_movimentos.empty:
        return pd.DataFrame(columns=colunas)

    desvios = df_movimentos[
        df_movimentos["Tipo_Desvio"] != "Sync Move (Sincronizado)"
    ].copy()

    if desvios.empty:
        return pd.DataFrame(columns=colunas)

    def atividade_alvo(row: pd.Series) -> str:
        if row["Tipo_Desvio"].startswith("Model Move"):
            return row["Atividade_Modelo"]
        if row["Tipo_Desvio"].startswith("Log Move"):
            return row["Atividade_Log"]

        return (
            f"{row['Atividade_Log']} | {row['Atividade_Modelo']}"
        ).strip(" |")

    desvios["Atividade_Alvo"] = desvios.apply(atividade_alvo, axis=1)

    return (
        desvios.groupby(
            ["Atividade_Alvo", "Tipo_Desvio"],
            dropna=False,
        )
        .size()
        .reset_index(name="Quantidade")
        .sort_values(
            by=["Quantidade", "Atividade_Alvo"],
            ascending=[False, True],
        )
        .reset_index(drop=True)
    )


def exportar_resultados(
    output_dir: Path,
    df_metricas: pd.DataFrame,
    df_movimentos: pd.DataFrame,
) -> None:
    """Exporta os três arquivos estáveis utilizados pelo painel."""
    logging.info("Exportando resultados para %s...", output_dir)

    df_estatisticas = gerar_estatisticas_desvios(df_movimentos)

    salvar_csv_atomico(
        df_metricas,
        output_dir / "metricas_conformidade.csv",
    )
    salvar_csv_atomico(
        df_movimentos,
        output_dir / "desvios_conformidade.csv",
    )
    salvar_csv_atomico(
        df_estatisticas,
        output_dir / "estatisticas_desvios.csv",
    )

    logging.info("Arquivos de conformidade atualizados com sucesso.")


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

def main() -> int:
    """Orquestra a execução completa do conformance checking."""
    try:
        cfg = carregar_configuracao()

        validar_arquivo(cfg.mapping_file, "Arquivo de parametrização")
        validar_arquivo(cfg.bpmn_file, "Modelo BPMN de referência")

        df_bruto = carregar_dados(cfg)
        event_df = preparar_event_log_dataframe(cfg, df_bruto)

        if event_df.empty:
            logging.warning("Nenhum evento válido disponível para análise.")
            return 0

        net, initial_marking, final_marking = carregar_modelo_referencia(cfg)

        event_log, alignments, precision = executar_conformance(
            event_df,
            net,
            initial_marking,
            final_marking,
        )

        df_metricas, df_movimentos = extrair_metricas_e_movimentos(
            event_log,
            alignments,
            precision,
            net,
            cfg.fitness_alert_threshold,
        )

        exportar_resultados(
            cfg.output_dir,
            df_metricas,
            df_movimentos,
        )

        logging.info(
            "Análise concluída: %d processo(s), precision global %.4f.",
            len(df_metricas),
            precision,
        )
        return 0

    except KeyboardInterrupt:
        logging.warning("Execução interrompida pelo usuário.")
        return 130

    except Exception:
        logging.exception("Falha durante a análise de conformidade.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
