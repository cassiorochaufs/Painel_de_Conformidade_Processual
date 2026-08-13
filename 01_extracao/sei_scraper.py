"""
Coletor de históricos de tramitação do SEI via Selenium.

Segurança:
- Não armazena credenciais no código-fonte.
- Credenciais podem ser informadas interativamente ou via variáveis de ambiente.
- O arquivo .env local não deve ser versionado.
- Logs não registram senhas, cookies, tokens ou URLs de sessão.

Observação:
Os seletores HTML refletem a interface do SEI utilizada na Prova de Conceito
e podem exigir ajustes em outras versões ou customizações do sistema.
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
import re
import sys
from dataclasses import dataclass
from getpass import getpass
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from sqlalchemy import create_engine


@dataclass(frozen=True)
class Config:
    """Parâmetros da execução."""

    sei_base_url: str
    usuario: str
    senha: str
    data_inicio: str
    data_fim: str
    output_dir: Path
    timeout: int = 30
    headless: bool = False
    database_url: str | None = None
    database_table: str = "historico"


def configurar_logging(verbose: bool = False) -> None:
    """Configura mensagens de execução sem expor dados sensíveis."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def validar_data(valor: str) -> str:
    """Valida DD/MM/AAAA."""
    try:
        return dt.datetime.strptime(valor, "%d/%m/%Y").strftime("%d/%m/%Y")
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Data inválida: {valor!r}. Use DD/MM/AAAA."
        ) from exc


def solicitar_data(mensagem: str) -> str:
    """Solicita uma data até receber um valor válido."""
    while True:
        valor = input(f"{mensagem} (DD/MM/AAAA): ").strip()
        try:
            return validar_data(valor)
        except argparse.ArgumentTypeError as exc:
            print(exc)


def carregar_configuracao() -> Config:
    """Carrega argumentos, variáveis de ambiente e entradas interativas."""
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Extrai históricos de tramitação do SEI via Selenium."
    )
    parser.add_argument("--inicio", type=validar_data, help="Data inicial DD/MM/AAAA.")
    parser.add_argument("--fim", type=validar_data, help="Data final DD/MM/AAAA.")
    parser.add_argument(
        "--saida",
        default=os.getenv("OUTPUT_DIR", "./dados"),
        help="Diretório de saída. Padrão: ./dados",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=int(os.getenv("SELENIUM_TIMEOUT", "30")),
        help="Tempo máximo das esperas Selenium.",
    )
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    configurar_logging(args.verbose)

    sei_base_url = os.getenv("SEI_BASE_URL", "").strip()
    if not sei_base_url:
        sei_base_url = input(
            "URL base do SEI (ex.: https://sei.instituicao.br/): "
        ).strip()
    if not sei_base_url:
        raise SystemExit("A URL base do SEI é obrigatória.")
    if not sei_base_url.endswith("/"):
        sei_base_url += "/"

    usuario = os.getenv("SEI_USERNAME", "").strip()
    if not usuario:
        usuario = input("Usuário do SEI: ").strip()

    senha = os.getenv("SEI_PASSWORD", "")
    if not senha:
        senha = getpass("Senha do SEI: ")

    data_inicio = args.inicio or solicitar_data("Data de início")
    data_fim = args.fim or solicitar_data("Data de fim")

    if (
        dt.datetime.strptime(data_inicio, "%d/%m/%Y")
        > dt.datetime.strptime(data_fim, "%d/%m/%Y")
    ):
        raise SystemExit("A data inicial não pode ser posterior à data final.")

    output_dir = Path(args.saida).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    return Config(
        sei_base_url=sei_base_url,
        usuario=usuario,
        senha=senha,
        data_inicio=data_inicio,
        data_fim=data_fim,
        output_dir=output_dir,
        timeout=args.timeout,
        headless=args.headless,
        database_url=os.getenv("DATABASE_URL") or None,
        database_table=os.getenv("DATABASE_TABLE", "historico").strip() or "historico",
    )


def criar_driver(headless: bool) -> webdriver.Chrome:
    """Inicializa o Chrome usando o Selenium Manager."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    return webdriver.Chrome(options=options)


def fazer_login(driver: webdriver.Chrome, wait: WebDriverWait, cfg: Config) -> None:
    """Autentica o usuário no SEI."""
    logging.info("Acessando o SEI...")
    driver.get(cfg.sei_base_url)

    usuario = wait.until(EC.presence_of_element_located((By.ID, "txtUsuario")))
    usuario.clear()
    usuario.send_keys(cfg.usuario)

    senha = driver.find_element(By.ID, "pwdSenha")
    senha.clear()
    senha.send_keys(cfg.senha)

    driver.find_element(By.ID, "Acessar").click()

    wait.until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))
    logging.info("Autenticação concluída.")


def abrir_pesquisa_avancada(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    data_inicio: str,
    data_fim: str,
) -> None:
    """Abre a pesquisa avançada e aplica o período informado."""
    caixa = wait.until(EC.element_to_be_clickable((By.ID, "txtPesquisaRapida")))
    caixa.send_keys("\n")

    wait.until(EC.presence_of_element_located((By.ID, "fldPesquisarEm")))

    wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "label.infraRadioLabel[for='optProcessos']")
        )
    ).click()

    campo_inicio = wait.until(EC.element_to_be_clickable((By.ID, "txtDataInicio")))
    campo_inicio.clear()
    campo_inicio.send_keys(data_inicio)

    campo_fim = wait.until(EC.element_to_be_clickable((By.ID, "txtDataFim")))
    campo_fim.clear()
    campo_fim.send_keys(data_fim)

    wait.until(EC.element_to_be_clickable((By.ID, "sbmPesquisar"))).click()
    wait.until(EC.presence_of_element_located((By.ID, "conteudo")))

    logging.info("Pesquisa executada para %s a %s.", data_inicio, data_fim)


def coletar_lista_processos(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
) -> list[dict[str, str]]:
    """Percorre a paginação dos resultados e devolve os processos encontrados."""
    processos: list[dict[str, str]] = []
    urls_vistas: set[str] = set()
    pagina = 1

    while True:
        logging.info("Lendo página %d dos resultados...", pagina)

        try:
            wait.until(
                EC.presence_of_element_located(
                    (By.CLASS_NAME, "pesquisaTituloEsquerda")
                )
            )
        except TimeoutException:
            if pagina == 1:
                logging.warning("Nenhum processo encontrado.")
            break

        for container in driver.find_elements(
            By.CLASS_NAME, "pesquisaTituloEsquerda"
        ):
            try:
                link = container.find_element(By.TAG_NAME, "a")
                url = (link.get_attribute("href") or "").strip()

                numero = container.find_element(
                    By.CLASS_NAME, "protocoloNormal"
                ).text.strip()

                span = container.find_element(By.TAG_NAME, "span")
                tipo = span.text.replace("Nº", "").strip()

                if url and url not in urls_vistas:
                    processos.append(
                        {"url": url, "tipo": tipo, "numero": numero}
                    )
                    urls_vistas.add(url)
            except NoSuchElementException:
                continue

        try:
            atual = driver.find_element(By.CLASS_NAME, "pesquisaPaginaSelecionada")
            proxima = atual.find_element(By.XPATH, "./following-sibling::a[1]")
            if not proxima.text.strip().isdigit():
                break
            pagina += 1
            proxima.click()
            wait.until(EC.staleness_of(atual))
        except NoSuchElementException:
            break

    logging.info("Total de processos encontrados: %d.", len(processos))
    return processos


def limpar_texto(valor: str) -> str:
    """Normaliza espaços sem alterar o conteúdo do registro."""
    return re.sub(r"\s+", " ", valor or "").strip()


def extrair_historico_processo(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    processo: dict[str, str],
    base_url: str,
) -> list[dict[str, str]]:
    """Extrai todas as páginas do histórico de um processo."""
    registros: list[dict[str, str]] = []
    url_absoluta = urljoin(base_url, processo["url"])

    driver.switch_to.new_window("tab")
    driver.get(url_absoluta)

    try:
        wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrArvore"))
        )

        marcador = "Sem marcador"
        try:
            elemento = driver.find_element(By.CSS_SELECTOR, "[id^='iconMC']")
            marcador = limpar_texto(
                elemento.get_attribute("title") or ""
            ) or "Sem marcador"
        except NoSuchElementException:
            pass

        wait.until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "#divConsultarAndamento a[onclick='consultarAndamento();']",
                )
            )
        ).click()

        driver.switch_to.default_content()

        wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao"))
        )
        wait.until(EC.presence_of_element_located((By.ID, "tblHistorico")))

        pagina = 1
        while True:
            tabela = driver.find_element(By.ID, "tblHistorico")

            for linha in tabela.find_elements(By.TAG_NAME, "tr")[1:]:
                colunas = linha.find_elements(By.TAG_NAME, "td")
                if not colunas:
                    continue

                valores = [limpar_texto(col.text) for col in colunas]

                registros.append(
                    {
                        "Processo": processo["numero"],
                        "Tipo": processo["tipo"],
                        "Marcador": marcador,
                        "Data": valores[0] if len(valores) > 0 else "",
                        "Unidade": valores[1] if len(valores) > 1 else "",
                        "Descricao": valores[2] if len(valores) > 2 else "",
                        "Obs": valores[3] if len(valores) > 3 else "",
                    }
                )

            try:
                proxima = driver.find_element(
                    By.CSS_SELECTOR, "img[title='Próxima Página']"
                )
            except NoSuchElementException:
                break

            html_anterior = tabela.get_attribute("innerHTML")
            proxima.click()
            pagina += 1

            wait.until(
                lambda d: d.find_element(
                    By.ID, "tblHistorico"
                ).get_attribute("innerHTML")
                != html_anterior
            )

        logging.debug(
            "Processo %s: %d registro(s) extraído(s) em %d página(s).",
            processo["numero"],
            len(registros),
            pagina,
        )
        return registros

    finally:
        driver.switch_to.default_content()


def salvar_screenshot_erro(
    driver: webdriver.Chrome,
    diretorio: Path,
    indice: int,
) -> None:
    """Salva screenshot de diagnóstico sem usar dados do processo no nome."""
    pasta = diretorio / "logs"
    pasta.mkdir(parents=True, exist_ok=True)
    caminho = pasta / f"erro_processo_{indice:04d}.png"

    try:
        driver.save_screenshot(str(caminho))
        logging.error("Screenshot de diagnóstico salvo em %s.", caminho)
    except WebDriverException:
        logging.exception("Não foi possível salvar o screenshot de erro.")


def preparar_dataframe(registros: Iterable[dict[str, str]]) -> pd.DataFrame:
    """Monta o DataFrame final e remove duplicidades no lote."""
    colunas = [
        "Processo",
        "Tipo",
        "Marcador",
        "Data",
        "Unidade",
        "Descricao",
        "Obs",
    ]
    df = pd.DataFrame(registros, columns=colunas)

    if df.empty:
        return df

    return df.drop_duplicates(
        subset=["Processo", "Data", "Descricao"],
        keep="first",
    ).reset_index(drop=True)


def salvar_csv(df: pd.DataFrame, diretorio: Path) -> Path:
    """Salva CSV UTF-8 com BOM para facilitar consumo por Excel/Power BI."""
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = diretorio / f"historico_sei_{timestamp}.csv"
    df.to_csv(caminho, index=False, encoding="utf-8-sig")
    logging.info("CSV salvo em %s.", caminho)
    return caminho


def salvar_banco_opcional(
    df: pd.DataFrame,
    database_url: str | None,
    tabela: str,
) -> None:
    """
    Persiste no banco somente se DATABASE_URL estiver configurada.

    A URL pode apontar para qualquer SGBD suportado pelo SQLAlchemy e pelo driver
    instalado no ambiente. Para impedir duplicidades entre execuções, recomenda-se
    criar no banco uma restrição ou índice único adequado ao modelo local.
    """
    if not database_url:
        logging.info("DATABASE_URL não configurada. Banco ignorado.")
        return

    logging.info("Gravando dados na tabela %s...", tabela)
    engine = create_engine(database_url, pool_pre_ping=True)

    try:
        df.to_sql(
            tabela,
            con=engine,
            if_exists="append",
            index=False,
            method="multi",
            chunksize=1000,
        )
        logging.info("Persistência em banco concluída.")
    finally:
        engine.dispose()


def main() -> int:
    """Fluxo principal da extração."""
    cfg = carregar_configuracao()
    driver: webdriver.Chrome | None = None
    registros_finais: list[dict[str, str]] = []

    try:
        driver = criar_driver(cfg.headless)
        wait = WebDriverWait(driver, cfg.timeout)

        fazer_login(driver, wait, cfg)
        abrir_pesquisa_avancada(
            driver,
            wait,
            cfg.data_inicio,
            cfg.data_fim,
        )

        processos = coletar_lista_processos(driver, wait)
        if not processos:
            return 0

        janela_principal = driver.current_window_handle

        for indice, processo in enumerate(processos, start=1):
            logging.info(
                "Processando %s (%d/%d)...",
                processo["numero"],
                indice,
                len(processos),
            )

            try:
                registros_finais.extend(
                    extrair_historico_processo(
                        driver,
                        wait,
                        processo,
                        cfg.sei_base_url,
                    )
                )
            except (NoSuchElementException, TimeoutException, WebDriverException):
                logging.exception(
                    "Falha ao processar o item %d de %d.",
                    indice,
                    len(processos),
                )
                salvar_screenshot_erro(driver, cfg.output_dir, indice)
            finally:
                try:
                    if driver.current_window_handle != janela_principal:
                        driver.close()
                except WebDriverException:
                    pass

                driver.switch_to.window(janela_principal)

        df = preparar_dataframe(registros_finais)
        if df.empty:
            logging.warning("Nenhum registro de histórico foi extraído.")
            return 0

        salvar_csv(df, cfg.output_dir)
        salvar_banco_opcional(
            df,
            cfg.database_url,
            cfg.database_table,
        )

        logging.info(
            "Extração concluída: %d registro(s) após deduplicação.",
            len(df),
        )
        return 0

    except KeyboardInterrupt:
        logging.warning("Execução interrompida pelo usuário.")
        return 130
    except Exception:
        logging.exception("Falha não tratada durante a execução.")
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except WebDriverException:
                pass


if __name__ == "__main__":
    sys.exit(main())
