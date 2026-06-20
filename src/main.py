DATABASE_URL = (
    "postgresql+psycopg2://alison@localhost:5432/payroll_analytics"
)

QUANTIDADE_FUNCIONARIOS = 2000

from faker import Faker
from pathlib import Path
from sqlalchemy import create_engine, text
import pandas as pd
import random
import logging
from datetime import datetime
from time import perf_counter

fake = Faker("pt_BR")

BASE_DIR = Path(__file__).resolve().parent.parent

# ======================================
# CONEXÃO POSTGRESQL
# ======================================

engine = create_engine(DATABASE_URL)

# ======================================
# CRIA DIRETÓRIOS
# ======================================

DATA_DIR = BASE_DIR / "data"

DATA_DIR.mkdir(
    parents=True,
    exist_ok=True
)

LOG_DIR = BASE_DIR / "logs"

LOG_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ======================================
# CRIA DIRETÓRIO EXPORTAÇÃO
# ======================================
EXPORT_DIR = BASE_DIR / "exports"

EXPORT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

ARQUIVO_LOG = (
    LOG_DIR
    / f"pipeline_{datetime.now():%Y%m%d}.log"
)

logging.basicConfig(
    level=logging.INFO,
    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(message)s"
    ),
    handlers=[
        logging.FileHandler(
            ARQUIVO_LOG,
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

def validar_dataframe(
        df,
        nome_tabela
):

    if df.empty:
        raise ValueError(
            f"{nome_tabela} está vazio."
        )

    logging.info(
        f"{nome_tabela}: "
        f"{len(df)} registros"
    )

    logging.info(
        f"{nome_tabela}: "
        f"{len(df.columns)} colunas"
    )

    return True

# ======================================
# GERA FUNCIONÁRIOS
# ======================================

def gerar_funcionarios():

    departamentos = [
        "RH",
        "Financeiro",
        "TI",
        "Operações",
        "Comercial",
        "Diretoria"
    ]

    cargos = {
        "RH": [
            "Assistente RH",
            "Analista RH",
            "Coordenador RH"
        ],
        "Financeiro": [
            "Assistente Financeiro",
            "Analista Financeiro",
            "Coordenador Financeiro"
        ],
        "TI": [
            "Desenvolvedor",
            "Analista de Sistemas",
            "Especialista TI"
        ],
        "Operações": [
            "Assistente Operacional",
            "Analista Operacional",
            "Supervisor Operacional"
        ],
        "Comercial": [
            "Executivo de Vendas",
            "Coordenador Comercial"
        ],
        "Diretoria": [
            "Diretor"
        ]
    }

    dados = []

    for id_funcionario in range(
            1,
            QUANTIDADE_FUNCIONARIOS + 1
    ):

        sexo = random.choice(["M", "F"])

        if sexo == "M":
            nome = (
                fake.first_name_male()
                + " "
                + fake.last_name()
            )
        else:
            nome = (
                fake.first_name_female()
                + " "
                + fake.last_name()
            )

        departamento = random.choices(
            population=[
                "Operações",
                "Comercial",
                "TI",
                "Financeiro",
                "RH",
                "Diretoria"
            ],
            weights=[
                45,
                20,
                15,
                10,
                7,
                3
            ],
            k=1
        )[0]
        # Distribuição aproximada para 2000 colaboradores:
        # Operações: 900
        # Comercial: 400
        # TI: 300
        # Financeiro: 200
        # RH: 140
        # Diretoria: 60

        cargo = random.choice(
            cargos[departamento]
        )

        dados.append({
            "id_funcionario": id_funcionario,
            "nome": nome,
            "sexo": sexo,
            "idade": random.randint(22, 60),
            "data_admissao": fake.date_between(
                start_date="-8y",
                end_date="today"
            ),
            "departamento": departamento,
            "cargo": cargo,
            "status": "Ativo"
        })

    df = pd.DataFrame(dados)

    caminho = (
        BASE_DIR
        / "data"
        / "funcionarios.xlsx"
    )

    df.to_excel(
        caminho,
        index=False
    )

    validar_dataframe(
        df,
        "Funcionários"
    )

    logging.info(
        "Arquivo funcionarios.xlsx gerado"
    )

    return df


# ======================================
# GERA SALÁRIOS
# ======================================

def gerar_salarios(df_funcionarios):

    faixas = {
        "Assistente RH": (3000, 5000),
        "Analista RH": (5000, 8000),
        "Coordenador RH": (8000, 12000),

        "Assistente Financeiro": (3500, 5500),
        "Analista Financeiro": (6000, 9000),
        "Coordenador Financeiro": (10000, 14000),

        "Desenvolvedor": (6000, 10000),
        "Analista de Sistemas": (7000, 11000),
        "Especialista TI": (10000, 15000),

        "Assistente Operacional": (2500, 4000),
        "Analista Operacional": (4000, 7000),
        "Supervisor Operacional": (7000, 10000),

        "Executivo de Vendas": (4000, 9000),
        "Coordenador Comercial": (9000, 14000),

        "Diretor": (20000, 35000)
    }

    salarios = []

    for _, linha in df_funcionarios.iterrows():

        minimo, maximo = faixas[
            linha["cargo"]
        ]

        salario = round(
            random.uniform(minimo, maximo),
            2
        )

        bonus = round(
            salario * random.uniform(0.05, 0.20),
            2
        )

        salarios.append({
            "id_funcionario": linha["id_funcionario"],
            "salario_base": salario,
            "bonus": bonus
        })

    df = pd.DataFrame(salarios)

    caminho = (
        BASE_DIR
        / "data"
        / "salarios.xlsx"
    )

    df.to_excel(
        caminho,
        index=False
    )

    validar_dataframe(
        df,
        "Salários"
    )

    logging.info(
        "Arquivo salarios.xlsx gerado"
    )

    return df


# ======================================
# GERA BENEFÍCIOS
# ======================================

def gerar_beneficios(df_funcionarios):

    beneficios = []

    for _, linha in df_funcionarios.iterrows():

        beneficios.append({
            "id_funcionario": linha["id_funcionario"],
            "vale_refeicao": random.choice([800, 1000, 1200]),
            "vale_transporte": random.choice([200, 300, 400]),
            "plano_saude": random.choice([350, 450, 600])
        })

    df = pd.DataFrame(beneficios)

    caminho = (
        BASE_DIR
        / "data"
        / "beneficios.xlsx"
    )

    df.to_excel(
        caminho,
        index=False
    )

    validar_dataframe(
        df,
        "Benefícios"
    )

    logging.info(
        "Arquivo beneficios.xlsx gerado"
    )

    return df


# ======================================
# CARGA POSTGRESQL
# ======================================

def carregar_postgres(
        funcionarios,
        salarios,
        beneficios
):

    with engine.begin() as conexao:

        conexao.execute(
            text("""
                TRUNCATE TABLE payroll_beneficios
                RESTART IDENTITY CASCADE;
            """)
        )

        conexao.execute(
            text("""
                TRUNCATE TABLE payroll_salarios
                RESTART IDENTITY CASCADE;
            """)
        )

        conexao.execute(
            text("""
                TRUNCATE TABLE payroll_funcionarios
                RESTART IDENTITY CASCADE;
            """)
        )

    validar_dataframe(
        funcionarios,
        "Funcionários"
    )

    validar_dataframe(
        salarios,
        "Salários"
    )

    validar_dataframe(
        beneficios,
        "Benefícios"
    )

    funcionarios.to_sql(
        "payroll_funcionarios",
        engine,
        if_exists="append",
        index=False
    )

    salarios.to_sql(
        "payroll_salarios",
        engine,
        if_exists="append",
        index=False
    )

    beneficios.to_sql(
        "payroll_beneficios",
        engine,
        if_exists="append",
        index=False
    )

    with engine.connect() as conexao:

        qtd_funcionarios = conexao.execute(
            text(
                """
                SELECT COUNT(*)
                FROM payroll_funcionarios
                """
            )
        ).scalar()

        qtd_salarios = conexao.execute(
            text(
                """
                SELECT COUNT(*)
                FROM payroll_salarios
                """
            )
        ).scalar()

        qtd_beneficios = conexao.execute(
            text(
                """
                SELECT COUNT(*)
                FROM payroll_beneficios
                """
            )
        ).scalar()

    logging.info(
        f"Funcionários carregados: {qtd_funcionarios}"
    )

    logging.info(
        f"Salários carregados: {qtd_salarios}"
    )

    logging.info(
        f"Benefícios carregados: {qtd_beneficios}"
    )


# ======================================
# EXPORTAÇÃO TABLEAU
# ======================================

def exportar_views_tableau():

    views = {
        "vw_payroll_resumo": "payroll_resumo",
        "vw_custo_departamento": "custo_departamento",
        "vw_ranking_salarios": "ranking_salarios",
        "vw_diversidade": "diversidade",
        "vw_faixa_etaria": "faixa_etaria",
        "vw_tempo_empresa_faixa": "tempo_empresa",
        "vw_headcount_departamento": "headcount_departamento",
        "vw_salario_medio_departamento": "salario_medio_departamento",
        "vw_folha_departamento": "folha_departamento",
        "vw_faixa_salarial": "faixa_salarial",
    }

    tableau_dir = EXPORT_DIR / "tableau"

    tableau_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    with engine.connect() as conexao:

        for view, nome_arquivo in views.items():

            df = pd.read_sql(
                f"SELECT * FROM {view}",
                conexao
            )

            caminho_excel = EXPORT_DIR / f"{nome_arquivo}.xlsx"

            df.to_excel(
                caminho_excel,
                index=False
            )

            caminho_csv = tableau_dir / f"{nome_arquivo}.csv"

            df.to_csv(
                caminho_csv,
                index=False,
                sep=",",
                encoding="utf-8"
            )

            logging.info(
                f"View exportada: {nome_arquivo} ({len(df)} registros)"
            )


# ======================================
# MAIN
# ======================================

def main():
    inicio = perf_counter()

    print("=" * 50)
    print("PAYROLL ANALYTICS PIPELINE")
    print("=" * 50)

    funcionarios = gerar_funcionarios()

    salarios = gerar_salarios(
        funcionarios
    )

    beneficios = gerar_beneficios(
        funcionarios
    )

    carregar_postgres(
        funcionarios,
        salarios,
        beneficios
    )
    exportar_views_tableau()

    fim = perf_counter()

    logging.info(
        f"Tempo total: {fim - inicio:.2f} segundos"
    )

    logging.info(
        "Arquivos Excel exportados para pasta exports"
    )

    logging.info(
        "Arquivos CSV exportados para exports/tableau"
    )

    print()
    print("🚀 Pipeline executado com sucesso!")


if __name__ == "__main__":

    try:

        main()

    except Exception as erro:

        logging.exception(
            "Erro durante execução"
        )

        raise erro