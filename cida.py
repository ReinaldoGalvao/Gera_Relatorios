import pandas as pd
from datetime import datetime, timedelta
from oracle_db import DBOracle

db = DBOracle()

def executar_query(inicio_mes, fim_mes, filial):
    query = f"""
        SELECT filial, SUM(soma1) + SUM(soma2) AS total
        FROM (
            SELECT f3_filial FILIAL, f3_cfo,
            CASE WHEN f3_cfo = '1102' or f3_cfo = '2102' THEN SUM(f3_valcont) ELSE 0 END AS soma1,
            CASE WHEN f3_cfo = '1403' or f3_cfo = '2403' THEN SUM(f3_valcont) ELSE 0 END AS soma2
            FROM sf3080 f3
            WHERE  f3_filial = '{filial}'
            AND f3_entrada >= '{inicio_mes}' AND f3_entrada <= '{fim_mes}'
            AND (f3_codrsef = ' ' OR f3_codrsef = '100')
            AND f3_dtcanc = ' ' AND (
                (f3_cfo = '1102' or f3_cfo = '2102') OR (f3_cfo = '1403' or f3_cfo = '2403')
            )
            AND f3.d_e_l_e_t_ = ' '
            GROUP BY f3_filial, f3_cfo
        ) t
        GROUP BY filial
        ORDER BY filial;
    """

    db.cursor.execute(query)
    resultado = db.cursor.fetchall()

    return resultado

def gerar_xlsx(primeiro_mes, primeiro_ano, ultimo_mes, ultimo_ano):
    df_final = pd.DataFrame()

    try:
        periodos = pd.date_range(start=f"{primeiro_ano}-{primeiro_mes}-01", end=f"{ultimo_ano}-{ultimo_mes}-01", freq='MS')

        for periodo in periodos:
            inicio_mes = periodo.strftime("%Y%m%d")
            fim_mes = (periodo + pd.offsets.MonthEnd(0)).strftime("%Y%m%d")

            df_mensal = pd.DataFrame()

            for filial in [f"{i:02}" for i in range(100)]:
                resultado_filial = executar_query(inicio_mes, fim_mes, filial)
                df_filial = pd.DataFrame(resultado_filial, columns=['filial', 'total'])
                df_mensal = pd.concat([df_mensal, df_filial])

            df_final = pd.concat([df_final, df_mensal])

    except Exception as err:
        print(err)

    finally:
        db.commit_db()
        db.close_db()

if __name__ == "__main__":
    gerar_xlsx(primeiro_mes='06', primeiro_ano='2023', ultimo_mes='10', ultimo_ano='2023')
