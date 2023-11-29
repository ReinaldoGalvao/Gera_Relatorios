import pandas as pd
from oracle_db import DBOracle
import shutil
import sys
#print(sys.path)

print("""
      
                                                                                                      
                                                                                                
        CCCCCCCCCCCCC                                            tttt                           
     CCC::::::::::::C                                         ttt:::t                           
   CC:::::::::::::::C                                         t:::::t                           
  C:::::CCCCCCCC::::C                                         t:::::t                           
 C:::::C       CCCCCCuuuuuu    uuuuuu      ssssssssss   ttttttt:::::ttttttt       ooooooooooo   
C:::::C              u::::u    u::::u    ss::::::::::s  t:::::::::::::::::t     oo:::::::::::oo 
C:::::C              u::::u    u::::u  ss:::::::::::::s t:::::::::::::::::t    o:::::::::::::::o
C:::::C              u::::u    u::::u  s::::::ssss:::::stttttt:::::::tttttt    o:::::ooooo:::::o
C:::::C              u::::u    u::::u   s:::::s  ssssss       t:::::t          o::::o     o::::o
C:::::C              u::::u    u::::u     s::::::s            t:::::t          o::::o     o::::o
C:::::C              u::::u    u::::u        s::::::s         t:::::t          o::::o     o::::o
 C:::::C       CCCCCCu:::::uuuu:::::u  ssssss   s:::::s       t:::::t    tttttto::::o     o::::o
  C:::::CCCCCCCC::::Cu:::::::::::::::uus:::::ssss::::::s      t::::::tttt:::::to:::::ooooo:::::o
   CC:::::::::::::::C u:::::::::::::::us::::::::::::::s       tt::::::::::::::to:::::::::::::::o
     CCC::::::::::::C  uu::::::::uu:::u s:::::::::::ss          tt:::::::::::tt oo:::::::::::oo 
        CCCCCCCCCCCCC    uuuuuuuu  uuuu  sssssssssss              ttttttttttt     ooooooooooo   
                                                                                                
                                                                                                
                                                                                                
      
      """)

db = DBOracle()

def criar_arquivo(filial, estado):
    # Query SQL
    query = f"""
        SELECT b1_cod CODIGO, b1_codbar COD_BAR, b1_prv1 PMC, b1_desc DESCRICAO,
        (SELECT z00_prv1 FROM z00010 WHERE z00_filial = b1_filial AND b1_cod = z00_codprd AND z00_uf = '{estado}' AND d_e_l_e_t_ = ' ') PRECO,
        (SELECT z00_prv1 * (1-(SELECT (acp_perdes/100)
        FROM acp010 acp
        WHERE acp_filial = '{filial}'  AND acp.d_e_l_e_t_ = ' ' AND acp_codreg = '000001'
        AND acp_codpro = z00_codprd))
        FROM z00010 WHERE z00_filial = ' ' AND z00_uf = '{estado}' AND d_e_l_e_t_ = ' ' AND z00_codprd = b1_cod) PREÇO_LIQUIDO,
        (SELECT acp_perdes FROM acp010 cp WHERE acp_filial = '{filial}' AND acp_codreg = '000001'  AND d_e_l_e_t_ = ' ' AND cp.acp_codpro = b1_cod) PERC_DESC,
        (SELECT b2_cm1 FROM sb2010 b2 WHERE b2_filial = '{filial}' AND b2_cod = b1_cod AND b2_local = '01' AND b2.d_e_l_e_t_ = ' ') CUSTO
        FROM sb1010 b1 WHERE b1_filial = ' ' /* and (B1_MSBLQL != '1' AND B1_SUSPENC != '1' AND B1_SUSPVEN != 'T')*/ AND b1.d_e_l_e_t_ = ' '
    """

    db.cursor.execute(query)

    resultado = db.cursor.fetchall()

    df = pd.DataFrame(resultado, columns=['CODIGO', 'COD_BAR', 'PMC', 'DESCRICAO', 'PRECO', 'PRECO_LIQUIDO', 'PERC_DESC', 'CUSTO'])

    caminho_arquivo = f'Preço e Custo F{filial}.xlsx'

    df.to_excel(caminho_arquivo, index=False)

    # Copiar o arquivo para publico_geral
    destino = r'\\172.20.153.59\publico_geral\Tecnologia da Informacao\Camila'
    shutil.copy(caminho_arquivo, destino)

    # Print loja feita
    print(f'Loja {filial} feita com sucesso!')

def processar_lojas(filiais, estado):
    for filial in filiais:
        criar_arquivo(filial, estado)

# Lista de filiais do Rio de Janeiro
filiais_rj = ['02', '04', '05', '06', '08', '10', '14', '15', '18', '25', '29', '32', '33', '35', '38', '40', '41', '43', '50', '51', '52', '66', '67', '69', '73', '78', '79']

# Lista de filiais de São Paulo
filiais_sp = ['54', '55', '56', '58', '64', '65']

# Processar lojas do Rio de Janeiro
processar_lojas(filiais_rj, 'RJ')

# Processar lojas de São Paulo
processar_lojas(filiais_sp, 'SP')

db.close_db()

