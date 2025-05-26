import os
import pyexcel as p
from dotenv import load_dotenv
import win32com.client


def letra_para_indice(letra):
    indice = 0
    for c in letra:
        indice = indice * 26 + (ord(c.upper()) - ord('A') + 1)
    return indice - 1


def converter_xls_para_ods(pasta_entrada, pasta_saida):
    os.makedirs(pasta_saida, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # rodar Excel em background

    arquivos_xls = [f for f in os.listdir(pasta_entrada) if f.lower().endswith('.xls')]

    if not arquivos_xls:
        print("Nenhum arquivo .xls encontrado em", pasta_entrada)
        excel.Quit()
        return

    for arquivo in arquivos_xls:
        try:
            caminho_entrada = os.path.abspath(os.path.join(pasta_entrada, arquivo))
            nome_sem_ext = os.path.splitext(arquivo)[0]
            caminho_saida = os.path.abspath(os.path.join(pasta_saida, f"{nome_sem_ext}.ods"))
            print(f"Convertendo {arquivo} para {nome_sem_ext}.ods ...")

            wb = excel.Workbooks.Open(caminho_entrada)
            # FileFormat=60 significa ODS
            wb.SaveAs(caminho_saida, FileFormat=60)
            wb.Close()
            print(f"Salvo: {caminho_saida}")
        except Exception as e:
            print(f"Erro ao converter {arquivo}: {e}")

    excel.Quit()


def extrair_colunas_e_unir(pasta_ods, colunas_letras, arquivo_saida):
    os.makedirs(os.path.dirname(arquivo_saida) or '.', exist_ok=True)
    indices = [letra_para_indice(c) for c in colunas_letras]

    planilha_unica = p.Sheet()

    arquivos_ods = [f for f in os.listdir(pasta_ods) if f.lower().endswith('.ods')]
    if not arquivos_ods:
        print("Nenhum arquivo .ods encontrado em", pasta_ods)
        return

    for arquivo in arquivos_ods:
        caminho = os.path.join(pasta_ods, arquivo)
        print(f"Lendo {arquivo} ...")
        try:
            sheet = p.get_sheet(file_name=caminho)
            for linha in sheet:
                valores = [linha[i] if i < len(linha) else None for i in indices]
                planilha_unica.row += [valores]
        except Exception as e:
            print(f"Erro ao ler {arquivo}: {e}")

    planilha_unica.save_as(arquivo_saida)
    print(f"\nArquivo final salvo em: {arquivo_saida}")


def main():
    # Carregar variáveis do .env
    load_dotenv()
    colunas_env = os.getenv("COLUNAS")
    pasta_xls = os.getenv("PASTA_XLS") or "pasta_xls"
    pasta_ods = os.getenv("PASTA_ODS") or "pasta_ods"
    arquivo_saida = os.getenv("ARQUIVO_SAIDA") or "saida_unica.ods"

    if not colunas_env:
        print("Variável COLUNAS não definida no arquivo .env. Exemplo: COLUNAS=A,F,J,T,AA,AB,AD")
        return

    colunas_letras = [c.strip() for c in colunas_env.split(",")]

    os.makedirs(pasta_xls, exist_ok=True)
    os.makedirs(pasta_ods, exist_ok=True)

    print("=== Convertendo arquivos XLS para ODS ===")
    converter_xls_para_ods(pasta_xls, pasta_ods)

    print("\n=== Extraindo colunas e unindo arquivos ODS ===")
    extrair_colunas_e_unir(pasta_ods, colunas_letras, arquivo_saida)

    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    main()
