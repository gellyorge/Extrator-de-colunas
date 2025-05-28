import os
from dotenv import load_dotenv
import win32com.client
import pyexcel as p

load_dotenv()

# Pega o caminho raiz da aplicação (diretório onde está o script principal)
raiz_app = os.path.dirname(os.path.abspath(__file__))

def caminho_absoluto_relativo(caminho_relativo):
    """Transforma um caminho relativo em absoluto baseado na raiz do app"""
    if not caminho_relativo:
        return None
    return os.path.abspath(os.path.join(raiz_app, caminho_relativo))

def converter_xls_para_ods(pasta_origem, pasta_destino):
    extensoes_validas = (".xls", ".xlsx")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    print(f"Iniciando conversão em: {pasta_origem}")
    for root, _, files in os.walk(pasta_origem):
        subpasta_relativa = os.path.relpath(root, pasta_origem)
        destino_atual = os.path.join(pasta_destino, subpasta_relativa)
        os.makedirs(destino_atual, exist_ok=True)

        for arquivo in files:
            if not arquivo.lower().endswith(extensoes_validas):
                print(f"Ignorando (formato não suportado): {arquivo}")
                continue  # pula esse arquivo

            caminho_entrada = os.path.join(root, arquivo)
            nome_sem_ext = os.path.splitext(arquivo)[0]
            caminho_saida = os.path.join(destino_atual, f"{nome_sem_ext}.ods")

            print(f"Convertendo: {caminho_entrada} → {caminho_saida}")
            try:
                wb = excel.Workbooks.Open(caminho_entrada)
                wb.SaveAs(caminho_saida, FileFormat=60)  # 60 = ODS
                wb.Close()
                print("Conversão concluída")
            except Exception as e:
                print(f"Erro ao converter {arquivo}: {e}")

    excel.Quit()
    print("Processo finalizado.")


def filtrar_colunas_ods(pasta_ods, pasta_filtrados, indices_colunas, ordem_colunas=None):
    for root, _, arquivos in os.walk(pasta_ods):
        subpasta_relativa = os.path.relpath(root, pasta_ods)
        destino_atual = os.path.join(pasta_filtrados, subpasta_relativa)
        os.makedirs(destino_atual, exist_ok=True)

        for arquivo in arquivos:
            if arquivo.lower().endswith(".ods") and not arquivo.lower().endswith("_filtrado.ods"):
                caminho_entrada = os.path.join(root, arquivo)
                nome_sem_ext = os.path.splitext(arquivo)[0]
                caminho_saida = os.path.join(destino_atual, f"{nome_sem_ext}_filtrado.ods")

                try:
                    print(f"Filtrando colunas em: {arquivo}")
                    sheet = p.get_sheet(file_name=caminho_entrada)
                    planilha_filtrada = p.Sheet()

                    for linha in sheet:
                        # Primeiro pega as colunas de interesse na ordem normal
                        colunas_extraidas = []
                        for i in indices_colunas:
                            if i < len(linha):
                                colunas_extraidas.append(linha[i])
                            else:
                                colunas_extraidas.append(None)
                        
                        # Agora rearranja conforme ordem_colunas (que é uma lista de índices referente à colunas_extraidas)
                        if ordem_colunas:
                            nova_linha = []
                            for idx in ordem_colunas:
                                if idx < len(colunas_extraidas):
                                    nova_linha.append(colunas_extraidas[idx])
                                else:
                                    nova_linha.append(None)
                        else:
                            nova_linha = colunas_extraidas

                        planilha_filtrada.row += [nova_linha]

                    planilha_filtrada.save_as(caminho_saida)
                    print(f"Salvo: {caminho_saida}")
                except Exception as e:
                    print(f"Erro ao filtrar {arquivo}: {e}")



if __name__ == "__main__":
    pasta_xls_rel = os.getenv("PASTA_XLS")
    pasta_ods_rel = os.getenv("PASTA_ODS")
    pasta_filtrados_rel = os.getenv("PASTA_FILTRADOS")
    indices_str = os.getenv("INDICES", "")
    ordem_str = os.getenv("ORDEM", "")

    pasta_xls = caminho_absoluto_relativo(pasta_xls_rel)
    pasta_ods = caminho_absoluto_relativo(pasta_ods_rel)
    pasta_filtrados = caminho_absoluto_relativo(pasta_filtrados_rel)

    indices = [int(i) for i in indices_str.split(",") if i.strip().isdigit()]
    ordem = [int(i) for i in ordem_str.split(",") if i.strip().isdigit()]

    converter_xls_para_ods(pasta_xls, pasta_ods)
    filtrar_colunas_ods(pasta_ods, pasta_filtrados, indices, ordem_colunas=ordem)
