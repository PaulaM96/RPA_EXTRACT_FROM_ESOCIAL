import os
import logging
import shutil

# Configuração do log
logging.basicConfig(level=logging.INFO, filename="renomear_arquivos.log", format="%(asctime)s - %(levelname)s - %(message)s")

def renomear_e_mover_arquivos(base_path):
    try:
        # Percorre todos os diretórios dentro do caminho base
        for dirpath, dirnames, filenames in os.walk(base_path):
            for filename in filenames:
                # Verifica se o arquivo contém "eSocial_Evento" no nome
                if "eSocial_Evento" in filename:
                    # Nome do funcionário é o nome da pasta
                    nome_funcionario = os.path.basename(dirpath).strip()
                    # Sanitizar o nome do funcionário
                    nome_funcionario = ''.join(e for e in nome_funcionario if e.isalnum() or e.isspace()).strip()
                    # Novo nome do arquivo
                    novo_nome = f"{nome_funcionario}.xml"
                    caminho_arquivo_atual = os.path.join(dirpath, filename)
                    caminho_arquivo_novo = os.path.join(base_path, novo_nome)

                    try:
                        # Renomeia e move o arquivo para a raiz da pasta base
                        os.rename(caminho_arquivo_atual, caminho_arquivo_novo)
                        logging.info(f"Arquivo '{caminho_arquivo_atual}' renomeado e movido para '{caminho_arquivo_novo}'")
                    except Exception as e:
                        logging.error(f"Erro ao renomear e mover o arquivo '{caminho_arquivo_atual}': {e}")
            # Após processar todos os arquivos na pasta, excluir a pasta do funcionário se estiver vazia
            if not os.listdir(dirpath):
                try:
                    os.rmdir(dirpath)
                    logging.info(f"Pasta '{dirpath}' excluída")
                except Exception as e:
                    logging.error(f"Erro ao excluir a pasta '{dirpath}': {e}")
    except Exception as e:
        logging.error(f"Erro ao processar o caminho base '{base_path}': {e}")

def main():
    # Lista de centros de custo
    ccs = ['cc1','cc2']
    for cc in ccs:
        # Caminho base onde as pastas dos funcionários estão localizadas
        base_path = f"\\\\servidor\\caminho\\CP\\{cc}"
        # Log para verificar se está processando o caminho base
        logging.info(f"Processando o caminho base: {base_path}")
        # Renomeia e move os arquivos conforme necessário
        renomear_e_mover_arquivos(base_path)

if __name__ == "__main__":
    main()
