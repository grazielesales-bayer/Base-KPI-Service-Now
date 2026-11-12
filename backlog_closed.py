import time
import os
import json
from datetime import date, timedelta
from playwright.sync_api import sync_playwright

def carregar_json():
    #Carrega o caminho do download a partir do arquivo pasta.json.
    try:
        # Busca o JSON no mesmo diretório do script
        caminho_json = os.path.join(os.path.dirname(__file__), 'pasta.json')
        with open(caminho_json, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get("caminho_download", r"C:\temp")
    except Exception as e:
        print(f"⚠️ Erro ao ler JSON: {e}. Usando C:\\temp por padrão.")
        return r"C:\temp"


def rodar_servicenow_closed():
    with sync_playwright() as p:
        context = None
        try:

            context = p.chromium.launch_persistent_context(
                user_data_dir= r"C:\temp",
                channel="chrome",
                headless=False,
                accept_downloads=True
            )
            page = context.pages[0] if context.pages else context.new_page()

            print("Acessando o site...")
            page.goto("https://bayersi.service-now.com/now/psm/workspace/list")
            page.get_by_role("treeitem", name='Backlog Tickets - Closed').click()
            time.sleep(2)
            page.get_by_role("button", name="filtro Atualizado", exact=True).click()
            time.sleep(2)
            page.locator(".sn-select-icon").click()
            page.get_by_role("option", name="em ou após", exact=True).click()
            campo_data = page.locator('input.now-input-native[placeholder="DD/MM/YYYY"]')
            campo_data.click()
            campo_data.fill((date.today() - timedelta(days=1)).strftime("%d/%m/%Y"))
            page.get_by_role("button", name="Ok", exact=True).click()
            page.get_by_role(
                "button",
                name="Filtro Aplicar na coluna Atualizado",
                exact=True,
            ).click()

            time.sleep(3)

            #print("clica no primeiro exportar")
            page.get_by_text("Exportar").click()
            time.sleep(3)

            try:
                with page.expect_download(timeout=60000) as download_info:

                    page.get_by_role("button", name="Exportar").click()

                download = download_info.value

                pasta_download = carregar_json()

                nome_arquivo = f"backlog_closed_{download.suggested_filename}"
                path = os.path.join(pasta_download, nome_arquivo)
                download.save_as(path)

                print(f"Sucesso! Arquivo salvo em: {path}")

            except Exception as e:
                print(f"O download não iniciou ou demorou demais: {e}")


        finally:
            #fecha sempre, mesmo se der erro
            if context:
                context.close()


if __name__ == "__main__":
    rodar_servicenow_closed()
