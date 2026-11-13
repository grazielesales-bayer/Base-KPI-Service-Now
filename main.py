from pathlib import Path

from openpyxl import load_workbook

from backlog_closed import carregar_json, rodar_servicenow_closed
from backlog_open import rodar_servicenow_open


ARQUIVO_DESTINO = Path(r"C:\temp\kpi-snow\pasta teste\Base Snow.xlsx")


def _snapshot_arquivos(pasta, prefixo):
    return {
        arquivo: (arquivo.stat().st_mtime_ns, arquivo.stat().st_size)
        for arquivo in pasta.glob(f"{prefixo}*.xlsx")
    }


def _localizar_download_atualizado(pasta, prefixo, snapshot):
    arquivos_atualizados = [
        arquivo
        for arquivo in pasta.glob(f"{prefixo}*.xlsx")
        if snapshot.get(arquivo)
        != (arquivo.stat().st_mtime_ns, arquivo.stat().st_size)
    ]
    if not arquivos_atualizados:
        raise FileNotFoundError(
            f"Nenhum download novo encontrado para o prefixo {prefixo!r} "
            f"em {pasta}"
        )
    return max(arquivos_atualizados, key=lambda arquivo: arquivo.stat().st_mtime_ns)


def _mapear_colunas(cabecalho_origem, cabecalho_destino):
    colunas_destino = {
        str(cabecalho).strip().casefold(): coluna
        for coluna, cabecalho in enumerate(cabecalho_destino, start=2)
        if cabecalho not in (None, "")
    }
    aliases = {"subcategory": "subcategoria"}
    mapeamento = []
    colunas_sem_destino = []

    for coluna_origem, cabecalho in enumerate(cabecalho_origem):
        if cabecalho in (None, ""):
            raise ValueError("O relatório contém um cabeçalho vazio.")
        nome = str(cabecalho).strip().casefold()
        nome_destino = aliases.get(nome, nome)
        coluna_destino = colunas_destino.get(nome_destino)
        if coluna_destino is None:
            colunas_sem_destino.append(cabecalho)
        else:
            mapeamento.append((coluna_origem, coluna_destino))

    if colunas_sem_destino:
        raise ValueError(
            "Não há colunas correspondentes no destino para: "
            + ", ".join(str(coluna) for coluna in colunas_sem_destino)
        )
    return mapeamento


def _combinar_planilhas(arquivo_closed, arquivo_open, arquivo_destino):
    workbook_closed = load_workbook(arquivo_closed, data_only=True)
    workbook_open = None
    workbook_destino = None
    try:
        workbook_open = load_workbook(arquivo_open, data_only=True)
        planilha_closed = workbook_closed.active
        planilha_open = workbook_open.active

        workbook_destino = load_workbook(arquivo_destino)
        planilha_destino = workbook_destino.active

        cabecalho_closed = next(
            planilha_closed.iter_rows(min_row=1, max_row=1, values_only=True),
            None,
        )
        cabecalho_open = next(
            planilha_open.iter_rows(min_row=1, max_row=1, values_only=True),
            None,
        )
        cabecalho_destino = tuple(
            celula.value
            for celula in planilha_destino[1][1:]
        )
        if not cabecalho_closed or not cabecalho_open:
            raise ValueError("Um dos relatórios está sem cabeçalho.")
        mapeamento_closed = _mapear_colunas(cabecalho_closed, cabecalho_destino)
        mapeamento_open = _mapear_colunas(cabecalho_open, cabecalho_destino)

        linha_destino = 2
        for linha in range(planilha_destino.max_row, 1, -1):
            numero_chamado = planilha_destino.cell(row=linha, column=2).value
            if numero_chamado not in (None, ""):
                linha_destino = linha + 1
                break

        for planilha_origem, mapeamento in (
            (planilha_closed, mapeamento_closed),
            (planilha_open, mapeamento_open),
        ):
            for linha in planilha_origem.iter_rows(min_row=2, values_only=True):
                for coluna_origem, coluna_destino in mapeamento:
                    planilha_destino.cell(
                        row=linha_destino,
                        column=coluna_destino,
                        value=linha[coluna_origem],
                    )
                linha_destino += 1

        arquivo_destino.parent.mkdir(parents=True, exist_ok=True)
        workbook_destino.save(arquivo_destino)
    finally:
        workbook_closed.close()
        if workbook_open is not None:
            workbook_open.close()
        if workbook_destino is not None:
            workbook_destino.close()


def main():
    pasta_download = Path(carregar_json())

    snapshot_closed = _snapshot_arquivos(pasta_download, "backlog_closed_")
    rodar_servicenow_closed()
    arquivo_closed = _localizar_download_atualizado(
        pasta_download, "backlog_closed_", snapshot_closed
    )

    snapshot_open = _snapshot_arquivos(pasta_download, "backlog_open_")
    rodar_servicenow_open()
    arquivo_open = _localizar_download_atualizado(
        pasta_download, "backlog_open_", snapshot_open
    )

    _combinar_planilhas(arquivo_closed, arquivo_open, ARQUIVO_DESTINO)
    print(f"Planilhas combinadas em: {ARQUIVO_DESTINO}")


if __name__ == "__main__":
    main()
