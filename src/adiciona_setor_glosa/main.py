"""
Lê o relatório de pagamento/glosa, filtra as linhas com Situação Glosa
"GLOSADO" ou "GLOSADO PARCIAL", consulta o setor de cada Cód. Guia no Tasy
usando a query do arquivo .sql e grava uma nova planilha com a coluna "Setor".

Uso:
    uv run adiciona-setor-glosa "Relatórios_pg_e_glosa_2026-09.xlsx"
"""
import sys
from collections import defaultdict
from pathlib import Path

import oracledb
import pandas as pd
from robo_utils.banco_dados.bancodadostasy import BancoDadosTasy  # type: ignore[import-untyped]

# ---------------- Configuração ----------------
ARQUIVO_SQL = Path(__file__).resolve().parents[2] / "sql" / "consulta_setor.sql"
LINHA_CABECALHO = 1            # cabeçalho está na 2ª linha (a 1ª tem totais)
COL_SITUACAO = "Situação Glosa"
COL_GUIA = "Cód. Guia"
COL_NOVA = "Setor"
SITUACOES = {"GLOSADO", "GLOSADO PARCIAL"}
USUARIO_TASY = "TASY"
TAMANHO_LOTE = 1000           # máximo de itens no IN do Oracle
# ----------------------------------------------


def ler_query(caminho: Path) -> str:
    sql = caminho.read_text(encoding="utf-8").strip()
    return sql.rstrip(";")  # o driver não aceita ';' no final


def consultar_setores(guias, sql_modelo: str) -> dict:
    """Consulta as guias em lotes de até 1000 (limite do IN no Oracle),
    usando bind variables e UMA conexão. 1.668 guias = 2 consultas."""
    bd = BancoDadosTasy()
    _, senha = bd.dcconnectasy()
    setores = defaultdict(set)
    com_erro = {}

    with oracledb.connect(user=USUARIO_TASY, password=senha, dsn=bd.dsn_tasy_prd) as conn:
        with conn.cursor() as cur:
            cur.arraysize = 5000  # busca mais linhas por ida ao banco
            for inicio in range(0, len(guias), TAMANHO_LOTE):
                lote = list(guias[inicio:inicio + TAMANHO_LOTE])
                binds = {f"g{i}": guia for i, guia in enumerate(lote)}
                sql = sql_modelo.format(guias=", ".join(f":{nome}" for nome in binds))
                try:
                    cur.execute(sql, binds)
                    for guia, setor in cur:
                        if setor is not None:
                            setores[str(guia).strip()].add(str(setor))
                except oracledb.Error as e:
                    bd.logger.error(f"Lote {inicio}-{inicio + len(lote)}: {e}")
                    com_erro.update({g: f"ERRO: {e}" for g in lote})
                print(f"  {min(inicio + TAMANHO_LOTE, len(guias))}/{len(guias)} guias consultadas")

    # Uma guia pode ter itens em mais de um setor: junta todos
    return {
        g: com_erro.get(g) or (" / ".join(sorted(setores[g])) if setores.get(g) else "NÃO ENCONTRADO")
        for g in guias
    }


def main(arquivo_entrada: str | None = None) -> None:
    # Quando chamado pelo "uv run" (entry point), o arquivo vem da linha de comando
    if arquivo_entrada is None:
        if len(sys.argv) < 2:
            sys.exit('Uso: uv run adiciona-setor-glosa "arquivo.xlsx"')
        arquivo_entrada = sys.argv[1]

    entrada = Path(arquivo_entrada)
    saida = entrada.with_name(f"{entrada.stem}_glosas_com_setor.xlsx")

    print(f"Lendo {entrada.name}...")
    df = pd.read_excel(entrada, header=LINHA_CABECALHO, dtype={COL_GUIA: str})

    situacao = df[COL_SITUACAO].astype(str).str.strip().str.upper()
    filtrado = df[situacao.isin(SITUACOES)].copy()
    guias = filtrado[COL_GUIA].dropna().str.strip().unique()
    print(f"{len(filtrado)} linhas glosadas, {len(guias)} guias distintas.")

    setor_por_guia = consultar_setores(guias, ler_query(ARQUIVO_SQL))

    # Cada linha recebe o setor correspondente ao seu Cód. Guia
    filtrado[COL_NOVA] = filtrado[COL_GUIA].str.strip().map(setor_por_guia)

    print(f"Gravando {saida.name}...")
    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        filtrado.to_excel(writer, index=False, sheet_name="Glosas")
        ws = writer.sheets["Glosas"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            largura = max(len(str(c.value or "")) for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(largura + 2, 50)

    print("Concluído!")


if __name__ == "__main__":
    main()