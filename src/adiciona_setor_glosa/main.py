"""
Lê o relatório de pagamento/glosa, filtra as linhas com Situação Glosa
"GLOSADO" ou "GLOSADO PARCIAL", consulta o setor de cada Cód. Guia no Tasy
usando a query do arquivo .sql e grava uma nova planilha com a coluna "Setor".

Uso:
    uv run adiciona-setor-glosa "entrada.xlsx" ["saida.xlsx"]     (linha de comando)
    uv run adiciona-setor-glosa-gui                                 (interface)
"""
import sys
from collections import defaultdict
from pathlib import Path

import oracledb
import pandas as pd
from robo_utils.banco_dados.bancodadostasy import BancoDadosTasy  # type: ignore[import-untyped]

# ---------------- Configuração ----------------
NOME_SQL = "consulta_setor.sql"
LINHA_CABECALHO = 1            # cabeçalho está na 2ª linha (a 1ª tem totais)
COL_SITUACAO = "Situação Glosa"
COL_GUIA = "Cód. Guia"
COL_MATERIAL = "Cód. Item"     # comparado com cd_material_tiss
COL_DATA = "Data Realização"   # comparada com dt_prescricao
FORMATO_DATA = "%d/%m/%Y %H:%M:%S"  # mesmo formato do to_char da query
COL_NOVA = "Setor"
SITUACOES = {"GLOSADO", "GLOSADO PARCIAL"}
USUARIO_TASY = "TASY"
TAMANHO_LOTE = 1000           # máximo de itens no IN do Oracle
# ----------------------------------------------

COLUNAS_OBRIGATORIAS = [COL_SITUACAO, COL_GUIA, COL_MATERIAL, COL_DATA]


def caminho_sql() -> Path:
    """Localiza a query no projeto ou ao lado do executável."""
    if getattr(sys, "frozen", False):  # rodando como .exe (PyInstaller)
        externo = Path(sys.executable).parent / "sql" / NOME_SQL
        if externo.exists():  # permite editar a query sem recompilar
            return externo
        return Path(sys._MEIPASS) / "sql" / NOME_SQL  # cópia embutida no .exe
    return Path(__file__).resolve().parents[2] / "sql" / NOME_SQL


def nome_saida_padrao(entrada: Path) -> Path:
    return entrada.with_name(f"{entrada.stem}_glosas_com_setor.xlsx")


def ler_query(caminho: Path) -> str:
    sql = caminho.read_text(encoding="utf-8").strip().rstrip(";")  # sem ';' no final
    if "{guias}" not in sql:
        raise ValueError(
            f"A query em {caminho} precisa ter 'in ({{guias}})' no lugar de '= :cod_guia'.\n"
            f"Conteúdo atual:\n{sql}"
        )
    return sql


def normalizar(valor) -> str:
    """Padroniza códigos para comparar planilha x banco (texto, sem zeros à esquerda)."""
    texto = str(valor).strip()
    return texto.lstrip("0") or texto if texto.isdigit() else texto


def juntar(setores: set) -> str:
    # Pode haver mais de um setor: junta todos
    return " / ".join(sorted(setores))


def consultar_setores(guias, sql_modelo: str, log=print, progresso=None):
    """Consulta as guias em lotes de até 1000 (limite do IN no Oracle).
    Retorna os setores por (guia, material), os setores por guia e as guias com erro."""
    bd = BancoDadosTasy()
    _, senha = bd.dcconnectasy()
    por_data = defaultdict(set)   # (guia, cd_material_tiss, dt_prescricao) -> setores
    por_item = defaultdict(set)   # (guia, cd_material_tiss) -> setores
    por_guia = defaultdict(set)   # guia -> setores (usado quando o material não existe)
    com_erro = {}

    with oracledb.connect(user=USUARIO_TASY, password=senha, dsn=bd.dsn_tasy_prd) as conn:
        with conn.cursor() as cur:
            cur.arraysize = 5000  # busca mais linhas por ida ao banco
            for inicio in range(0, len(guias), TAMANHO_LOTE):
                lote = list(guias[inicio:inicio + TAMANHO_LOTE])
                # binds posicionais :1, :2, ... na mesma ordem da lista
                sql = sql_modelo.replace("{guias}", ", ".join(f":{i}" for i in range(1, len(lote) + 1)))
                try:
                    cur.execute(sql, lote)
                    for guia, material, data, setor in cur:
                        if setor is None:
                            continue
                        guia = normalizar(guia)
                        por_guia[guia].add(str(setor))
                        if material is not None:
                            material = normalizar(material)
                            por_item[(guia, material)].add(str(setor))
                            if data is not None:
                                por_data[(guia, material, data)].add(str(setor))
                except oracledb.Error as e:
                    bd.logger.error(f"Lote {inicio}-{inicio + len(lote)}: {e}")
                    log(f"Erro no lote {inicio}-{inicio + len(lote)}: {e}")
                    com_erro.update({normalizar(g): f"ERRO: {e}" for g in lote})
                feitas = min(inicio + TAMANHO_LOTE, len(guias))
                log(f"  {feitas}/{len(guias)} guias consultadas")
                if progresso:
                    progresso(feitas / len(guias))

    return por_data, por_item, por_guia, com_erro


def setor_da_linha(guia, material, data, por_data, por_item, por_guia, com_erro) -> str:
    guia = normalizar(guia)
    if guia in com_erro:
        return com_erro[guia]
    # 1º guia + material + data; 2º guia + material; 3º só a guia
    if pd.notna(material):
        material = normalizar(material)
        if data and (guia, material, data) in por_data:
            return juntar(por_data[(guia, material, data)])
        if (guia, material) in por_item:
            return juntar(por_item[(guia, material)])
    if guia in por_guia:
        return juntar(por_guia[guia])
    return "NÃO ENCONTRADO"


def processar(entrada, saida, log=print, progresso=None) -> dict:
    """Executa todo o fluxo. log(msg) recebe mensagens; progresso(0..100) o andamento."""
    entrada, saida = Path(entrada), Path(saida)
    avancar = progresso or (lambda pct: None)

    log(f"Lendo {entrada.name}...")
    avancar(5)
    df = pd.read_excel(entrada, header=LINHA_CABECALHO, dtype={COL_GUIA: str, COL_MATERIAL: str})
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise ValueError(
            "A planilha não tem as colunas esperadas: " + ", ".join(faltando)
            + ".\nVerifique se é o relatório de pagamento e glosa correto."
        )

    situacao = df[COL_SITUACAO].astype(str).str.strip().str.upper()
    filtrado = df[situacao.isin(SITUACOES)].copy()
    guias = filtrado[COL_GUIA].dropna().str.strip().unique()
    log(f"{len(filtrado)} linhas glosadas, {len(guias)} guias distintas.")
    if len(filtrado) == 0:
        raise ValueError("Nenhuma linha com Situação Glosa GLOSADO ou GLOSADO PARCIAL.")
    avancar(15)

    log("Consultando setores no Tasy...")
    por_data, por_item, por_guia, com_erro = consultar_setores(
        guias, ler_query(caminho_sql()), log=log,
        progresso=lambda fracao: avancar(15 + 70 * fracao),  # consultas: 15% a 85%
    )

    # Data Realização no mesmo texto do to_char da query (dd/mm/yyyy hh24:mi:ss)
    datas = pd.to_datetime(filtrado[COL_DATA], errors="coerce").dt.strftime(FORMATO_DATA)

    # Cada linha recebe o setor mais específico que existir no Tasy
    filtrado[COL_NOVA] = [
        setor_da_linha(guia, material, data if pd.notna(data) else None,
                       por_data, por_item, por_guia, com_erro)
        for guia, material, data in zip(filtrado[COL_GUIA], filtrado[COL_MATERIAL], datas)
    ]

    log(f"Gravando {saida.name}...")
    avancar(90)
    try:
        with pd.ExcelWriter(saida, engine="openpyxl") as writer:
            filtrado.to_excel(writer, index=False, sheet_name="Glosas")
            ws = writer.sheets["Glosas"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                largura = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(largura + 2, 50)
    except PermissionError:
        raise PermissionError(
            f"Não foi possível gravar {saida.name}.\n"
            "Se o arquivo estiver aberto no Excel, feche-o e tente novamente."
        ) from None

    setores = filtrado[COL_NOVA]
    resumo = {
        "linhas": len(filtrado),
        "guias": len(guias),
        "nao_encontrado": int((setores == "NÃO ENCONTRADO").sum()),
        "erros": int(setores.str.startswith("ERRO").sum()),
        "saida": str(saida),
    }
    avancar(100)
    log("Concluído!")
    return resumo


def main() -> None:
    """Linha de comando: uv run adiciona-setor-glosa entrada.xlsx [saida.xlsx]"""
    if len(sys.argv) < 2:
        sys.exit('Uso: uv run adiciona-setor-glosa "entrada.xlsx" ["saida.xlsx"]')
    entrada = Path(sys.argv[1])
    saida = Path(sys.argv[2]) if len(sys.argv) > 2 else nome_saida_padrao(entrada)
    processar(entrada, saida)


if __name__ == "__main__":
    main()