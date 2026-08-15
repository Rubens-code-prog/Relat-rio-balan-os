# -*- coding: utf-8 -*-
"""Relatório de previsões x reais a partir da planilha de balanços."""

import glob
import os
import sys
import unicodedata
import webbrowser
from collections import defaultdict
from datetime import datetime

import openpyxl

MAPA_COLUNAS = {
    "EMPRESA": "empresa",
    "BALANCO": "balanco",
    "DATA BALANCO": "data",
    "RECEITA PREVISTA": "receita_prev",
    "RECEITA REAL": "receita_real",
    "EBITDA PREVISTO": "ebitda_prev",
    "EBITDA TOTAL": "ebitda_real",
    "LUCRO PREVISTO": "lucro_prev",
    "LUCRO REAL": "lucro_real",
    "LPA PREVISTO": "lpa_prev",
    "LPA TOTAL": "lpa_real",
    "LANCAMENTO DO BALANCO": "lancamento",
    "TELECONFERENCIA": "teleconf",
}

MAX_LINHAS_POR_PAGINA = 12


def encontrar_planilha():
    padroes = ["*.xlsx", "*.xlsm", "*.xls"]
    for padrao in padroes:
        arquivos = glob.glob(padrao)
        if arquivos:
            return arquivos[0]
    base = os.path.dirname(os.path.abspath(__file__))
    for padrao in padroes:
        arquivos = glob.glob(os.path.join(base, padrao))
        if arquivos:
            return arquivos[0]
    raise FileNotFoundError("Nenhum arquivo .xlsx/.xlsm/.xls encontrado na pasta.")


def normalizar_cabecalho(texto):
    texto = str(texto or "").strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def mapear_colunas(cabecalho):
    colunas = {}
    for i, valor in enumerate(cabecalho):
        chave = MAPA_COLUNAS.get(normalizar_cabecalho(valor))
        if chave:
            colunas[chave] = i
    return colunas


def obter(row, colunas, nome):
    idx = colunas.get(nome)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def fmt_num(valor, casas=2, sinal=False):
    if valor is None:
        return "—"
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        return "—"
    if sinal:
        prefixo = "+" if valor > 0 else ("-" if valor < 0 else "")
    else:
        prefixo = ""
    corpo = f"{abs(valor):,.{casas}f}"
    corpo = corpo.replace(",", "X").replace(".", ",").replace("X", ".")
    return prefixo + corpo


def fmt_brl(valor):
    v = fmt_num(valor, casas=2)
    return "—" if v == "—" else "R$ " + v


def fmt_data(valor):
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    try:
        return datetime.fromisoformat(str(valor)).strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        return str(valor)


def variacao_pct(previsto, real):
    if previsto is None or real is None:
        return None
    try:
        previsto = float(previsto)
        real = float(real)
    except (TypeError, ValueError):
        return None
    if previsto == 0:
        return None
    return (real - previsto) / abs(previsto) * 100.0


def grupo_lancamento(valor):
    if valor is None:
        return 2
    texto = str(valor).strip().upper()
    if texto in ("ANTES DA ABERTURA", "ANTES", "ANTES DO PREGÃO", "ANTES DO PREGAO", "PRE-MARKET"):
        return 0
    if texto in ("APÓS O FECHAMENTO", "APOS O FECHAMENTO", "APÓS", "APOS", "APÓS O PREGÃO", "APOS O PREGAO", "AFTER-MARKET"):
        return 1
    return 2


def fmt_lancamento(valor):
    grupo = grupo_lancamento(valor)
    if grupo == 0:
        return '<span class="lanc-antes">Antes da abertura</span>'
    if grupo == 1:
        return '<span class="lanc-apos">Após o fechamento</span>'
    return "—"


def ler_dados(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.worksheets[0]

    colunas = None
    linhas = []
    for row in ws.iter_rows(values_only=True):
        if not row or not any(v is not None for v in row):
            continue
        primeira = str(row[0]).strip().upper() if row[0] is not None else ""
        if primeira == "EMPRESA":
            colunas = mapear_colunas(row)
            continue
        if colunas is None or primeira == "":
            continue
        linhas.append(row)

    if colunas is None or "empresa" not in colunas:
        raise ValueError("Não foi possível identificar o cabeçalho (linha com 'EMPRESA').")

    por_data = defaultdict(list)
    por_teleconf = defaultdict(list)
    for row in linhas:
        data = obter(row, colunas, "data")
        if data is not None:
            por_data[data].append(row)
        tc = obter(row, colunas, "teleconf")
        if tc is not None:
            if isinstance(tc, datetime):
                por_teleconf[tc.replace(hour=0, minute=0, second=0, microsecond=0)].append(row)
            else:
                por_teleconf[tc].append(row)
    return wb, colunas, por_data, por_teleconf


def tem_resultado_real(linhas, colunas):
    for row in linhas:
        for nome in ("receita_real", "ebitda_real", "lucro_real", "lpa_real"):
            if obter(row, colunas, nome) is not None:
                return True
    return False


def periodo_de(linhas, colunas):
    for row in linhas:
        valor = obter(row, colunas, "balanco")
        if valor:
            return str(valor)
    return "—"


def menu_escolha(titulo, opcoes):
    """opcoes: lista de tuplas (descrição, retorno)."""
    print(f"\n{titulo}")
    print("  0) Sair")
    for i, (desc, _) in enumerate(opcoes, start=1):
        print(f"  {i}) {desc}")
    while True:
        try:
            opcao = input("\nEscolha: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            sys.exit(0)
        if opcao == "0":
            print("Saindo.")
            sys.exit(0)
        try:
            indice = int(opcao) - 1
            if 0 <= indice < len(opcoes):
                return opcoes[indice][1]
        except ValueError:
            pass
        print("Opção inválida. Tente novamente.")


def menu_tipo():
    return menu_escolha(
        "O que deseja gerar?",
        [
            ("Relatório de Previsões x Reais", "prev"),
            ("Relatório de Teleconferências", "teleconf"),
        ],
    )


def menu_datas(por_data, colunas):
    datas = sorted(por_data.keys())
    opcoes = []
    for data in datas:
        linhas = por_data[data]
        periodo = periodo_de(linhas, colunas)
        marcador = "com dados reais" if tem_resultado_real(linhas, colunas) else "apenas previsões"
        opcoes.append((f"{periodo} — {fmt_data(data)}  [{marcador}]", data))
    return menu_escolha("Datas de balanço disponíveis na planilha:", opcoes)


def menu_datas_teleconf(por_teleconf):
    datas = sorted(por_teleconf.keys())
    opcoes = []
    for data in datas:
        n = len(por_teleconf[data])
        plural = "teleconferência" if n == 1 else "teleconferências"
        opcoes.append((f"{fmt_data(data)}  ({n} {plural})", data))
    return menu_escolha("Dias com teleconferência na planilha:", opcoes)


def css_estilo():
    return """
        :root {
            --bg: #0B0E11;
            --panel: #12171C;
            --panel-2: #161D23;
            --border: #222A31;
            --border-soft: #1A2128;
            --accent: #4EA98B;
            --accent-claro: #6FC79E;
            --accent-mint: #60A080;
            --teal: #208060;
            --texto: #E9EDF0;
            --muted: #8B959E;
            --pos: #46C28A;
            --neg: #E06A7E;
            --ambra: #E3C46F;
            --logo-fundo: #292929;
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            padding: 0 32px 44px;
            min-height: 100vh;
            background:
                radial-gradient(1100px 520px at 88% -8%, rgba(96, 160, 128, .10), transparent 60%),
                radial-gradient(900px 440px at -8% 0%, rgba(96, 160, 128, .06), transparent 55%),
                var(--bg);
            color: var(--texto);
            font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
            font-size: 14px;
            line-height: 1.5;
        }
        ::selection { background: rgba(96, 160, 128, .35); color: #fff; }
        .masthead {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 44px;
            margin: 0 -32px 30px;
            padding: 46px 56px 38px;
            background: linear-gradient(180deg, #0E1419 0%, var(--bg) 100%);
            border-top: 1px solid rgba(96, 160, 128, .35);
            border-bottom: 1px solid var(--border);
            overflow: hidden;
        }
        .masthead::before {
            content: "";
            position: absolute;
            left: 56px;
            top: -1px;
            width: 90px;
            height: 1px;
            background: rgba(96, 160, 128, .8);
        }
        .masthead::after {
            content: "";
            position: absolute;
            left: 56px;
            right: 56px;
            bottom: -1px;
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(96, 160, 128, .6), transparent);
        }
        .mast-rotulo {
            margin: 0 0 12px;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .34em;
            text-transform: uppercase;
            color: var(--accent-claro);
        }
        .mast-titulo {
            margin: 0;
            font-family: 'Playfair Display', Georgia, 'Times New Roman', serif;
            font-size: 34px;
            font-weight: 600;
            line-height: 1.15;
            letter-spacing: .01em;
            color: var(--texto);
        }
        .mast-sub {
            margin: 14px 0 0;
            max-width: 680px;
            color: var(--muted);
            font-size: 13.5px;
        }
        .mast-direita {
            flex-shrink: 0;
            padding: 12px;
            background: var(--logo-fundo);
            border: 1px solid rgba(96, 160, 128, .45);
            border-radius: 20px;
            box-shadow:
                0 16px 38px rgba(0, 0, 0, .5),
                0 2px 8px rgba(0, 0, 0, .35),
                inset 0 0 0 1px rgba(255, 255, 255, .04);
        }
        .mast-direita .mast-sub { text-align: right; }
        .logo {
            display: block;
            height: 118px;
            width: auto;
            border-radius: 12px;
        }
        .badge {
            display: inline-block;
            vertical-align: middle;
            margin-left: 16px;
            padding: 5px 15px;
            border-radius: 999px;
            font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
            font-size: 10.5px;
            font-weight: 600;
            letter-spacing: .1em;
            text-transform: uppercase;
            border: 1px solid transparent;
        }
        .badge-real {
            color: var(--pos);
            background: rgba(70, 194, 138, .10);
            border-color: rgba(70, 194, 138, .28);
        }
        .badge-prev {
            color: var(--ambra);
            background: rgba(201, 162, 39, .10);
            border-color: rgba(201, 162, 39, .30);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border);
            border-radius: 14px;
            overflow: hidden;
            box-shadow: 0 20px 54px rgba(0, 0, 0, .48), 0 2px 6px rgba(0, 0, 0, .32);
            font-size: 13px;
        }
        thead th {
            background: #0D1217;
            color: var(--muted);
            padding: 13px 14px;
            text-align: right;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .12em;
            text-transform: uppercase;
            white-space: nowrap;
            border-bottom: 1px solid var(--border);
        }
        thead th.grupo {
            background: #101720;
            color: var(--accent-claro);
            text-align: center;
            letter-spacing: .16em;
            font-family: 'Playfair Display', Georgia, serif;
            font-size: 13.5px;
            font-weight: 600;
        }
        thead th.rotulo {
            text-align: left;
            color: var(--texto);
        }
        tbody td {
            padding: 11px 14px;
            border-top: 1px solid var(--border-soft);
            text-align: right;
            white-space: nowrap;
            color: var(--texto);
        }
        td.empresa { text-align: left; font-weight: 600; color: var(--texto); }
        tr:hover td { background: rgba(96, 160, 128, .06); }
        .gp { border-left: 1px solid #2A343C; }
        .pos { color: var(--pos); font-weight: 600; }
        .neg { color: var(--neg); font-weight: 600; }
        .neutro { color: var(--muted); }
        .lanc-antes { color: #7FB6EC; font-weight: 600; white-space: nowrap; }
        .lanc-apos { color: var(--ambra); font-weight: 600; white-space: nowrap; }
        .rodape {
            margin-top: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 18px;
            color: var(--muted);
            font-size: 11px;
            letter-spacing: .06em;
        }
        @media (max-width: 760px) {
            .masthead { flex-direction: column; text-align: center; padding: 36px 24px 30px; }
            .mast-direita { order: -1; }
            .logo { height: 96px; }
            body { padding: 0 14px 30px; }
        }
        @media print {
            * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
    """


def encontrar_logo(diretorio):
    for padrao in ("Logo.*", "logo.*", "LOGO.*"):
        arquivos = glob.glob(os.path.join(diretorio, padrao))
        if arquivos:
            return os.path.abspath(arquivos[0])
    return None


def escrever_html(caminho_saida, titulo, subtitulo, thead, corpo, rodape, etiqueta="", pasta_logo=None):
    origem = os.path.dirname(os.path.abspath(caminho_saida))
    logo = encontrar_logo(pasta_logo or origem)
    if not logo and not pasta_logo:
        for _ in range(2):
            origem = os.path.dirname(origem)
            logo = encontrar_logo(origem)
            if logo:
                break
    if logo:
        rel = os.path.relpath(logo, os.path.dirname(os.path.abspath(caminho_saida)))
        direita = f'<div class="mast-direita"><img class="logo" src="{rel}" alt=""></div>'
    else:
        direita = ""

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>{titulo}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>{css_estilo()}</style>
</head>
<body>
    <header class="masthead">
        <div>
            {f'<p class="mast-rotulo">{etiqueta}</p>' if etiqueta else ""}
            <h1 class="mast-titulo">{titulo}</h1>
            <p class="mast-sub">{subtitulo}</p>
        </div>
        {direita}
    </header>
    <table>
        {thead}
        <tbody>
            {corpo}
        </tbody>
    </table>
    <p class="rodape">
        <span>{rodape}</span>
    </p>
</body>
</html>
"""
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html)


def localizar_navegador():
    candidatos = [
        os.environ.get("CHROME_PATH"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in candidatos:
        if c and os.path.isfile(c):
            return c
    return None


def gerar_png(caminho_html, largura=1280, escala=2.0, pasta_png=None):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("AVISO: 'playwright' não está instalado. Rode: pip install playwright")
        return None

    executavel = localizar_navegador()
    if not executavel:
        print("AVISO: Chrome/Edge não encontrado; a imagem não foi gerada.")
        return None

    if pasta_png:
        os.makedirs(pasta_png, exist_ok=True)
        caminho_png = os.path.join(pasta_png, os.path.splitext(os.path.basename(caminho_html))[0] + ".png")
    else:
        caminho_png = os.path.splitext(caminho_html)[0] + ".png"
    url = "file:///" + os.path.abspath(caminho_html).replace("\\", "/")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path=executavel, headless=True)
            page = browser.new_page(
                viewport={"width": largura, "height": 1000},
                device_scale_factor=escala,
            )
            page.goto(url, wait_until="load")
            page.evaluate("document.fonts && document.fonts.ready")
            altura = page.evaluate(
                """() => {
                    const b = document.body;
                    const ultimo = b.lastElementChild;
                    const pb = parseFloat(getComputedStyle(b).paddingBottom) || 0;
                    return Math.ceil(ultimo.getBoundingClientRect().bottom + pb) + 2;
                }"""
            )
            page.set_viewport_size({"width": largura, "height": altura})
            page.screenshot(path=caminho_png)
            browser.close()
    except Exception as e:
        print(f"AVISO: falha ao gerar a imagem: {e}")
        return None
    return caminho_png


def gerar_html(datas_linhas, colunas, caminho_saida):
    data, linhas = datas_linhas
    tem_real = tem_resultado_real(linhas, colunas)
    periodo = periodo_de(linhas, colunas)
    if tem_real:
        linhas_ord = sorted(linhas, key=lambda r: str(obter(r, colunas, "empresa")).lower())
    else:
        linhas_ord = sorted(
            linhas,
            key=lambda r: (
                grupo_lancamento(obter(r, colunas, "lancamento")),
                str(obter(r, colunas, "empresa")).lower(),
            ),
        )

    def var_html(prev, real):
        pct = variacao_pct(prev, real)
        if pct is None:
            return '<span class="neutro">—</span>'
        if pct == 0:
            return '<span class="neutro">0,00%</span>'
        cls = "pos" if pct > 0 else "neg"
        return f'<span class="{cls}">{fmt_num(pct, casas=2, sinal=True)}%</span>'

    def linha_empresa(row, com_real):
        empresa = obter(row, colunas, "empresa")
        receita_prev = obter(row, colunas, "receita_prev")
        receita_real = obter(row, colunas, "receita_real")
        ebitda_prev = obter(row, colunas, "ebitda_prev")
        ebitda_real = obter(row, colunas, "ebitda_real")
        lucro_prev = obter(row, colunas, "lucro_prev")
        lucro_real = obter(row, colunas, "lucro_real")
        lpa_prev = obter(row, colunas, "lpa_prev")
        lpa_real = obter(row, colunas, "lpa_real")
        if com_real:
            cols = [
                f'<td class="empresa">{empresa}</td>',
                f'<td class="gp">{fmt_brl(receita_prev)}</td>',
                f"<td>{fmt_brl(receita_real)}</td>",
                f"<td>{var_html(receita_prev, receita_real)}</td>",
                f'<td class="gp">{fmt_brl(ebitda_prev)}</td>',
                f"<td>{fmt_brl(ebitda_real)}</td>",
                f"<td>{var_html(ebitda_prev, ebitda_real)}</td>",
                f'<td class="gp">{fmt_brl(lucro_prev)}</td>',
                f"<td>{fmt_brl(lucro_real)}</td>",
                f"<td>{var_html(lucro_prev, lucro_real)}</td>",
                f'<td class="gp">{fmt_num(lpa_prev, casas=4)}</td>',
                f"<td>{fmt_num(lpa_real, casas=4)}</td>",
                f"<td>{var_html(lpa_prev, lpa_real)}</td>",
            ]
            return "<tr>" + "".join(cols) + "</tr>"
        cols = [
            f'<td class="empresa">{empresa}</td>',
            f"<td>{fmt_lancamento(obter(row, colunas, 'lancamento'))}</td>",
            f'<td class="gp">{fmt_brl(receita_prev)}</td>',
            f'<td class="gp">{fmt_brl(ebitda_prev)}</td>',
            f'<td class="gp">{fmt_brl(lucro_prev)}</td>',
            f'<td class="gp">{fmt_num(lpa_prev, casas=4)}</td>',
        ]
        return "<tr>" + "".join(cols) + "</tr>"

    if tem_real:
        thead = """
            <thead>
                <tr>
                    <th class="rotulo" rowspan="2">Empresa</th>
                    <th class="grupo gp" colspan="3">Receita</th>
                    <th class="grupo gp" colspan="3">EBITDA</th>
                    <th class="grupo gp" colspan="3">Lucro</th>
                    <th class="grupo gp" colspan="3">LPA</th>
                </tr>
                <tr>
                    <th class="gp">Prevista</th><th>Real</th><th>Var.</th>
                    <th class="gp">Previsto</th><th>Real</th><th>Var.</th>
                    <th class="gp">Previsto</th><th>Real</th><th>Var.</th>
                    <th class="gp">Previsto</th><th>Real</th><th>Var.</th>
                </tr>
            </thead>
        """
    else:
        thead = """
            <thead>
                <tr>
                    <th class="rotulo">Empresa</th>
                    <th>Lançamento</th>
                    <th class="gp">Receita Prevista</th>
                    <th class="gp">EBITDA Previsto</th>
                    <th class="gp">Lucro Previsto</th>
                    <th class="gp">LPA Previsto</th>
                </tr>
            </thead>
        """

    if tem_real:
        badge = '<span class="badge badge-real">Resultados divulgados</span>'
        nota = "Comparação entre o previsto e o realizado. A variação é calculada sobre o valor previsto."
    else:
        badge = '<span class="badge badge-prev">Previsões</span>'
        nota = "Previsões para este período. O horário de divulgação dos resultados (antes da abertura ou após o fechamento do pregão) está indicado na tabela."

    titulo = f"Relatório de Previsões x Reais — {periodo} ({fmt_data(data)}){badge}"
    rodape = f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} · {len(linhas)} empresa(s)"
    return _escrever_paginas(
        caminho_saida, titulo, nota, thead, [linha_empresa(r, tem_real) for r in linhas_ord], rodape
    )


def _escrever_paginas(caminho_saida, titulo, subtitulo, thead, linhas_html, rodape, etiqueta=""):
    """Escreve o relatório paginado e devolve a lista de arquivos HTML gerados."""
    base, ext = os.path.splitext(caminho_saida)
    paginas = []
    total = max(1, -(-len(linhas_html) // MAX_LINHAS_POR_PAGINA))
    for i in range(total):
        if total == 1:
            saida = caminho_saida
        else:
            saida = f"{base}_p{i + 1:02d}{ext}"
        bloco = linhas_html[i * MAX_LINHAS_POR_PAGINA:(i + 1) * MAX_LINHAS_POR_PAGINA]
        pag_rodape = rodape if total == 1 else f"{rodape} · Página {i + 1} de {total}"
        escrever_html(saida, titulo, subtitulo, thead, "".join(bloco), pag_rodape, etiqueta=etiqueta)
        paginas.append(saida)
    return paginas


def gerar_html_teleconf(data_teleconf, linhas, colunas, caminho_saida):
    linhas_ord = sorted(linhas, key=lambda r: obter(r, colunas, "teleconf"))

    def linha(row):
        tc = obter(row, colunas, "teleconf")
        hora = tc.strftime("%H:%M") if isinstance(tc, datetime) else "—"
        return (
            "<tr>"
            f'<td class="empresa">{obter(row, colunas, "empresa")}</td>'
            f"<td>{periodo_de([row], colunas)}</td>"
            f"<td>{hora}</td>"
            "</tr>"
        )

    thead = """
            <thead>
                <tr>
                    <th class="rotulo">Empresa</th>
                    <th>Balanço</th>
                    <th>Horário</th>
                </tr>
            </thead>
        """
    titulo = f"Teleconferências de Resultados — {fmt_data(data_teleconf)}"
    nota = "Agenda de teleconferências das empresas nesta data, em ordem de horário."
    rodape = f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} · {len(linhas)} empresa(s)"
    return _escrever_paginas(
        caminho_saida, titulo, nota, thead, [linha(r) for r in linhas_ord], rodape, etiqueta="AGENDA CORPORATIVA"
    )


def main():
    try:
        caminho = encontrar_planilha()
    except FileNotFoundError as e:
        print(e)
        sys.exit(1)

    print(f"Lendo: {caminho}")
    try:
        wb, colunas, por_data, por_teleconf = ler_dados(caminho)
    except PermissionError:
        print("ERRO: a planilha está aberta no Excel. Feche-a e tente novamente.")
        sys.exit(1)
    except ValueError as e:
        print(f"ERRO: {e}")
        sys.exit(1)

    base = os.path.dirname(os.path.abspath(caminho))
    dir_html = os.path.join(base, "html")
    dir_img = os.path.join(base, "imagens")
    os.makedirs(dir_html, exist_ok=True)
    os.makedirs(dir_img, exist_ok=True)
    tipo = menu_tipo()

    if tipo == "teleconf":
        if not por_teleconf:
            print("Nenhuma data de teleconferência encontrada na planilha.")
            sys.exit(1)
        data = menu_datas_teleconf(por_teleconf)
        linhas = por_teleconf[data]
        nome = f"teleconferencias_{data.strftime('%Y%m%d')}.html"
        saida = os.path.join(dir_html, nome)
        paginas = gerar_html_teleconf(data, linhas, colunas, saida)
    else:
        if not por_data:
            print("Nenhuma data de balanço encontrada na planilha.")
            sys.exit(1)
        data = menu_datas(por_data, colunas)
        linhas = por_data[data]
        periodo = periodo_de(linhas, colunas)
        nome = f"relatorio_{periodo}_{data.strftime('%Y%m%d')}.html"
        saida = os.path.join(dir_html, nome)
        paginas = gerar_html((data, linhas), colunas, saida)

    pngs = []
    for i, pagina in enumerate(paginas, start=1):
        png = gerar_png(pagina, pasta_png=dir_img)
        if png:
            pngs.append(png)
            print(f"Imagem {i} de {len(paginas)}: {png}")
        else:
            print(f"Relatório gerado: {pagina}")
    if not pngs:
        print("\nNenhuma imagem pôde ser gerada.")
    try:
        webbrowser.open("file://" + os.path.abspath(paginas[0]).replace("\\", "/"))
    except Exception:
        pass


if __name__ == "__main__":
    main()
