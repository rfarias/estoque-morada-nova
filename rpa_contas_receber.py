# -*- coding: utf-8 -*-
"""
RPA - Lancamento de Contas a Receber no Fortes AG (Nuvem) a partir da Planilha1.

Uso (no terminal, dentro desta pasta):
    python rpa_contas_receber.py listar          -> so mostra o que seria lancado (nao mexe no Fortes)
    python rpa_contas_receber.py teste           -> lanca 1 pedido e PARA antes do Ok final (voce confere e aperta F9)
    python rpa_contas_receber.py teste 144       -> idem, para a linha 144 (ou um nº de pedido)
    python rpa_contas_receber.py gravar 144      -> aperta Ok (F9) no documento preenchido e apaga J da linha 144
    python rpa_contas_receber.py rodar           -> lanca todos os pedidos que ainda tem a coluna J preenchida
    python rpa_contas_receber.py retomar 8 1     -> termina o documento ja aberto da linha 8 (1 item ja lancado) e segue com os demais

Seguranca:
  - Leve o mouse para o CANTO SUPERIOR ESQUERDO da tela para abortar na hora.
  - Depois de gravar cada documento o robo APAGA a coluna J daquela linha (= lancado).
  - Antes do Ok final de cada documento e salvo um print em prints_rpa/.

Pre-requisitos na tela:
  - Excel aberto com a planilha SAIDA ... (aba Planilha1).
  - Fortes AG maximizado, na tela "Contas a Receber - Documentos" (lista, sem documento aberto).
"""
import datetime as dt
import os
import re
import sys
import time

import pyautogui
import pyperclip
import win32api
import win32com.client
import win32con
import win32gui
import win32process

# ----------------------------------------------------------------- CONFIGURACAO
ESTABELECIMENTO = "9"
CENTRO_RESULTADOS = "003"
RECEITA = "4.03.001"
DIAS_VENCIMENTO = None      # None = mesmo dia do mes seguinte a emissao; ou um numero de dias (ex.: 30)
AGENTE_COBRADOR = "1"       # 0001 CICRAC
TIPO_DOCUMENTO = "6"        # 006 Nota Promissoria

PAUSA = 0.5                 # pausa entre acoes (s) - aumente se o acesso remoto estiver lento
PAUSA_LOOKUP = 1.5          # espera apos digitar um codigo (cliente, item...) para o Fortes buscar
PAUSA_TELA = 2.5            # espera para abrir/salvar formularios

# Coordenadas (tela 1920x1080, Fortes maximizado) - tiradas dos prints
P = {
    # Botao Novo (F2) da lista (mesma posicao do Ok (F9) quando o documento esta aberto)
    "novo_documento": (1862, 132),
    # Cabecalho do documento
    "estabelecimento": (31, 178),
    "cliente": (297, 178),
    "centro": (55, 223),
    "receita": (313, 222),
    "emissao": (50, 265),
    "obs": (230, 463),
    "ok_documento": (1867, 132),
}
# Painel de baixo: a altura do cabecalho varia, entao o painel sobe/desce.
# As posicoes abaixo valem com as abas em y=500; o robo procura as abas na tela
# (imagem ref_abas.png, topo em y=493 nessa posicao) e soma a diferenca.
PAINEL_BASE = {
    # Abas de baixo
    "aba_itens": (89, 500),
    "aba_venc": (153, 500),
    # Botao Novo / Ok do painel de baixo (mesma posicao)
    "novo_baixo": (1767, 531),
    "ok_baixo": (1767, 531),
    # Formulario do item
    "item_codigo": (55, 553),
    "item_sta": (32, 705),
    "item_stb": (244, 705),
    "item_cfop": (40, 661),
    "item_unidade": (38, 793),
    "item_qtd": (302, 793),
    "item_valor": (393, 793),
    # Formulario do vencimento
    "venc_data": (50, 553),
    "venc_valor": (135, 553),
    "venc_agente": (186, 594),
    "venc_tipo": (429, 594),
}
REGIAO_OK_BAIXO = None  # regiao do icone do botao Novo/Ok do painel de baixo

PASTA = os.path.dirname(os.path.abspath(__file__))
PASTA_PRINTS = os.path.join(PASTA, "prints_rpa")
REF_ABAS = os.path.join(PASTA, "ref_abas.png")
REF_ABAS_TOPO = 493
REF_MSG_STA = os.path.join(PASTA, "ref_msg_sta.png")      # texto: Campo "STA" em branco.
REF_MSG_CAMPO_ST = os.path.join(PASTA, "ref_msg_campo_st.png")  # texto: Campo "ST
VALORES_ST = {"item_sta": "0", "item_stb": "40"}  # o que digitar quando o Fortes reclama do campo em branco
REF_MSG_NAO_CADASTRADO = os.path.join(PASTA, "ref_msg_nao_cadastrado.png")  # texto: Não cadastrado.
CFOP_NAO_CADASTRADO = "6949"  # quando o Fortes diz 'Não cadastrado.' no CFOP -> Ok e digita este CFOP


def ajustar_painel():
    """Procura as abas do painel de baixo na tela e acerta as posicoes."""
    global REGIAO_OK_BAIXO
    import pyscreeze
    try:
        achou = pyscreeze.locateOnScreen(REF_ABAS, region=(0, 300, 700, 600))
    except pyscreeze.ImageNotFoundException:
        achou = None
    if achou is None:
        raise TelaErrada("Nao encontrei as abas Itens / Venc. Diversos na tela.")
    dy = achou.top - REF_ABAS_TOPO
    for nome, (x, y) in PAINEL_BASE.items():
        P[nome] = (x, y + dy)
    REGIAO_OK_BAIXO = (1733, 525 + dy, 16, 14)
    acao(f"painel de baixo encontrado (deslocamento {dy:+d} px)")

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


# ----------------------------------------------------------------- LEITURA DA PLANILHA
def num(txt):
    """'1.720,00' -> 1720.0"""
    return float(str(txt).strip().replace(".", "").replace(",", "."))


def br(valor, casas=2):
    s = f"{valor:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", "")


def normaliza(nome):
    import unicodedata
    nome = unicodedata.normalize("NFKD", nome.upper()).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", nome).strip()


RE_TITULO = re.compile(r"^SAIDA DE (.+?)\s*\(", re.I)
RE_CABEC = re.compile(r"^(.*?) - (\d+) - (\d{2}/\d{2}/\d{2,4}) - (.*)$")
RE_ITEM = re.compile(r"^([\d.,]+) (SACOS?|KG) DE (.+?) A R\$ ([\d.,]+)$", re.I)
# Produtos vendidos por KG: forca a unidade KG mesmo se a descricao disser SACOS.
PALAVRAS_KG = {"CAROCO", "OLEO", "KG"}


def eh_boleto(texto):
    return "BOLET" in normaliza(texto)


def abrir_planilha():
    xl = win32com.client.GetActiveObject("Excel.Application")
    for wb in xl.Workbooks:
        for ws in wb.Worksheets:
            if ws.Name.strip().lower() == "planilha1":
                return wb.Name, ws
    raise SystemExit("Nao encontrei nenhuma planilha aberta com a aba 'Planilha1'.")


def ler_pedidos():
    nome_wb, ws = abrir_planilha()
    ultima = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    dados = ws.Range(f"A1:J{ultima}").Value  # tupla de tuplas

    def txt(v):
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).strip()

    # 1) mapa produto -> codigo do item (linhas de titulo "SAIDA DE ...", codigo na coluna I)
    produtos = {}
    for linha in dados:
        m = RE_TITULO.match(txt(linha[0]))
        if m and txt(linha[8]):
            produtos[normaliza(m.group(1))] = txt(linha[8])

    # 2) pedidos (coluna J)
    pedidos, problemas = [], []
    for i, linha in enumerate(dados, start=1):
        j = txt(linha[9])
        if not j:
            continue
        cod_cli = txt(linha[8])
        m = RE_CABEC.match(j)
        if not m:
            problemas.append((i, "?", "Nao consegui ler a coluna J: " + j))
            continue
        nome, pedido, data, resto = m.groups()
        if "CANC" in normaliza(nome):
            continue  # cancelado: ignora
        partes = [p.strip() for p in resto.split(" - ")]
        forma = partes[-1] if partes and not RE_ITEM.match(partes[-1]) else ""
        if eh_boleto(forma) or eh_boleto(txt(linha[7])):
            problemas.append((i, pedido, f"{nome.strip()}: pago com BOLETO - nao lancar"))
            continue
        itens, erro = [], None
        for p in partes:
            mi = RE_ITEM.match(p)
            if not mi:
                continue
            qtd, un, prod, preco = mi.groups()
            cod_item = produtos.get(normaliza(prod))
            if not cod_item:
                erro = f"Produto sem codigo na planilha: {prod}"
            itens.append({
                "codigo": cod_item,
                "produto": prod,
                "unidade": "KG" if un.upper() == "KG" or PALAVRAS_KG & set(normaliza(prod).split()) else "SC",
                "qtd": num(qtd),
                "preco": num(preco),
            })
        if not cod_cli:
            erro = "Sem codigo de cliente na coluna I"
        elif not itens:
            erro = "Nenhum item encontrado na coluna J"
        d, mth, a = data.split("/")
        emissao = dt.date(2000 + int(a) if len(a) == 2 else int(a), int(mth), int(d))
        ped = {
            "linha": i, "pedido": pedido, "cliente": cod_cli, "nome": nome.strip(),
            "emissao": emissao, "forma": forma, "itens": itens, "obs": j,
            "total": round(sum(it["qtd"] * it["preco"] for it in itens), 2),
        }
        if erro:
            problemas.append((i, pedido, f"{nome.strip()}: {erro}"))
        else:
            pedidos.append(ped)
    return nome_wb, ws, pedidos, problemas


def vencimento(emissao):
    if DIAS_VENCIMENTO is not None:
        return emissao + dt.timedelta(days=DIAS_VENCIMENTO)
    ano, mes = (emissao.year + 1, 1) if emissao.month == 12 else (emissao.year, emissao.month + 1)
    import calendar
    dia = min(emissao.day, calendar.monthrange(ano, mes)[1])
    return dt.date(ano, mes, dia)


def apagar_j(ws, linha):
    ws.Cells(linha, 10).ClearContents()


# ----------------------------------------------------------------- ACOES NO FORTES
FORTES_PID = None


def ativar_fortes():
    alvo = []
    win32gui.EnumWindows(lambda h, _: alvo.append(h) if "Fortes AG" in win32gui.GetWindowText(h)
                         and win32gui.IsWindowVisible(h) else None, None)
    if not alvo:
        raise SystemExit("Janela do Fortes AG nao encontrada. Abra o Fortes pelo acesso remoto.")
    h = alvo[0]
    global FORTES_PID
    FORTES_PID = pid_da_janela(h)
    win32gui.ShowWindow(h, win32con.SW_MAXIMIZE)
    # ATENCAO: nao apertar Alt aqui - no Fortes o Alt abre o menu principal.
    if not fortes_na_frente():
        try:
            win32gui.SetForegroundWindow(h)
        except Exception:
            pass
    time.sleep(1.5)
    if not fortes_na_frente():
        # minimiza a janela que esta na frente (ex.: terminal) e clica na barra de titulo do Fortes
        frente = win32gui.GetForegroundWindow()
        if frente and frente != h:
            win32gui.ShowWindow(frente, win32con.SW_MINIMIZE)
            time.sleep(1.0)
        x, y = 958, 19  # barra de titulo "CARNEIRO IND E COM LTDA"
        if pid_da_janela(win32gui.WindowFromPoint((x, y))) == FORTES_PID:
            pyautogui.click(x, y)
            time.sleep(1.0)
    if not fortes_na_frente():
        raise SystemExit("Nao consegui trazer o Fortes para a frente. Clique na janela do Fortes, "
                         "minimize o terminal e rode de novo.")


class Parada(Exception):
    pass


def checar_parada():
    """ESC pressionado a qualquer momento = parar o robo."""
    if win32api.GetAsyncKeyState(win32con.VK_ESCAPE) & 0x8001:
        raise Parada()


def esperar(seg):
    fim = time.time() + seg
    while time.time() < fim:
        checar_parada()
        time.sleep(0.05)


def pid_da_janela(h):
    return win32process.GetWindowThreadProcessId(h)[1]


def fortes_na_frente():
    return "Fortes AG" in win32gui.GetWindowText(win32gui.GetForegroundWindow())


def janela_passageira(h):
    """Janelinha do Fortes sem titulo e pequena ou invisivel (dica/hint): pode ignorar.
    O menu principal tambem nao tem titulo, mas e grande -> NAO e passageira."""
    if not FORTES_PID or pid_da_janela(h) != FORTES_PID:
        return False
    if win32gui.GetWindowText(h).replace("(Remoto)", "").strip():
        return False  # tem titulo (ex.: 'Informação') -> e uma mensagem de verdade
    if not win32gui.IsWindowVisible(h):
        return True
    l, t, r, b = win32gui.GetWindowRect(h)
    return (r - l) * (b - t) < 600 * 300


def fortes_ok():
    """Fortes na frente, ou so uma dica passageira do Fortes por cima."""
    h = win32gui.GetForegroundWindow()
    return "Fortes AG" in win32gui.GetWindowText(h) or janela_passageira(h)


def mensagem_do_fortes():
    """Retorna o titulo se a janela da frente for uma mensagem do Fortes com titulo
    (ex.: 'Informação (Remoto)'); senao None."""
    h = win32gui.GetForegroundWindow()
    titulo = win32gui.GetWindowText(h)
    if "Fortes AG" in titulo or not FORTES_PID or pid_da_janela(h) != FORTES_PID:
        return None
    if not titulo.replace("(Remoto)", "").strip():
        return None
    return titulo


def confirmar_mensagem(tentativas=6, espera_max=0.0):
    """Se aparecer a mensagem do Fortes (ex.: limite de credito depois do cliente), aperta Ok
    (Enter). espera_max > 0: com o acesso remoto lento a mensagem pode demorar, entao espera
    ate esse tempo por ela antes de seguir (sai logo depois de responder)."""
    fim = time.time() + espera_max
    respondeu = False
    n = 0
    while True:
        esperar(0.5)
        n += 1
        tit = mensagem_do_fortes()
        if tit is None:
            if fortes_ok() and (respondeu or time.time() >= fim):
                return
            if n >= tentativas and time.time() >= fim:
                return
            continue
        if "Informação" not in tit:
            raise TelaErrada(f"Apareceu a janela '{tit}' e nao sei o que responder - parei.")
        print(f"[mensagem '{tit}' -> Ok] ", end="", flush=True)
        pyautogui.press("enter")
        respondeu = True
        esperar(PAUSA)


DIAGNOSTICO = False   # ligado no modo teste: registra cada acao e confere a janela depois dela
ULTIMAS_ACOES = []


def acao(descricao):
    ULTIMAS_ACOES.append(descricao)
    if DIAGNOSTICO:
        print(f"\n    > {descricao}", end="", flush=True)


def conferir_janela(depois_de, permitir_mensagem=False):
    if DIAGNOSTICO:
        time.sleep(0.25)
    if permitir_mensagem and mensagem_do_fortes() is not None:
        return
    # janelinhas passageiras do Fortes (dicas, listas de busca) somem sozinhas: espera ate 3 s
    for _ in range(30):
        if fortes_ok():
            return
        esperar(0.1)
    if not fortes_na_frente() and mensagem_do_fortes() is None:
        raise TelaErrada(f"A janela da frente deixou de ser o Fortes logo depois de: {depois_de}")
    if not fortes_na_frente():
        h = win32gui.GetForegroundWindow()
        raise TelaErrada(f"Apareceu a janela '{win32gui.GetWindowText(h)}' "
                         f"(visivel={win32gui.IsWindowVisible(h)}, ret={win32gui.GetWindowRect(h)})"
                         f" logo depois de: {depois_de}")


def tecla(k, vezes=1, permitir_mensagem=False):
    checar_parada()
    acao(f"tecla {k}" + (f" x{vezes}" if vezes > 1 else ""))
    pyautogui.press(k, presses=vezes, interval=0.02)
    conferir_janela(f"tecla {k}", permitir_mensagem)


def clicar(nome, espera=PAUSA):
    checar_parada()
    conferir_janela(f"(antes de clicar em {nome})")
    acao(f"clique em {nome} {P[nome]}")
    pyautogui.click(*P[nome])
    esperar(espera)
    conferir_janela(f"clique em {nome}")


def digitar(texto):
    acao(f"digita '{texto}'")
    for ch in str(texto):
        checar_parada()
        pyautogui.typewrite(ch)
        time.sleep(0.03)
    conferir_janela(f"digitar '{texto}'")


SEM_LEITURA = "<<rpa-sem-leitura>>"


def ler_campo():
    """Copia o conteudo do campo atual (Home, Shift+End, Ctrl+C) e devolve o texto,
    ou None se nao conseguiu copiar."""
    pyperclip.copy(SEM_LEITURA)
    acao("le o campo (home, shift+end, ctrl+c)")
    pyautogui.press("home")
    pyautogui.hotkey("shift", "end")
    pyautogui.hotkey("ctrl", "c")
    conferir_janela("ler o campo")
    global LEITURA_OK
    tentativas = 25 if LEITURA_OK is None else 10  # 1a vez espera ate 2,5 s
    for _ in range(tentativas):  # a area de transferencia do acesso remoto demora um pouco
        time.sleep(0.1)
        lido = pyperclip.paste()
        if lido != SEM_LEITURA:
            LEITURA_OK = True
            return lido
    if LEITURA_OK is None:
        LEITURA_OK = False
        acao("NAO consegui ler o campo pelo Ctrl+C - desligo a verificacao e sigo apagando/digitando")
    return None


LEITURA_OK = None  # None = ainda nao testou; False = Ctrl+C nao funciona no acesso remoto


def mesmo_valor(atual, desejado):
    """Compara ignorando espacos, zeros a esquerda (codigos) e barras (datas)."""
    if atual is None:
        return False
    a, d = atual.strip().upper(), str(desejado).strip().upper()
    if a == d:
        return True
    a2, d2 = a.replace("/", "").replace(".", ""), d.replace("/", "").replace(".", "")
    if a2.isdigit() and d2.isdigit():
        return a2.lstrip("0") == d2.lstrip("0") and (len(a2) == len(d2) or "/" not in a)
    try:  # numeros com virgula: 1,80000 == 1,80
        return abs(num(a) - num(d)) < 1e-9
    except ValueError:
        return False


def preencher(nome, texto, tab=True, espera=PAUSA, apagar=12, mensagem=False, verificar=True):
    """Clica no campo; se ja tiver o valor certo so da Tab, senao apaga e digita."""
    clicar(nome, 0.3)
    if verificar and LEITURA_OK is not False:
        atual = ler_campo()
        if mesmo_valor(atual, texto):
            acao(f"campo {nome} ja esta com '{atual}' -> so Tab")
            if tab:
                tecla("tab", permitir_mensagem=mensagem)
                esperar(espera)
            return
        acao(f"campo {nome} tinha {atual!r} -> vou digitar '{texto}'")
    tecla("end")
    tecla("backspace", apagar)
    tecla("delete", apagar)
    digitar(texto)
    if tab:
        tecla("tab", permitir_mensagem=mensagem)
    esperar(espera)


def data_digitos(d):
    return d.strftime("%d%m%Y")  # campo com mascara  /  /


def colar(nome, texto):
    pyperclip.copy(texto)
    clicar(nome, 0.3)
    acao("ctrl+a, ctrl+v (cola a coluna J)")
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("ctrl", "v")
    conferir_janela("colar a OBS")
    time.sleep(PAUSA)


def print_tela(ped):
    os.makedirs(PASTA_PRINTS, exist_ok=True)
    arq = os.path.join(PASTA_PRINTS, f"{ped['pedido']}_{dt.datetime.now():%Y%m%d_%H%M%S}.png")
    pyautogui.screenshot(arq)
    return arq


def em_inclusao():
    """True se o documento esta aberto em inclusao: o botao do canto superior direito
    vira 'Ok (F9)' com o icone de check verde (na lista ele e 'Novo (F2)', azul/branco)."""
    return tem_verde((1830, 126, 16, 14))


def subform_aberto():
    """True se o formulario de item/vencimento esta aberto (botao Ok verde no painel de baixo)."""
    return tem_verde(REGIAO_OK_BAIXO)


def abrir_subform():
    clicar("novo_baixo", PAUSA_TELA)
    for _ in range(40):  # espera ate +8 s o formulario abrir
        if subform_aberto():
            break
        esperar(0.2)
    if not subform_aberto():
        raise TelaErrada("Cliquei em Novo no painel de baixo mas o formulario nao abriu.")


def achar_janela(trecho_titulo):
    achadas = []
    win32gui.EnumWindows(lambda h, _: achadas.append(h) if trecho_titulo in win32gui.GetWindowText(h)
                         and win32gui.IsWindowVisible(h) else None, None)
    return achadas[0] if achadas else None


def responder_nao(motivo, espera_max=4.0, desistir=None):
    """Confirmacao do Fortes (recalcular tributos retidos? / gerar nota fiscal?) -> clica em Nao.
    O botao Nao fica a (+230, +92) do canto da janela 'Confirmação' (383 x 125).
    desistir: funcao opcional; se devolver True, para de esperar pela pergunta."""
    fim = time.time() + espera_max
    while time.time() < fim:
        h = achar_janela("Confirmação")
        if h:
            l, t, r, b = win32gui.GetWindowRect(h)
            acao(f"confirmacao '{motivo}' -> clique em Nao ({l + 230}, {t + 92})")
            pyautogui.click(l + 230, t + 92)
            esperar(PAUSA)
            for _ in range(50):  # o acesso remoto as vezes demora: espera ate +10 s a janela fechar
                if not achar_janela("Confirmação"):
                    break
                esperar(0.2)
            if achar_janela("Confirmação"):
                raise TelaErrada("Cliquei em Nao mas a Confirmacao continuou aberta.")
            return True
        if desistir is not None and desistir():
            return False
        esperar(0.2)
    return False


def fechar_subform(recalcular=False):
    acao("clique em ok_baixo (pode abrir confirmacao)")
    checar_parada()
    pyautogui.click(*P["ok_baixo"])
    if recalcular:
        # o Fortes as vezes demora varios segundos para perguntar; se o formulario ja fechou, nao vem pergunta
        responder_nao('recalcular tributos retidos?', espera_max=15.0,
                      desistir=lambda: not subform_aberto())
    esperar(PAUSA_TELA)
    confirmar_mensagem(tentativas=2)
    for _ in range(40):  # Fortes as vezes demora: espera ate +8 s o formulario fechar
        if not subform_aberto():
            break
        esperar(0.2)
    if subform_aberto():
        raise TelaErrada("Cliquei em Ok no item/vencimento mas o formulario continuou aberto "
                         "(algum campo pode estar invalido).")


def tem_verde(regiao):
    im = pyautogui.screenshot(region=regiao)
    for x in range(im.width):
        for y in range(im.height):
            r, g, b = im.getpixel((x, y))[:3]
            if g > r + 40 and g > b + 40:
                return True
    return False


class TelaErrada(Exception):
    pass


def texto_na_mensagem(ref, regiao, tolerancia=40):
    """Procura a imagem ref na regiao aceitando pequena diferenca de cor (o acesso
    remoto muda alguns tons, e a busca exata do pyscreeze falha)."""
    import numpy as np
    from PIL import Image
    tela = np.asarray(pyautogui.screenshot(region=regiao).convert("RGB"), dtype=np.int16)
    alvo = np.asarray(Image.open(ref).convert("RGB"), dtype=np.int16)
    h, w = alvo.shape[:2]
    H, W = tela.shape[:2]
    if h > H or w > W:
        return False
    # filtra posicoes pelo pixel do canto e depois confere a imagem inteira
    canto = np.abs(tela[:H - h + 1, :W - w + 1] - alvo[0, 0]).max(axis=2) <= tolerancia
    for y, x in zip(*np.nonzero(canto)):
        if np.abs(tela[y:y + h, x:x + w] - alvo).max() <= tolerancia:
            return True
    return False


def campo_st_da_mensagem():
    """Olha o texto da mensagem 'Atencao!' na frente: devolve item_sta / item_stb, ou None
    se for qualquer outra mensagem. O texto do STB nunca foi visto: vale o inicio 'Campo "ST'."""
    h = win32gui.GetForegroundWindow()
    l, t, r, b = win32gui.GetWindowRect(h)
    regiao = (l, t, r - l, b - t)
    if texto_na_mensagem(REF_MSG_STA, regiao):
        return "item_sta"
    if texto_na_mensagem(REF_MSG_CAMPO_ST, regiao):
        return "item_stb"
    if texto_na_mensagem(REF_MSG_NAO_CADASTRADO, regiao):
        return "item_cfop"
    return None


def aguardar_mensagem(segundos):
    fim = time.time() + segundos
    while time.time() < fim:
        tit = mensagem_do_fortes()
        if tit is not None:
            return tit
        esperar(0.1)
    return mensagem_do_fortes()


def tratar_st_em_branco():
    """Sai do bloco de situacao tributaria clicando em Unidade. Se o Fortes avisar
    'Campo "STA" em branco' -> Ok e digita 0; se avisar do campo seguinte -> Ok e digita 40.
    Se avisar 'Não cadastrado.' (CFOP) -> Ok e digita o CFOP 6949."""
    # A leitura do texto da mensagem as vezes confunde STA com STB (acesso remoto muda
    # os tons). Por isso, a qualquer aviso de campo ST em branco, preenche os DOIS campos.
    cfop_trocado = False
    for _ in range(4):
        if mensagem_do_fortes() is None:
            conferir_janela("(antes de sair do bloco de situacao tributaria)")
            checar_parada()
            acao("clique em item_unidade (confere STA/STB)")
            pyautogui.click(*P["item_unidade"])
        tit = aguardar_mensagem(4.0)  # acesso remoto lento: o aviso pode demorar 2-3 s
        if tit is None:
            return
        campo = campo_st_da_mensagem() if "Atenção" in tit else None
        if campo is None:
            raise TelaErrada(f"Apareceu a janela '{tit}' e nao sei o que responder - parei.")
        if campo == "item_cfop":
            if cfop_trocado:
                raise TelaErrada(f"O Fortes continuou dizendo 'Nao cadastrado' depois do CFOP {CFOP_NAO_CADASTRADO}.")
            print(f"[CFOP nao cadastrado -> {CFOP_NAO_CADASTRADO}] ", end="", flush=True)
            pyautogui.press("enter")
            esperar(PAUSA)
            if mensagem_do_fortes() is not None:
                raise TelaErrada("Apertei Ok mas a mensagem 'Nao cadastrado' continuou aberta.")
            preencher("item_cfop", CFOP_NAO_CADASTRADO, espera=PAUSA_LOOKUP, mensagem=True, verificar=False)
            cfop_trocado = True
            continue  # volta a clicar em Unidade e confere de novo
        print(f"[ST em branco -> STA {VALORES_ST['item_sta']} / STB {VALORES_ST['item_stb']}] ",
              end="", flush=True)
        pyautogui.press("enter")
        esperar(PAUSA)
        if mensagem_do_fortes() is not None:
            raise TelaErrada("Apertei Ok mas a mensagem do campo em branco continuou aberta.")
        for c in ("item_sta", "item_stb"):
            preencher(c, VALORES_ST[c], espera=PAUSA_LOOKUP, mensagem=True, verificar=False)
            if mensagem_do_fortes() is not None:
                break  # o Fortes reclamou de novo: o laco responde
    raise TelaErrada("O Fortes continuou reclamando de STA/STB depois de preencher.")


def abrir_e_preencher_cabecalho(ped):
    # Novo documento: precisa estar na LISTA de Contas a Receber - Documentos
    if em_inclusao():
        raise TelaErrada("Ja existe um documento aberto no Fortes. Grave ou cancele e rode de novo.")
    clicar("novo_documento", PAUSA_TELA)
    for _ in range(40):  # espera ate +8 s o documento abrir
        if em_inclusao():
            break
        esperar(0.2)
    if not em_inclusao():
        raise TelaErrada("Cliquei em Novo mas o documento nao abriu. Deixe o Fortes maximizado "
                         "na tela 'Contas a Receber - Documentos' (lista).")

    # Cabecalho
    preencher("estabelecimento", ESTABELECIMENTO, espera=PAUSA_LOOKUP)
    preencher("cliente", ped["cliente"], espera=PAUSA_LOOKUP, mensagem=True)
    confirmar_mensagem(espera_max=6.0)  # mensagem (limite de credito) depois do cliente -> Ok
    preencher("centro", CENTRO_RESULTADOS, espera=PAUSA_LOOKUP)
    preencher("receita", RECEITA, espera=PAUSA_LOOKUP)
    preencher("emissao", data_digitos(ped["emissao"]))
    colar("obs", ped["obs"])
    esperar(0.5)


def lancar(ped, salvar=True, itens_feitos=0):
    """itens_feitos > 0: retoma um documento que ja esta aberto, pulando o cabecalho e esses itens."""
    if itens_feitos:
        if not em_inclusao():
            raise TelaErrada("Mandei retomar, mas nao ha documento aberto no Fortes.")
        acao(f"retomando documento aberto, pulando {itens_feitos} item(ns)")
    else:
        abrir_e_preencher_cabecalho(ped)
    ajustar_painel()

    # Itens
    clicar("aba_itens", PAUSA_TELA / 2)
    for it in ped["itens"][itens_feitos:]:
        abrir_subform()
        preencher("item_codigo", it["codigo"], espera=PAUSA_LOOKUP)
        tratar_st_em_branco()
        preencher("item_unidade", it["unidade"])
        preencher("item_qtd", br(it["qtd"], 0) if it["qtd"].is_integer() else br(it["qtd"]))
        preencher("item_valor", br(it["preco"]))
        fechar_subform(recalcular=True)

    # Vencimento
    clicar("aba_venc", PAUSA_TELA / 2)
    abrir_subform()
    preencher("venc_data", data_digitos(vencimento(ped["emissao"])))
    preencher("venc_valor", br(ped["total"]))
    if AGENTE_COBRADOR:
        preencher("venc_agente", AGENTE_COBRADOR, espera=PAUSA_LOOKUP)
    if TIPO_DOCUMENTO:
        preencher("venc_tipo", TIPO_DOCUMENTO, espera=PAUSA_LOOKUP)
    fechar_subform()

    arq = print_tela(ped)
    if salvar:
        gravar_documento()
    return arq


def gravar_documento():
    """Ok (F9) do documento -> 'Deseja gerar nota fiscal?' -> Nao -> confere que fechou."""
    if not em_inclusao():
        raise TelaErrada("Nao ha documento aberto no Fortes para gravar.")
    if achar_janela("Confirmação"):
        acao("pergunta 'gerar nota fiscal?' ja estava aberta: nao clico em Ok de novo")
    else:
        conferir_janela("(antes do Ok final)")
        acao("clique em Ok (F9) do documento")
        checar_parada()
        pyautogui.click(*P["ok_documento"])
    if not responder_nao("gerar nota fiscal?", espera_max=30.0):
        raise TelaErrada("Depois do Ok nao apareceu a pergunta 'gerar nota fiscal?' - confira o Fortes.")
    for _ in range(150):  # espera ate 30 s o documento fechar e voltar para a lista
        esperar(0.2)
        if not em_inclusao() and fortes_na_frente():
            acao("documento gravado (voltou para a lista)")
            return
    frente = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    raise TelaErrada(f"Depois de responder Nao o documento nao fechou (janela da frente: '{frente}').")


# ----------------------------------------------------------------- PRINCIPAL
def mostrar(pedidos, problemas):
    for p in pedidos:
        marca = f"  (linha {p['linha']})"
        print(f"\nPedido {p['pedido']} | cliente {p['cliente']} {p['nome']} | emissao "
              f"{p['emissao']:%d/%m/%Y} | venc {vencimento(p['emissao']):%d/%m/%Y} | "
              f"total R$ {br(p['total'])} | {p['forma']}{marca}")
        for it in p["itens"]:
            print(f"    item {it['codigo']:>6} {it['unidade']} {br(it['qtd'], 0 if it['qtd'].is_integer() else 2):>7}"
                  f" x R$ {br(it['preco']):>8}  ({it['produto']})")
    if problemas:
        print("\n=== NAO SERAO LANCADOS (corrija na planilha) ===")
        for linha, ped, msg in problemas:
            print(f"  linha {linha} | pedido {ped} | {msg}")


def main():
    modo = sys.argv[1] if len(sys.argv) > 1 else "listar"
    global DIAGNOSTICO
    nome_wb, ws, pedidos, problemas = ler_pedidos()
    print(f"Planilha: {nome_wb} | {len(pedidos)} pedidos validos | {len(problemas)} com problema")

    if modo == "listar":
        mostrar(pedidos, problemas)
        print(f"\nTotal geral: R$ {br(sum(p['total'] for p in pedidos))}")
        return

    pendentes = pedidos
    linha_retomar, itens_feitos = None, 0
    if modo == "retomar":  # retomar <linha> <itens ja lancados>: termina o documento aberto e segue com os demais
        linha_retomar, itens_feitos = int(sys.argv[2]), int(sys.argv[3])
        if not any(p["linha"] == linha_retomar for p in pedidos):
            raise SystemExit(f"Nao achei pedido valido na linha {linha_retomar}.")
        pendentes = sorted(pedidos, key=lambda p: p["linha"] != linha_retomar)

    alvos = [] if modo == "retomar" else sys.argv[2:]  # uma ou mais linhas (ou nº de pedido), na ordem dada
    if alvos:
        pendentes = []
        for a in alvos:
            achou = [p for p in pedidos if str(p["linha"]) == a or p["pedido"] == a]
            if not achou:
                raise SystemExit(f"Nao achei pedido valido na linha/pedido {a} (J vazia ou com problema?).")
            pendentes.append(achou[0])
        DIAGNOSTICO = True  # lancamento escolhido a dedo: mostra cada acao

    if modo == "gravar":  # grava um documento ja preenchido pelo 'teste'
        ped = pendentes[0]
        ativar_fortes()
        try:
            gravar_documento()
        except (TelaErrada, Parada, pyautogui.FailSafeException) as e:
            os.makedirs(PASTA_PRINTS, exist_ok=True)
            arq = os.path.join(PASTA_PRINTS, f"erro_gravar_{dt.datetime.now():%H%M%S}.png")
            pyautogui.screenshot(arq)
            print(f"PAREI: {e or 'parado pelo usuario'}\nNAO apaguei a coluna J. Print: {arq}")
            return
        apagar_j(ws, ped["linha"])
        print(f"Gravado pedido {ped['pedido']} e apagada a coluna J da linha {ped['linha']}.")
        return

    if modo == "teste":
        pendentes = pendentes[:1]
        DIAGNOSTICO = True
    if not pendentes:
        print("Nada a lancar.")
        return

    print(f"\nVou lancar {len(pendentes)} pedido(s). Em 5 segundos comeco - NAO mexa no mouse/teclado.")
    print(">>> PARA PARAR NA HORA: aperte ESC (ou leve o mouse ao canto superior esquerdo) <<<")
    win32api.GetAsyncKeyState(win32con.VK_ESCAPE)  # zera um ESC antigo
    try:
        esperar(5)
    except Parada:
        print("PARADO pelo ESC.")
        return
    ativar_fortes()

    for n, ped in enumerate(pendentes, 1):
        print(f"[{n}/{len(pendentes)}] Linha {ped['linha']} pedido {ped['pedido']} - {ped['nome']} - "
              f"R$ {br(ped['total'])} ... ", end="", flush=True)
        try:
            arq = lancar(ped, salvar=(modo != "teste"),
                         itens_feitos=itens_feitos if ped["linha"] == linha_retomar else 0)
        except (pyautogui.FailSafeException, Parada):
            print("\nPARADO pelo usuario (ESC / mouse no canto). Confira o Fortes: o documento "
                  "ficou aberto sem gravar - cancele ou termine a mao.")
            return
        except TelaErrada as e:
            print(f"\nPAREI: {e}")
            print("Ultimas acoes: " + " | ".join(ULTIMAS_ACOES[-6:]))
            os.makedirs(PASTA_PRINTS, exist_ok=True)
            arq = os.path.join(PASTA_PRINTS, f"erro_{dt.datetime.now():%Y%m%d_%H%M%S}.png")
            pyautogui.screenshot(arq)
            print("Print do erro:", arq)
            return
        if modo == "teste":
            print(f"preenchido. NAO gravei.\nPrint: {arq}")
            return
        apagar_j(ws, ped["linha"])
        print("ok")
    print("\nConcluido.")


if __name__ == "__main__":
    main()
