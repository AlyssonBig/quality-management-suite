# Aplicativo da Qualidade
# Arquivo principal: main.py
# Dependencias: kivy e openpyxl  ->  pip install kivy openpyxl
#
# Interface sem logomarcas corporativas.
# O arquivo dados_qualidade.xlsx e criado automaticamente na pasta do aplicativo.

from pathlib import Path
from datetime import datetime
from uuid import uuid4
import os
import sys
import tempfile
import traceback

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from kivy.app import App
from kivy.clock import Clock
from kivy.metrics import dp, sp
from kivy.core.window import Window
from kivy.graphics import (Color, RoundedRectangle, Rectangle, Line, Ellipse,
                           Triangle)
from kivy.graphics.texture import Texture
from kivy.uix.anchorlayout import AnchorLayout
from kivy.uix.behaviors import ButtonBehavior
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner, SpinnerOption
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget
from kivy.utils import escape_markup


APP_TITLE = "Aplicativo da Qualidade"

# ----------------------------------------------------------------------------
# Identidade visual: roxo corporativo + branco
# ----------------------------------------------------------------------------
ROXO = (0.44, 0.10, 0.52, 1)
ROXO_ESCURO = (0.25, 0.06, 0.31, 1)
ROXO_MEDIO = (0.47, 0.13, 0.55, 1)
LILAS = (0.95, 0.91, 0.97, 1)
BORDA = (0.86, 0.80, 0.90, 1)
FUNDO = (0.968, 0.958, 0.978, 1)
BRANCO = (1, 1, 1, 1)
TEXTO = (0.17, 0.12, 0.20, 1)
CINZA = (0.45, 0.41, 0.48, 1)
CINZA_CLARO = (0.62, 0.58, 0.65, 1)
VERDE = (0.10, 0.50, 0.30, 1)
VERDE_BG = (0.90, 0.97, 0.93, 1)
AMBAR = (0.78, 0.47, 0.03, 1)
AMBAR_BG = (1.0, 0.96, 0.88, 1)
VERMELHO = (0.74, 0.14, 0.16, 1)
VERMELHO_BG = (0.99, 0.91, 0.91, 1)
AZUL = (0.16, 0.36, 0.72, 1)
AZUL_BG = (0.91, 0.94, 0.99, 1)

PASTA_APP = Path(__file__).resolve().parent
ARQUIVO_EXCEL = PASTA_APP / "dados_qualidade.xlsx"
# Cada aba recebe um conjunto de campos. As colunas de ID e data/hora
# sao preenchidas automaticamente pelo aplicativo.
MODULOS = {
    "Abertura de seleção": {
        "aba": "Abertura_Selecao",
        "prefixo": "SEL",
        "campos": [
            ("Responsável", "texto", True),
            ("Setor", "opcoes", True, ["Produção", "Qualidade", "Logística", "Manutenção", "PCP", "Outro"]),
            ("Ordem", "texto", True),
            ("Produto / código", "texto", True),
            ("Descrição", "texto", True),
            ("Quantidade", "numero", True),
            ("Motivo da seleção", "opcoes", True, ["Inspeção adicional", "Suspeita de não conformidade", "Solicitação da produção", "Solicitação da qualidade", "Outro"]),
            ("Máquina / processo", "texto", False),
            ("Prioridade", "opcoes", True, ["Baixa", "Normal", "Alta", "Urgente"]),
            ("Observação", "multilinha", False),
        ],
    },
    "Disposição de seleção": {
        "aba": "Disposicao_Selecao",
        "prefixo": "DSP",
        "campos": [
            ("ID da seleção (se houver)", "texto", False),
            ("Responsável", "texto", True),
            ("Ordem", "texto", True),
            ("Produto / código", "texto", True),
            ("Quantidade avaliada", "numero", True),
            ("Quantidade aprovada", "numero", True),
            ("Quantidade reprovada", "numero", True),
            ("Decisão", "opcoes", True, ["Liberar", "Reinspecionar", "Segregar", "Retrabalhar", "Devolver", "Aguardar decisão"]),
            ("Justificativa / motivo", "multilinha", True),
            ("Observação", "multilinha", False),
        ],
    },
    "Chamado para metrologia": {
        "aba": "Chamado_Metrologia",
        "prefixo": "MET",
        "campos": [
            ("Solicitante / inspetor", "texto", True),
            ("Setor", "opcoes", True, ["Produção", "Qualidade", "Logística", "Manutenção", "PCP", "Outro"]),
            ("Ordem", "texto", False),
            ("Máquina / equipamento", "texto", False),
            ("Instrumento / código", "texto", True),
            ("Tipo de solicitação", "opcoes", True, ["Calibração", "Verificação", "Manutenção", "Dúvida de medição", "Empréstimo", "Outro"]),
            ("Prioridade", "opcoes", True, ["Baixa", "Normal", "Alta", "Urgente"]),
            ("Descrição da necessidade", "multilinha", True),
            ("Evidência / caminho do arquivo", "texto", False),
            ("Status", "opcoes", True, ["Aberto", "Em atendimento", "Aguardando", "Concluído", "Cancelado"]),
        ],
    },
    "Relatório de bugs": {
        "aba": "Relatorio_Bugs",
        "prefixo": "BUG",
        "campos": [
            ("Nome do usuário", "texto", True),
            ("Setor / turno", "texto", True),
            ("Tablet / equipamento", "texto", False),
            ("Tela / módulo", "opcoes", True, ["Menu principal", "Abertura de seleção", "Disposição de seleção", "Chamado para metrologia", "Relatório do turno", "Fechamento de ordens", "Fechamento de turno", "Excel / salvamento", "Outro"]),
            ("Tipo de problema", "opcoes", True, ["Erro de tela", "Não salva", "Dados incorretos", "Lentidão", "Botão não funciona", "Sugestão de melhoria", "Outro"]),
            ("Descrição do problema", "multilinha", True),
            ("Evidência / caminho do arquivo", "texto", False),
            ("Impacto", "opcoes", True, ["Baixo", "Médio", "Alto", "Crítico"]),
            ("Prioridade", "opcoes", True, ["Baixa", "Normal", "Alta", "Urgente"]),
            ("Status", "opcoes", True, ["Aberto", "Em análise", "Em correção", "Resolvido", "Fechado"]),
        ],
    },
    "Relatório do turno": {
        "aba": "Relatorio_Turno",
        "prefixo": "TUR",
        "campos": [
            ("Responsável pelo relatório", "texto", True),
            ("Data do turno", "texto", True),
            ("Turno", "opcoes", True, ["1º turno", "2º turno", "3º turno", "Administrativo", "Outro"]),
            ("Setor / processo", "texto", True),
            ("Ordens trabalhadas", "multilinha", False),
            ("Seleções realizadas", "multilinha", False),
            ("Não conformidades / ocorrências", "multilinha", False),
            ("Chamados de metrologia", "multilinha", False),
            ("Bugs / problemas do aplicativo", "multilinha", False),
            ("Pendências para o próximo turno", "multilinha", True),
            ("Observações gerais", "multilinha", False),
        ],
    },
    "Fechamento de ordens": {
        "aba": "Fechamento_Ordens",
        "prefixo": "ORD",
        "campos": [
            ("Responsável", "texto", True),
            ("Ordem", "texto", True),
            ("Produto / código", "texto", True),
            ("Descrição", "texto", False),
            ("Quantidade produzida", "numero", False),
            ("Quantidade inspecionada", "numero", False),
            ("Resultado da inspeção", "opcoes", True, ["Aprovada", "Reprovada", "Aprovada com ressalva", "Pendente"]),
            ("Não conformidade / registro relacionado", "texto", False),
            ("Checklist - documentação conferida", "simnao", True),
            ("Checklist - inspeção concluída", "simnao", True),
            ("Checklist - registros atualizados", "simnao", True),
            ("Pendências / observações", "multilinha", False),
        ],
    },
    "Fechamento de turno": {
        "aba": "Fechamento_Turno",
        "prefixo": "FTR",
        "campos": [
            ("Responsável pelo fechamento", "texto", True),
            ("Data do turno", "texto", True),
            ("Turno", "opcoes", True, ["1º turno", "2º turno", "3º turno", "Administrativo", "Outro"]),
            ("Setor / processo", "texto", True),
            ("Ordens finalizadas", "multilinha", False),
            ("Ordens pendentes", "multilinha", False),
            ("Seleções em andamento", "multilinha", False),
            ("Equipamentos / instrumentos com pendência", "multilinha", False),
            ("Ocorrências relevantes", "multilinha", False),
            ("Passagem de turno realizada?", "opcoes", True, ["Sim", "Não", "Parcial"]),
            ("Responsável do próximo turno", "texto", False),
            ("Observações finais", "multilinha", False),
        ],
    },
}


def cabecalhos_do_modulo(config):
    return ["ID", "Data_Hora"] + [campo[0] for campo in config["campos"]]


def salvar_arquivo_seguro(wb):
    """Grava em arquivo temporario e troca no final: se algo falhar no meio,
    a planilha original nao fica corrompida. Se o Excel estiver com o arquivo
    aberto, o Windows recusa a troca e levantamos PermissionError."""
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=str(PASTA_APP))
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, ARQUIVO_EXCEL)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def estilizar_aba(ws, config):
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="164B73")
        celula.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_idx, titulo in enumerate(cabecalhos_do_modulo(config), start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(str(titulo)) + 3, 15), 32)


def preparar_planilha():
    """Cria o arquivo e as abas que faltam. Nunca apaga registros.
    So grava no disco quando algo realmente mudou (assim nao falha por o
    Excel estar aberto nem reescreve o arquivo a toda consulta).

    Se uma aba ja existir com cabecalho diferente do esperado (por exemplo,
    criada por outro script), ela e RENOMEADA para '<nome>_antiga' e uma aba
    nova e criada. Assim os dados antigos ficam guardados e as colunas novas
    nunca ficam desalinhadas."""
    alterado = False
    if ARQUIVO_EXCEL.exists():
        wb = load_workbook(ARQUIVO_EXCEL)
    else:
        wb = Workbook()
        if wb.sheetnames == ["Sheet"]:
            wb.remove(wb["Sheet"])
        alterado = True

    for nome, config in MODULOS.items():
        aba = config["aba"]
        esperado = cabecalhos_do_modulo(config)
        if aba in wb.sheetnames:
            ws = wb[aba]
            atual = [c.value for c in ws[1]] if ws.max_row >= 1 else []
            vazia = ws.max_row <= 1 and all(v is None for v in atual)
            if vazia:
                wb.remove(ws)
            elif atual[:len(esperado)] != esperado or len(atual) != len(esperado):
                novo_nome, n = f"{aba}_antiga", 1
                while novo_nome in wb.sheetnames:
                    n += 1
                    novo_nome = f"{aba}_antiga{n}"
                ws.title = novo_nome[:31]
                alterado = True
        if aba not in wb.sheetnames:
            ws = wb.create_sheet(aba)
            ws.append(esperado)
            estilizar_aba(ws, config)
            alterado = True

    if alterado:
        salvar_arquivo_seguro(wb)
    wb.close()


def gerar_id(prefixo):
    return f"{prefixo}-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:4].upper()}"


def salvar_registro(nome_modulo, valores):
    config = MODULOS[nome_modulo]
    preparar_planilha()
    wb = load_workbook(ARQUIVO_EXCEL)
    ws = wb[config["aba"]]
    novo_id = gerar_id(config["prefixo"])
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ws.append([novo_id, data_hora] + valores)
    for celula in ws[ws.max_row]:
        celula.alignment = Alignment(vertical="top", wrap_text=True)
    ws.auto_filter.ref = ws.dimensions
    salvar_arquivo_seguro(wb)
    wb.close()
    return novo_id, data_hora


def ler_registros(nome_modulo):
    """Devolve (cabecalho, linhas) da aba, sem gravar nada no disco."""
    config = MODULOS[nome_modulo]
    if not ARQUIVO_EXCEL.exists():
        preparar_planilha()
    wb = load_workbook(ARQUIVO_EXCEL, read_only=True, data_only=True)
    try:
        linhas = list(wb[config["aba"]].iter_rows(values_only=True))
    finally:
        wb.close()
    if not linhas:
        return cabecalhos_do_modulo(config), []
    return list(linhas[0]), [l for l in linhas[1:] if any(v not in (None, "") for v in l)]



# ============================================================================
# INTERFACE (visual roxo + branco)
# ============================================================================
DESCRICOES = {
    "Abertura de seleção": ("Registre uma nova seleção de produto.", "novo"),
    "Disposição de seleção": ("Informe a decisão da seleção.", "ok"),
    "Chamado para metrologia": ("Solicite atendimento de metrologia.", "metro"),
    "Relatório de bugs": ("Reporte problemas no sistema.", "bug"),
    "Relatório do turno": ("Envie o relatório do seu turno.", "relatorio"),
    "Fechamento de ordens": ("Finalize as ordens do turno.", "bandeira"),
    "Fechamento de turno": ("Encerrar o turno atual.", "power"),
}

# Campo que vira "etiqueta" colorida nos cartoes da consulta (o primeiro que existir).
CAMPOS_ETIQUETA = ["Decisão", "Status", "Resultado da inspeção",
                   "Passagem de turno realizada?", "Prioridade"]


def cor_etiqueta(valor):
    """Devolve (cor do texto, cor de fundo) conforme o significado do valor."""
    v = str(valor).strip().lower()
    if v in ("liberar", "aprovada", "concluído", "concluido", "resolvido", "fechado", "sim", "baixa", "baixo"):
        return VERDE, VERDE_BG
    if v in ("segregar", "reprovada", "crítico", "critico", "urgente", "alto", "alta", "devolver", "não", "nao", "cancelado"):
        return VERMELHO, VERMELHO_BG
    if v in ("reinspecionar", "retrabalhar", "aguardar decisão", "aguardando", "pendente", "parcial",
             "em atendimento", "em análise", "em correção", "aprovada com ressalva", "médio", "normal"):
        return AMBAR, AMBAR_BG
    return AZUL, AZUL_BG


# ----------------------------------------------------------------------------
# Blocos visuais reutilizaveis
# ----------------------------------------------------------------------------
class Painel(BoxLayout):
    """Caixa com fundo arredondado, borda e sombra suave."""

    def __init__(self, bg=BRANCO, borda=None, raio=14, sombra=True, **kwargs):
        super().__init__(**kwargs)
        self.raio = raio
        self.bg_normal = bg
        with self.canvas.before:
            Color(0.25, 0.06, 0.31, 0.07 if sombra else 0)
            self._sombra = RoundedRectangle(radius=[dp(raio)])
            self._cor_bg = Color(*bg)
            self._fundo = RoundedRectangle(radius=[dp(raio)])
            self._cor_borda = Color(*(borda or (0, 0, 0, 0)))
            self._linha = Line(rounded_rectangle=(0, 0, 1, 1, dp(raio)), width=1)
        self.bind(pos=self._atualizar, size=self._atualizar)
        self._atualizar()

    def _atualizar(self, *_):
        self._fundo.pos = self.pos
        self._fundo.size = self.size
        self._sombra.pos = (self.x, self.y - dp(2))
        self._sombra.size = self.size
        self._linha.rounded_rectangle = (self.x, self.y, self.width, self.height, dp(self.raio))

    def definir_bg(self, cor):
        self._cor_bg.rgba = cor


class Gradiente(BoxLayout):
    """Faixa roxa com degrade, cantos de baixo arredondados e formas suaves."""

    def __init__(self, raio_baixo=26, **kwargs):
        super().__init__(**kwargs)
        self.raio_baixo = raio_baixo
        n = 64
        tex = Texture.create(size=(1, n), colorfmt="rgba")
        dados = []
        for i in range(n):          # i = 0 e a base; i = n-1 e o topo
            t = i / (n - 1)
            for a, b in zip(ROXO_MEDIO[:3], ROXO_ESCURO[:3]):
                dados.append(int((a + (b - a) * t) * 255))
            dados.append(255)
        tex.blit_buffer(bytes(dados), colorfmt="rgba", bufferfmt="ubyte")
        tex.mag_filter = "linear"
        tex.min_filter = "linear"
        with self.canvas.before:
            Color(1, 1, 1, 1)
            self._fundo = RoundedRectangle(texture=tex,
                                           radius=[0, 0, dp(raio_baixo), dp(raio_baixo)])
            Color(1, 1, 1, 0.05)
            self._t1 = Triangle()
            Color(1, 1, 1, 0.035)
            self._t2 = Triangle()
        self.bind(pos=self._atualizar, size=self._atualizar)

    def _atualizar(self, *_):
        x, y, w, h = self.x, self.y, self.width, self.height
        self._fundo.pos = (x, y)
        self._fundo.size = (w, h)
        self._t1.points = [x + w * 0.55, y + h, x + w, y + h, x + w, y + h * 0.35]
        self._t2.points = [x + w * 0.72, y + h, x + w, y + h, x + w, y + h * 0.7]


class Icone(Widget):
    """Icones desenhados em vetor (nao dependem de fontes nem de imagens)."""

    def __init__(self, nome, cor=ROXO, **kwargs):
        kwargs.setdefault("size_hint", (None, None))
        kwargs.setdefault("size", (dp(24), dp(24)))
        super().__init__(**kwargs)
        self.nome = nome
        self.cor = cor
        self.bind(pos=self.redesenhar, size=self.redesenhar)
        self.redesenhar()

    def definir_cor(self, cor):
        self.cor = cor
        self.redesenhar()

    def redesenhar(self, *_):
        self.canvas.clear()
        s = min(self.width, self.height)
        if s <= 0:
            return
        x0, y0 = self.center_x - s / 2, self.center_y - s / 2
        tw = max(dp(1.3), s * 0.065)

        def P(a, b):
            return (x0 + a * s, y0 + b * s)

        def pts(*pares):
            out = []
            for a, b in pares:
                out.extend(P(a, b))
            return out

        def linha(*pares, fechar=False):
            Line(points=pts(*pares), width=tw, cap="round", joint="round", close=fechar)

        def caixa(a, b, w, h, r=0.06):
            x, y = P(a, b)
            Line(rounded_rectangle=(x, y, w * s, h * s, r * s), width=tw)

        with self.canvas:
            Color(*self.cor)
            n = self.nome
            if n == "novo":
                caixa(0.18, 0.08, 0.52, 0.84)
                linha((0.30, 0.68), (0.58, 0.68))
                linha((0.30, 0.54), (0.58, 0.54))
                linha((0.30, 0.40), (0.46, 0.40))
                Ellipse(pos=P(0.56, 0.04), size=(0.42 * s, 0.42 * s))
                Color(1, 1, 1, 1)
                linha((0.67, 0.25), (0.87, 0.25))
                linha((0.77, 0.15), (0.77, 0.35))
            elif n == "ok":
                Line(circle=(self.center_x, self.center_y, 0.40 * s), width=tw)
                linha((0.30, 0.50), (0.44, 0.36), (0.71, 0.66))
            elif n == "metro":
                caixa(0.06, 0.28, 0.88, 0.44, 0.07)
                for i, a in enumerate((0.20, 0.34, 0.48, 0.62, 0.76)):
                    linha((a, 0.72), (a, 0.56 if i % 2 == 0 else 0.62))
            elif n == "bug":
                Ellipse(pos=P(0.31, 0.12), size=(0.38 * s, 0.50 * s))
                Ellipse(pos=P(0.38, 0.60), size=(0.24 * s, 0.22 * s))
                for lado in (1, -1):
                    c = 0.5
                    linha((c + lado * 0.17, 0.50), (c + lado * 0.36, 0.58))
                    linha((c + lado * 0.19, 0.38), (c + lado * 0.40, 0.38))
                    linha((c + lado * 0.17, 0.26), (c + lado * 0.34, 0.16))
                    linha((c + lado * 0.07, 0.80), (c + lado * 0.14, 0.92))
            elif n == "relatorio":
                caixa(0.20, 0.06, 0.60, 0.80)
                RoundedRectangle(pos=P(0.35, 0.80), size=(0.30 * s, 0.15 * s), radius=[dp(2)])
                linha((0.32, 0.62), (0.68, 0.62))
                linha((0.32, 0.48), (0.68, 0.48))
                linha((0.32, 0.34), (0.54, 0.34))
            elif n == "bandeira":
                linha((0.22, 0.08), (0.22, 0.94))
                linha((0.22, 0.90), (0.88, 0.90), (0.88, 0.48), (0.22, 0.48), fechar=True)
                q = 0.165
                for i in range(4):
                    for j in range(2):
                        if (i + j) % 2 == 0:
                            Rectangle(pos=P(0.22 + i * q, 0.48 + j * 0.21), size=(q * s, 0.21 * s))
            elif n == "power":
                Line(circle=(self.center_x, y0 + 0.46 * s, 0.34 * s, 38, 322), width=tw, cap="round")
                linha((0.50, 0.50), (0.50, 0.92))
            elif n == "voltar":
                linha((0.64, 0.12), (0.30, 0.50), (0.64, 0.88))
            elif n == "seta":
                linha((0.36, 0.14), (0.70, 0.50), (0.36, 0.86))
            elif n == "baixo":
                linha((0.14, 0.62), (0.50, 0.28), (0.86, 0.62))
            elif n == "casa":
                linha((0.08, 0.50), (0.50, 0.90), (0.92, 0.50))
                linha((0.22, 0.52), (0.22, 0.10), (0.78, 0.10), (0.78, 0.52))
                linha((0.42, 0.10), (0.42, 0.34), (0.58, 0.34), (0.58, 0.10))
            elif n == "busca":
                Line(circle=(x0 + 0.44 * s, y0 + 0.56 * s, 0.30 * s), width=tw)
                linha((0.66, 0.34), (0.90, 0.10))
            elif n == "relogio":
                Line(circle=(self.center_x, self.center_y, 0.40 * s), width=tw)
                linha((0.50, 0.50), (0.50, 0.76))
                linha((0.50, 0.50), (0.68, 0.38))
            elif n == "pasta":
                linha((0.08, 0.20), (0.08, 0.80), (0.40, 0.80), (0.50, 0.66),
                      (0.92, 0.66), (0.92, 0.20), fechar=True)
            elif n == "salvar":
                caixa(0.12, 0.10, 0.76, 0.80, 0.07)
                caixa(0.30, 0.60, 0.40, 0.30, 0.02)
                linha((0.28, 0.10), (0.28, 0.36), (0.72, 0.36), (0.72, 0.10))
            elif n == "limpar":
                linha((0.22, 0.22), (0.78, 0.78))
                linha((0.22, 0.78), (0.78, 0.22))


class IconeBotao(ButtonBehavior, Icone):
    def __init__(self, nome, **kwargs):
        super().__init__(nome=nome, **kwargs)


class Texto(Label):
    """Label com quebra de linha automatica e altura que acompanha o texto."""

    def __init__(self, texto="", tam=14, cor=TEXTO, negrito=False, halign="left", **kwargs):
        kwargs.setdefault("size_hint_y", None)
        super().__init__(text=texto, markup=True, font_size=sp(tam), color=cor,
                         bold=negrito, halign=halign, valign="middle", **kwargs)
        self.bind(width=lambda i, w: setattr(i, "text_size", (w, None)))
        self.bind(texture_size=lambda i, t: setattr(i, "height", t[1]))


class Cartao(ButtonBehavior, Painel):
    """Painel clicavel com efeito ao tocar."""

    def __init__(self, bg_press=None, **kwargs):
        super().__init__(**kwargs)
        self.bg_press = bg_press or LILAS

    def on_state(self, _inst, valor):
        self.definir_bg(self.bg_press if valor == "down" else self.bg_normal)


class Botao(ButtonBehavior, Painel):
    """Botao arredondado com icone opcional. estilo: primario | escuro | contorno."""

    def __init__(self, texto, callback=None, icone=None, estilo="primario", altura=50, **kwargs):
        if estilo == "contorno":
            bg, borda, fg, press = BRANCO, ROXO, ROXO, LILAS
        elif estilo == "escuro":
            bg, borda, fg, press = ROXO_ESCURO, None, BRANCO, ROXO
        else:
            bg, borda, fg, press = ROXO, None, BRANCO, ROXO_ESCURO
        kwargs.setdefault("size_hint_y", None)
        kwargs.setdefault("height", dp(altura))
        super().__init__(bg=bg, borda=borda, raio=12, **kwargs)
        self.bg_press = press
        centro = AnchorLayout(anchor_x="center", anchor_y="center")
        linha = BoxLayout(size_hint=(None, None), height=dp(24), spacing=dp(8))
        linha.bind(minimum_width=linha.setter("width"))
        if icone:
            linha.add_widget(Icone(icone, cor=fg, size=(dp(20), dp(20)),
                                   pos_hint={"center_y": 0.5}))
        rotulo = Label(text=texto, font_size=sp(15), bold=True, color=fg, size_hint_x=None)
        rotulo.bind(texture_size=lambda i, t: setattr(i, "width", t[0]))
        linha.add_widget(rotulo)
        centro.add_widget(linha)
        self.add_widget(centro)
        if callback:
            self.bind(on_release=callback)

    def on_state(self, _inst, valor):
        self.definir_bg(self.bg_press if valor == "down" else self.bg_normal)


class Chip(ButtonBehavior, Painel):
    def __init__(self, texto, selecionado=False, **kwargs):
        super().__init__(bg=ROXO if selecionado else BRANCO,
                         borda=ROXO if selecionado else BORDA, raio=18, sombra=False,
                         size_hint=(None, None), height=dp(36), **kwargs)
        rotulo = Label(text=texto, font_size=sp(13), bold=selecionado,
                       color=BRANCO if selecionado else CINZA, size_hint_x=None)
        rotulo.bind(texture_size=lambda i, t: (setattr(i, "width", t[0]),
                                                setattr(self, "width", t[0] + dp(32))))
        self.padding = [dp(16), 0]
        self.add_widget(rotulo)


def criar_etiqueta(texto, fg, bg):
    caixa = Painel(bg=bg, borda=fg, raio=10, sombra=False, size_hint=(None, None),
                   height=dp(24), padding=[dp(10), 0])
    rotulo = Label(text=str(texto), font_size=sp(11), bold=True, color=fg, size_hint_x=None)
    rotulo.bind(texture_size=lambda i, t: (setattr(i, "width", t[0]),
                                            setattr(caixa, "width", t[0] + dp(20))))
    caixa.add_widget(rotulo)
    return caixa


class Divisor(Widget):
    def __init__(self, cor=BORDA, **kwargs):
        kwargs.setdefault("size_hint_y", None)
        kwargs.setdefault("height", dp(1))
        super().__init__(**kwargs)
        with self.canvas:
            Color(*cor)
            self._r = Rectangle()
        self.bind(pos=lambda *_: setattr(self._r, "pos", self.pos),
                  size=lambda *_: setattr(self._r, "size", self.size))


class NavItem(ButtonBehavior, BoxLayout):
    def __init__(self, texto, icone, ativo=False, **kwargs):
        super().__init__(orientation="vertical", padding=[0, dp(7), 0, dp(4)], spacing=dp(2), **kwargs)
        cor = ROXO if ativo else CINZA_CLARO
        self.add_widget(ancorar(Icone(icone, cor=cor, size=(dp(24), dp(24)))))
        self.add_widget(Label(text=texto, font_size=sp(11), bold=ativo, color=cor,
                              size_hint_y=None, height=dp(16)))


class StyledInput(TextInput):
    def __init__(self, **kwargs):
        multi = kwargs.get("multiline", False)
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_active", "")
        kwargs.setdefault("background_disabled_normal", "")
        kwargs.setdefault("background_color", (0, 0, 0, 0))
        kwargs.setdefault("foreground_color", TEXTO)
        kwargs.setdefault("hint_text_color", (0.56, 0.52, 0.60, 1))
        kwargs.setdefault("cursor_color", ROXO)
        kwargs.setdefault("selection_color", (0.44, 0.10, 0.52, 0.25))
        kwargs.setdefault("font_size", sp(15))
        kwargs.setdefault("write_tab", False)
        super().__init__(**kwargs)
        self._multi = multi
        with self.canvas.before:
            Color(*BRANCO)
            self._bg = RoundedRectangle(radius=[dp(10)])
            self._cor_borda = Color(*BORDA)
            self._borda = Line(rounded_rectangle=(0, 0, 1, 1, dp(10)), width=1)
        self.bind(pos=self._atualizar, size=self._atualizar, focus=self._foco)
        self._atualizar()

    def _atualizar(self, *_):
        self._bg.pos = self.pos
        self._bg.size = self.size
        self._borda.rounded_rectangle = (self.x, self.y, self.width, self.height, dp(10))
        if not self._multi:
            pad_y = max(dp(6), (self.height - self.font_size * 1.25) / 2)
            self.padding = [self.padding[0], pad_y, self.padding[2], pad_y]

    def _foco(self, _inst, valor):
        self._cor_borda.rgba = ROXO if valor else BORDA
        self._borda.width = 1.6 if valor else 1


class OpcaoSpinner(SpinnerOption):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = BRANCO
        self.color = TEXTO
        self.font_size = sp(15)
        self.size_hint_y = None
        self.height = dp(46)
        self.halign = "left"
        self.valign = "middle"
        self.bind(size=lambda i, s: setattr(i, "text_size", (s[0] - dp(24), s[1])))
        self.bind(state=lambda i, v: setattr(i, "background_color", LILAS if v == "down" else BRANCO))


class StyledSpinner(Spinner):
    def __init__(self, **kwargs):
        kwargs.setdefault("background_normal", "")
        kwargs.setdefault("background_down", "")
        kwargs.setdefault("background_color", (0, 0, 0, 0))
        kwargs.setdefault("option_cls", OpcaoSpinner)
        kwargs.setdefault("font_size", sp(15))
        kwargs.setdefault("color", CINZA_CLARO)
        super().__init__(**kwargs)
        self.halign = "left"
        self.valign = "middle"
        self.bind(size=lambda i, s: setattr(i, "text_size", (s[0] - dp(36), s[1])))
        with self.canvas.before:
            Color(*BRANCO)
            self._bg = RoundedRectangle(radius=[dp(10)])
            Color(*BORDA)
            self._borda = Line(rounded_rectangle=(0, 0, 1, 1, dp(10)), width=1)
            Color(*ROXO)
            self._seta = Line(width=dp(1.6), cap="round", joint="round")
        self._retrabalhar = Clock.create_trigger(self._atualizar)
        self.bind(pos=self._retrabalhar, size=self._retrabalhar)
        self._atualizar()

    def _atualizar(self, *_):
        self._bg.pos = self.pos
        self._bg.size = self.size
        self._borda.rounded_rectangle = (self.x, self.y, self.width, self.height, dp(10))
        cx, cy = self.x + self.width - dp(22), self.y + self.height / 2
        self._seta.points = [cx - dp(5), cy + dp(2.5), cx, cy - dp(2.5), cx + dp(5), cy + dp(2.5)]

    def on_text(self, _inst, valor):
        self.color = CINZA_CLARO if valor == "Selecione..." else TEXTO


class CampoFormulario(BoxLayout):
    def __init__(self, definicao, **kwargs):
        super().__init__(orientation="vertical", spacing=dp(5), size_hint_y=None, **kwargs)
        self.definicao = definicao
        self.nome = definicao[0]
        tipo = definicao[1]
        obrigatorio = definicao[2]
        titulo = escape_markup(self.nome)
        if obrigatorio:
            titulo += " [color=#B8232A]*[/color]"
        rotulo = Label(text=titulo, markup=True, size_hint_y=None, height=dp(22),
                       halign="left", valign="middle", color=TEXTO, bold=True,
                       font_size=sp(13.5))
        rotulo.bind(size=lambda inst, tam: setattr(inst, "text_size", tam))
        self.add_widget(rotulo)
        if tipo == "opcoes":
            self.entrada = StyledSpinner(text="Selecione...", values=definicao[3],
                                         size_hint_y=None, height=dp(46))
            altura_entrada = 46
        elif tipo == "simnao":
            self.entrada = StyledSpinner(text="Selecione...", values=["Sim", "Não"],
                                         size_hint_y=None, height=dp(46))
            altura_entrada = 46
        else:
            multi = tipo == "multilinha"
            altura_entrada = 92 if multi else 46
            self.entrada = StyledInput(
                hint_text="Digite aqui..." if not multi else "Descreva aqui...",
                multiline=multi,
                input_filter="int" if tipo == "numero" else None,
                size_hint_y=None, height=dp(altura_entrada),
                padding=[dp(12), dp(10), dp(12), dp(10)],
            )
        self.height = dp(22 + 5 + altura_entrada)
        self.padrao = self.valor_padrao()
        self.limpar()
        self.add_widget(self.entrada)

    def valor_padrao(self):
        # "Data do turno" ja vem preenchida com a data de hoje (pode ser editada).
        if self.nome == "Data do turno":
            return datetime.now().strftime("%d/%m/%Y")
        return ""

    def obter_valor(self):
        if isinstance(self.entrada, Spinner):
            return "" if self.entrada.text == "Selecione..." else self.entrada.text
        return self.entrada.text.strip()

    def limpar(self):
        if isinstance(self.entrada, Spinner):
            self.entrada.text = "Selecione..."
        else:
            self.entrada.text = self.padrao


class CartaoRegistro(ButtonBehavior, Painel):
    """Cartao de um registro salvo; toque para ver todos os campos."""

    def __init__(self, cabecalho, linha, **kwargs):
        super().__init__(orientation="vertical", raio=14, borda=BORDA,
                         padding=[dp(20), dp(12), dp(14), dp(12)], spacing=dp(6),
                         size_hint_y=None, **kwargs)
        self.bind(minimum_height=self.setter("height"))
        self.expandido = False
        self.cabecalho = [str(c) for c in cabecalho]
        self.linha = list(linha)
        with self.canvas.before:
            Color(*ROXO)
            self._barra = RoundedRectangle(radius=[dp(14), 0, 0, dp(14)])
        self.bind(pos=self._barra_atualizar, size=self._barra_atualizar)
        self._barra_atualizar()

        valores = dict(zip(self.cabecalho, self.linha))
        self.campo_etiqueta = next((c for c in CAMPOS_ETIQUETA if valores.get(c) not in (None, "")), None)

        topo = BoxLayout(size_hint_y=None, height=dp(28), spacing=dp(10))
        id_txt = Label(text=f"[b]{escape_markup(str(valores.get('ID', '')))}[/b]", markup=True,
                       font_size=sp(14), color=ROXO, size_hint_x=None, halign="left", valign="middle")
        id_txt.bind(texture_size=lambda i, t: setattr(i, "width", t[0]))
        topo.add_widget(id_txt)
        if self.campo_etiqueta:
            fg, bg = cor_etiqueta(valores[self.campo_etiqueta])
            topo.add_widget(criar_etiqueta(valores[self.campo_etiqueta], fg, bg))
            topo.children[0].pos_hint = {"center_y": 0.5}
        topo.add_widget(Widget())
        data = Label(text=escape_markup(str(valores.get("Data_Hora", ""))), markup=True,
                     font_size=sp(12), color=CINZA, size_hint_x=None, halign="right", valign="middle")
        data.bind(texture_size=lambda i, t: setattr(i, "width", t[0]))
        topo.add_widget(data)
        self.add_widget(topo)

        self.detalhes = Texto("", tam=13.5, cor=TEXTO)
        self.add_widget(self.detalhes)
        self.dica = Texto("", tam=11, cor=CINZA_CLARO)
        self.add_widget(self.dica)
        self.atualizar_texto()

    def _barra_atualizar(self, *_):
        self._barra.pos = self.pos
        self._barra.size = (dp(6), self.height)

    def atualizar_texto(self):
        ignorar = {"ID", "Data_Hora", self.campo_etiqueta}
        itens = [(c, v) for c, v in zip(self.cabecalho, self.linha)
                 if c not in ignorar and v not in (None, "")]
        if not self.expandido:
            itens_vis = itens[:4]
        else:
            itens_vis = itens
        partes = []
        for c, v in itens_vis:
            txt = str(v)
            if not self.expandido and len(txt) > 70:
                txt = txt[:67] + "..."
            partes.append(f"[color=#736A7A]{escape_markup(c)}:[/color]  {escape_markup(txt)}")
        self.detalhes.text = "\n".join(partes) if partes else "Sem detalhes."
        if len(itens) > 4 or any(len(str(v)) > 70 for _, v in itens):
            self.dica.text = "Toque para recolher" if self.expandido else "Toque para ver todos os campos"
        else:
            self.dica.text = ""
            self.dica.height = 0

    def on_release(self):
        self.expandido = not self.expandido
        self.atualizar_texto()


def ancorar(filho, **kwargs):
    """AnchorLayout que ja contem 'filho' (o parametro children nao adiciona de verdade)."""
    a = AnchorLayout(**kwargs)
    a.add_widget(filho)
    return a


def centralizado(filho, largura_max=900):
    """Centraliza 'filho' com largura maxima (bom para tablets e telas largas)."""
    ancora = AnchorLayout(anchor_x="center", anchor_y="top", size_hint_y=None,
                          padding=[dp(14), dp(6), dp(14), dp(6)])
    filho.size_hint_x = None

    def ajustar(*_):
        filho.width = max(dp(200), min(ancora.width - dp(28), dp(largura_max)))
    ancora.bind(width=ajustar)
    filho.bind(height=lambda i, h: setattr(ancora, "height", h + dp(12)))
    ajustar()
    ancora.add_widget(filho)
    ancora.height = filho.height + dp(12)
    return ancora


class QualidadeApp(App):
    def build(self):
        self.title = APP_TITLE
        Window.minimum_width = 720
        Window.minimum_height = 560
        Window.clearcolor = FUNDO
        try:
            preparar_planilha()
        except Exception:
            pass
        self.modulo_consulta = list(MODULOS.keys())[0]
        self.raiz = BoxLayout(orientation="vertical")
        self.mostrar_menu()
        return self.raiz

    def limpar_tela(self):
        self.raiz.clear_widgets()

    # ------------------------------------------------------------------ cabecalhos
    def faixa_logos(self, altura):
        linha = BoxLayout(size_hint=(None, None), height=altura, spacing=dp(8))
        linha.bind(minimum_width=linha.setter("width"))
        linha.add_widget(Texto("Gestão da Qualidade", tam=20, cor=BRANCO, negrito=True))
        return linha

    def cabecalho_menu(self):
        hero = Gradiente(orientation="vertical", size_hint_y=None, height=dp(232),
                         padding=[0, dp(16), 0, dp(14)])
        coluna = BoxLayout(orientation="vertical", spacing=dp(4))
        coluna.add_widget(self.faixa_logos(dp(48)))
        coluna.add_widget(Widget(size_hint_y=None, height=dp(8)))
        coluna.add_widget(Texto("Aplicativo da", tam=22, cor=BRANCO))
        coluna.add_widget(Texto("Qualidade", tam=40, cor=BRANCO, negrito=True))
        coluna.add_widget(Texto("Gestão e controle de processos", tam=17, cor=(1, 1, 1, 0.92)))
        coluna.add_widget(Widget(size_hint_y=None, height=dp(2)))
        coluna.add_widget(Texto("Mais agilidade, organização e controle para os processos da qualidade.",
                                tam=13.5, cor=(1, 1, 1, 0.85)))
        coluna.size_hint_y = 1
        # Conteudo da faixa fica alinhado ao mesmo miolo do restante da tela.
        coluna.size_hint_x = None
        coluna.bind(minimum_height=lambda *_: None)
        hero.add_widget(ancorar(coluna, anchor_x="center", anchor_y="top",
                                padding=[dp(18), 0, dp(18), 0]))

        def ajustar(*_):
            coluna.width = max(dp(200), min(hero.width - dp(36), dp(900)))
        hero.bind(width=ajustar)
        ajustar()
        return hero

    def cabecalho_interno(self):
        barra = Gradiente(orientation="horizontal", size_hint_y=None, height=dp(78),
                          padding=[dp(14), dp(10), dp(14), dp(14)], raio_baixo=22)
        voltar = IconeBotao("voltar", cor=BRANCO, size=(dp(30), dp(30)), pos_hint={"center_y": 0.5})
        voltar.bind(on_release=self.mostrar_menu)
        casa = IconeBotao("casa", cor=BRANCO, size=(dp(30), dp(30)), pos_hint={"center_y": 0.5})
        casa.bind(on_release=self.mostrar_menu)
        esq = ancorar(voltar, anchor_x="left", size_hint_x=None, width=dp(48))
        dire = ancorar(casa, anchor_x="right", size_hint_x=None, width=dp(48))
        centro = ancorar(self.faixa_logos(dp(40)), anchor_x="center", anchor_y="center")
        barra.add_widget(esq)
        barra.add_widget(centro)
        barra.add_widget(dire)
        return barra

    def titulo_tela(self, titulo, subtitulo, icone):
        linha = BoxLayout(size_hint_y=None, height=dp(62), spacing=dp(14))
        quadro = Painel(bg=ROXO_ESCURO, raio=14, sombra=False, size_hint=(None, None),
                        size=(dp(54), dp(54)), pos_hint={"center_y": 0.5})
        quadro.add_widget(ancorar(Icone(icone, cor=BRANCO, size=(dp(30), dp(30)))))
        linha.add_widget(quadro)
        textos = BoxLayout(orientation="vertical", spacing=dp(2))
        textos.add_widget(Texto(titulo, tam=23, cor=ROXO_ESCURO, negrito=True))
        textos.add_widget(Texto(subtitulo, tam=13.5, cor=CINZA))
        linha.add_widget(textos)
        return centralizado(linha)

    def barra_nav(self, ativo):
        self.raiz.add_widget(Divisor())
        barra = Painel(bg=BRANCO, raio=0, sombra=False, size_hint_y=None, height=dp(62))
        itens = [("Início", "casa", "inicio", self.mostrar_menu),
                 ("Consultar", "busca", "consulta", self.mostrar_consulta),
                 ("Pasta", "pasta", "pasta", self.abrir_pasta)]
        for texto, icone, chave, acao in itens:
            item = NavItem(texto, icone, ativo=(chave == ativo))
            item.bind(on_release=acao)
            barra.add_widget(item)
        self.raiz.add_widget(barra)

    # ------------------------------------------------------------------ tela: menu
    def cartao_menu(self, nome, descricao, icone):
        card = Cartao(orientation="horizontal", borda=BORDA, raio=16,
                      padding=[dp(16), dp(14), dp(12), dp(14)], spacing=dp(6))
        coluna = BoxLayout(orientation="vertical", spacing=dp(6))
        coluna.add_widget(Icone(icone, cor=ROXO, size=(dp(38), dp(38))))
        coluna.add_widget(Widget())
        coluna.add_widget(Texto(nome, tam=15.5, cor=ROXO_ESCURO, negrito=True))
        coluna.add_widget(Texto(descricao, tam=12, cor=CINZA))
        card.add_widget(coluna)
        seta = ancorar(Icone("seta", cor=ROXO_ESCURO, size=(dp(20), dp(20))),
                       anchor_x="center", anchor_y="center", size_hint_x=None, width=dp(24))
        card.add_widget(seta)
        card.bind(on_release=lambda *_: self.mostrar_formulario(nome))
        return card

    def mostrar_menu(self, *_):
        self.limpar_tela()
        self.raiz.add_widget(self.cabecalho_menu())

        coluna = BoxLayout(orientation="vertical", spacing=dp(12), size_hint_y=None)
        coluna.bind(minimum_height=coluna.setter("height"))
        grade = GridLayout(cols=2, spacing=dp(12), size_hint_y=None,
                           row_default_height=dp(126), row_force_default=True)
        grade.bind(minimum_height=grade.setter("height"))
        nomes = list(MODULOS.keys())
        for nome in nomes[:-1]:
            descricao, icone = DESCRICOES.get(nome, ("Acessar rotina da qualidade.", "novo"))
            grade.add_widget(self.cartao_menu(nome, descricao, icone))
        coluna.add_widget(grade)

        ultimo = nomes[-1]
        descricao, icone = DESCRICOES[ultimo]
        faixa = Cartao(bg=ROXO, bg_press=ROXO_ESCURO, raio=16, sombra=True, size_hint_y=None,
                       height=dp(70), padding=[dp(20), 0, dp(18), 0], spacing=dp(16))
        faixa.add_widget(Icone(icone, cor=BRANCO, size=(dp(32), dp(32)), pos_hint={"center_y": 0.5}))
        textos = BoxLayout(orientation="vertical", padding=[0, dp(12)])
        textos.add_widget(Texto(ultimo, tam=16, cor=BRANCO, negrito=True))
        textos.add_widget(Texto(descricao, tam=12, cor=(1, 1, 1, 0.8)))
        faixa.add_widget(textos)
        faixa.add_widget(Icone("seta", cor=BRANCO, size=(dp(20), dp(20)), pos_hint={"center_y": 0.5}))
        faixa.bind(on_release=lambda *_: self.mostrar_formulario(ultimo))
        coluna.add_widget(faixa)

        rolagem = ScrollView(do_scroll_x=False, bar_width=dp(4), bar_color=(0.44, 0.10, 0.52, 0.4))
        rolagem.add_widget(centralizado(coluna, 920))
        self.raiz.add_widget(rolagem)
        self.barra_nav("inicio")

    # ------------------------------------------------------------------ tela: formulario
    def mostrar_formulario(self, nome_modulo, *_):
        self.limpar_tela()
        config = MODULOS[nome_modulo]
        descricao, icone = DESCRICOES.get(nome_modulo, ("", "novo"))
        self.raiz.add_widget(self.cabecalho_interno())
        self.raiz.add_widget(Widget(size_hint_y=None, height=dp(8)))
        self.raiz.add_widget(self.titulo_tela(nome_modulo, descricao, icone))

        cartao = Painel(orientation="vertical", bg=BRANCO, borda=BORDA, raio=16,
                        padding=[dp(18), dp(14), dp(18), dp(18)], spacing=dp(12),
                        size_hint_y=None)
        cartao.bind(minimum_height=cartao.setter("height"))
        topo = BoxLayout(size_hint_y=None, height=dp(26))
        topo.add_widget(Texto("Campos marcados com [color=#B8232A]*[/color] são obrigatórios",
                              tam=12, cor=CINZA, size_hint_y=1))
        topo.add_widget(ancorar(criar_etiqueta(f"NOVO • {config['prefixo']}", BRANCO, ROXO),
                                anchor_x="right", size_hint_x=None, width=dp(130)))
        cartao.add_widget(topo)

        self.campos_formulario = []
        for definicao in config["campos"]:
            campo = CampoFormulario(definicao)
            self.campos_formulario.append(campo)
            cartao.add_widget(campo)

        rolagem = ScrollView(do_scroll_x=False, bar_width=dp(4), bar_color=(0.44, 0.10, 0.52, 0.4))
        rolagem.add_widget(centralizado(cartao, 820))
        self.raiz.add_widget(rolagem)

        self.raiz.add_widget(Divisor())
        rodape = Painel(bg=BRANCO, raio=0, sombra=False, size_hint_y=None, height=dp(76))
        botoes = BoxLayout(spacing=dp(12), size_hint_y=None, height=dp(50))
        botoes.add_widget(Botao("Limpar", lambda *_: self.limpar_campos(), icone="limpar",
                                estilo="contorno", size_hint_x=0.38))
        botoes.add_widget(Botao("Salvar registro", lambda *_: self.salvar_formulario(nome_modulo),
                                icone="salvar", size_hint_x=0.62))
        rodape.add_widget(ancorar(centralizado(botoes, 820), anchor_y="center"))
        self.raiz.add_widget(rodape)

    def limpar_campos(self):
        for campo in self.campos_formulario:
            campo.limpar()

    def salvar_formulario(self, nome_modulo):
        valores = []
        faltantes = []
        for campo in self.campos_formulario:
            valor = campo.obter_valor()
            definicao = campo.definicao
            if definicao[2] and not valor:
                faltantes.append(campo.nome)
            if definicao[1] == "numero" and valor:
                try:
                    valor = int(valor)
                    if valor < 0:
                        raise ValueError
                except ValueError:
                    self.mensagem("Campo inválido", f"O campo '{campo.nome}' deve conter um número inteiro igual ou maior que zero.")
                    return
            valores.append(valor)

        if faltantes:
            self.mensagem("Campos obrigatórios", "Preencha os seguintes campos:\n\n" + "\n".join(faltantes))
            return

        # Validacao da disposicao: aprovadas + reprovadas nao podem passar do avaliado.
        if nome_modulo == "Disposição de seleção":
            try:
                avaliada = int(valores[4])
                aprovada = int(valores[5])
                reprovada = int(valores[6])
                if aprovada + reprovada > avaliada:
                    self.mensagem("Verifique as quantidades", "A soma de aprovadas e reprovadas não pode superar a quantidade avaliada.")
                    return
            except (ValueError, IndexError):
                self.mensagem("Verifique as quantidades", "Informe quantidades inteiras válidas.")
                return

        try:
            novo_id, data_hora = salvar_registro(nome_modulo, valores)
            self.mensagem("Registro salvo", f"Registro salvo com sucesso!\n\nID: {novo_id}\nData/hora: {data_hora}\n\nPlanilha:\n{ARQUIVO_EXCEL}")
            self.limpar_campos()
        except PermissionError:
            self.mensagem("Planilha em uso", "Não foi possível salvar. Feche o arquivo dados_qualidade.xlsx no Excel e tente novamente.")
        except Exception as erro:
            self.mensagem("Erro ao salvar", f"{erro}\n\nConfira se a planilha está acessível e tente novamente.")

    # ------------------------------------------------------------------ tela: consulta
    def mostrar_consulta(self, *_):
        self.limpar_tela()
        self.raiz.add_widget(self.cabecalho_interno())
        self.raiz.add_widget(Widget(size_hint_y=None, height=dp(8)))
        self.raiz.add_widget(self.titulo_tela("Consulta de registros",
                                              "Veja o que foi salvo na planilha Excel.", "busca"))

        # Campo de busca com lupa
        busca_box = FloatLayout(size_hint_y=None, height=dp(48))
        self.busca = StyledInput(hint_text="Pesquisar por ordem, produto, ID, responsável...",
                                 multiline=False, size_hint=(1, 1), pos_hint={"x": 0, "y": 0},
                                 padding=[dp(42), dp(10), dp(12), dp(10)])
        self.busca.bind(text=lambda *_: self.carregar_registros())
        lupa = Icone("busca", cor=CINZA, size=(dp(20), dp(20)))
        busca_box.add_widget(self.busca)
        busca_box.add_widget(lupa)
        busca_box.bind(pos=lambda i, p: setattr(lupa, "pos", (i.x + dp(13), i.center_y - dp(10))),
                       size=lambda i, s: setattr(lupa, "pos", (i.x + dp(13), i.center_y - dp(10))))
        self.raiz.add_widget(centralizado(busca_box, 900))

        # Filtros por rotina (chips)
        self.chips_area = ScrollView(do_scroll_y=False, bar_width=0, size_hint_y=None, height=dp(44))
        self.raiz.add_widget(centralizado(self.chips_area, 900))
        self.montar_chips()

        self.area_registros = BoxLayout(orientation="vertical", size_hint_y=1)
        self.raiz.add_widget(self.area_registros)
        self.barra_nav("consulta")
        self.carregar_registros()

    def montar_chips(self):
        linha = BoxLayout(size_hint=(None, None), height=dp(40), spacing=dp(8), padding=[0, dp(2)])
        linha.bind(minimum_width=linha.setter("width"))
        for nome in MODULOS:
            chip = Chip(nome, selecionado=(nome == self.modulo_consulta))
            chip.bind(on_release=lambda _c, n=nome: self.trocar_modulo(n))
            linha.add_widget(chip)
        self.chips_area.clear_widgets()
        self.chips_area.add_widget(linha)

    def trocar_modulo(self, nome):
        self.modulo_consulta = nome
        self.montar_chips()
        self.carregar_registros()

    def carregar_registros(self):
        if not hasattr(self, "area_registros"):
            return
        self.area_registros.clear_widgets()
        nome = self.modulo_consulta
        termo = self.busca.text.strip().lower() if hasattr(self, "busca") else ""
        try:
            cabecalho, linhas = ler_registros(nome)
        except Exception as erro:
            self.area_registros.add_widget(Label(
                text=f"Não foi possível consultar a planilha:\n{erro}", color=VERMELHO))
            return

        total = len(linhas)
        # Mais recentes primeiro; filtro opcional por texto em qualquer coluna.
        linhas = linhas[::-1]
        if termo:
            linhas = [l for l in linhas if termo in " ".join(str(v or "") for v in l).lower()]
        linhas = linhas[:100]

        coluna = BoxLayout(orientation="vertical", spacing=dp(10), size_hint_y=None)
        coluna.bind(minimum_height=coluna.setter("height"))
        if not linhas:
            msg = "Nenhum registro encontrado para a busca." if termo else "Ainda não há registros nesta rotina."
            coluna.add_widget(Texto(msg, tam=14, cor=CINZA, halign="center", padding=[0, dp(30)]))
        for linha in linhas:
            coluna.add_widget(CartaoRegistro(cabecalho, linha))

        rolagem = ScrollView(do_scroll_x=False, bar_width=dp(4), bar_color=(0.44, 0.10, 0.52, 0.4))
        rolagem.add_widget(centralizado(coluna, 900))
        self.area_registros.add_widget(rolagem)
        self.area_registros.add_widget(Label(
            text=f"Exibindo {len(linhas)} de {total} registro(s) desta rotina (até 100, mais recentes primeiro)",
            size_hint_y=None, height=dp(26), font_size=sp(11), color=CINZA))

    def abrir_pasta(self, *_):
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(PASTA_APP))
            elif sys.platform == "darwin":
                os.system(f'open "{PASTA_APP}"')
            else:
                os.system(f'xdg-open "{PASTA_APP}"')
        except Exception as erro:
            self.mensagem("Pasta do aplicativo", f"Local da pasta:\n{PASTA_APP}\n\n{erro}")

    # ------------------------------------------------------------------ mensagens
    def mensagem(self, titulo, texto):
        conteudo = BoxLayout(orientation="vertical", spacing=dp(12), padding=dp(14))
        rolagem = ScrollView()
        rotulo = Texto(escape_markup(texto), tam=14, cor=TEXTO)
        rolagem.add_widget(rotulo)
        conteudo.add_widget(rolagem)
        popup = Popup(title=titulo, content=conteudo, size_hint=(0.86, 0.66), auto_dismiss=False,
                      background="", background_color=BRANCO, title_color=ROXO,
                      title_size=sp(17), separator_color=ROXO)
        conteudo.add_widget(Botao("OK", lambda *_: popup.dismiss(), altura=46))
        popup.open()


if __name__ == "__main__":
    try:
        QualidadeApp().run()
    except Exception:
        print("O aplicativo encontrou um erro:")
        traceback.print_exc()
        input("Pressione Enter para fechar...")
