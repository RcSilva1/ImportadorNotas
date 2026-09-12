import customtkinter as ctk
import pandas as pd
import pyautogui
import keyboard
import threading
import time
import os
import re
import unicodedata
from datetime import datetime
from tkinter import filedialog, messagebox
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from atualizacao import iniciar_verificacao


 
URL_SISTEMA = "https://colegiogenoma.professor.gvdasa.com.br/"

QUANTIDADE_PREVIEW = 5                                                       


COLUNA_ALUNO = "Nome"

MAPA_COLUNAS_NOTAS = {
    "Atividade Formativa": "AT FORMAT",
    "Projeto institucional": "PROJ INST",
    "Trilha Virtual": "TRILHA VIR",
    "Ciclo Discussivo": "CICLO DISC",
    "Ciclo Objetivo": "CICLO OBJE",
    "Simulado COC": "SIM. COC",
    "Conselho": "CONSELHO",
}

                                                         
COLUNAS_PLANILHA_OBRIGATORIAS = [COLUNA_ALUNO] + list(MAPA_COLUNAS_NOTAS.keys())
                                                            

PASTA_LOGS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")

try:
    os.makedirs(PASTA_LOGS, exist_ok=True)
except Exception:
    pass

CAMINHO_LOG_ATUAL = os.path.join(
    PASTA_LOGS, datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".log"
)


def registrar_log(mensagem):
    """Grava uma linha com data/hora no arquivo de log da execução
    atual. Nunca interrompe o programa se o log falhar."""

    try:
        linha = f"[{datetime.now().strftime('%H:%M:%S')}] {mensagem}"

        with open(CAMINHO_LOG_ATUAL, "a", encoding="utf-8") as arquivo_log:
            arquivo_log.write(linha + "\n")

    except Exception:
        pass


COR_LARANJA = "#FF7A1A"
COR_LARANJA_ESCURO = "#E4650A"
COR_LARANJA_CLARO = "#FFE3CC"
COR_BRANCO = "#FFFFFF"
COR_FUNDO = "#F7F4F1"
COR_TEXTO = "#2B2B2B"
COR_TEXTO_SECUNDARIO = "#8A8A8A"
COR_ALERTA = "#C0392B"

FONTE_TITULO = ("Segoe UI", 17, "bold")
FONTE_TITULO_GRANDE = ("Segoe UI", 25, "bold")
FONTE_PADRAO = ("Segoe UI", 13)
FONTE_NEGRITO = ("Segoe UI", 13, "bold")
                                                         

df_atual = None
arquivo_selecionado = None

                                                           
                          
playwright_ativo = False
lancamento_em_andamento = False

parar_evento = threading.Event()

configuracao_concluida = False

programa_fechando = False
                                                      

evento_executar_notas = threading.Event()
evento_reiniciar_lancamento = threading.Event()
evento_reiniciar_concluido = threading.Event()
parametros_notas = None                                                         
                                                                
                                                            
thread_automacao_atual = None
                                                       

sedes_disponiveis = []
ciclos_disponiveis = []
turmas_disponiveis = []
materias_disponiveis = []

                                                                        
dados_grid = []

                                                                        
                                                                             
                                                       
cache_grid_por_sede = {}

                                                     

sede = ""
estabelecimento_selecionado = ""
ciclo = ""
turma = ""
materia = ""

sede_confirmada = False
ciclo_confirmado = False
turma_confirmada = False
materia_confirmada = False
modulo = ""
modulo_confirmado = False
modulos_disponiveis = []

navegador = None
contexto = None
pagina = None

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3


def normalizar(texto):
    """Remove acentos, espaços nas pontas e diferenças de maiúsculas."""

    texto = str(texto).strip()

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return texto.casefold()


def textos_iguais(texto1, texto2):
    return normalizar(texto1) == normalizar(texto2)

                                                

class AlunoNaoEncontrado(Exception):
    pass


class AlunoDuplicado(Exception):
    pass


class ColunaNaoEncontrada(Exception):
    pass

                                                  

def normalizar_nota(valor):

    if valor is None:
        return None

    if isinstance(valor, float) and pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):

                                                                  
                                            
        if float(valor).is_integer():
            texto = str(int(valor))
        else:
            texto = f"{valor}".rstrip("0").rstrip(".")

    else:
        texto = str(valor).strip()

    texto = texto.strip()

    if not texto:
        return None

    texto = texto.replace(".", ",")

    return texto


def encontrar_coluna(df, nome_procurado):
    """Encontra o nome real de uma coluna no DataFrame, ignorando
    acentos, maiúsculas/minúsculas e espaços extras."""

    procurado = normalizar(nome_procurado)

    for coluna in df.columns:
        if normalizar(coluna) == procurado:
            return coluna

    return None
                                                 

def carregar_dataframe(caminho):

    extensao = os.path.splitext(caminho)[1].lower()

    try:

        if extensao == ".csv":

            try:
                df = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")
            except UnicodeDecodeError:
                df = pd.read_csv(caminho, sep=";", encoding="latin1")

                                                                            
            if len(df.columns) == 1:

                try:
                    df = pd.read_csv(caminho, sep=",", encoding="utf-8-sig")
                except UnicodeDecodeError:
                    df = pd.read_csv(caminho, sep=",", encoding="latin1")

        elif extensao in [".xlsx", ".xlsm", ".xls"]:

            df = pd.read_excel(caminho)

        else:

            raise ValueError("Formato não suportado.\nUse CSV ou Excel.")

        df = df.dropna(how="all")

        if len(df.columns) == 0:
            raise ValueError("A planilha não possui colunas.")

        return df

    except Exception as erro:
        raise Exception(f"Não foi possível ler a planilha:\n\n{erro}")

                                             

def executar_na_interface(funcao):

    if programa_fechando:
        return

    def callback_seguro():

        if programa_fechando:
            return

        try:
            funcao()
        except Exception:
            pass

    try:
        janela.after(0, callback_seguro)
    except Exception:
        pass


def atualizar_status(texto):

    executar_na_interface(lambda: status_label.configure(text=texto))


def atualizar_progresso(valor):

    executar_na_interface(lambda: progresso.set(valor))
                                                   

def verificar_configuracao_completa():

    global configuracao_concluida

    tudo_pronto = (
        arquivo_selecionado is not None
        and df_atual is not None
        and sede_combo.get()
        and ciclo_combo.get()
        and turma_combo.get()
        and materia_combo.get()
        and modulo_combo.get()
        and modulo_confirmado
    )

    if tudo_pronto:

        configuracao_concluida = True

        botao_iniciar.configure(state="normal")

        atualizar_status("Tudo pronto. Você pode iniciar a automação.")

    else:

        configuracao_concluida = False

        botao_iniciar.configure(state="disabled")                                      
                                                              

def atualizar_preview():

    for widget in preview_frame.winfo_children():
        widget.destroy()

    if df_atual is None:
        return

    coluna_nome_real = encontrar_coluna(df_atual, COLUNA_ALUNO)

    if not coluna_nome_real:
        return
                                                  

    ctk.CTkLabel(
        preview_frame, text="Aluno", font=FONTE_NEGRITO, text_color=COR_TEXTO
    ).grid(row=0, column=0, sticky="w", padx=10, pady=8)

    ctk.CTkLabel(
        preview_frame,
        text="Notas preenchidas",
        font=FONTE_NEGRITO,
        text_color=COR_TEXTO,
    ).grid(row=0, column=1, sticky="w", padx=10, pady=8)
                                                    

    for linha, (_, registro) in enumerate(
        df_atual.head(QUANTIDADE_PREVIEW).iterrows(), start=1
    ):

        nome = registro[coluna_nome_real]

        if pd.isna(nome):
            nome = ""

        nome = str(nome)

        notas_preenchidas = []

        for coluna_planilha in MAPA_COLUNAS_NOTAS:

            coluna_real = encontrar_coluna(df_atual, coluna_planilha)

            if not coluna_real:
                continue

            valor = registro[coluna_real]

            if pd.isna(valor):
                continue

            if str(valor).strip() == "":
                continue

            notas_preenchidas.append(f"{coluna_planilha}: {valor}")

        ctk.CTkLabel(
            preview_frame, text=nome, anchor="w", text_color=COR_TEXTO
        ).grid(row=linha, column=0, sticky="w", padx=10, pady=4)

        ctk.CTkLabel(
            preview_frame,
            text=(
                " | ".join(notas_preenchidas)
                if notas_preenchidas
                else "Nenhuma nota preenchida"
            ),
            anchor="w",
            text_color=(COR_TEXTO if notas_preenchidas else COR_TEXTO_SECUNDARIO),
        ).grid(row=linha, column=1, sticky="w", padx=10, pady=4)
                                                 

def baixar_modelo_planilha():
    """Gera e salva um arquivo Excel em branco, já com os cabeçalhos
    corretos (coluna do aluno + todas as colunas de nota configuradas
    em MAPA_COLUNAS_NOTAS), para o usuário preencher."""

    caminho = filedialog.asksaveasfilename(
        title="Salvar modelo da planilha",
        defaultextension=".xlsx",
        initialfile="modelo_notas.xlsx",
        filetypes=[("Excel", "*.xlsx")],
    )

    if not caminho:
        return

    try:

        df_modelo = pd.DataFrame(columns=COLUNAS_PLANILHA_OBRIGATORIAS)

        df_modelo.to_excel(caminho, index=False)

        atualizar_status(f"Modelo salvo em: {caminho}")

    except Exception as erro:

        atualizar_status(f"Erro ao salvar o modelo: {erro}")

def selecionar_arquivo():

    global arquivo_selecionado
    global df_atual

    caminho = filedialog.askopenfilename(
        title="Selecionar planilha",
        filetypes=[
            ("Planilhas", "*.xlsx *.xlsm *.xls *.csv"),
            ("Excel", "*.xlsx *.xlsm *.xls"),
            ("CSV", "*.csv"),
        ],
    )

    if not caminho:
        return

    try:

        df = carregar_dataframe(caminho)

        colunas_encontradas = {normalizar(str(coluna)) for coluna in df.columns}

        colunas_faltando = [
            coluna_obrigatoria
            for coluna_obrigatoria in COLUNAS_PLANILHA_OBRIGATORIAS
            if normalizar(coluna_obrigatoria) not in colunas_encontradas
        ]

        if colunas_faltando:

            raise Exception(
                "A planilha não está no formato correto.\n\n"
                "Colunas ausentes:\n\n"
                + "\n".join(f"• {coluna}" for coluna in colunas_faltando)
            )

        arquivo_selecionado = caminho
        df_atual = df
                                                

        nome_arquivo = os.path.basename(caminho)

        arquivo_label.configure(text=f"📄 {nome_arquivo}")

        quantidade_label.configure(text=f"{len(df)} aluno(s) encontrado(s)")

        preview_container.grid()
        atualizar_preview()

        atualizar_status("Planilha carregada com sucesso.")

        verificar_configuracao_completa()

    except Exception as erro:

        arquivo_selecionado = None
        df_atual = None

        arquivo_label.configure(text="Nenhum arquivo selecionado")
        quantidade_label.configure(text="")

        atualizar_status(f"Erro: {erro}")
                    
                                                              

def iniciar_preparacao():

    global playwright_ativo, thread_automacao_atual

    usuario = usuario_entry.get().strip()
    senha = senha_entry.get()

    if not usuario:
        atualizar_status("Digite o usuário.")
        return

    if not senha:
        atualizar_status("Digite a senha.")
        return

    playwright_ativo = True

    botao_continuar_login.configure(state="disabled")

    atualizar_status("Conectando ao sistema...")

    progresso.set(0)

    parar_evento.clear()

    thread_automacao_atual = threading.Thread(
        target=automacao, args=(usuario, senha), daemon=True
    )

    thread_automacao_atual.start()
                                    
                                                              

def filtrar_por_ciclo(ciclo_selecionado):

    return [
        registro
        for registro in dados_grid
        if textos_iguais(registro["ciclo"], ciclo_selecionado)
    ]


def filtrar_por_ciclo_turma(ciclo_selecionado, turma_selecionada):

    return [
        registro
        for registro in dados_grid
        if textos_iguais(registro["ciclo"], ciclo_selecionado)
        and textos_iguais(registro["turma"], turma_selecionada)
    ]


def filtrar_por_selecao(ciclo_selecionado, turma_selecionada, materia_selecionada):

    return [
        registro
        for registro in dados_grid
        if textos_iguais(registro["ciclo"], ciclo_selecionado)
        and textos_iguais(registro["turma"], turma_selecionada)
        and textos_iguais(registro["materia"], materia_selecionada)
    ]


def obter_turmas_do_ciclo(ciclo_selecionado):

    registros = filtrar_por_ciclo(ciclo_selecionado)

    turmas = []

    for registro in registros:

        turma_atual = registro["turma"]

        if not turma_atual:
            continue

        if not any(textos_iguais(turma_atual, existente) for existente in turmas):
            turmas.append(turma_atual)

    return turmas


def obter_materias_da_turma(ciclo_selecionado, turma_selecionada):

    registros = filtrar_por_ciclo_turma(ciclo_selecionado, turma_selecionada)

    materias = []

    for registro in registros:

        materia_atual = registro["materia"]

        if not materia_atual:
            continue

        if not any(
            textos_iguais(materia_atual, existente) for existente in materias
        ):
            materias.append(materia_atual)

    return materias
                                          

def confirmar_sede():

    global sede
    global sede_confirmada, ciclo_confirmado, turma_confirmada, materia_confirmada

    sede_selecionada = sede_combo.get()

    if not sede_selecionada:
        atualizar_status("Selecione uma sede antes de confirmar.")
        return

    sede = sede_selecionada

    sede_confirmada = True
    ciclo_confirmado = False
    turma_confirmada = False
    materia_confirmada = False

    ciclo_combo.set("")
    turma_combo.set("")
    materia_combo.set("")

    ciclo_combo.configure(values=ciclos_disponiveis, state="readonly")
    turma_combo.configure(values=[], state="disabled")
    materia_combo.configure(values=[], state="disabled")

    botao_confirmar_ciclo.configure(state="normal")
    botao_confirmar_turma.configure(state="disabled")
    botao_confirmar_materia.configure(state="disabled")
    botao_continuar_configuracao.configure(state="disabled")

    atualizar_status(f"Sede confirmada: {sede}. Agora escolha o ciclo/curso.")
                                               

def confirmar_ciclo():

    global ciclo
    global ciclo_confirmado, turma_confirmada, materia_confirmada
    global modulo, modulo_confirmado, modulos_disponiveis

    if not sede_confirmada:
        atualizar_status("Confirme a sede primeiro.")
        return

    ciclo_selecionado = ciclo_combo.get()

    if not ciclo_selecionado:
        atualizar_status("Selecione um ciclo/curso antes de confirmar.")
        return

    ciclo = ciclo_selecionado

    ciclo_confirmado = True
    turma_confirmada = False
    materia_confirmada = False

    turma_combo.set("")
    materia_combo.set("")

    registros_filtrados = filtrar_por_ciclo(ciclo_selecionado)
    turmas_do_ciclo = obter_turmas_do_ciclo(ciclo_selecionado)

    turma_combo.configure(values=turmas_do_ciclo, state="readonly")
    materia_combo.configure(values=[], state="disabled")

    botao_confirmar_turma.configure(
        state="normal" if turmas_do_ciclo else "disabled"
    )
    botao_confirmar_materia.configure(state="disabled")
    botao_continuar_configuracao.configure(state="disabled")

    if turmas_do_ciclo:
        atualizar_status(
            f"Ciclo/Curso confirmado: {ciclo}. "
            f"{len(registros_filtrados)} registro(s) encontrado(s) em "
            f"{len(turmas_do_ciclo)} turma(s)."
        )
    else:
        atualizar_status(f"Nenhuma turma encontrada para '{ciclo}'.")
                                              

def confirmar_turma():

    global turma
    global turma_confirmada, materia_confirmada

    if not ciclo_confirmado:
        atualizar_status("Confirme o ciclo/curso primeiro.")
        return

    turma_selecionada = turma_combo.get()

    if not turma_selecionada:
        atualizar_status("Selecione uma turma antes de confirmar.")
        return

    turma = turma_selecionada

    turma_confirmada = True
    materia_confirmada = False

    materia_combo.set("")

    registros_filtrados = filtrar_por_ciclo_turma(ciclo, turma)
    materias_da_turma = obter_materias_da_turma(ciclo, turma)

    materia_combo.configure(values=materias_da_turma, state="readonly")

    botao_confirmar_materia.configure(
        state="normal" if materias_da_turma else "disabled"
    )
    botao_continuar_configuracao.configure(state="disabled")

    if materias_da_turma:
        atualizar_status(
            f"Turma confirmada: {turma}. "
            f"{len(registros_filtrados)} registro(s) e "
            f"{len(materias_da_turma)} matéria(s) encontrada(s)."
        )
    else:
        atualizar_status(f"Nenhuma matéria encontrada para {ciclo} | {turma}.")

def confirmar_modulo():

    global modulo, modulo_confirmado

    if not materia_confirmada:
        atualizar_status("Confirme a matéria primeiro.")
        return

    modulo_selecionado = modulo_combo.get()

    if not modulo_selecionado:
        atualizar_status("Selecione um módulo/trimestre antes de confirmar.")
        return

    modulo = modulo_selecionado
    modulo_confirmado = True
    atualizar_status(f"Módulo confirmado: {modulo}.")

                                                                          
                                                                          
                                                                        
    botao_lancar_novamente.configure(state="normal")

    verificar_configuracao_completa()
                                                 

def confirmar_materia():

    global materia
    global materia_confirmada, modulo_confirmado, modulo

    if not turma_confirmada:
        atualizar_status("Confirme a turma primeiro.")
        return

    materia_selecionada = materia_combo.get()

    if not materia_selecionada:
        atualizar_status("Selecione uma matéria antes de confirmar.")
        return

    materia = materia_selecionada
    materia_confirmada = True
    modulo_confirmado = False
    modulo = ""
    modulo_combo.set("")
    modulo_combo.configure(values=[], state="disabled")
    botao_confirmar_modulo.configure(state="disabled")

    registros_filtrados = filtrar_por_selecao(ciclo, turma, materia)

    if len(registros_filtrados) == 1:

        registro = registros_filtrados[0]

        atualizar_status(
            f"Matéria confirmada: {materia}. Linha alvo: {registro['rowindex']}."
        )

        botao_continuar_configuracao.configure(state="normal")

    elif len(registros_filtrados) > 1:

        atualizar_status(
            f"Atenção: existem {len(registros_filtrados)} registros "
            f"iguais para essa seleção."
        )

        botao_continuar_configuracao.configure(state="disabled")

    else:

        atualizar_status("Nenhuma linha corresponde à seleção.")

        botao_continuar_configuracao.configure(state="disabled")
                                        

def ler_opcoes_grid():

    global ciclos_disponiveis, turmas_disponiveis, materias_disponiveis, dados_grid

    registros_por_indice = {}

    scroller = pagina.locator(".MuiDataGrid-virtualScroller").first
    scroller.wait_for(state="visible", timeout=15000)

    dimensoes = scroller.evaluate(
        "el => ({scrollHeight: el.scrollHeight, clientHeight: el.clientHeight})"
    )

    scroll_height = dimensoes["scrollHeight"]
    client_height = dimensoes["clientHeight"]

    scroller.evaluate("el => el.scrollTop = 0")
    pagina.wait_for_timeout(300)

                                                              
                                                                     
    passo = max(int(client_height * 0.8), 100)

    posicao = 0
    ultima_posicao = -1
    ultima_quantidade_registros = -1
    confirmacoes_de_fim = 0

    while True:

        verificar_parada()

        scroller.evaluate("(el, posicao) => { el.scrollTop = posicao; }", posicao)
        pagina.wait_for_timeout(250)

        linhas = pagina.locator('[role="row"][data-rowindex]')
        quantidade_linhas = linhas.count()

        for indice in range(quantidade_linhas):

            verificar_parada()

            linha = linhas.nth(indice)

            try:
                rowindex_texto = linha.get_attribute("data-rowindex")

                if rowindex_texto is None:
                    continue

                rowindex = int(rowindex_texto)

            except Exception:
                continue

            curso = linha.locator('[data-field="cursoCiclo"]')
            turma_linha = linha.locator('[data-field="nomeTurma"]')
            materia_linha = linha.locator('[data-field="nomeComponenteCurricular"]')

            if not curso.count() or not turma_linha.count() or not materia_linha.count():
                continue

            try:
                curso_texto = curso.inner_text().strip()
                turma_texto = turma_linha.inner_text().strip()
                materia_texto = materia_linha.inner_text().strip()
            except Exception:
                continue

            if not (curso_texto and turma_texto and materia_texto):
                continue

            registros_por_indice[rowindex] = {
                "rowindex": rowindex,
                "ciclo": curso_texto,
                "turma": turma_texto,
                "materia": materia_texto,
            }

        posicao_real = scroller.evaluate("el => el.scrollTop")
        altura_total = scroller.evaluate("el => el.scrollHeight")
        altura_visivel = scroller.evaluate("el => el.clientHeight")

        max_scroll = max(altura_total - altura_visivel, 0)

        quantidade_registros_atual = len(registros_por_indice)

        chegou_ao_fim = (
            posicao_real >= max_scroll - 5 or posicao_real == ultima_posicao
        )

        if chegou_ao_fim:
                                                  
                                                          
            if quantidade_registros_atual == ultima_quantidade_registros:

                confirmacoes_de_fim += 1

                if confirmacoes_de_fim >= 2:
                    break

            else:
                confirmacoes_de_fim = 0

        else:
            confirmacoes_de_fim = 0

        ultima_quantidade_registros = quantidade_registros_atual
        ultima_posicao = posicao_real

        posicao = min(posicao + passo, max_scroll)

    scroller.evaluate("el => el.scrollTop = 0")
    pagina.wait_for_timeout(300)

    dados_grid = sorted(
        registros_por_indice.values(), key=lambda registro: registro["rowindex"]
    )

    def valores_unicos(chave):

        valores = []

        for registro in dados_grid:

            valor = registro[chave]

            if not any(textos_iguais(valor, existente) for existente in valores):
                valores.append(valor)

        return valores

    ciclos_disponiveis = valores_unicos("ciclo")
    turmas_disponiveis = valores_unicos("turma")
    materias_disponiveis = valores_unicos("materia")

    return dados_grid

def obter_dados_grid_para_sede(codigo_sede):
    """Garante que dados_grid/ciclos/turmas/materias estejam preenchidos
    para a sede informada. Se essa sede já foi analisada antes nesta
    mesma execução, reaproveita os dados em cache em vez de rolar e
    raspar a grade inteira de novo (o que é lento). Retorna True se
    usou o cache, e False se precisou raspar a grade agora."""

    global dados_grid, ciclos_disponiveis, turmas_disponiveis, materias_disponiveis

    if codigo_sede in cache_grid_por_sede:

        dados_em_cache = cache_grid_por_sede[codigo_sede]

        dados_grid = dados_em_cache["dados_grid"]
        ciclos_disponiveis = dados_em_cache["ciclos"]
        turmas_disponiveis = dados_em_cache["turmas"]
        materias_disponiveis = dados_em_cache["materias"]

        registrar_log(
            f"Sede '{codigo_sede}': usando dados em cache "
            f"({len(dados_grid)} registro(s)), sem raspar a grade de novo."
        )

        return True

    ler_opcoes_grid()

    cache_grid_por_sede[codigo_sede] = {
        "dados_grid": dados_grid,
        "ciclos": ciclos_disponiveis,
        "turmas": turmas_disponiveis,
        "materias": materias_disponiveis,
    }

    registrar_log(
        f"Sede '{codigo_sede}': grade raspada e guardada em cache "
        f"({len(dados_grid)} registro(s))."
    )

    return False


def encontrar_linha_por_rowindex(rowindex_alvo):

    scroller = pagina.locator(".MuiDataGrid-virtualScroller").first
    scroller.wait_for(state="visible", timeout=10000)

                                                             
    linha = pagina.locator(f'[role="row"][data-rowindex="{rowindex_alvo}"]')

    if linha.count():
        return linha.first

    dados_scroll = scroller.evaluate(
        "el => ({scrollHeight: el.scrollHeight, clientHeight: el.clientHeight})"
    )

    scroll_height = dados_scroll["scrollHeight"]
    client_height = dados_scroll["clientHeight"]

                                                                            
    altura_linha = None

    linhas_visiveis = pagina.locator('[role="row"][data-rowindex]')
    quantidade = linhas_visiveis.count()

    if quantidade >= 2:

        try:
            caixas = []

            for i in range(min(quantidade, 5)):
                caixa = linhas_visiveis.nth(i).bounding_box()
                if caixa:
                    caixas.append(caixa["height"])

            if caixas:
                altura_linha = sum(caixas) / len(caixas)

        except Exception:
            altura_linha = None

    if not altura_linha or altura_linha <= 0:
        altura_linha = 52                                                   

    max_scroll = max(scroll_height - client_height, 0)

    posicao_estimada = min(max(rowindex_alvo * altura_linha, 0), max_scroll)

    scroller.evaluate(
        "(el, posicao) => { el.scrollTop = posicao; }", posicao_estimada
    )

    pagina.wait_for_timeout(300)

    linha = pagina.locator(f'[role="row"][data-rowindex="{rowindex_alvo}"]')

    if linha.count():
        return linha.first

                                                              
    deslocamentos = [
        -client_height,
        client_height,
        -client_height * 2,
        client_height * 2,
    ]

    for deslocamento in deslocamentos:

        verificar_parada()

        posicao = min(max(posicao_estimada + deslocamento, 0), max_scroll)

        scroller.evaluate("(el, posicao) => { el.scrollTop = posicao; }", posicao)
        pagina.wait_for_timeout(250)

        linha = pagina.locator(f'[role="row"][data-rowindex="{rowindex_alvo}"]')

        if linha.count():
            return linha.first

                                                               
    passo = max(int(client_height * 0.8), 100)
    posicao = 0

    while posicao <= max_scroll:

        verificar_parada()

        scroller.evaluate("(el, posicao) => { el.scrollTop = posicao; }", posicao)
        pagina.wait_for_timeout(200)

        linha = pagina.locator(f'[role="row"][data-rowindex="{rowindex_alvo}"]')

        if linha.count():
            return linha.first

        posicao += passo

    raise Exception(
        f"Não foi possível localizar no grid a linha com rowindex={rowindex_alvo}."
    )
                                                

def clicar_no_diario(ciclo_selecionado, turma_selecionada, materia_selecionada):

    registros = filtrar_por_selecao(
        ciclo_selecionado, turma_selecionada, materia_selecionada
    )

    if not registros:
        raise Exception(
            "O filtro não encontrou nenhum registro para:\n"
            f"{ciclo_selecionado} | {turma_selecionada} | {materia_selecionada}"
        )

    if len(registros) > 1:

        detalhes = "\n".join(
            f"rowindex={registro['rowindex']} | {registro['ciclo']} | "
            f"{registro['turma']} | {registro['materia']}"
            for registro in registros
        )

        raise Exception(
            "O filtro encontrou mais de uma linha correspondente.\n\n" f"{detalhes}"
        )

    registro = registros[0]
    rowindex_alvo = registro["rowindex"]

    atualizar_status("Localizando a linha correta no grid...")

    linha = encontrar_linha_por_rowindex(rowindex_alvo)

                                                                           
    curso = linha.locator('[data-field="cursoCiclo"]')
    turma_linha = linha.locator('[data-field="nomeTurma"]')
    materia_linha = linha.locator('[data-field="nomeComponenteCurricular"]')

    curso_texto = curso.inner_text().strip()
    turma_texto = turma_linha.inner_text().strip()
    materia_texto = materia_linha.inner_text().strip()

    if not (
        textos_iguais(curso_texto, ciclo_selecionado)
        and textos_iguais(turma_texto, turma_selecionada)
        and textos_iguais(materia_texto, materia_selecionada)
    ):

        raise Exception(
            "A linha localizada pelo rowindex não corresponde à seleção.\n\n"
            f"Esperado:\n{ciclo_selecionado} | {turma_selecionada} | "
            f"{materia_selecionada}\n\n"
            f"Encontrado:\n{curso_texto} | {turma_texto} | {materia_texto}"
        )

    botao_no = linha.locator('[data-field="##"] button', has_text="NO")

    if not botao_no.count():
        botao_no = linha.get_by_role("button", name="NO")

    if not botao_no.count():

        botoes = linha.locator("button")

        for i in range(botoes.count()):

            try:
                texto_botao = botoes.nth(i).inner_text().strip()

                if textos_iguais(texto_botao, "NO"):
                    botao_no = botoes.nth(i)
                    break

            except Exception:
                continue

    if not botao_no.count():
        raise Exception(
            "A linha correta foi encontrada, mas o botão NO não foi "
            "encontrado dentro dela."
        )

    atualizar_status("Linha correta encontrada. Clicando no botão NO...")

    botao_no.first.scroll_into_view_if_needed()
    pagina.wait_for_timeout(150)
    botao_no.first.click()

    atualizar_status(
        f"Diário selecionado: {ciclo_selecionado} | {turma_selecionada} | "
        f"{materia_selecionada}. Aguardando tela de lançamento de notas..."
    )

  
    try:
        pagina.locator("#note-typing-iframe").wait_for(
            state="attached", timeout=20000
        )
    except Exception:
        raise Exception(
            "Cliquei em 'NO', mas a tela de lançamento de notas "
            "(iframe '#note-typing-iframe') não apareceu em até 20 "
            "segundos. Confira se o diário realmente abriu na tela e "
            "se o id do iframe ainda é esse (inspecione com o DevTools "
            "do navegador, botão direito > Inspecionar, sobre a área "
            "da tabela de notas)."
        )

    atualizar_status(
        f"Diário selecionado: {ciclo_selecionado} | {turma_selecionada} | "
        f"{materia_selecionada}"
    )
                                                        

def abrir_lista_estabelecimentos():
    """Abre o menu do usuário e o combobox de 'Estabelecimento', e
    devolve a lista de nomes disponíveis (sem selecionar nada ainda).
    O dropdown fica aberto — a seleção real acontece depois, em
    selecionar_estabelecimento().

    O nome do combobox muda de acordo com o estabelecimento
    atualmente selecionado (ex.: "Estabelecimento 1.11 - DNA"), então
    o match é feito só pelo início do nome, sem depender do valor
    atual."""

    pagina.get_by_test_id("user-menu-trigger").click()
    pagina.wait_for_timeout(300)

    combo_estabelecimento = pagina.get_by_role(
        "combobox", name=re.compile(r"^Estabelecimento", re.IGNORECASE)
    ).first

    combo_estabelecimento.wait_for(state="visible", timeout=10000)
    combo_estabelecimento.click()

    opcoes = pagina.get_by_role("option")
    opcoes.first.wait_for(state="visible", timeout=10000)

    nomes = [texto.strip() for texto in opcoes.all_text_contents() if texto.strip()]

    return nomes


def fechar_menu_flutuante_se_houver():
    """Fecha qualquer menu/popover da MUI que ainda esteja aberto.

    Necessário porque, se o estabelecimento clicado já era o que
    estava ativo (ex.: o site abre logado na última sede usada), o
    valor não muda e o componente às vezes não fecha o menu sozinho,
    deixando um backdrop bloqueando cliques futuros na página.

    O clique precisa ser especificamente no backdrop que é filho
    direto do Popover/Menu (".MuiPopover-root.MuiMenu-root.MuiModal-root
    > .MuiBackdrop-root") — é o único que realmente fecha esse menu
    (a classe "css-xxxx" do exemplo original é instável entre builds
    do site, por isso não é usada aqui)."""

    try:
        backdrops = pagina.locator(
            ".MuiPopover-root.MuiMenu-root.MuiModal-root > .MuiBackdrop-root"
        )

        for i in range(backdrops.count()):

            try:
                alvo = backdrops.nth(i)

                if alvo.is_visible():
                    alvo.click(force=True, timeout=3000)
                    pagina.wait_for_timeout(200)

            except Exception:
                continue

    except Exception:
        pass


def selecionar_estabelecimento(codigo_sede):
    """Clica na opção do estabelecimento escolhido. O dropdown deve
    continuar aberto desde abrir_lista_estabelecimentos() — só
    tenta reabrir se, por algum motivo, a opção já não existir mais
    na tela (evita clicar no menu de novo à toa, o que pode deixar um
    backdrop de um popup antigo bloqueando a tela)."""

    opcao = pagina.locator(f'li[data-value="{codigo_sede}"]')

    if opcao.count() == 0:

        registrar_log(
            "AVISO: a lista de estabelecimentos não estava mais aberta; "
            "reabrindo antes de selecionar."
        )

        pagina.get_by_test_id("user-menu-trigger").click()
        pagina.wait_for_timeout(300)

        combo_estabelecimento = pagina.get_by_role(
            "combobox", name=re.compile(r"^Estabelecimento", re.IGNORECASE)
        ).first
        combo_estabelecimento.wait_for(state="visible", timeout=10000)
        combo_estabelecimento.click()

        opcao = pagina.locator(f'li[data-value="{codigo_sede}"]')

    opcao.wait_for(state="visible", timeout=10000)
    opcao.click()

    pagina.wait_for_timeout(500)

                                                                          
                                                                          
                                                                        
    fechar_menu_flutuante_se_houver()

    pagina.wait_for_timeout(500)


def automacao(usuario, senha):

    global navegador, contexto, pagina
    global playwright_ativo, estabelecimento_selecionado
    global sedes_disponiveis, ciclos_disponiveis, turmas_disponiveis
    global materias_disponiveis, dados_grid

    registrar_log("Iniciando sessão do Playwright...")

    try:

        with sync_playwright() as pw:

            navegador = pw.chromium.launch(headless=False)
            contexto = navegador.new_context()
            pagina = contexto.new_page()
                                                        
            pagina.goto(URL_SISTEMA, wait_until="domcontentloaded", timeout=30000)

            atualizar_status("Carregando tela de login...")

            campo_usuario = pagina.get_by_role("textbox", name="USUÁRIO")
            campo_usuario.wait_for(state="visible", timeout=20000)

            atualizar_status("Realizando login...")

    
            campo_usuario.fill(usuario)
            pagina.get_by_role("textbox", name="Senha").fill(senha)
            pagina.get_by_role("button", name="ENTRAR").click()
                                                    
                                                                   
                                                                          
                                                                          
                                                                          
            mensagem_erro_login = pagina.get_by_text(
                re.compile(r"inv[aá]lid", re.IGNORECASE)
            )
            menu_usuario = pagina.get_by_test_id("user-menu-trigger")

            login_ok = False
            erro_login_detectado = None
            tempo_limite_login = time.time() + 20

            while time.time() < tempo_limite_login:

                verificar_parada()

                if mensagem_erro_login.count() > 0:

                    try:
                        erro_login_detectado = (
                            mensagem_erro_login.first.inner_text().strip()
                        )
                    except Exception:
                        erro_login_detectado = "Usuário ou senha inválida."

                    break

                if menu_usuario.is_visible():
                    login_ok = True
                    break

                pagina.wait_for_timeout(150)

            if not login_ok:

                if erro_login_detectado:
                    mensagem_popup = erro_login_detectado
                    atualizar_status(f"Login falhou: {erro_login_detectado}")
                    registrar_log(f"Login falhou: {erro_login_detectado}")
                else:
                    mensagem_popup = (
                        "Não foi possível confirmar o login. Verifique o "
                        "usuário e a senha e tente novamente."
                    )
                    atualizar_status(
                        "Não foi possível confirmar o login (usuário/senha "
                        "podem estar incorretos, ou a tela demorou a "
                        "carregar). Tente novamente."
                    )
                    registrar_log(
                        "Login falhou ou tela pós-login não apareceu a tempo."
                    )

                                                                          
                executar_na_interface(
                    lambda: messagebox.showerror("Login inválido", mensagem_popup)
                )

                return

            atualizar_status("Login realizado com sucesso.")
            registrar_log("Login realizado com sucesso.")
                                              

            sedes_disponiveis = abrir_lista_estabelecimentos()

            registrar_log(f"{len(sedes_disponiveis)} sede(s) encontrada(s).")
                                                              

            executar_na_interface(mostrar_configuracao)

                                                              

            while not sede_confirmada:
                if parar_evento.is_set() or programa_fechando:
                    return
                time.sleep(0.1)

            verificar_parada()

            sede_selecionada = sede_combo.get()
            estabelecimento_selecionado = sede_selecionada
            codigo_sede = sede_selecionada.split(" - ")[0].strip()

            selecionar_estabelecimento(codigo_sede)

              
            atualizar_status("Analisando ciclos, turmas e matérias...")

            pagina.locator('[data-field="cursoCiclo"]').first.wait_for(
                state="visible", timeout=15000
            )

                                                 
            pagina.locator('[role="row"][data-rowindex]').first.wait_for(
                state="visible", timeout=10000
            )

            veio_do_cache = obter_dados_grid_para_sede(codigo_sede)

            if not dados_grid:
                raise Exception("Nenhum registro foi encontrado no grid.")

            if veio_do_cache:
                atualizar_status(
                    f"Sede já conhecida. Reaproveitando {len(dados_grid)} "
                    "registro(s) já lidos anteriormente."
                )

            executar_na_interface(
                lambda: ciclo_combo.configure(
                    values=ciclos_disponiveis, state="readonly"
                )
            )

            atualizar_status(
                f"Sistema analisado. {len(dados_grid)} registro(s) encontrado(s). "
                f"{len(ciclos_disponiveis)} ciclo(s)/curso(s). Escolha o ciclo/curso."
            )
                                                

            while not ciclo_confirmado:
                if parar_evento.is_set() or programa_fechando:
                    return
                time.sleep(0.1)

            verificar_parada()

            while not turma_confirmada:
                if parar_evento.is_set() or programa_fechando:
                    return
                time.sleep(0.1)

            verificar_parada()

            while not materia_confirmada:
                if parar_evento.is_set() or programa_fechando:
                    return
                time.sleep(0.1)

            verificar_parada()
                                                 

            atualizar_status("Aplicando filtro final...")

            clicar_no_diario(ciclo, turma, materia)

            iframe = obter_iframe_notas()
            combo_modulo = iframe.get_by_role("combobox", name="Módulo")
            combo_modulo.click()
            opcoes_modulo = iframe.get_by_role("option")
            opcoes_modulo.first.wait_for(state="visible", timeout=10000)
            modulos = [texto.strip() for texto in opcoes_modulo.all_text_contents() if texto.strip()]
                                                                          
                                                                          
            combo_modulo.press("Escape")

            executar_na_interface(lambda: modulo_combo.configure(values=modulos, state="readonly"))
            executar_na_interface(lambda: botao_confirmar_modulo.configure(state="normal" if modulos else "disabled"))
            atualizar_status("Diário aberto. Escolha o módulo/trimestre.")

            while not modulo_confirmado:
                if parar_evento.is_set() or programa_fechando:
                    return
                time.sleep(0.1)

            verificar_parada()                                                 

            atualizar_status(
                "Diário aberto. Selecione a planilha e clique em "
                "'Iniciar automação' quando estiver pronto."
            )
                                             

            while not parar_evento.is_set() and not programa_fechando:

                if evento_reiniciar_lancamento.is_set():
                    try:
                        iframe = obter_iframe_notas()
                        iframe.get_by_test_id("botao-cancelar-footer").click()
                        pagina.wait_for_timeout(700)

                        estabelecimentos = abrir_lista_estabelecimentos()

                        executar_na_interface(mostrar_configuracao)
                        executar_na_interface(
                            lambda: sede_combo.configure(values=estabelecimentos, state="readonly")
                        )

                                                                          
                                                                        
                                                                        
                        valor_pre_selecionado = next(
                            (
                                opcao
                                for opcao in estabelecimentos
                                if textos_iguais(opcao, estabelecimento_selecionado)
                            ),
                            None,
                        )

                        if valor_pre_selecionado:
                            executar_na_interface(
                                lambda valor=valor_pre_selecionado: sede_combo.set(valor)
                            )
                            atualizar_status(
                                "Sessão mantida. Confira o estabelecimento já "
                                "selecionado e clique em CONFIRMAR para continuar."
                            )
                        else:
                            atualizar_status(
                                "Sessão mantida. Escolha o estabelecimento para continuar."
                            )

                        while not sede_confirmada:
                            if parar_evento.is_set() or programa_fechando:
                                return
                            time.sleep(0.1)

                        verificar_parada()

                                                                          
                                                                          
                        snapshot_linha_anterior = None
                        if dados_grid:
                            snapshot_linha_anterior = (
                                dados_grid[0]["ciclo"],
                                dados_grid[0]["turma"],
                                dados_grid[0]["materia"],
                            )

                        sede_nova = sede_combo.get()
                        estabelecimento_selecionado = sede_nova
                        codigo_sede_nova = sede_nova.split(" - ")[0].strip()
                        selecionar_estabelecimento(codigo_sede_nova)

                        atualizar_status("Analisando o novo estabelecimento...")
                        pagina.locator('[data-field="cursoCiclo"]').first.wait_for(
                            state="visible", timeout=20000
                        )
                        pagina.locator('[role="row"][data-rowindex]').first.wait_for(
                            state="visible", timeout=10000
                        )

                                                                          
                                                                          
                                                                       
                        if (
                            codigo_sede_nova not in cache_grid_por_sede
                            and snapshot_linha_anterior is not None
                        ):

                            tempo_limite_troca = time.time() + 8

                            while time.time() < tempo_limite_troca:

                                verificar_parada()

                                try:
                                    primeira_linha_atual = pagina.locator(
                                        '[role="row"][data-rowindex]'
                                    ).first
                                    curso_atual = primeira_linha_atual.locator(
                                        '[data-field="cursoCiclo"]'
                                    ).inner_text().strip()
                                    turma_atual = primeira_linha_atual.locator(
                                        '[data-field="nomeTurma"]'
                                    ).inner_text().strip()
                                    materia_atual = primeira_linha_atual.locator(
                                        '[data-field="nomeComponenteCurricular"]'
                                    ).inner_text().strip()
                                except Exception:
                                    pagina.wait_for_timeout(300)
                                    continue

                                if (
                                    curso_atual,
                                    turma_atual,
                                    materia_atual,
                                ) != snapshot_linha_anterior:
                                    break

                                pagina.wait_for_timeout(300)

                        veio_do_cache = obter_dados_grid_para_sede(codigo_sede_nova)
                        if not dados_grid:
                            raise Exception("Nenhum registro foi encontrado no estabelecimento selecionado.")

                        if veio_do_cache:
                            atualizar_status(
                                f"Estabelecimento já conhecido. Reaproveitando "
                                f"{len(dados_grid)} registro(s) já lidos "
                                "anteriormente."
                            )

                        executar_na_interface(
                            lambda: ciclo_combo.configure(
                                values=ciclos_disponiveis, state="readonly"
                            )
                        )
                        atualizar_status(
                            f"Estabelecimento selecionado: {sede_nova}. "
                            "Escolha o ciclo/curso."
                        )

                        while not ciclo_confirmado:
                            if parar_evento.is_set() or programa_fechando:
                                return
                            time.sleep(0.1)

                        verificar_parada()

                        while not turma_confirmada:
                            if parar_evento.is_set() or programa_fechando:
                                return
                            time.sleep(0.1)

                        verificar_parada()

                        while not materia_confirmada:
                            if parar_evento.is_set() or programa_fechando:
                                return
                            time.sleep(0.1)

                        verificar_parada()
                        atualizar_status("Abrindo o diário selecionado...")
                        clicar_no_diario(ciclo, turma, materia)

                        iframe = obter_iframe_notas()
                        combo_modulo = iframe.get_by_role("combobox", name="Módulo")
                        combo_modulo.click()
                        opcoes_modulo = iframe.get_by_role("option")
                        opcoes_modulo.first.wait_for(state="visible", timeout=10000)
                        modulos = [
                            texto.strip()
                            for texto in opcoes_modulo.all_text_contents()
                            if texto.strip()
                        ]
                                                                          
                                                                          
                        combo_modulo.press("Escape")

                        executar_na_interface(
                            lambda: modulo_combo.configure(values=modulos, state="readonly")
                        )
                        executar_na_interface(
                            lambda: botao_confirmar_modulo.configure(
                                state="normal" if modulos else "disabled"
                            )
                        )
                        atualizar_status("Diário aberto. Escolha o módulo/trimestre.")

                    except InterruptedError:
                        raise

                    except Exception as erro:
                        atualizar_status(f"Erro ao lançar nova nota: {erro}")
                        registrar_log(f"ERRO ao lançar nova nota: {erro}")
                        print("=" * 60)
                        print("ERRO AO LANÇAR NOVA NOTA (reiniciar_lancamento):")
                        print(erro)
                        print("=" * 60)

                    finally:
                        evento_reiniciar_lancamento.clear()
                        evento_reiniciar_concluido.set()
                    continue
                pedido_recebido = evento_executar_notas.wait(timeout=0.1)

                if not pedido_recebido:
                    continue

                evento_executar_notas.clear()

                if parametros_notas is None:
                    continue

                sede_exec, ciclo_exec, turma_exec, materia_exec, modulo_exec = parametros_notas

                executar_automacao(sede_exec, ciclo_exec, turma_exec, materia_exec, modulo_exec)

    except InterruptedError:
        atualizar_status("Automação interrompida pelo usuário.")
        registrar_log("Sessão interrompida pelo usuário.")

    except PlaywrightTimeoutError as erro:
        atualizar_status(f"Tempo limite excedido: {erro}")
        registrar_log(f"ERRO (timeout) na sessão: {erro}")
        print("=" * 60)
        print("TIMEOUT NO PLAYWRIGHT (automacao):")
        print(erro)
        print("=" * 60)

    except Exception as erro:
        atualizar_status(f"Erro no Playwright: {erro}")
        registrar_log(f"ERRO na sessão: {erro}")
        print("=" * 60)
        print("ERRO NO PLAYWRIGHT (automacao):")
        print(erro)
        print("=" * 60)

    finally:

        try:
            if contexto:
                contexto.close()
        except Exception:
            pass

        try:
            if navegador:
                navegador.close()
        except Exception:
            pass

        navegador = None
        contexto = None
        pagina = None

        playwright_ativo = False

        registrar_log("Sessão do Playwright encerrada. Navegador fechado.")

        if not programa_fechando:
            executar_na_interface(
                lambda: botao_continuar_login.configure(state="normal")
            )
                                                  

def mostrar_configuracao():

    global sede_confirmada, ciclo_confirmado, turma_confirmada, materia_confirmada
    global sede, ciclo, turma, materia, modulo, modulo_confirmado

    if programa_fechando:
        return

    sede_confirmada = False
    ciclo_confirmado = False
    turma_confirmada = False
    materia_confirmada = False

    sede = ""
    ciclo = ""
    turma = ""
    materia = ""
    modulo = ""
    modulo_confirmado = False

    login_frame.grid_remove()
    modelo_frame.grid_remove()
    sistema_frame.grid()

    sede_combo.configure(values=sedes_disponiveis, state="readonly")
    sede_combo.set("")
    botao_confirmar_sede.configure(state="normal")

    ciclo_combo.configure(values=[], state="disabled")
    ciclo_combo.set("")
    botao_confirmar_ciclo.configure(state="disabled")

    turma_combo.configure(values=[], state="disabled")
    turma_combo.set("")
    botao_confirmar_turma.configure(state="disabled")

    materia_combo.configure(values=[], state="disabled")
    materia_combo.set("")
    botao_confirmar_materia.configure(state="disabled")

    botao_continuar_configuracao.configure(state="disabled")

    atualizar_status("Escolha uma sede e confirme.")


def continuar_configuracao():

    global sede, ciclo, turma, materia

    sede = sede_combo.get()
    ciclo = ciclo_combo.get()
    turma = turma_combo.get()
    materia = materia_combo.get()

    if not sede:
        atualizar_status("Selecione uma sede.")
        return

    if not ciclo:
        atualizar_status("Selecione um ciclo/curso.")
        return

    if not turma:
        atualizar_status("Selecione uma turma.")
        return

    if not materia:
        atualizar_status("Selecione uma matéria.")
        return

    sistema_frame.grid_remove()

    arquivo_frame.grid()
    modulo_frame.grid()
    preview_container.grid()
    bottom.grid()

    botao_iniciar.configure(
        text="▶ INICIAR AUTOMAÇÃO",
        command=iniciar_automacao_final,
        state="disabled",
    )

                                                                          
                                                                          
    botao_lancar_novamente.configure(state="disabled")

    atualizar_status(
        f"Sede: {sede} | Ciclo: {ciclo} | Turma: {turma} | Matéria: {materia}"
    )

    verificar_configuracao_completa()
                                           

def reiniciar_lancamento():

    if lancamento_em_andamento:
        atualizar_status(
            "Aguarde o lançamento atual terminar (ou pressione F8) antes de "
            "lançar notas novamente."
        )
        return

                                                                          
                                                                          
                                                                          
    if not modulo_confirmado:
        atualizar_status(
            "Aguarde a etapa atual terminar (o diário ainda está sendo "
            "preparado) antes de lançar notas novamente."
        )
        return

                                                                          
    botao_lancar_novamente.configure(state="disabled")
    botao_iniciar.configure(state="disabled")

    atualizar_status("Preparando para lançar outra nota, aguarde...")

    evento_reiniciar_concluido.clear()
    evento_reiniciar_lancamento.set()

    _aguardar_reiniciar_lancamento(150)


def _aguardar_reiniciar_lancamento(tentativas_restantes):
    """Substitui o antigo laço bloqueante por um agendamento assíncrono
    via janela.after(), para nunca travar a interface enquanto espera a
    thread do Playwright processar o pedido de reiniciar o lançamento."""

    if programa_fechando:
        return

    if evento_reiniciar_concluido.is_set() or tentativas_restantes <= 0:
        _concluir_reiniciar_lancamento()
        return

    janela.after(
        100,
        lambda: _aguardar_reiniciar_lancamento(tentativas_restantes - 1),
    )


def _concluir_reiniciar_lancamento():

    global arquivo_selecionado, df_atual
    global modulo, modulo_confirmado, modulos_disponiveis

    if programa_fechando:
        return

    arquivo_selecionado = None
    df_atual = None
    modulo = ""
    modulo_confirmado = False
    modulos_disponiveis = []

    arquivo_label.configure(text="Nenhum arquivo selecionado")
    quantidade_label.configure(text="")

    for widget in preview_frame.winfo_children():
        widget.destroy()

    modulo_combo.set("")
    modulo_combo.configure(values=[], state="disabled")
    botao_confirmar_modulo.configure(state="disabled")

    sistema_frame.grid()
    arquivo_frame.grid_remove()
    modulo_frame.grid_remove()
    preview_container.grid_remove()
    bottom.grid()

    botao_iniciar.configure(
        text="▶ INICIAR AUTOMAÇÃO",
        command=iniciar_automacao_final,
        state="disabled",
    )

                                                                          
                                                                       
    botao_lancar_novamente.configure(state="disabled")

    atualizar_status(
        "Destino mantido. Selecione outra planilha e o módulo/trimestre para "
        "lançar novamente."
    )
                                                

def iniciar_automacao_final():

    global lancamento_em_andamento, parametros_notas

    if df_atual is None:
        atualizar_status("Selecione uma planilha primeiro.")
        return

    sede_selecionada = sede_combo.get()
    ciclo_selecionado = ciclo_combo.get()
    turma_selecionada = turma_combo.get()
    materia_selecionada = materia_combo.get()
    modulo_selecionado = modulo_combo.get()

    if not sede_selecionada:
        atualizar_status("Selecione a sede.")
        return

    if not ciclo_selecionado:
        atualizar_status("Selecione um ciclo/curso.")
        return

    if not turma_selecionada:
        atualizar_status("Selecione uma turma.")
        return

    if not materia_selecionada:
        atualizar_status("Selecione uma matéria.")
        return

    if not modulo_selecionado or not modulo_confirmado:
        atualizar_status("Selecione e confirme o módulo/trimestre.")
        return

    resposta = ctk.CTkInputDialog(
        text=(
            f"Confira antes de iniciar:\n\n"
            f"Sede: {sede_selecionada}\n"
            f"Ciclo/Curso: {ciclo_selecionado}\n"
            f"Turma: {turma_selecionada}\n"
            f"Matéria: {materia_selecionada}\n"
            f"Módulo: {modulo_selecionado}\n\n"
            f"Digite SIM para iniciar:"
        ),
        title="Confirmar automação",
    )

    if resposta is None:
        return

    if resposta.get_input().strip().upper() != "SIM":
        atualizar_status("Automação cancelada.")
        return

    lancamento_em_andamento = True

    botao_iniciar.configure(state="disabled")
    botao_lancar_novamente.configure(state="disabled")

    parar_evento.clear()
                                           
                                                                    
                                                          
    parametros_notas = (
        sede_selecionada,
        ciclo_selecionado,
        turma_selecionada,
        materia_selecionada,
        modulo_selecionado,
    )

    evento_executar_notas.set()                        
                                                              

def solicitar_parada():

    if playwright_ativo:
        parar_evento.set()
        atualizar_status(
            "Parada solicitada... (vai parar assim que a ação atual do "
            "navegador terminar)"
        )


def verificar_parada():

    if parar_evento.is_set():
        raise InterruptedError("Automação interrompida.")
                                                           

def obter_iframe_notas(tentativas=3, timeout_por_tentativa=20000, espera_entre_tentativas=1500):

    ultimo_erro = None

    for tentativa in range(tentativas):

        try:
            pagina.locator("#note-typing-iframe").wait_for(
                state="attached", timeout=timeout_por_tentativa
            )

            iframe = pagina.locator("#note-typing-iframe").content_frame

            iframe.get_by_label("sticky table").wait_for(
                state="visible", timeout=timeout_por_tentativa
            )

            return iframe

        except Exception as erro:
            ultimo_erro = erro

            if tentativa < tentativas - 1:
                pagina.wait_for_timeout(espera_entre_tentativas)                                         
                                                              

    caminho_print = None

    try:
        caminho_print = os.path.join(
            os.path.expanduser("~"), "Desktop", "erro_iframe_notas.png"
        )
        pagina.screenshot(path=caminho_print, full_page=True)
    except Exception:
        caminho_print = None

    try:
        frames_abertos = "\n".join(f"  - {f.url}" for f in pagina.frames)
    except Exception:
        frames_abertos = "  (não foi possível listar os frames)"

    try:
        url_atual = pagina.url
    except Exception:
        url_atual = "(não foi possível obter a URL)"

    mensagem = (
        "A tela de lançamento de notas (iframe '#note-typing-iframe') não "
        f"ficou pronta após {tentativas} tentativa(s).\n"
        f"Detalhe do erro: {ultimo_erro}\n\n"
        f"URL da página no momento do erro:\n  {url_atual}\n\n"
        f"Frames abertos no momento do erro:\n{frames_abertos}"
    )

    if caminho_print:
        mensagem += f"\n\nCaptura de tela salva em: {caminho_print}"

    print("=" * 60)
    print("ERRO AO OBTER O IFRAME DE NOTAS:")
    print(mensagem)
    print("=" * 60)

    registrar_log(f"ERRO ao obter iframe de notas: {ultimo_erro}")

    raise Exception(mensagem)                                                                           
                                                              

def coluna_existe_na_tabela(coluna_sistema, timeout=3000):
    """Verifica rapidamente se a coluna existe no cabeçalho do diário
    atual. Diferentes tipos de turma (ex.: turmas que não são do
    ensino médio) podem não ter todas as colunas do mapeamento padrão
    — nesse caso, a coluna deve ser pulada em vez de travar a
    automação inteira."""

    try:
        iframe = obter_iframe_notas()
        cabecalho = iframe.get_by_role("columnheader", name=coluna_sistema)
        cabecalho.first.wait_for(state="visible", timeout=timeout)
        return True
    except Exception:
        return False


def habilitar_edicao_coluna(coluna_sistema):

    try:
        iframe = obter_iframe_notas()
                                                       
                                      
        cabecalho = iframe.get_by_role("columnheader", name=coluna_sistema)

        cabecalho.wait_for(state="visible", timeout=10000)

        checkbox = cabecalho.get_by_role("checkbox")

        checkbox.wait_for(state="visible", timeout=10000)

        if not checkbox.is_checked():

            checkbox.check()
                                                      
                                                        
            for _ in range(20):
                if checkbox.is_checked():
                    break
                pagina.wait_for_timeout(100)

            pagina.wait_for_timeout(300)

        registrar_log(f"Coluna '{coluna_sistema}' habilitada para edição.")

    except Exception as erro:
        raise Exception(
            f"Não consegui habilitar a edição da coluna '{coluna_sistema}' "
            f"(caixa 'Edit.' no cabeçalho). Detalhe: {erro}"
        )
                      
                                                              

def mapear_indices_colunas(iframe, colunas_sistema):

    tabela = iframe.get_by_label("sticky table")                                                   
                                                              
                                                   
    cabecalhos = tabela.locator("thead tr").last.locator("th")

    indices = {}

    total_cabecalhos = cabecalhos.count()

    for i in range(total_cabecalhos):

        texto = cabecalhos.nth(i).inner_text().strip()                                                   
                                                                 
                                                       
        primeira_linha = texto.splitlines()[0].strip() if texto else ""

        for coluna_sistema in colunas_sistema:

            if coluna_sistema in indices:
                continue

            if textos_iguais(primeira_linha, coluna_sistema):
                indices[coluna_sistema] = i

    return indices                           
                                                              

def alterar_nota_aluno(nome_aluno, nota_texto, coluna_sistema, indices_colunas):

    iframe = obter_iframe_notas()
    tabela = iframe.get_by_label("sticky table")
                                                           
                                                        
    padrao_nome = re.compile(rf"^\s*{re.escape(nome_aluno)}\s*$", re.IGNORECASE)

    aluno = tabela.get_by_text(padrao_nome)

    quantidade_encontrada = aluno.count()

    if quantidade_encontrada == 0:
        raise AlunoNaoEncontrado(
            f"Aluno '{nome_aluno}' não foi encontrado na tabela do sistema."
        )

    if quantidade_encontrada > 1:
        raise AlunoDuplicado(
            f"Existe mais de um aluno chamado '{nome_aluno}' na tabela do "
            "sistema — pulei para não arriscar lançar a nota na pessoa errada."
        )

    aluno.first.wait_for(state="visible", timeout=5000)

    linha = aluno.first.locator("xpath=ancestor::tr")

    indice_coluna = indices_colunas.get(coluna_sistema)

    if indice_coluna is None:
        raise ColunaNaoEncontrada(
            f"Coluna '{coluna_sistema}' não foi encontrada no cabeçalho da tabela."
        )

    celulas = linha.locator("td")
    celula_nota = celulas.nth(indice_coluna)

    celula_nota.wait_for(state="visible", timeout=5000)                                                        
                                             
    alvo_clique = celula_nota.locator(".MuiTypography-root").first

    if alvo_clique.count() == 0:
        alvo_clique = celula_nota

    alvo_clique.scroll_into_view_if_needed()

                                
    alvo_clique.click()

    textbox = linha.get_by_role("textbox")
    textbox.first.wait_for(state="visible", timeout=5000)

                                                                  
                                                                   
                                                              
                  
    textbox.first.fill(nota_texto)

    return True
                                                       

def selecionar_modulo(modulo_selecionado, tentativas=3):

    texto_alvo = modulo_selecionado.strip()

                                                                        
                                                                       
                                                                          
                                                                          
    prefixo = texto_alvo[:15] if len(texto_alvo) > 15 else texto_alvo
    padrao_prefixo = re.compile("^" + re.escape(prefixo), re.IGNORECASE)

    ultimo_valor_lido = ""

    for tentativa in range(1, tentativas + 1):

        verificar_parada()

        iframe = obter_iframe_notas()
        combo = iframe.get_by_role("combobox", name="Módulo")

        combo.click()
        pagina.wait_for_timeout(150)

        opcoes = iframe.get_by_role("option")

        try:
            opcoes.first.wait_for(state="visible", timeout=5000)
        except Exception:
                                                                          
                                                                          
                                                                     
            botao_abrir = iframe.get_by_role("button", name="Abrir")

            if botao_abrir.count():
                botao_abrir.first.click()
                opcoes.first.wait_for(state="visible", timeout=10000)
            else:
                raise

                                                                          
                                                                          
        pagina.wait_for_timeout(250)

        opcao_alvo = None

        opcao_por_nome = iframe.get_by_role("option", name=padrao_prefixo)

        if opcao_por_nome.count() == 1:
            opcao_alvo = opcao_por_nome.first
        else:
            quantidade_opcoes = opcoes.count()
            for i in range(quantidade_opcoes):
                try:
                    texto_opcao = opcoes.nth(i).inner_text().strip()
                except Exception:
                    continue
                if textos_iguais(texto_opcao, texto_alvo):
                    opcao_alvo = opcoes.nth(i)
                    break

        if opcao_alvo is None:
            raise Exception(
                f"Não encontrei o módulo/trimestre '{modulo_selecionado}' na "
                "lista de opções do sistema."
            )

        opcao_alvo.scroll_into_view_if_needed()
        opcao_alvo.click()
        pagina.wait_for_timeout(400)

                                                                          
                                                                          
        try:
            ultimo_valor_lido = combo.input_value().strip()
        except Exception:
            ultimo_valor_lido = ""

        if textos_iguais(ultimo_valor_lido, texto_alvo) or normalizar(
            ultimo_valor_lido
        ).startswith(normalizar(prefixo)):
            registrar_log(f"Módulo '{modulo_selecionado}' selecionado com sucesso.")
            return

        registrar_log(
            f"AVISO: tentativa {tentativa}/{tentativas} de selecionar o "
            f"módulo '{modulo_selecionado}' não confirmou no campo (valor "
            f"lido: '{ultimo_valor_lido}'). Tentando de novo..."
        )
        pagina.wait_for_timeout(400)

    raise Exception(
        f"Não consegui confirmar a seleção do módulo/trimestre "
        f"'{modulo_selecionado}' no sistema após {tentativas} tentativa(s) "
        f"(último valor lido no campo: '{ultimo_valor_lido}')."
    )


def salvar_lancamento():
    #iframe = obter_iframe_notas()
    #iframe.locator("div").filter(has_text=re.compile(r"^Salvar$")).last.click()
    #pagina.wait_for_timeout(500)
    print("Salvo")


def executar_automacao(sede, ciclo, turma, materia, modulo_selecionado):

    global lancamento_em_andamento

    lancados = 0
    ignorados = 0
    erros = 0
    detalhes_problemas = []

    try:

        atualizar_status(f"Preparando {sede} - {ciclo} - Turma {turma} - {materia} - {modulo_selecionado}...")
        atualizar_progresso(0)
        registrar_log(f"Módulo selecionado: {modulo_selecionado}")
        selecionar_modulo(modulo_selecionado)
        registrar_log(f"Lançamento iniciado: {sede} | {ciclo} | {turma} | {materia}")

        coluna_aluno_real = encontrar_coluna(df_atual, COLUNA_ALUNO)

        if not coluna_aluno_real:
            raise Exception(f"A coluna '{COLUNA_ALUNO}' não foi encontrada.")
                                              

        colunas_ativas = []

        for coluna_planilha, coluna_sistema in MAPA_COLUNAS_NOTAS.items():

            coluna_real = encontrar_coluna(df_atual, coluna_planilha)

            if not coluna_real:
                continue

            if df_atual[coluna_real].notna().sum() == 0:
                continue

            colunas_ativas.append((coluna_planilha, coluna_real, coluna_sistema))                                               
                                                        
                                                              

        atualizar_status("Verificando quais colunas existem neste diário...")

                                                                          
                                                                          
                                                                        
        colunas_existentes = []
        colunas_puladas = []

        for item in colunas_ativas:

            verificar_parada()

            coluna_planilha, coluna_real, coluna_sistema = item

            if coluna_existe_na_tabela(coluna_sistema):
                colunas_existentes.append(item)
            else:
                colunas_puladas.append(coluna_planilha)

        colunas_ativas = colunas_existentes

        if colunas_puladas:

            mensagem_puladas = (
                "Colunas não encontradas neste diário (turma pode ser de "
                f"outro tipo/segmento) e que serão ignoradas: "
                f"{', '.join(colunas_puladas)}."
            )

            atualizar_status(mensagem_puladas)
            registrar_log(f"AVISO: {mensagem_puladas}")

        atualizar_status("Habilitando edição das colunas de notas...")

        for _, _, coluna_sistema in colunas_ativas:

            verificar_parada()

            habilitar_edicao_coluna(coluna_sistema)

                                                    

        iframe = obter_iframe_notas()

        indices_colunas = mapear_indices_colunas(
            iframe, [coluna_sistema for _, _, coluna_sistema in colunas_ativas]
        )

        for _, _, coluna_sistema in colunas_ativas:
            if coluna_sistema not in indices_colunas:
                registrar_log(
                    f"AVISO: coluna '{coluna_sistema}' não foi encontrada no "
                    "cabeçalho da tabela — notas dessa coluna serão marcadas "
                    "como erro."
                )

        alunos = df_atual[coluna_aluno_real]
        total_alunos = len(df_atual)

 
        total_notas = 0

        for indice in range(total_alunos):

            nome_bruto = alunos.iloc[indice]

            if pd.isna(nome_bruto) or not str(nome_bruto).strip():
                continue

            for _, coluna_real, _ in colunas_ativas:

                if normalizar_nota(df_atual[coluna_real].iloc[indice]) is not None:
                    total_notas += 1

        notas_processadas = 0

        for indice in range(total_alunos):

            verificar_parada()

            aluno = alunos.iloc[indice]

            if pd.isna(aluno):
                aluno = ""

            aluno = str(aluno).strip()

            if not aluno:
                atualizar_status(f"Aluno {indice + 1}/{total_alunos} sem nome. Ignorando.")
                continue

            atualizar_status(f"Processando {indice + 1}/{total_alunos}: {aluno}")

            aluno_com_problema = False

            for coluna_planilha, coluna_real, coluna_sistema in colunas_ativas:

                verificar_parada()

                nota_texto = normalizar_nota(df_atual[coluna_real].iloc[indice])

                if nota_texto is None:
                    continue

                atualizar_status(f"{aluno} → {coluna_planilha}: {nota_texto}")

                try:
                    alterar_nota_aluno(aluno, nota_texto, coluna_sistema, indices_colunas)

                    lancados += 1
                    registrar_log(f"OK: {aluno} → {coluna_planilha} = {nota_texto}")

                except (AlunoNaoEncontrado, AlunoDuplicado) as erro:
                                                           
                    ignorados += 1
                    aluno_com_problema = True

                    mensagem = f"{aluno}: {erro}"
                    detalhes_problemas.append(mensagem)
                    registrar_log(f"IGNORADO: {mensagem}")

                    break

                except PlaywrightTimeoutError as erro:

                    erros += 1

                    mensagem = (
                        f"{aluno} → {coluna_planilha}: tempo limite excedido "
                        f"({erro})"
                    )
                    detalhes_problemas.append(mensagem)
                    registrar_log(f"ERRO (timeout): {mensagem}")

                except Exception as erro:

                    erros += 1

                    mensagem = f"{aluno} → {coluna_planilha}: {erro}"
                    detalhes_problemas.append(mensagem)
                    registrar_log(f"ERRO: {mensagem}")

                notas_processadas += 1

                if total_notas:
                    atualizar_progresso(notas_processadas / total_notas)

            if aluno_com_problema:
                continue

        verificar_parada()
        atualizar_status("Salvando lançamento...")
        #salvar_lancamento()
        registrar_log("Lançamento salvo no sistema.")
                                                      

        resumo = (
            f"Concluído! Lançadas: {lancados} | Ignoradas: {ignorados} | "
            f"Erros: {erros}"
        )

        if colunas_puladas:
            resumo += f" | Colunas não existentes nesta turma: {len(colunas_puladas)}"

        atualizar_status(resumo)
        registrar_log(resumo)

        if detalhes_problemas:
            registrar_log("Detalhes dos problemas encontrados:")
            for linha_problema in detalhes_problemas:
                registrar_log(f"  - {linha_problema}")

    except InterruptedError:
        atualizar_status(
            f"Interrompido pelo usuário. Lançadas até aqui: {lancados}."
        )
        atualizar_progresso(0)
        registrar_log(f"Lançamento interrompido pelo usuário. Lançadas: {lancados}.")

    except PlaywrightTimeoutError as erro:
        atualizar_status(f"Tempo limite excedido durante o lançamento: {erro}")
        registrar_log(f"ERRO (timeout) fatal no lançamento: {erro}")
        print("=" * 60)
        print("TIMEOUT DURANTE AUTOMAÇÃO (executar_automacao):")
        print(erro)
        print("=" * 60)

    except Exception as erro:
        atualizar_status(f"Erro durante automação: {erro}")
        registrar_log(f"ERRO fatal no lançamento: {erro}")
        print("=" * 60)
        print("ERRO DURANTE AUTOMAÇÃO (executar_automacao):")
        print(erro)
        print("=" * 60)

    finally:

        lancamento_em_andamento = False
        parar_evento.clear()

        if not programa_fechando:
            executar_na_interface(
                lambda: botao_iniciar.configure(state="normal")
            )
            executar_na_interface(
                lambda: botao_lancar_novamente.configure(state="normal")
            )                                                    
                        
                                                              

def fechar_programa():

    global programa_fechando, playwright_ativo, lancamento_em_andamento

    if programa_fechando:
        return

    programa_fechando = True
    playwright_ativo = False
    lancamento_em_andamento = False
                                                
                                      
    parar_evento.set()
    evento_executar_notas.set()

    try:
        keyboard.unhook_all_hotkeys()
    except Exception:
        pass
                                                       
                                                                    
                                                           
    if thread_automacao_atual is not None and thread_automacao_atual.is_alive():

                                                                     
                                                                      
                                                    
        try:
            status_label.configure(text="Encerrando o navegador...")
            janela.update()
        except Exception:
            pass
                                                         
                                                           
        thread_automacao_atual.join(timeout=8)

        if thread_automacao_atual.is_alive():
            print(
                "Aviso: a thread do Playwright não encerrou a tempo "
                "(8s). O programa vai fechar mesmo assim; o navegador "
                "pode levar um instante a mais para sumir da tela."
            )

    registrar_log("Programa encerrado pelo usuário.")

    janela.destroy()
               
                                                              

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

janela = ctk.CTk()
janela.title("Importador de Notas")
janela.geometry("900x780")
janela.minsize(760, 620)
janela.configure(fg_color=COR_FUNDO)

janela.grid_rowconfigure(1, weight=1)
janela.grid_columnconfigure(0, weight=1)     
                                                              

header = ctk.CTkFrame(janela, corner_radius=0, fg_color=COR_LARANJA)
header.grid(row=0, column=0, sticky="ew")

ctk.CTkLabel(
    header,
    text="IMPORTADOR DE NOTAS",
    font=FONTE_TITULO_GRANDE,
    text_color=COR_BRANCO,
).pack(pady=(18, 2))

ctk.CTkLabel(
    header,
    text="Transfira as notas da sua planilha para o sistema da escola",
    text_color=COR_LARANJA_CLARO,
).pack(pady=(0, 18))                                                         
                        
                                                              

scroll = ctk.CTkScrollableFrame(janela, corner_radius=0, fg_color=COR_FUNDO)
scroll.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)
scroll.grid_columnconfigure(0, weight=1)


def criar_secao(titulo):
    """Cria um cartão branco padrão com um título em laranja."""

    frame = ctk.CTkFrame(scroll, fg_color=COR_BRANCO, corner_radius=12)

    ctk.CTkLabel(
        frame, text=titulo, font=FONTE_TITULO, text_color=COR_LARANJA
    ).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(15, 10))

    return frame                                                        
               
                                                              

passos_frame = criar_secao("COMO FUNCIONA")
passos_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 10))

passos = [
    "① Baixe o modelo da planilha e preencha as notas",
    "② Informe seu usuário e senha para acessar o sistema",
    "③ Escolha a sede, o ciclo/curso, a turma e a matéria",
    "④ Selecione o módulo/trimestre e confirme",
    "⑤ Selecione a planilha e confira a prévia dos dados",
    "⑥ Clique em INICIAR AUTOMAÇÃO e confirme com SIM",
    "⑦ O programa seleciona o módulo, lança as notas e salva",
]

for passo in passos:
    ctk.CTkLabel(
        passos_frame, text=passo, anchor="w", text_color=COR_TEXTO
    ).grid(row=passos.index(passo) + 1, column=0, columnspan=2, sticky="w", padx=25, pady=3)

ctk.CTkLabel(
    passos_frame,
    text="⚠ Durante a automação, pressione F8 para interromper (a parada acontece assim que a ação em andamento no navegador terminar).",
    text_color=COR_ALERTA,
).grid(row=len(passos) + 1, column=0, columnspan=2, sticky="w", padx=25, pady=(10, 15))
    
                                                              

modelo_frame = criar_secao("MODELO DA PLANILHA")
modelo_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=10)
modelo_frame.grid_columnconfigure(0, weight=1)

ctk.CTkLabel(
    modelo_frame,
    text=(
        f"Use a planilha padrão do programa. A coluna '{COLUNA_ALUNO}' identifica "
        "o aluno e cada uma das colunas abaixo é lançada automaticamente:"
    ),
    justify="left",
    anchor="w",
    text_color=COR_TEXTO,
    wraplength=560,
).grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

ctk.CTkLabel(
    modelo_frame,
    text=" • " + "\n • ".join(MAPA_COLUNAS_NOTAS.keys()),
    justify="left",
    anchor="w",
    text_color=COR_TEXTO,
).grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))

ctk.CTkLabel(
    modelo_frame,
    text="Você pode deixar as notas que não deseja alterar em branco.",
    text_color=COR_TEXTO_SECUNDARIO,
    anchor="w",
).grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 10))

ctk.CTkButton(
    modelo_frame,
    text="⬇ Baixar modelo padrão",
    command=baixar_modelo_planilha,
    width=210,
    fg_color=COR_LARANJA,
    hover_color=COR_LARANJA_ESCURO,
    text_color=COR_BRANCO,
).grid(row=4, column=0, sticky="w", padx=20, pady=(0, 15))
                                                              

login_frame = criar_secao("ACESSO AO SISTEMA")
login_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=10)
login_frame.grid_columnconfigure(0, weight=1)

ctk.CTkLabel(login_frame, text="Usuário", text_color=COR_TEXTO).grid(
    row=1, column=0, sticky="w", padx=20
)

usuario_entry = ctk.CTkEntry(
    login_frame,
    placeholder_text="Digite seu usuário",
    height=38,
    fg_color=COR_FUNDO,
    border_color=COR_LARANJA,
    text_color=COR_TEXTO,
)
usuario_entry.grid(row=2, column=0, sticky="ew", padx=20, pady=(5, 15))

ctk.CTkLabel(login_frame, text="Senha", text_color=COR_TEXTO).grid(
    row=3, column=0, sticky="w", padx=20
)

senha_entry = ctk.CTkEntry(
    login_frame,
    placeholder_text="Digite sua senha",
    show="•",
    height=38,
    fg_color=COR_FUNDO,
    border_color=COR_LARANJA,
    text_color=COR_TEXTO,
)
senha_entry.grid(row=4, column=0, sticky="ew", padx=20, pady=(5, 20))

botao_continuar_login = ctk.CTkButton(
    login_frame,
    text="▶ CONTINUAR",
    command=iniciar_preparacao,
    height=42,
    font=FONTE_NEGRITO,
    fg_color=COR_LARANJA,
    hover_color=COR_LARANJA_ESCURO,
    text_color=COR_BRANCO,
)
botao_continuar_login.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 20))
                                             
                                                              

sistema_frame = criar_secao("DESTINO")
sistema_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=10)
sistema_frame.grid_columnconfigure(0, weight=1)
sistema_frame.grid_columnconfigure(1, weight=0)


def criar_linha_selecao(frame, linha_label, linha_combo, texto_label):

    ctk.CTkLabel(frame, text=texto_label, text_color=COR_TEXTO).grid(
        row=linha_label, column=0, sticky="w", padx=20, pady=(0, 5)
    )

    combo = ctk.CTkComboBox(
        frame,
        values=[],
        state="disabled",
        fg_color=COR_FUNDO,
        border_color=COR_LARANJA,
        button_color=COR_LARANJA,
        button_hover_color=COR_LARANJA_ESCURO,
        text_color=COR_TEXTO,
        dropdown_fg_color=COR_BRANCO,
        dropdown_text_color=COR_TEXTO,
    )

    combo.grid(row=linha_combo, column=0, sticky="ew", padx=(20, 10), pady=(0, 20))

    return combo


def criar_botao_confirmar(frame, linha, comando):

    botao = ctk.CTkButton(
        frame,
        text="✓ CONFIRMAR",
        command=comando,
        width=130,
        height=36,
        state="disabled",
        fg_color=COR_LARANJA,
        hover_color=COR_LARANJA_ESCURO,
        text_color=COR_BRANCO,
    )

    botao.grid(row=linha, column=1, padx=(10, 20), pady=(0, 20))

    return botao


sede_combo = criar_linha_selecao(sistema_frame, 1, 2, "Sede")
sede_combo.configure(state="readonly")
botao_confirmar_sede = criar_botao_confirmar(sistema_frame, 2, confirmar_sede)
botao_confirmar_sede.configure(state="normal")

ciclo_combo = criar_linha_selecao(sistema_frame, 3, 4, "Ciclo/Curso")
botao_confirmar_ciclo = criar_botao_confirmar(sistema_frame, 4, confirmar_ciclo)

turma_combo = criar_linha_selecao(sistema_frame, 5, 6, "Turma")
botao_confirmar_turma = criar_botao_confirmar(sistema_frame, 6, confirmar_turma)

materia_combo = criar_linha_selecao(
    sistema_frame, 7, 8, "Componente curricular (Matéria)"
)
botao_confirmar_materia = criar_botao_confirmar(sistema_frame, 8, confirmar_materia)

botao_continuar_configuracao = ctk.CTkButton(
    sistema_frame,
    text="CONTINUAR →",
    command=continuar_configuracao,
    height=40,
    state="disabled",
    fg_color=COR_LARANJA,
    hover_color=COR_LARANJA_ESCURO,
    text_color=COR_BRANCO,
)
botao_continuar_configuracao.grid(
    row=9, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20)
)

sistema_frame.grid_remove()
                                                       

arquivo_frame = criar_secao("PLANILHA")
arquivo_frame.grid(row=4, column=0, sticky="ew", padx=5, pady=10)
arquivo_frame.grid_columnconfigure(1, weight=1)

ctk.CTkButton(
    arquivo_frame,
    text="📂 Selecionar arquivo",
    command=selecionar_arquivo,
    width=180,
    fg_color=COR_LARANJA,
    hover_color=COR_LARANJA_ESCURO,
    text_color=COR_BRANCO,
).grid(row=1, column=0, padx=20, pady=(5, 15))

arquivo_label = ctk.CTkLabel(
    arquivo_frame, text="Nenhum arquivo selecionado", anchor="w", text_color=COR_TEXTO
)
arquivo_label.grid(row=1, column=1, sticky="ew", padx=(5, 20), pady=(5, 15))

quantidade_label = ctk.CTkLabel(
    arquivo_frame, text="", text_color=COR_TEXTO_SECUNDARIO
)
quantidade_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 15))
                                                        

modulo_frame = criar_secao("MÓDULO / TRIMESTRE")
modulo_frame.grid(row=5, column=0, sticky="ew", padx=5, pady=10)
modulo_frame.grid_columnconfigure(0, weight=1)

ctk.CTkLabel(
    modulo_frame,
    text="Escolha o módulo/trimestre que receberá as notas:",
    text_color=COR_TEXTO,
).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 5))

modulo_combo = ctk.CTkComboBox(
    modulo_frame,
    values=[],
    state="disabled",
    fg_color=COR_FUNDO,
    border_color=COR_LARANJA,
    button_color=COR_LARANJA,
    button_hover_color=COR_LARANJA_ESCURO,
    text_color=COR_TEXTO,
    dropdown_fg_color=COR_BRANCO,
    dropdown_text_color=COR_TEXTO,
)
modulo_combo.grid(row=2, column=0, sticky="ew", padx=(20, 10), pady=(0, 10))

botao_confirmar_modulo = ctk.CTkButton(
    modulo_frame,
    text="✓ CONFIRMAR MÓDULO",
    command=confirmar_modulo,
    height=38,
    state="disabled",
    fg_color=COR_LARANJA,
    hover_color=COR_LARANJA_ESCURO,
    text_color=COR_BRANCO,
)
botao_confirmar_modulo.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 15))
                                                 

preview_container = criar_secao("PRÉVIA DOS DADOS")
preview_container.grid(row=6, column=0, sticky="ew", padx=5, pady=10)
preview_container.grid_columnconfigure(0, weight=1)

preview_frame = ctk.CTkFrame(preview_container, fg_color=COR_BRANCO)
preview_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 15))
preview_frame.grid_columnconfigure(0, weight=1)
preview_frame.grid_columnconfigure(1, weight=2)
                                               

arquivo_frame.grid_remove()
modulo_frame.grid_remove()
preview_container.grid_remove()
                                                

bottom = ctk.CTkFrame(janela, fg_color=COR_BRANCO, corner_radius=12)
bottom.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 15))
bottom.grid_columnconfigure(0, weight=1)

status_label = ctk.CTkLabel(
    bottom,
    text="Informe seu usuário e senha.",
    anchor="w",
    text_color=COR_TEXTO,
)
status_label.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="ew")

progresso = ctk.CTkProgressBar(bottom, progress_color=COR_LARANJA)
progresso.set(0)
progresso.grid(row=1, column=0, padx=15, pady=(0, 12), sticky="ew")

botoes_frame = ctk.CTkFrame(bottom, fg_color="transparent")
botoes_frame.grid(row=0, column=1, rowspan=2, padx=15, pady=10)

botao_iniciar = ctk.CTkButton(
    botoes_frame,
    text="■ INICIAR",
    command=solicitar_parada,
    state="disabled",
    width=160,
    fg_color=COR_LARANJA,
    hover_color=COR_LARANJA_ESCURO,
    text_color=COR_BRANCO,
)
botao_iniciar.pack(side="left", padx=5)

botao_lancar_novamente = ctk.CTkButton(
    botoes_frame,
    text="🔁 Lançar nota novamente",
    command=reiniciar_lancamento,
    state="disabled",
    width=190,
    fg_color=COR_BRANCO,
    hover_color=COR_LARANJA_CLARO,
    text_color=COR_LARANJA,
    border_width=2,
    border_color=COR_LARANJA,
)
botao_lancar_novamente.pack(side="left", padx=5)

botao_fechar_app = ctk.CTkButton(
    botoes_frame,
    text="✕ Fechar app",
    command=fechar_programa,
    width=130,
    fg_color=COR_TEXTO_SECUNDARIO,
    hover_color=COR_TEXTO,
    text_color=COR_BRANCO,
)
botao_fechar_app.pack(side="left", padx=5)

bottom.grid_remove()                                             
                                                              

def configurar_atalho():

    try:
        keyboard.add_hotkey("f8", solicitar_parada)
    except Exception as erro:
        print("Não foi possível registrar o F8:", erro)


configurar_atalho()


janela.protocol("WM_DELETE_WINDOW", fechar_programa)
                     
iniciar_verificacao(app)                                         

janela.mainloop()