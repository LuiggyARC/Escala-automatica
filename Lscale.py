import customtkinter as ctk
import pandas as pd
from datetime import datetime, timedelta, time # Adicionado time
import calendar
import os
import random
import re # Importar regex para validacao de horario
import webbrowser
import json # Para salvar/carregar dados
from tkcalendar import DateEntry # Para seletor de data
import traceback # Para debug

# Configuracao da aparencia
ctk.set_appearance_mode("System")  # Modos: "System" (padrao), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Temas: "blue" (padrao), "green", "dark-blue"

# --- Constantes --- 
CIDADES_PADRAO = [
    "MANAUS", "PIN - ITA", "MUR - CRI - HUM", "PORTO VELHO", 
    "JIP - ARI - COA", "RORAIMA", "AMAPA", "AM SAT E RIO BRANCO"
]
HORARIOS_PADRAO = [
    "06:00 - 12:00", "12:00 - 18:00", "18:00 - 00:00", "00:00 - 06:00", "Folguista"
]
OPCOES_EDICAO = HORARIOS_PADRAO + ["Folga", "Folga (Banco de Horas)", "-"]
ARQUIVO_DADOS = "pessoas_data.json" # Nome do arquivo para salvar dados

class EscalaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # CORRIGIDO: Versao e Titulo
        self.title("Gerador de Escalas v9.2.1 - Correção de Erros JS") 
        self.geometry("1280x800") 
        
        self.pessoas = [] 
        self.carregar_pessoas() # Carregar dados ao iniciar
        self.df_escala = None # Inicializar DataFrame da escala
        
        # Configurar tema e cores
        self.configure(fg_color=("#f0f0f0", "#2b2b2b"))
        
        # Criar abas
        self.tabview = ctk.CTkTabview(self, fg_color=("#e0e0e0", "#333333"))
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)
        
        self.tab_config = self.tabview.add("Configuracoes")
        self.tab_result = self.tabview.add("Resultado")
        self.tabview.set("Configuracoes")
        
        self.setup_config_tab()
        self.setup_result_tab()
        
        # Salvar dados ao fechar
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        self.salvar_pessoas()
        self.destroy()

    def setup_config_tab(self):
        # Frame principal com grid para melhor responsividade
        config_frame = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        config_frame.pack(fill="both", expand=True)
        config_frame.grid_columnconfigure(0, weight=1)
        config_frame.grid_rowconfigure(3, weight=1) # Permitir que a lista expanda

        # --- Cabecalho --- 
        header_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(5, 15))
        ctk.CTkLabel(
            header_frame, 
            text="Configuracao da Escala", 
            font=("Arial", 20, "bold"),
            text_color=("navy", "lightblue")
        ).pack(side="left", padx=10)
        
        # --- Secao de Periodo (com DateEntry) --- 
        period_frame = ctk.CTkFrame(config_frame, corner_radius=10)
        period_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        period_frame.grid_columnconfigure(1, weight=1)
        period_frame.grid_columnconfigure(3, weight=1)
        
        ctk.CTkLabel(
            period_frame, 
            text="Periodo da Escala", 
            font=("Arial", 14, "bold"),
            anchor="w"
        ).grid(row=0, column=0, columnspan=4, sticky="ew", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(period_frame, text="Data Inicio:", font=("Arial", 12)).grid(row=1, column=0, padx=(15, 2), pady=5, sticky="w")
        # CORRIGIDO: Removido barras invertidas dos argumentos
        self.data_inicio_entry = DateEntry(
            period_frame, 
            width=12, 
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR'
        )
        self.data_inicio_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(period_frame, text="Data Fim:", font=("Arial", 12)).grid(row=1, column=2, padx=(15, 2), pady=5, sticky="w")
        # CORRIGIDO: Removido barras invertidas dos argumentos
        self.data_fim_entry = DateEntry(
            period_frame, 
            width=12, 
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR'
        )
        self.data_fim_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        # --- Secao de Adicionar Pessoa --- 
        add_frame = ctk.CTkFrame(config_frame, corner_radius=10)
        add_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        add_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            add_frame, 
            text="Adicionar Pessoa", 
            font=("Arial", 14, "bold"),
            anchor="w"
        ).grid(row=0, column=0, columnspan=3, sticky="ew", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(add_frame, text="Nome:", font=("Arial", 12)).grid(row=1, column=0, padx=(15, 2), pady=5, sticky="w")
        self.nome_entry = ctk.CTkEntry(add_frame, font=("Arial", 12))
        self.nome_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        add_button = ctk.CTkButton(
            add_frame, 
            text="Adicionar", 
            command=self.adicionar_pessoa,
            font=("Arial", 12, "bold"),
            fg_color=("#3a7ebf", "#1f538d"),
            hover_color=("#2a6da9", "#14375e")
        )
        add_button.grid(row=1, column=2, padx=15, pady=5)

        # --- Secao de Lista de Pessoas --- 
        list_section_frame = ctk.CTkFrame(config_frame, corner_radius=10)
        list_section_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)
        list_section_frame.grid_rowconfigure(1, weight=1)
        list_section_frame.grid_columnconfigure(0, weight=1)
        
        list_header = ctk.CTkFrame(list_section_frame, fg_color="transparent")
        list_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 5))
        list_header.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(
            list_header, 
            text="Pessoas Cadastradas", 
            font=("Arial", 14, "bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(list_header, text="Filtrar por Cidade:", font=("Arial", 12)).grid(row=0, column=1, padx=(30, 5), sticky="e")
        self.cidade_filtro_var = ctk.StringVar(value="Todas as Cidades")
        opcoes_filtro_cidade = ["Todas as Cidades"] + CIDADES_PADRAO
        self.cidade_filtro_menu = ctk.CTkOptionMenu(
            list_header, 
            values=opcoes_filtro_cidade, 
            variable=self.cidade_filtro_var, 
            command=self.atualizar_lista_pessoas,
            width=180,
            font=("Arial", 12),
            dropdown_font=("Arial", 12)
        ) 
        self.cidade_filtro_menu.grid(row=0, column=2, padx=5, sticky="e")
        
        list_table_frame = ctk.CTkFrame(list_section_frame)
        list_table_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=(5, 10))
        list_table_frame.grid_columnconfigure(0, weight=1)
        list_table_frame.grid_rowconfigure(1, weight=1)
        
        header_frame = ctk.CTkFrame(list_table_frame, fg_color=("#e0e0e0", "#333333"))
        header_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        # Ajustar pesos das colunas do cabecalho
        header_frame.grid_columnconfigure(0, weight=2) # Nome
        header_frame.grid_columnconfigure(1, weight=2) # Cidade
        header_frame.grid_columnconfigure(2, weight=2) # Horario
        header_frame.grid_columnconfigure(3, weight=1) # Folga
        header_frame.grid_columnconfigure(4, weight=1) # Banco Horas
        header_frame.grid_columnconfigure(5, weight=1) # Acoes
        
        ctk.CTkLabel(header_frame, text="Nome", anchor="w", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="ew", padx=5)
        ctk.CTkLabel(header_frame, text="Cidade", anchor="w", font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="ew", padx=5) 
        ctk.CTkLabel(header_frame, text="Horario / Status", anchor="w", font=("Arial", 12, "bold")).grid(row=0, column=2, sticky="ew", padx=5) 
        ctk.CTkLabel(header_frame, text="Folga Inicial", anchor="w", font=("Arial", 12, "bold")).grid(row=0, column=3, sticky="ew", padx=5) 
        ctk.CTkLabel(header_frame, text="Banco Horas", anchor="w", font=("Arial", 12, "bold")).grid(row=0, column=4, sticky="ew", padx=5)
        ctk.CTkLabel(header_frame, text="Acoes", anchor="w", font=("Arial", 12, "bold")).grid(row=0, column=5, sticky="ew", padx=5)
        
        self.pessoas_list_frame = ctk.CTkScrollableFrame(list_table_frame)
        self.pessoas_list_frame.grid(row=1, column=0, sticky="nsew", padx=2, pady=2)
        self.pessoas_list_frame.grid_columnconfigure(0, weight=1)
        
        # --- Botoes de acao (agora no final) --- 
        action_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        # Usar grid e sticky="ew" para ocupar a largura
        action_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=(5, 10))
        action_frame.grid_columnconfigure(1, weight=1) # Coluna do meio expansivel
        
        ctk.CTkButton(
            action_frame, 
            text="Limpar Tudo", 
            command=self.limpar_dados,
            font=("Arial", 12),
            fg_color=("#d32f2f", "#b71c1c"),
            hover_color=("#c62828", "#7f0000"),
            width=120
        ).grid(row=0, column=0, padx=10, pady=5) # Alinhar a esquerda
        
        # Botao Gerar Escala alinhado a direita
        self.gerar_button = ctk.CTkButton(
            action_frame, 
            text="Gerar Escala", 
            command=self.gerar_escala,
            font=("Arial", 14, "bold"),
            fg_color=("#4caf50", "#2e7d32"),
            hover_color=("#43a047", "#1b5e20"),
            width=150,
            height=40
        )
        self.gerar_button.grid(row=0, column=2, padx=10, pady=5) # Alinhar a direita
        
        # Atualizar lista ao iniciar
        self.atualizar_lista_pessoas()
    
    def setup_result_tab(self):
        result_frame = ctk.CTkFrame(self.tab_result)
        result_frame.pack(fill="both", expand=True, padx=15, pady=15)
        result_frame.grid_columnconfigure(0, weight=1)
        result_frame.grid_rowconfigure(1, weight=1)
        
        header_frame = ctk.CTkFrame(result_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(5, 15))
        
        ctk.CTkLabel(
            header_frame, 
            text="Escala Gerada", 
            font=("Arial", 20, "bold"),
            text_color=("navy", "lightblue")
        ).pack(side="left", padx=10)
        
        ctk.CTkLabel(
            header_frame, 
            text="(Exporte para HTML para editar a escala)", 
            font=("Arial", 12, "italic"),
            text_color=("gray40", "gray60")
        ).pack(side="left", padx=10)
        
        table_frame = ctk.CTkFrame(result_frame, corner_radius=10)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        self.result_table_frame = ctk.CTkScrollableFrame(table_frame)
        self.result_table_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.result_table_frame.grid_columnconfigure(0, weight=1)
        
        action_frame = ctk.CTkFrame(result_frame, fg_color="transparent")
        action_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))
        action_frame.grid_columnconfigure(0, weight=1) # Coluna esquerda expansivel
        
        ctk.CTkButton(
            action_frame, 
            text="Exportar para Excel", 
            command=self.exportar_excel,
            font=("Arial", 12),
            fg_color=("#3a7ebf", "#1f538d"),
            hover_color=("#2a6da9", "#14375e"),
            width=150
        ).grid(row=0, column=1, padx=10, pady=5) # Alinhar a direita
        
        ctk.CTkButton(
            action_frame, 
            text="Exportar para HTML", 
            command=self.exportar_html,
            font=("Arial", 12),
            fg_color=("#ff9800", "#e65100"),
            hover_color=("#fb8c00", "#d84315"),
            width=150
        ).grid(row=0, column=2, padx=10, pady=5) # Alinhar a direita
    
    # --- Funcoes de Persistencia --- 
    def salvar_pessoas(self):
        try:
            with open(ARQUIVO_DADOS, "w", encoding="utf-8") as f:
                json.dump(self.pessoas, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Erro ao salvar dados: {e}") # Logar erro no console

    def carregar_pessoas(self):
        if os.path.exists(ARQUIVO_DADOS):
            try:
                with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
                    self.pessoas = json.load(f)
                # Adicionar campo banco_horas para pessoas existentes que não o possuem
                for pessoa in self.pessoas:
                    if "banco_horas" not in pessoa:
                        pessoa["banco_horas"] = False
            except Exception as e:
                print(f"Erro ao carregar dados: {e}")
                self.pessoas = [] # Resetar se o arquivo estiver corrompido
        else:
            self.pessoas = []

    # --- Funcoes de Configuracao --- 
    def adicionar_pessoa(self):
        nome = self.nome_entry.get().strip()
        if not nome:
            self.mostrar_erro("O nome nao pode estar vazio.")
            return
        # Verificar se nome ja existe
        if any(p["nome"].lower() == nome.lower() for p in self.pessoas):
            self.mostrar_erro(f"O nome '{nome}' ja existe.")
            return
            
        cidade_padrao = CIDADES_PADRAO[0]
        horario_padrao = HORARIOS_PADRAO[0] 
        folga_inicial_padrao = 1
        self.pessoas.append({
            "nome": nome, 
            "cidade": cidade_padrao,
            "horario": horario_padrao, 
            "folga_inicial": folga_inicial_padrao,
            "banco_horas": False  # Novo campo para banco de horas
        })
        self.nome_entry.delete(0, "end")
        self.atualizar_lista_pessoas()
        self.salvar_pessoas() # Salvar apos adicionar
        self.mostrar_info(f"Pessoa '{nome}' adicionada. Edite Cidade, Horario/Status e Folga na lista.")

    def atualizar_lista_pessoas(self, filtro_selecionado=None):
        for widget in self.pessoas_list_frame.winfo_children():
            widget.destroy()
        filtro_cidade = self.cidade_filtro_var.get()
        pessoas_filtradas = [p for p in self.pessoas if filtro_cidade == "Todas as Cidades" or p["cidade"] == filtro_cidade]
        
        cores_linhas = [("white", "#2d2d2d"), ("#f5f5f5", "#333333")]
        
        for i, pessoa in enumerate(pessoas_filtradas):
            try: original_index = self.pessoas.index(pessoa)
            except ValueError: continue
            
            cor_linha = cores_linhas[i % 2]
            pessoa_frame = ctk.CTkFrame(self.pessoas_list_frame, fg_color=cor_linha)
            # Usar grid para layout interno da linha
            pessoa_frame.grid(row=i, column=0, sticky="ew", padx=2, pady=1)
            pessoa_frame.grid_columnconfigure(0, weight=2) # Nome
            pessoa_frame.grid_columnconfigure(1, weight=2) # Cidade
            pessoa_frame.grid_columnconfigure(2, weight=2) # Horario
            pessoa_frame.grid_columnconfigure(3, weight=1) # Folga
            pessoa_frame.grid_columnconfigure(4, weight=1) # Banco Horas
            pessoa_frame.grid_columnconfigure(5, weight=1) # Acoes
            
            ctk.CTkLabel(
                pessoa_frame, 
                text=pessoa["nome"], 
                anchor="w",
                font=("Arial", 12)
            ).grid(row=0, column=0, sticky="ew", padx=5)
            
            # Campo Cidade
            cidade_var = ctk.StringVar(value=pessoa["cidade"])
            cidade_menu = ctk.CTkOptionMenu(
                pessoa_frame,
                values=CIDADES_PADRAO,
                variable=cidade_var,
                width=120,
                font=("Arial", 11),
                dropdown_font=("Arial", 11),
                command=lambda value, idx=original_index: self.salvar_campo(idx, "cidade", value)
            )
            cidade_menu.grid(row=0, column=1, sticky="ew", padx=5)
            
            # Campo Horario/Status
            horario_var = ctk.StringVar(value=pessoa["horario"])
            horario_menu = ctk.CTkOptionMenu(
                pessoa_frame,
                values=HORARIOS_PADRAO,
                variable=horario_var,
                width=120,
                font=("Arial", 11),
                dropdown_font=("Arial", 11),
                command=lambda value, idx=original_index: self.salvar_campo_horario(idx, value)
            )
            horario_menu.grid(row=0, column=2, sticky="ew", padx=5)
            
            # Campo Folga Inicial
            folga_entry = ctk.CTkEntry(
                pessoa_frame,
                width=60,
                font=("Arial", 11),
                justify="center"
            )
            folga_entry.insert(0, str(pessoa["folga_inicial"]) if pessoa["horario"] != "Folguista" else "-")
            folga_entry.bind("<Return>", lambda event, idx=original_index, entry=folga_entry: self.salvar_folga_inicial(idx, entry))
            folga_entry.bind("<FocusOut>", lambda event, idx=original_index, entry=folga_entry: self.salvar_folga_inicial(idx, entry))
            folga_entry.grid(row=0, column=3, sticky="ew", padx=5)
            
            # Campo Banco de Horas
            banco_horas_var = ctk.BooleanVar(value=pessoa.get("banco_horas", False))
            banco_horas_checkbox = ctk.CTkCheckBox(
                pessoa_frame,
                text="",
                variable=banco_horas_var,
                width=20,
                command=lambda idx=original_index, var=banco_horas_var: self.salvar_campo(idx, "banco_horas", var.get())
            )
            banco_horas_checkbox.grid(row=0, column=4, sticky="ew", padx=5)
            
            ctk.CTkButton(
                pessoa_frame, 
                text="Remover", 
                width=80,
                font=("Arial", 12),
                fg_color=("#d32f2f", "#b71c1c"),
                hover_color=("#c62828", "#7f0000"),
                command=lambda idx=original_index: self.remover_pessoa(idx)
            ).grid(row=0, column=5, sticky="ew", padx=5)

    def salvar_campo(self, index, campo, novo_valor):
         if 0 <= index < len(self.pessoas):
            valor_antigo = self.pessoas[index].get(campo)
            self.pessoas[index][campo] = novo_valor
            self.salvar_pessoas() # Salvar apos editar campo
            if campo == "cidade" and self.cidade_filtro_var.get() != "Todas as Cidades" and valor_antigo != novo_valor:
                 self.atualizar_lista_pessoas()
         else:
             self.mostrar_erro("Erro ao salvar: Indice da pessoa invalido.")
    
    def salvar_campo_horario(self, index, novo_horario):
        if 0 <= index < len(self.pessoas):
            horario_antigo = self.pessoas[index]["horario"]
            self.pessoas[index]["horario"] = novo_horario
            if novo_horario == "Folguista" and horario_antigo != "Folguista":
                self.pessoas[index]["folga_inicial"] = 0
                self.mostrar_info(f"Folga inicial removida para o folguista {self.pessoas[index]['nome']}.")
            if horario_antigo == "Folguista" and novo_horario != "Folguista":
                self.pessoas[index]["folga_inicial"] = 1
                self.mostrar_info(f"Folga inicial definida como 1 para {self.pessoas[index]['nome']}.")
            self.salvar_pessoas() # Salvar apos editar horario
            self.atualizar_lista_pessoas()
        else:
            self.mostrar_erro("Erro ao salvar: Indice da pessoa invalido.")

    def salvar_folga_inicial(self, index, entry_widget):
        if self.pessoas[index]["horario"] == "Folguista":
            self.mostrar_erro("Folguistas nao possuem folga inicial definida.")
            entry_widget.delete(0, "end"); entry_widget.insert(0, "-"); return
            
        novo_valor_str = entry_widget.get().strip()
        if not novo_valor_str.isdigit():
            self.mostrar_erro("A folga inicial deve ser um numero.")
            entry_widget.delete(0, "end"); entry_widget.insert(0, str(self.pessoas[index]["folga_inicial"])); return
        novo_valor = int(novo_valor_str)
        
        # Validar folga inicial com base no periodo selecionado (se possivel)
        try:
            # CORRIGIDO: Converter para datetime para validacao
            data_inicio_date = self.data_inicio_entry.get_date()
            data_inicio = datetime.combine(data_inicio_date, time.min)
            
            # Validacao basica do dia (1-31)
            if novo_valor < 1 or novo_valor > 31: 
                 self.mostrar_erro(f"A folga inicial deve ser um dia valido (1-31).")
                 entry_widget.delete(0, "end"); entry_widget.insert(0, str(self.pessoas[index]["folga_inicial"])); return
            # Validacao mais especifica (dia existe no mes de inicio)
            try:
                primeiro_dia_mes_inicio = data_inicio.replace(day=1)
                ultimo_dia_mes_inicio = primeiro_dia_mes_inicio.replace(day=calendar.monthrange(data_inicio.year, data_inicio.month)[1])
                if novo_valor > ultimo_dia_mes_inicio.day:
                    self.mostrar_erro(f"O dia {novo_valor} nao existe no mes de inicio ({data_inicio.strftime('%m/%Y')}).")
                    entry_widget.delete(0, "end"); entry_widget.insert(0, str(self.pessoas[index]["folga_inicial"])); return
            except Exception as date_val_err:
                print(f"Aviso: Nao foi possivel validar o dia da folga com o mes de inicio: {date_val_err}")
                # Continuar mesmo sem validacao de dia no mes

        except Exception as get_date_err: 
             print(f"Aviso: Nao foi possivel obter as datas para validar a folga inicial: {get_date_err}")
             # Continuar mesmo sem validacao de data especifica

        self.salvar_campo(index, "folga_inicial", novo_valor)
        self.salvar_pessoas() # Salvar apos editar folga
        self.mostrar_info(f"Folga inicial de {self.pessoas[index]['nome']} atualizada para {novo_valor}.") 
        self.pessoas_list_frame.focus()

    def remover_pessoa(self, index):
        if 0 <= index < len(self.pessoas):
            nome_removido = self.pessoas[index]["nome"]
            del self.pessoas[index]
            self.atualizar_lista_pessoas() 
            self.salvar_pessoas() # Salvar apos remover
            self.mostrar_info(f"Pessoa '{nome_removido}' removida.")
    
    def limpar_dados(self):
        self.pessoas = []; self.cidade_filtro_var.set("Todas as Cidades") 
        self.atualizar_lista_pessoas(); self.nome_entry.delete(0, "end")
        self.salvar_pessoas() # Salvar apos limpar
        self.mostrar_info("Todos os dados de pessoas foram limpos.")
    
    # --- Geracao da Escala (com Date Range e Debugging) --- 
    def gerar_escala(self):
        print("Botao Gerar Escala clicado.") # DEBUG
        try:
            # CORRIGIDO: Obter como date e converter para datetime
            data_inicio_date = self.data_inicio_entry.get_date()
            data_fim_date = self.data_fim_entry.get_date()
            data_inicio = datetime.combine(data_inicio_date, time.min) # Converte para datetime 00:00:00
            data_fim = datetime.combine(data_fim_date, time.min)     # Converte para datetime 00:00:00
            print(f"Datas obtidas (convertidas para datetime): Inicio={data_inicio}, Fim={data_fim}") # DEBUG
        except Exception as e:
            self.mostrar_erro(f"Erro ao obter datas: {e}. Verifique o formato (DD/MM/AAAA) e se as datas sao validas.")
            print(f"Erro ao obter datas: {e}") # DEBUG
            return

        if data_inicio > data_fim:
            self.mostrar_erro("A data de inicio nao pode ser posterior a data de fim.")
            print("Erro: Data de inicio posterior a data de fim.") # DEBUG
            return

        filtro_cidade = self.cidade_filtro_var.get()
        pessoas_para_escala = [p for p in self.pessoas if filtro_cidade == "Todas as Cidades" or p["cidade"] == filtro_cidade]
        print(f"Pessoas filtradas ({filtro_cidade}): {len(pessoas_para_escala)}") # DEBUG

        if not pessoas_para_escala:
            self.mostrar_erro(f"Nenhuma pessoa encontrada para a cidade '{filtro_cidade}'.")
            print("Erro: Nenhuma pessoa encontrada para o filtro.") # DEBUG
            return

        folguistas = [p for p in pessoas_para_escala if p["horario"] == "Folguista"]
        regulares = [p for p in pessoas_para_escala if p["horario"] != "Folguista"]
        print(f"Regulares: {len(regulares)}, Folguistas: {len(folguistas)}") # DEBUG

        if not regulares:
            self.mostrar_erro("Nenhum funcionario regular encontrado para gerar a escala base.")
            print("Erro: Nenhum funcionario regular encontrado.") # DEBUG
            return
        if not folguistas:
            self.mostrar_info("Nenhum folguista encontrado. A escala sera gerada apenas para os regulares.")
            print("Info: Nenhum folguista encontrado.") # DEBUG

        try:
            print("Chamando gerar_escala_final_daterange...") # DEBUG
            # Passar datas como datetime
            self.gerar_escala_final_daterange(data_inicio, data_fim, regulares, folguistas)
            print("gerar_escala_final_daterange concluido.") # DEBUG
            self.tabview.set("Resultado")
            print("Aba Resultado selecionada.") # DEBUG
        except Exception as e:
            # Adicionar um tratamento de erro mais generico aqui
            self.mostrar_erro(f"Ocorreu um erro inesperado ao gerar a escala: {e}")
            print(f"Erro inesperado em gerar_escala: {e}") # DEBUG
            traceback.print_exc() # Printar o traceback completo no console
    
    # CORRIGIDO: Garante que todas as datas sao datetime + SEPARAÇÃO POR CIDADE
    def gerar_escala_final_daterange(self, data_inicio, data_fim, regulares, folguistas):
        print("Iniciando gerar_escala_final_daterange com separação por cidade...") # DEBUG
        try: # Adicionar try-except geral aqui
            for widget in self.result_table_frame.winfo_children(): widget.destroy()
            print("Widgets antigos da aba Resultado destruidos.") # DEBUG

            # Gerar lista de datas no intervalo (ja sao datetime)
            datas_escala = []
            data_atual = data_inicio
            while data_atual <= data_fim:
                datas_escala.append(data_atual)
                data_atual += timedelta(days=1)
            print(f"Datas da escala geradas: {len(datas_escala)} dias.") # DEBUG

            if not datas_escala:
                self.mostrar_erro("Intervalo de datas invalido resultou em 0 dias.")
                print("Erro: Intervalo de datas invalido.") # DEBUG
                return

            # NOVA LÓGICA: Agrupar funcionários por cidade
            cidades = {}
            for pessoa in regulares + folguistas:
                cidade = pessoa.get("cidade", "SEM_CIDADE")
                if cidade not in cidades:
                    cidades[cidade] = {"regulares": [], "folguistas": []}
                
                if pessoa in regulares:
                    cidades[cidade]["regulares"].append(pessoa)
                else:
                    cidades[cidade]["folguistas"].append(pessoa)
            
            print(f"Cidades encontradas: {list(cidades.keys())}") # DEBUG
            for cidade, funcionarios in cidades.items():
                print(f"  {cidade}: {len(funcionarios['regulares'])} regulares, {len(funcionarios['folguistas'])} folguistas") # DEBUG

            # Criar DataFrame final com todos os funcionários
            todos_nomes = [p["nome"] for p in regulares + folguistas]
            df_escala_final = pd.DataFrame(index=todos_nomes, columns=datas_escala).fillna("")
            print("DataFrame final criado.") # DEBUG

            # Processar cada cidade independentemente
            for cidade, funcionarios_cidade in cidades.items():
                print(f"\n=== PROCESSANDO CIDADE: {cidade} ===") # DEBUG
                
                regulares_cidade = funcionarios_cidade["regulares"]
                folguistas_cidade = funcionarios_cidade["folguistas"]
                
                if not regulares_cidade:
                    print(f"  Cidade {cidade} sem funcionários regulares, pulando.") # DEBUG
                    continue
                
                # Processar escala para esta cidade específica
                df_cidade = self.processar_escala_cidade(
                    data_inicio, data_fim, datas_escala, 
                    regulares_cidade, folguistas_cidade, cidade
                )
                
                # Copiar resultados para o DataFrame final
                print(f"  Copiando resultados da cidade {cidade} para DataFrame final...") # DEBUG
                for nome in df_cidade.index:
                    if nome in df_escala_final.index:
                        for data in datas_escala:
                            valor_original = df_escala_final.at[nome, data]
                            valor_cidade = df_cidade.at[nome, data]
                            df_escala_final.at[nome, data] = valor_cidade
                            if valor_cidade == "Folga":
                                print(f"    Copiando folga: {nome} em {data.strftime('%d/%m')}") # DEBUG

            self.df_escala = df_escala_final # Armazenar para exportacao
            print("DataFrame final consolidado armazenado.") # DEBUG

            print("Chamando exibir_escala_em_blocos_seg_dom_daterange...") # DEBUG
            self.exibir_escala_em_blocos_seg_dom_daterange(df_escala_final)
            print("exibir_escala_em_blocos_seg_dom_daterange concluido.") # DEBUG

        except Exception as e:
            self.mostrar_erro(f"Erro interno ao calcular a escala: {e}")
            print(f"Erro interno em gerar_escala_final_daterange: {e}") # DEBUG
            traceback.print_exc() # Printar o traceback completo no console

    def processar_escala_cidade(self, data_inicio, data_fim, datas_escala, regulares, folguistas, cidade):
        """Processa a escala para uma cidade específica"""
        print(f"  Processando escala para cidade {cidade}...") # DEBUG
        
        nomes_regulares = [p["nome"] for p in regulares]
        nomes_folguistas = [p["nome"] for p in folguistas]
        todos_nomes = nomes_regulares + nomes_folguistas
        
        # Criar DataFrame para esta cidade
        df_escala = pd.DataFrame(index=todos_nomes, columns=datas_escala).fillna("")
        
        horario_por_nome = {p["nome"]: p["horario"] for p in regulares}
        folgas_por_data = {data: [] for data in datas_escala}
        folgas_regulares_map = {}
        
        turno_18_00_nomes = [p["nome"] for p in regulares if p["horario"] == "18:00 - 00:00"]
        turno_00_06_nomes = [p["nome"] for p in regulares if p["horario"] == "00:00 - 06:00"]
        folgas_apos_cobertura = {nome: set() for nome in turno_18_00_nomes}
        
        print(f"    Funcionários da cidade {cidade}:") # DEBUG
        print(f"      Regulares: {nomes_regulares}") # DEBUG
        print(f"      Folguistas: {nomes_folguistas}") # DEBUG
        print(f"      Turno 18:00-00:00: {turno_18_00_nomes}") # DEBUG
        print(f"      Turno 00:00-06:00: {turno_00_06_nomes}") # DEBUG

        # 1. Calcular folgas dos REGULARES desta cidade
        print(f"  Calculando folgas dos regulares da cidade {cidade}...") # DEBUG
        for pessoa in regulares:
            nome = pessoa["nome"]
            folga_inicial_dia = pessoa["folga_inicial"]
            horario = pessoa["horario"]
            tem_banco_horas = pessoa.get("banco_horas", False)
            
            # CORREÇÃO: Pular folguistas (não devem ter folgas calculadas)
            if horario == "Folguista":
                print(f"    Pulando {nome} (folguista) - não calcula folgas automaticamente.") # DEBUG
                continue
            print(f"    Calculando folgas para: {nome} (Folga inicial: {folga_inicial_dia})") # DEBUG
            
            # Encontrar a primeira data de folga no intervalo ou apos (usando datetime)
            primeira_folga_data = None
            try:
                # Tenta criar a data no mes/ano de inicio (como datetime)
                data_teste = datetime(data_inicio.year, data_inicio.month, folga_inicial_dia)
            except ValueError: # Dia invalido para o mes
                # Se invalido, pega o ultimo dia do mes de inicio (como datetime)
                ultimo_dia_mes_inicio = calendar.monthrange(data_inicio.year, data_inicio.month)[1]
                data_teste = datetime(data_inicio.year, data_inicio.month, ultimo_dia_mes_inicio)
                print(f"      Aviso: Dia {folga_inicial_dia} invalido para {data_inicio.strftime('%m/%Y')}. Usando {ultimo_dia_mes_inicio} como base inicial.")
            
            # A comparacao aqui agora e entre datetime e datetime
            while data_teste < data_inicio:
                 data_teste += timedelta(days=8) # Pula para proxima folga 7x1 (8 dias = avança 1 dia da semana)
            primeira_folga_data = data_teste
            
            # Gerar todas as folgas no intervalo (usando datetime)
            folgas_funcionario = set()
            data_folga_atual = primeira_folga_data
            while data_folga_atual <= data_fim:
                folgas_funcionario.add(data_folga_atual)
                print(f"      Folga programada: {nome} em {data_folga_atual.strftime('%d/%m')}") # DEBUG
                
                # Adicionar segundo dia de folga se tem banco de horas
                if tem_banco_horas:
                    segunda_folga = data_folga_atual + timedelta(days=1)
                    if segunda_folga <= data_fim:
                        folgas_funcionario.add(segunda_folga)
                        print(f"      Segunda folga (banco de horas): {nome} em {segunda_folga.strftime('%d/%m')}") # DEBUG
                
                data_folga_atual += timedelta(days=8) # Proxima folga 7x1
            
            folgas_regulares_map[nome] = folgas_funcionario
            
            # Aplicar folgas no DataFrame
            for data_folga in folgas_funcionario:
                if data_folga in datas_escala:
                    df_escala.at[nome, data_folga] = "Folga"
                    folgas_por_data[data_folga].append(nome)

        # 2. Aplicar horarios regulares (apenas para quem nao esta de folga)
        print(f"  Aplicando horários regulares da cidade {cidade}...") # DEBUG
        for pessoa in regulares:
            nome = pessoa["nome"]
            horario = pessoa["horario"]
            
            if horario == "Folguista":
                continue # Folguistas nao tem horario fixo
            
            for data in datas_escala:
                if df_escala.at[nome, data] == "": # Se nao esta de folga
                    df_escala.at[nome, data] = horario
                    print(f"    Aplicando horário: {nome} - {horario} em {data.strftime('%d/%m')}") # DEBUG

        # 3. Processar coberturas do turno 00:00-06:00 desta cidade
        print(f"  Processando coberturas 00:00-06:00 da cidade {cidade}...") # DEBUG
        for data in datas_escala:
            funcionarios_00_06_de_folga = []
            for nome in turno_00_06_nomes:
                if df_escala.at[nome, data] == "Folga":
                    funcionarios_00_06_de_folga.append(nome)
                    print(f"    {nome} (00:00-06:00) está de folga em {data.strftime('%d/%m')}") # DEBUG
            
            if funcionarios_00_06_de_folga:
                print(f"    COBERTURA NECESSÁRIA em {data.strftime('%d/%m')} (cidade {cidade}) para: {funcionarios_00_06_de_folga}") # DEBUG
                
                for funcionario_de_folga in funcionarios_00_06_de_folga:
                    # Encontrar cobertor do turno 18:00-00:00 disponível
                    cobertor = None
                    for nome_18_00 in turno_18_00_nomes:
                        if df_escala.at[nome_18_00, data] != "Folga":
                            cobertor = nome_18_00
                            break
                    
                    if cobertor:
                        print(f"      {cobertor} cobrira o turno 00:00-06:00 de {funcionario_de_folga}.") # DEBUG
                        df_escala.at[cobertor, data] = "18:00 - 00:00 + 00:00 - 06:00 (Cob)"
                        
                        # Verificar se há folguista disponível para cobrir o 18:00-00:00
                        if folguistas:
                            folguista_disponivel = None
                            for folguista in folguistas:
                                folguista_nome = folguista["nome"]
                                if folguista_nome in df_escala.index:
                                    folguista_status = df_escala.at[folguista_nome, data]
                                    if folguista_status == "-" or pd.isna(folguista_status) or folguista_status == "":
                                        folguista_disponivel = folguista_nome
                                        break
                            
                            if folguista_disponivel:
                                print(f"      {folguista_disponivel} (folguista) cobrira o turno 18:00-00:00 de {cobertor}.") # DEBUG
                                df_escala.at[folguista_disponivel, data] = "18:00 - 00:00"
                            else:
                                print(f"      Aviso: Nenhum folguista disponível para cobrir turno 18:00-00:00 de {cobertor}.") # DEBUG
                        
                        # Dar folga pós-cobertura
                        cobertor_pessoa = next((p for p in regulares if p["nome"] == cobertor), None)
                        tem_banco_horas_cobertor = cobertor_pessoa.get("banco_horas", False) if cobertor_pessoa else False
                        
                        # Verificar se é folga estendida
                        em_folga_estendida = False
                        if data.weekday() in [5, 6]:  # Sábado ou domingo
                            em_folga_estendida = True
                        elif data.weekday() == 0:  # Segunda
                            sabado_anterior = data - timedelta(days=2)
                            domingo_anterior = data - timedelta(days=1)
                            folgas_funcionario = folgas_regulares_map.get(funcionario_de_folga, set())
                            if sabado_anterior in folgas_funcionario or domingo_anterior in folgas_funcionario:
                                em_folga_estendida = True
                        elif data.weekday() == 1:  # Terça
                            domingo_anterior = data - timedelta(days=2)
                            folgas_funcionario = folgas_regulares_map.get(funcionario_de_folga, set())
                            funcionario_folga_pessoa = next((p for p in regulares if p["nome"] == funcionario_de_folga), None)
                            tem_banco_horas_folga = funcionario_folga_pessoa.get("banco_horas", False) if funcionario_folga_pessoa else False
                            if domingo_anterior in folgas_funcionario and tem_banco_horas_folga:
                                em_folga_estendida = True
                        
                        if em_folga_estendida:
                            # Durante folga estendida: folga apenas na terça
                            dias_para_terca = (1 - data.weekday() + 7) % 7
                            if dias_para_terca == 0:
                                dias_para_terca = 7
                            data_terca = data + timedelta(days=dias_para_terca)
                            
                            if data_terca <= data_fim and data_terca not in folgas_regulares_map.get(cobertor, set()):
                                print(f"      {cobertor} tera folga na terca: {data_terca.strftime('%d/%m')}.") # DEBUG
                                df_escala.at[cobertor, data_terca] = "Folga"
                                folgas_apos_cobertura.setdefault(cobertor, set()).add(data_terca)
                                folgas_por_data.setdefault(data_terca, []).append(cobertor)
                        else:
                            # Comportamento normal: folga no dia seguinte
                            data_seguinte = data + timedelta(days=1)
                            if data_seguinte <= data_fim and data_seguinte not in folgas_regulares_map.get(cobertor, set()):
                                print(f"      {cobertor} tera folga pos-cobertura em {data_seguinte.strftime('%d/%m')}.") # DEBUG
                                df_escala.at[cobertor, data_seguinte] = "Folga"
                                folgas_apos_cobertura.setdefault(cobertor, set()).add(data_seguinte)
                                folgas_por_data.setdefault(data_seguinte, []).append(cobertor)
                                
                                if tem_banco_horas_cobertor:
                                    data_segundo_dia = data_seguinte + timedelta(days=1)
                                    if data_segundo_dia <= data_fim and data_segundo_dia not in folgas_regulares_map.get(cobertor, set()):
                                        print(f"      {cobertor} tera segundo dia de folga pos-cobertura em {data_segundo_dia.strftime('%d/%m')} (banco de horas).") # DEBUG
                                        df_escala.at[cobertor, data_segundo_dia] = "Folga"
                                        folgas_apos_cobertura.setdefault(cobertor, set()).add(data_segundo_dia)
                                        folgas_por_data.setdefault(data_segundo_dia, []).append(cobertor)
            else:
                print(f"    NENHUMA FOLGA no turno 00-06 em {data.strftime('%d/%m')} (cidade {cidade}) - SEM COBERTURA NECESSÁRIA") # DEBUG

        # 3. Alocar FOLGUISTAS desta cidade
        print(f"  Alocando folguistas da cidade {cidade}...") # DEBUG
        if folguistas:
            for data in datas_escala:
                print(f"    Processando folguistas para {data.strftime('%d/%m')} (cidade {cidade})...") # DEBUG
                
                funcionarios_de_folga = []
                for nome in df_escala.index:
                    nomes_folguistas = [f["nome"] for f in folguistas]  # CORRIGIDO: Extrair nomes dos folguistas
                    if nome not in nomes_folguistas:
                        valor_celula = df_escala.at[nome, data]
                        if valor_celula == "Folga":
                            horario_funcionario = horario_por_nome.get(nome, "")
                            if horario_funcionario != "00:00 - 06:00":
                                funcionarios_de_folga.append((nome, horario_funcionario))
                                print(f"      {nome} ({horario_funcionario}) precisa de cobertura") # DEBUG
                
                if funcionarios_de_folga:
                    folguistas_disponiveis = []
                    for folguista in folguistas:
                        nome_folguista = folguista["nome"]
                        valor_folguista = df_escala.at[nome_folguista, data]
                        if valor_folguista == "" or valor_folguista == "-":
                            folguistas_disponiveis.append(folguista)
                    
                    for i, (nome_funcionario, horario_funcionario) in enumerate(funcionarios_de_folga):
                        if i < len(folguistas_disponiveis):
                            folguista = folguistas_disponiveis[i]
                            nome_folguista = folguista["nome"]
                            print(f"      Alocando {nome_folguista} para cobrir {nome_funcionario} ({horario_funcionario})") # DEBUG
                            df_escala.at[nome_folguista, data] = horario_funcionario

        # Preencher células vazias dos folguistas com '-'
        for pessoa_f in folguistas:
            nome_f = pessoa_f["nome"]
            for data in datas_escala:
                if df_escala.at[nome_f, data] == "":
                    df_escala.at[nome_f, data] = "-"

        print(f"  Escala da cidade {cidade} processada com sucesso.") # DEBUG
        return df_escala

    # --- Exibicao e Exportacao (Adaptadas para Date Range e Debugging) ---
    def exibir_escala_em_blocos_seg_dom_daterange(self, df):
        print("Iniciando exibir_escala_em_blocos_seg_dom_daterange...") # DEBUG
        try: # Adicionar try-except geral aqui
            for widget in self.result_table_frame.winfo_children(): widget.destroy()
            print("Widgets antigos da aba Resultado (exibicao) destruidos.") # DEBUG

            datas_escala = list(df.columns) # Ja sao datetime
            if not datas_escala:
                print("Erro: DataFrame vazio para exibir.") # DEBUG
                return
            data_inicio = datas_escala[0]
            data_fim = datas_escala[-1]
            print(f"Exibindo escala de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}") # DEBUG

            dias_semana_abrev = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex", 5: "Sab", 6: "Dom"}
            nomes_folguistas = [p["nome"] for p in self.pessoas if p["horario"] == "Folguista"] # Obter nomes dos folguistas

            # Logica para Blocos de Seg-Dom (usando datetime)
            blocos_seg_dom = []
            # Encontra a primeira segunda-feira no periodo ou logo apos o inicio
            dias_para_primeira_segunda = (0 - data_inicio.weekday() + 7) % 7
            primeira_segunda_data = data_inicio + timedelta(days=dias_para_primeira_segunda)
            
            # Adiciona dias iniciais (se houver) antes da primeira segunda
            dias_iniciais = [d for d in datas_escala if d < primeira_segunda_data]
            if dias_iniciais: blocos_seg_dom.append(dias_iniciais)
            
            # Adiciona blocos completos de segunda a domingo
            data_bloco_atual = primeira_segunda_data
            while data_bloco_atual <= data_fim:
                fim_bloco_data = data_bloco_atual + timedelta(days=6)
                bloco_atual = [d for d in datas_escala if data_bloco_atual <= d <= fim_bloco_data]
                if bloco_atual: blocos_seg_dom.append(bloco_atual)
                data_bloco_atual += timedelta(days=7)
            print(f"Numero de blocos semanais a exibir: {len(blocos_seg_dom)}") # DEBUG
            
            cor_cabecalho = ("#1a237e", "#0d47a1")
            cor_controlador = ("#263238", "#37474f")
            cores_linhas = [("white", "#2d2d2d"), ("#f5f5f5", "#333333")]
            cor_folga = ("#1f538d", "#1976d2")
            cor_folga_pos_cob = ("#42a5f5", "#64b5f6")
            cor_cobertura_00_06 = ("#ffa726", "#ffb74d")
            # MODIFICAÇÃO 1: Remover cor verde do folguista - usar cor padrão
            cor_folguista = cores_linhas[0]  # Usar cor padrão em vez de verde
            
            row_idx_global = 0
            for bloco_num, bloco_datas in enumerate(blocos_seg_dom):
                print(f"Exibindo bloco {bloco_num + 1} com {len(bloco_datas)} dias.") # DEBUG
                
                # Cabecalho do bloco
                bloco_header_frame = ctk.CTkFrame(self.result_table_frame, fg_color=cor_cabecalho)
                bloco_header_frame.grid(row=row_idx_global, column=0, sticky="ew", padx=5, pady=(10, 2))
                bloco_header_frame.grid_columnconfigure(0, weight=1)
                
                primeira_data_bloco = bloco_datas[0]
                ultima_data_bloco = bloco_datas[-1]
                titulo_bloco = f"Semana {bloco_num + 1}: {primeira_data_bloco.strftime('%d/%m')} - {ultima_data_bloco.strftime('%d/%m/%Y')}"
                ctk.CTkLabel(
                    bloco_header_frame, 
                    text=titulo_bloco, 
                    font=("Arial", 14, "bold"),
                    text_color="white"
                ).grid(row=0, column=0, pady=8)
                row_idx_global += 1
                
                # Cabecalho da tabela
                table_header_frame = ctk.CTkFrame(self.result_table_frame, fg_color=cor_controlador)
                table_header_frame.grid(row=row_idx_global, column=0, sticky="ew", padx=5, pady=2)
                table_header_frame.grid_columnconfigure(0, weight=2) # Nome
                for col_idx in range(len(bloco_datas)):
                    table_header_frame.grid_columnconfigure(col_idx + 1, weight=1)
                
                ctk.CTkLabel(
                    table_header_frame, 
                    text="Funcionario", 
                    font=("Arial", 11, "bold"),
                    text_color="white",
                    anchor="w"
                ).grid(row=0, column=0, sticky="ew", padx=10, pady=5)
                
                for col_idx, data in enumerate(bloco_datas):
                    dia_str = f"{data.day}/{data.month}\n({dias_semana_abrev[data.weekday()]})"
                    ctk.CTkLabel(
                        table_header_frame, 
                        text=dia_str, 
                        font=("Arial", 10, "bold"),
                        text_color="white",
                        anchor="center"
                    ).grid(row=0, column=col_idx + 1, sticky="ew", padx=2, pady=5)
                row_idx_global += 1
                
                # Linhas de dados
                for pessoa_idx, nome in enumerate(df.index):
                    is_folguista = nome in nomes_folguistas
                    cor_linha = cores_linhas[pessoa_idx % 2]
                    
                    linha_frame = ctk.CTkFrame(self.result_table_frame, fg_color=cor_linha)
                    linha_frame.grid(row=row_idx_global + pessoa_idx, column=0, sticky="ew", padx=5, pady=1)
                    linha_frame.grid_columnconfigure(0, weight=2) # Nome
                    for col_idx in range(len(bloco_datas)):
                        linha_frame.grid_columnconfigure(col_idx + 1, weight=1)
                    
                    # Nome do funcionario
                    nome_display = f"{nome} {'(Folguista)' if is_folguista else ''}"
                    nome_label = ctk.CTkLabel(
                        linha_frame,
                        text=nome_display,
                        font=("Arial", 11, "bold" if not is_folguista else "normal"),
                        anchor="w",
                        fg_color=cor_controlador[ctk.get_appearance_mode() == "Dark"],
                        text_color="white"
                    )
                    nome_label.grid(row=0, column=0, sticky="nsew", padx=2, pady=1)
                    
                    # Celulas de dados
                    for col_idx, data in enumerate(bloco_datas):
                        valor_celula = df.at[nome, data]
                        cor_fundo_celula = cor_linha[ctk.get_appearance_mode() == "Dark"]
                        cor_texto = ("black", "white")
                        
                        if valor_celula == "Folga":
                            cor_fundo_celula = cor_folga[ctk.get_appearance_mode() == "Dark"]
                            cor_texto = ("white", "white")
                        elif "(Cob)" in str(valor_celula):
                            cor_fundo_celula = cor_cobertura_00_06[ctk.get_appearance_mode() == "Dark"]
                            cor_texto = ("black", "black")
                        # MODIFICAÇÃO 1: Folguista trabalhando usa cor padrão (sem verde)
                        elif is_folguista and valor_celula != "-": # Folguista trabalhando
                            cor_fundo_celula = cor_folguista[ctk.get_appearance_mode() == "Dark"]
                            cor_texto = ("black", "white")

                        cell_label = ctk.CTkLabel(
                            linha_frame,
                            text=str(valor_celula), # Usar str() para seguranca
                            font=("Arial", 10),
                            anchor="center",
                            fg_color=cor_fundo_celula,
                            text_color=cor_texto[ctk.get_appearance_mode() == "Dark"]
                        )
                        cell_label.grid(row=0, column=col_idx + 1, sticky="nsew", padx=1, pady=1)

                row_idx_global += len(df.index) + 2 # Atualiza indice global para proximo bloco
            print("Exibicao da escala concluida.") # DEBUG

        except Exception as e:
            self.mostrar_erro(f"Erro ao exibir a escala na interface: {e}")
            print(f"Erro em exibir_escala_em_blocos_seg_dom_daterange: {e}") # DEBUG
            traceback.print_exc() # Printar o traceback completo no console

    def exportar_excel(self):
        if self.df_escala is None or self.df_escala.empty:
            self.mostrar_erro("Nenhuma escala gerada para exportar.")
            return
        try:
            filtro_cidade = self.cidade_filtro_var.get().replace(' ', '_').replace('-', '_')
            data_inicio_str = self.df_escala.columns[0].strftime('%d_%m_%Y')
            data_fim_str = self.df_escala.columns[-1].strftime('%d_%m_%Y')
            filename = f"escala_{filtro_cidade}_{data_inicio_str}_a_{data_fim_str}.xlsx"
            
            # Criar um novo Excel Writer
            writer = pd.ExcelWriter(filename, engine='openpyxl')
            
            # Converter datas para strings formatadas no DataFrame
            df_export = self.df_escala.copy()
            df_export.columns = [data.strftime('%d/%m/%Y') for data in df_export.columns]
            
            # Exportar DataFrame para Excel
            df_export.to_excel(writer, sheet_name='Escala', index=True)
            
            # Ajustar largura das colunas
            worksheet = writer.sheets['Escala']
            for i, col in enumerate(df_export.columns):
                # Ajustar largura com base no conteúdo
                max_len = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(66 + i)].width = max_len
            
            # Ajustar largura da coluna de nomes
            max_nome_len = max(len(str(nome)) for nome in df_export.index) + 5
            worksheet.column_dimensions['A'].width = max_nome_len
            
            # Salvar o arquivo
            writer.close()
            
            self.mostrar_info(f"Escala exportada para {filename}")
            # Abrir o arquivo Excel
            if os.path.exists(filename):
                webbrowser.open(f"file://{os.path.realpath(filename)}")
        except Exception as e:
            self.mostrar_erro(f"Erro ao exportar para Excel: {e}")
            print(f"Erro ao exportar Excel: {e}")
            traceback.print_exc()

    def exportar_html(self):
        if self.df_escala is None or self.df_escala.empty:
            self.mostrar_erro("Nenhuma escala gerada para exportar.")
            return
        try:
            filtro_cidade = self.cidade_filtro_var.get().replace(' ', '_').replace('-', '_')
            data_inicio_str = self.df_escala.columns[0].strftime('%d_%m_%Y')
            data_fim_str = self.df_escala.columns[-1].strftime('%d_%m_%Y')
            filename = f"escala_{filtro_cidade}_{data_inicio_str}_a_{data_fim_str}.html"
            
            html_content = self.gerar_html_editavel(self.df_escala)
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)
            self.mostrar_info(f"Escala exportada para {filename}")
            webbrowser.open(f"file://{os.path.realpath(filename)}")
        except Exception as e:
            self.mostrar_erro(f"Erro ao exportar para HTML: {e}")
            print(f"Erro ao exportar HTML: {e}")
            traceback.print_exc()

    def gerar_html_editavel(self, df):
        datas_escala = list(df.columns)
        data_inicio = datas_escala[0]
        data_fim = datas_escala[-1]
        dias_semana_abrev = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex", 5: "Sab", 6: "Dom"}
        nomes_folguistas = [p["nome"] for p in self.pessoas if p["horario"] == "Folguista"]

        # Obter cidades únicas dos funcionários na escala
        cidades_na_escala = set()
        for nome in df.index:
            pessoa = next((p for p in self.pessoas if p["nome"] == nome), None)
            if pessoa:
                cidades_na_escala.add(pessoa["cidade"])
        cidades_na_escala = sorted(list(cidades_na_escala))

        # Criar mapeamento nome -> cidade
        nome_para_cidade = {}
        for nome in df.index:
            pessoa = next((p for p in self.pessoas if p["nome"] == nome), None)
            if pessoa:
                nome_para_cidade[nome] = pessoa["cidade"]

        # Logica para Blocos de Seg-Dom (igual a da exibicao)
        blocos_seg_dom = []
        dias_para_primeira_segunda = (0 - data_inicio.weekday() + 7) % 7
        primeira_segunda_data = data_inicio + timedelta(days=dias_para_primeira_segunda)
        dias_iniciais = [d for d in datas_escala if d < primeira_segunda_data]
        if dias_iniciais: blocos_seg_dom.append(dias_iniciais)
        data_bloco_atual = primeira_segunda_data
        while data_bloco_atual <= data_fim:
            fim_bloco_data = data_bloco_atual + timedelta(days=6)
            bloco_atual = [d for d in datas_escala if data_bloco_atual <= d <= fim_bloco_data]
            if bloco_atual: blocos_seg_dom.append(bloco_atual)
            data_bloco_atual += timedelta(days=7)

        # CORREÇÃO: CSS e JavaScript melhorados para HTML editável
        css = """
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                margin: 20px; 
                background-color: #f0f0f0; 
                line-height: 1.4;
            }
            h1 { 
                color: #0d47a1; 
                text-align: center; 
                margin-bottom: 30px;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
            }
            .filter-container { 
                text-align: center; 
                margin: 20px 0; 
                padding: 20px; 
                background-color: #ffffff; 
                border-radius: 10px; 
                box-shadow: 0 4px 8px rgba(0,0,0,0.1); 
                border: 1px solid #e0e0e0;
            }
            .filter-container h3 {
                margin: 0 0 15px 0; 
                color: #0d47a1;
                font-size: 16px;
            }
            .filter-button { 
                display: inline-block; 
                margin: 5px 8px; 
                padding: 10px 18px; 
                background-color: #e8e8e8; 
                color: #333; 
                border: 2px solid #d0d0d0;
                border-radius: 8px; 
                cursor: pointer; 
                font-size: 13px; 
                font-weight: 500;
                transition: all 0.3s ease; 
                user-select: none;
            }
            .filter-button:hover { 
                background-color: #d8d8d8; 
                border-color: #b0b0b0;
                transform: translateY(-1px);
            }
            .filter-button.active { 
                background-color: #0d47a1; 
                color: white; 
                border-color: #0d47a1;
                box-shadow: 0 2px 4px rgba(13, 71, 161, 0.3);
            }
            .escala-container { 
                margin-bottom: 30px; 
                background-color: #ffffff; 
                border-radius: 10px; 
                box-shadow: 0 4px 8px rgba(0,0,0,0.1); 
                overflow: hidden; 
                border: 1px solid #e0e0e0;
            }
            table { 
                width: 100%; 
                border-collapse: collapse; 
            }
            th, td { 
                border: 1px solid #ddd; 
                padding: 10px 8px; 
                text-align: center; 
                font-size: 11px; 
                min-width: 85px; 
                height: 40px; 
                position: relative; 
                vertical-align: middle;
            }
            th { 
                background-color: #0d47a1; 
                color: white; 
                font-weight: bold; 
                font-size: 10px; 
                white-space: pre-wrap; 
                text-shadow: 1px 1px 1px rgba(0,0,0,0.2);
            }
            .nome-col { 
                background-color: #37474f; 
                color: white; 
                text-align: left; 
                font-weight: bold; 
                min-width: 160px; 
                padding-left: 12px;
            }
            .folguista-nome { 
                font-weight: normal; 
                font-style: italic;
            }
            .editable-cell { 
                cursor: pointer; 
                transition: all 0.2s ease;
                user-select: none;
            }
            .editable-cell:hover { 
                background-color: #f5f5f5 !important; 
                border-color: #0d47a1 !important;
                box-shadow: inset 0 0 0 2px #0d47a1;
            }
            .folga { 
                background-color: #1f538d !important; 
                color: white !important; 
                font-weight: bold;
            }
            .folga-banco-horas { 
                background-color: #1565c0 !important; 
                color: white !important; 
                font-weight: bold;
            }
            .cobertura-00-06 { 
                background-color: #ffa726 !important; 
                color: black !important; 
                font-weight: bold;
            }
            .folguista-trabalhando { 
                background-color: #e3f2fd !important; 
                color: #1565c0 !important; 
                font-weight: 500;
            }
            .funcionario-row { 
                transition: opacity 0.3s ease, transform 0.2s ease; 
            }
            .funcionario-row.hidden { 
                display: none; 
            }
            .dropdown-menu {
                display: none;
                position: absolute;
                background-color: white;
                border: 2px solid #0d47a1;
                border-radius: 8px;
                box-shadow: 0 8px 16px rgba(0,0,0,0.2);
                z-index: 1000;
                min-width: 140px;
                max-height: 250px;
                overflow-y: auto;
            }
            .dropdown-menu div { 
                padding: 12px 15px; 
                cursor: pointer; 
                font-size: 12px; 
                border-bottom: 1px solid #f0f0f0;
                transition: background-color 0.2s ease;
            }
            .dropdown-menu div:last-child {
                border-bottom: none;
            }
            .dropdown-menu div:hover { 
                background-color: #e3f2fd; 
                color: #0d47a1;
            }
            .action-buttons {
                text-align: center;
                margin: 25px 0;
                padding: 20px;
                background-color: #ffffff;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            }
            .print-button, .readonly-button { 
                display: inline-block; 
                margin: 0 10px; 
                padding: 12px 24px; 
                font-size: 14px; 
                font-weight: 600;
                cursor: pointer; 
                border: none; 
                border-radius: 8px; 
                transition: all 0.3s ease;
                text-decoration: none;
            }
            .print-button {
                background-color: #4caf50; 
                color: white;
            }
            .readonly-button {
                background-color: #2196F3; 
                color: white;
            }
            .print-button:hover { 
                background-color: #45a049; 
                transform: translateY(-2px);
                box-shadow: 0 4px 8px rgba(76, 175, 80, 0.3);
            }
            .readonly-button:hover { 
                background-color: #1976D2; 
                transform: translateY(-2px);
                box-shadow: 0 4px 8px rgba(33, 150, 243, 0.3);
            }
            
            /* Sistema de Alertas de Conflito Melhorado */
            .conflict-alert {
                background: linear-gradient(135deg, #f44336, #d32f2f);
                color: white;
                padding: 20px;
                margin: 25px 0;
                border-radius: 10px;
                box-shadow: 0 6px 12px rgba(244, 67, 54, 0.3);
                display: none;
                border-left: 5px solid #b71c1c;
            }
            .conflict-alert h3 {
                margin: 0 0 15px 0;
                font-size: 18px;
                display: flex;
                align-items: center;
                gap: 10px;
            }
            .conflict-item {
                background-color: rgba(255,255,255,0.15);
                padding: 12px;
                margin: 8px 0;
                border-radius: 6px;
                font-size: 13px;
                border-left: 3px solid rgba(255,255,255,0.5);
            }
            .conflict-cell {
                background-color: #f44336 !important;
                color: white !important;
                animation: conflictPulse 2s infinite;
                border: 3px solid #d32f2f !important;
                font-weight: bold !important;
            }
            @keyframes conflictPulse {
                0% { box-shadow: 0 0 0 0 rgba(244, 67, 54, 0.7); }
                70% { box-shadow: 0 0 0 15px rgba(244, 67, 54, 0); }
                100% { box-shadow: 0 0 0 0 rgba(244, 67, 54, 0); }
            }
            @media print {
                body { margin: 0; background-color: white; }
                .escala-container { box-shadow: none; border: 1px solid #ccc; margin-bottom: 15px; }
                .action-buttons, .filter-container, .conflict-alert { display: none; }
                h1 { font-size: 16pt; }
                th, td { font-size: 9pt; padding: 5px; }
            }
        </style>
        """ 

        html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Escala Editável - {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}</title>
    {css}
</head>
<body>
    <h1>📅 Escala Editável ({data_inicio.strftime('%d/%m/%Y')} - {data_fim.strftime('%d/%m/%Y')})</h1>
    
    <!-- Filtro por Cidade -->
    <div class="filter-container">
        <h3>🏙️ Filtrar por Cidade:</h3>
        <button class="filter-button active" onclick="filterByCidade('todas', event)">Todas as Cidades</button>"""
        
        for cidade in cidades_na_escala:
            cidade_id = cidade.replace(' ', '_').replace('-', '_')
            html += f"""
        <button class="filter-button" onclick="filterByCidade('{cidade_id}', event)">{cidade}</button>"""
        
        html += """
    </div>
    
    <!-- Alerta de Conflitos -->
    <div id="conflictAlert" class="conflict-alert">
        <h3>⚠️ CONFLITOS DETECTADOS - Múltiplas pessoas no mesmo horário</h3>
        <div id="conflictList"></div>
    </div>
    
    <div class="action-buttons">
        <button class="print-button" onclick="window.print()">🖨️ Imprimir Escala</button>
        <button class="readonly-button" onclick="gerarHTMLSomenteLeitura()">📄 Gerar Versão Somente Leitura</button>
    </div>
"""

        for bloco_num, bloco_datas in enumerate(blocos_seg_dom):
            html += f"""<div class="escala-container">
        <table>
            <thead>
                <tr>
                    <th class="nome-col">Funcionário</th>
"""
            for data in bloco_datas:
                dia_str_html = f"{data.day}/{data.month}<br>({dias_semana_abrev[data.weekday()]})"
                html += f"<th>{dia_str_html}</th>"
            html += "</tr></thead><tbody>"

            for nome in df.index:
                is_folguista = nome in nomes_folguistas
                nome_class = "folguista-nome" if is_folguista else ""
                cidade_funcionario = nome_para_cidade.get(nome, "")
                cidade_id = cidade_funcionario.replace(' ', '_').replace('-', '_')
                html += f"""<tr class="funcionario-row" data-cidade="{cidade_id}">
                    <td class="nome-col {nome_class}">{nome} {'(Folguista)' if is_folguista else ''}</td>
"""
                for data in bloco_datas:
                    valor_celula = df.at[nome, data]
                    cell_id = f"cell_{nome.replace(' ', '_')}_{data.strftime('%Y%m%d')}"
                    cell_class = "editable-cell"
                    if valor_celula == "Folga": 
                        cell_class += " folga"
                    elif valor_celula == "Folga (Banco de Horas)": 
                        cell_class += " folga-banco-horas"
                    elif valor_celula == "Folga (Pos-Cob)": 
                        cell_class += " folga-pos-cob"
                    elif "(Cob)" in str(valor_celula): 
                        cell_class += " cobertura-00-06"
                    elif is_folguista and valor_celula != "-": 
                        cell_class += " folguista-trabalhando"
                    
                    html += f"""<td id="{cell_id}" class="{cell_class}" onclick="showDropdown(event, '{cell_id}')">{str(valor_celula)}</td>"""
                html += "</tr>"
            html += "</tbody></table></div>"

        # Dropdown Menu (oculto inicialmente)
        html += '<div id="dropdown" class="dropdown-menu">'
        for option in OPCOES_EDICAO:
            html += f"""<div onclick="selectOption('{option}')">{option}</div>"""
        html += '</div>'

        # CORREÇÃO: JavaScript melhorado para edição e filtro por cidade
        js = f"""
        <script>
            let currentCellId = null;
            let dropdown = null;
            const horariosPadrao = {json.dumps(HORARIOS_PADRAO)};
            const nomesFolguistas = {json.dumps(nomes_folguistas)};

            // Inicializar quando DOM carregar
            document.addEventListener('DOMContentLoaded', function() {{
                console.log('DOM carregado, inicializando...');
                dropdown = document.getElementById('dropdown');
                if (!dropdown) {{
                    console.error('Elemento dropdown não encontrado!');
                    return;
                }}
                
                // Detectar conflitos iniciais
                setTimeout(() => {{
                    detectarConflitos();
                }}, 500);
                
                console.log('Inicialização concluída');
            }});

            // Função melhorada para filtrar por cidade
            function filterByCidade(cidadeId, event) {{
                console.log('Filtrando por cidade:', cidadeId);
                
                // Remover classe active de todos os botões
                const buttons = document.querySelectorAll('.filter-button');
                buttons.forEach(btn => btn.classList.remove('active'));
                
                // Adicionar classe active ao botão clicado
                if (event && event.target) {{
                    event.target.classList.add('active');
                }}
                
                // Mostrar/ocultar linhas baseado na cidade
                const rows = document.querySelectorAll('.funcionario-row');
                let visibleCount = 0;
                
                rows.forEach(row => {{
                    const rowCidade = row.getAttribute('data-cidade') || '';
                    if (cidadeId === 'todas' || rowCidade === cidadeId) {{
                        row.classList.remove('hidden');
                        visibleCount++;
                    }} else {{
                        row.classList.add('hidden');
                    }}
                }});
                
                console.log(`Filtro aplicado: ${{visibleCount}} funcionários visíveis`);
                
                // Redetectar conflitos após filtro
                setTimeout(() => {{
                    detectarConflitos();
                }}, 100);
            }}

            // Função melhorada para mostrar dropdown
            function showDropdown(event, cellId) {{
                console.log('Mostrando dropdown para célula:', cellId);
                
                if (!dropdown) {{
                    console.error('Dropdown não inicializado!');
                    return;
                }}
                
                // Fechar dropdown anterior se existir
                if (dropdown.style.display === 'block') {{
                    dropdown.style.display = 'none';
                }}
                
                currentCellId = cellId;
                const cell = document.getElementById(cellId);
                if (!cell) {{
                    console.error('Célula não encontrada:', cellId);
                    return;
                }}
                
                const rect = cell.getBoundingClientRect();
                dropdown.style.display = 'block';
                
                // Posicionamento melhorado do dropdown
                let top = rect.bottom + window.scrollY + 5;
                let left = rect.left + window.scrollX;
                
                // Ajustar se sair da tela horizontalmente
                if (left + dropdown.offsetWidth > window.innerWidth) {{
                    left = window.innerWidth - dropdown.offsetWidth - 10;
                }}
                if (left < 10) {{
                    left = 10;
                }}
                
                // Ajustar se sair da tela verticalmente
                if (top + dropdown.offsetHeight > window.innerHeight + window.scrollY) {{
                    top = rect.top + window.scrollY - dropdown.offsetHeight - 5;
                }}
                
                dropdown.style.top = top + 'px';
                dropdown.style.left = left + 'px';
                
                event.stopPropagation();
                console.log('Dropdown posicionado em:', {{ top, left }});
            }}

            // Função melhorada para selecionar opção
            function selectOption(value) {{
                console.log('Opção selecionada:', value, 'para célula:', currentCellId);
                
                if (!dropdown || !currentCellId) {{
                    console.error('Dropdown ou célula não definidos!');
                    return;
                }}
                
                const cell = document.getElementById(currentCellId);
                if (!cell) {{
                    console.error('Célula não encontrada:', currentCellId);
                    return;
                }}
                
                // Atualizar conteúdo da célula
                cell.textContent = value;
                
                // Atualizar estilo da célula
                updateCellStyle(cell, value);
                
                // Fechar dropdown
                dropdown.style.display = 'none';
                currentCellId = null;
                
                // Verificar conflitos após mudança
                setTimeout(() => {{
                    detectarConflitos();
                }}, 100);
                
                console.log('Célula atualizada com sucesso');
            }}

            // Função melhorada para atualizar estilo da célula
            function updateCellStyle(cell, value) {{
                const parts = cell.id.split('_');
                const nome = parts.slice(1, -1).join(' ');
                const isFolguista = nomesFolguistas.includes(nome);

                // Remover todas as classes de estilo
                cell.classList.remove('folga', 'folga-banco-horas', 'cobertura-00-06', 'folguista-trabalhando');

                // Aplicar nova classe baseada no valor
                if (value === "Folga") {{
                    cell.classList.add('folga');
                }} else if (value === "Folga (Banco de Horas)") {{
                    cell.classList.add('folga-banco-horas');
                }} else if (value.includes("(Cob)")) {{
                    cell.classList.add('cobertura-00-06');
                }} else if (isFolguista && value !== '-' && value !== '') {{
                    cell.classList.add('folguista-trabalhando');
                }}
                
                console.log('Estilo da célula atualizado:', value);
            }}

            // Função melhorada para detectar conflitos de horário
            function detectarConflitos() {{
                console.log('Detectando conflitos...');
                const conflitos = [];
                const tabelas = document.querySelectorAll('table');
                
                // Limpar marcações de conflito anteriores
                document.querySelectorAll('.conflict-cell').forEach(cell => {{
                    cell.classList.remove('conflict-cell');
                }});
                
                tabelas.forEach((tabela, tabelaIndex) => {{
                    const cabecalhos = tabela.querySelectorAll('thead th');
                    const datas = [];
                    
                    // Extrair datas dos cabeçalhos (pular primeira coluna que é nome)
                    for (let i = 1; i < cabecalhos.length; i++) {{
                        const textoHeader = cabecalhos[i].textContent.trim();
                        datas.push(textoHeader);
                    }}
                    
                    // Para cada data, verificar conflitos POR CIDADE
                    datas.forEach((data, colIndex) => {{
                        const horariosPorCidade = {{}}; // Agrupar por cidade
                        const linhas = tabela.querySelectorAll('tbody tr');
                        
                        linhas.forEach(linha => {{
                            // Verificar se a linha está visível (não filtrada)
                            if (linha.classList.contains('hidden')) return;
                            
                            const celulas = linha.querySelectorAll('td');
                            if (celulas.length > colIndex + 1) {{
                                const nomeCompleto = celulas[0].textContent.trim();
                                const nome = nomeCompleto.replace('(Folguista)', '').trim();
                                const horario = celulas[colIndex + 1].textContent.trim();
                                const cidade = linha.getAttribute('data-cidade') || 'DESCONHECIDA';
                                
                                // Ignorar folgas, células vazias e valores especiais
                                if (horario === 'Folga' || horario === 'Folga (Banco de Horas)' || 
                                    horario === '-' || horario === '' || 
                                    horario.includes('Folga') || !horario.includes(':')) {{
                                    return;
                                }}
                                
                                // Remover "(Cob)" para comparar horários base
                                const horarioBase = horario.replace(/\s*\(Cob\)\s*/g, '').trim();
                                
                                // Agrupar por cidade e horário
                                if (!horariosPorCidade[cidade]) {{
                                    horariosPorCidade[cidade] = {{}};
                                }}
                                if (!horariosPorCidade[cidade][horarioBase]) {{
                                    horariosPorCidade[cidade][horarioBase] = [];
                                }}
                                horariosPorCidade[cidade][horarioBase].push({{
                                    nome: nome,
                                    celula: celulas[colIndex + 1]
                                }});
                            }}
                        }});
                        
                        // Verificar conflitos DENTRO DE CADA CIDADE
                        Object.keys(horariosPorCidade).forEach(cidade => {{
                            const horariosDaCidade = horariosPorCidade[cidade];
                            Object.keys(horariosDaCidade).forEach(horario => {{
                                const pessoas = horariosDaCidade[horario];
                                if (pessoas.length > 1) {{
                                    const nomes = pessoas.map(p => p.nome);
                                    conflitos.push({{
                                        data: data,
                                        cidade: cidade.replace(/_/g, ' '), // Converter underscores para espaços
                                        horario: horario,
                                        funcionarios: nomes,
                                        tabela: tabelaIndex + 1
                                    }});
                                    
                                    // Marcar células em conflito
                                    pessoas.forEach(pessoa => {{
                                        pessoa.celula.classList.add('conflict-cell');
                                    }});
                                }}
                            }});
                        }});
                    }});
                }});
                
                // Exibir alertas de conflito
                exibirAlertasConflito(conflitos);
                console.log('Detecção de conflitos concluída:', conflitos.length, 'conflitos encontrados');
            }}

            // Função melhorada para exibir alertas de conflito
            function exibirAlertasConflito(conflitos) {{
                const alertContainer = document.getElementById('conflictAlert');
                const conflictList = document.getElementById('conflictList');
                
                if (!alertContainer || !conflictList) {{
                    console.error('Elementos de alerta não encontrados');
                    return;
                }}
                
                if (conflitos.length > 0) {{
                    conflictList.innerHTML = '';
                    conflitos.forEach((conflito, index) => {{
                        const item = document.createElement('div');
                        item.className = 'conflict-item';
                        item.innerHTML = `
                            <strong>Conflito #${{index + 1}}</strong><br>
                            📅 Data: ${{conflito.data}}<br>
                            🏙️ Cidade: ${{conflito.cidade}}<br>
                            🕐 Horário: ${{conflito.horario}}<br>
                            👥 Funcionários: ${{conflito.funcionarios.join(', ')}}
                            ${{conflito.tabela ? `<br>📊 Tabela: ${{conflito.tabela}}` : ''}}
                        `;
                        conflictList.appendChild(item);
                    }});
                    alertContainer.style.display = 'block';
                    console.log('Alertas de conflito exibidos');
                }} else {{
                    alertContainer.style.display = 'none';
                    console.log('Nenhum conflito encontrado');
                }}
            }}

            // Função para gerar HTML somente leitura (melhorada)
            function gerarHTMLSomenteLeitura() {{
                try {{
                    console.log('Gerando HTML somente leitura...');
                    
                    // Coletar dados atuais das tabelas
                    const dadosEscala = coletarDadosTabelas();
                    
                    // Gerar HTML somente leitura
                    const htmlSomenteLeitura = gerarHTMLVisualizacao(dadosEscala);
                    
                    // Criar e baixar arquivo
                    const blob = new Blob([htmlSomenteLeitura], {{ type: 'text/html;charset=utf-8' }});
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'escala_somente_leitura.html';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    
                    alert('✅ HTML somente leitura gerado com sucesso!\\n\\nArquivo: escala_somente_leitura.html');
                    console.log('HTML somente leitura gerado com sucesso');
                    
                }} catch (error) {{
                    console.error('Erro ao gerar HTML somente leitura:', error);
                    alert('❌ Erro ao gerar HTML somente leitura: ' + error.message);
                }}
            }}

            // Função melhorada para coletar dados das tabelas
            function coletarDadosTabelas() {{
                const dados = [];
                const tabelas = document.querySelectorAll('table');
                
                tabelas.forEach((tabela, tabelaIndex) => {{
                    const dadosTabela = {{
                        index: tabelaIndex,
                        cabecalhos: [],
                        linhas: []
                    }};
                    
                    // Coletar cabeçalhos
                    const cabecalhos = tabela.querySelectorAll('thead th');
                    cabecalhos.forEach(th => {{
                        dadosTabela.cabecalhos.push(th.innerHTML);
                    }});
                    
                    // Coletar dados das linhas visíveis
                    const linhas = tabela.querySelectorAll('tbody tr');
                    linhas.forEach(linha => {{
                        if (linha.classList.contains('hidden')) return; // Pular linhas ocultas
                        
                        const dadosLinha = {{
                            cidade: linha.getAttribute('data-cidade') || '',
                            celulas: []
                        }};
                        
                        const celulas = linha.querySelectorAll('td');
                        celulas.forEach(celula => {{
                            dadosLinha.celulas.push({{
                                conteudo: celula.textContent.trim(),
                                classes: celula.className
                            }});
                        }});
                        
                        dadosTabela.linhas.push(dadosLinha);
                    }});
                    
                    dados.push(dadosTabela);
                }});
                
                console.log('Dados das tabelas coletados:', dados.length, 'tabelas');
                return dados;
            }}

            // Função melhorada para gerar HTML de visualização
            function gerarHTMLVisualizacao(dadosEscala) {{
                const agora = new Date();
                const dataHora = agora.toLocaleString('pt-BR');
                
                // Coletar cidades únicas
                const cidadesUnicas = new Set();
                dadosEscala.forEach(tabela => {{
                    tabela.linhas.forEach(linha => {{
                        if (linha.cidade) {{
                            cidadesUnicas.add(linha.cidade);
                        }}
                    }});
                }});
                
                let html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Escala - Versão Final</title>
    <style>
        body {{ 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 20px; 
            background-color: #f0f0f0; 
            line-height: 1.4;
        }}
        h1 {{ 
            color: #0d47a1; 
            text-align: center; 
            margin-bottom: 15px;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        }}
        .subtitle {{
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-style: italic;
            font-size: 14px;
        }}
        .escala-container {{ 
            margin-bottom: 30px; 
            background-color: #ffffff; 
            border-radius: 10px; 
            box-shadow: 0 4px 8px rgba(0,0,0,0.1); 
            overflow: hidden; 
            border: 1px solid #e0e0e0;
        }}
        table {{ 
            width: 100%; 
            border-collapse: collapse; 
        }}
        th, td {{ 
            border: 1px solid #ddd; 
            padding: 10px 8px; 
            text-align: center; 
            font-size: 11px; 
            min-width: 85px; 
            height: 40px; 
            vertical-align: middle;
        }}
        th {{ 
            background-color: #0d47a1; 
            color: white; 
            font-weight: bold; 
            font-size: 10px; 
            white-space: pre-wrap; 
            text-shadow: 1px 1px 1px rgba(0,0,0,0.2);
        }}
        .nome-col {{ 
            background-color: #37474f; 
            color: white; 
            text-align: left; 
            font-weight: bold; 
            min-width: 160px; 
            padding-left: 12px;
        }}
        .folguista-nome {{ 
            font-weight: normal; 
            font-style: italic;
        }}
        .folga {{ 
            background-color: #1f538d; 
            color: white; 
            font-weight: bold;
        }}
        .folga-banco-horas {{ 
            background-color: #1565c0; 
            color: white; 
            font-weight: bold;
        }}
        .cobertura-00-06 {{ 
            background-color: #ffa726; 
            color: black; 
            font-weight: bold;
        }}
        .folguista-trabalhando {{ 
            background-color: #e3f2fd; 
            color: #1565c0; 
            font-weight: 500;
        }}
        .city-filter {{
            text-align: center;
            margin: 25px 0;
            padding: 20px;
            background-color: #ffffff;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            border: 1px solid #e0e0e0;
        }}
        .city-filter select {{
            padding: 10px 20px;
            font-size: 14px;
            border: 2px solid #0d47a1;
            border-radius: 8px;
            background-color: white;
            color: #0d47a1;
            cursor: pointer;
            min-width: 200px;
        }}
        .city-filter label {{
            font-weight: bold;
            margin-right: 15px;
            color: #0d47a1;
            font-size: 16px;
        }}
        .funcionario-row {{
            transition: opacity 0.3s ease;
        }}
        .funcionario-row.hidden {{
            display: none;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding: 25px;
            background-color: #ffffff;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            color: #666;
            font-size: 13px;
            border: 1px solid #e0e0e0;
        }}
        @media print {{
            body {{ margin: 0; background-color: white; }}
            .escala-container {{ box-shadow: none; border: 1px solid #ccc; margin-bottom: 15px; }}
            h1 {{ font-size: 16pt; }}
            th, td {{ font-size: 9pt; padding: 5px; }}
            .footer, .city-filter {{ display: none; }}
        }}
    </style>
</head>
<body>
    <h1>📅 Escala de Trabalho - Versão Final</h1>
    <div class="subtitle">Gerado em: ${{dataHora}} | Versão somente leitura com filtro por cidade</div>
    
    <!-- Filtro por Cidade Melhorado -->
    <div class="city-filter">
        <label for="cityFilter">🏙️ Filtrar por Cidade:</label>
        <select id="cityFilter" onchange="filtrarPorCidade()">
            <option value="todas">📍 Todas as Cidades</option>`;
                
                // Adicionar opções de cidade
                Array.from(cidadesUnicas).sort().forEach(cidade => {{
                    const cidadeDisplay = cidade.replace(/_/g, ' ');
                    html += `<option value="${{cidade}}">${{cidadeDisplay}}</option>`;
                }});
                
                html += `
        </select>
    </div>
`;
                
                // Adicionar tabelas
                dadosEscala.forEach((tabela, tabelaIndex) => {{
                    html += `<div class="escala-container">
        <table>
            <thead>
                <tr>`;
                    
                    tabela.cabecalhos.forEach(cabecalho => {{
                        html += `<th>${{cabecalho}}</th>`;
                    }});
                    
                    html += `</tr>
            </thead>
            <tbody>`;
                    
                    tabela.linhas.forEach(linha => {{
                        const cidadeAttr = linha.cidade ? `data-cidade="${{linha.cidade}}"` : '';
                        html += `<tr class="funcionario-row" ${{cidadeAttr}}>`;
                        linha.celulas.forEach(celula => {{
                            html += `<td class="${{celula.classes}}">${{celula.conteudo}}</td>`;
                        }});
                        html += `</tr>`;
                    }});
                    
                    html += `</tbody>
        </table>
    </div>`;
                }});
                
                html += `
    <script>
        // Função melhorada para filtrar por cidade
        function filtrarPorCidade() {{
            const filtro = document.getElementById('cityFilter').value;
            const linhas = document.querySelectorAll('.funcionario-row');
            let visibleCount = 0;
            
            console.log('Filtrando por cidade (visualização):', filtro);
            
            linhas.forEach(linha => {{
                const cidade = linha.getAttribute('data-cidade') || '';
                if (filtro === 'todas' || cidade === filtro) {{
                    linha.classList.remove('hidden');
                    visibleCount++;
                }} else {{
                    linha.classList.add('hidden');
                }}
            }});
            
            console.log('Funcionários visíveis após filtro:', visibleCount);
        }}
        
        // Aplicar filtro inicial
        document.addEventListener('DOMContentLoaded', function() {{
            console.log('HTML de visualização carregado');
        }});
    </script>
    
    <div class="footer">
        <p><strong>📊 Escala gerada pelo Sistema Gerador de Escalas v9.2.1</strong></p>
        <p>Esta é uma versão somente leitura com filtro por cidade. Para edições, utilize o arquivo original editável.</p>
        <p>Total de funcionários: ${{dadosEscala.reduce((total, tabela) => total + tabela.linhas.length, 0)}} | Cidades: ${{Array.from(cidadesUnicas).join(', ').replace(/_/g, ' ')}}</p>
    </div>
</body>
</html>`;
                
                return html;
            }}

            // Fechar dropdown ao clicar fora
            document.addEventListener('click', function(event) {{
                if (dropdown && dropdown.style.display === 'block' && 
                    !dropdown.contains(event.target) && 
                    !event.target.classList.contains('editable-cell')) {{
                    dropdown.style.display = 'none';
                    currentCellId = null;
                    console.log('Dropdown fechado por clique externo');
                }}
            }});

            // Fechar dropdown com tecla Escape
            document.addEventListener('keydown', function(event) {{
                if (event.key === 'Escape' && dropdown && dropdown.style.display === 'block') {{
                    dropdown.style.display = 'none';
                    currentCellId = null;
                    console.log('Dropdown fechado com Escape');
                }}
            }});
        </script>
        """
        html += js
        html += "</body></html>"
        return html

    # --- Funcoes Auxiliares --- 
    def mostrar_erro(self, mensagem):
        from tkinter import messagebox
        messagebox.showerror("Erro", mensagem)

    def mostrar_info(self, mensagem):
        from tkinter import messagebox
        messagebox.showinfo("Informacao", mensagem)

if __name__ == "__main__":
    app = EscalaApp()
    app.mainloop()

