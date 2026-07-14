import customtkinter as ctk
import pandas as pd
from datetime import datetime, timedelta, time
import calendar
import os
import random
import re
import webbrowser
import json
from tkcalendar import DateEntry
import traceback

# Configuracao da aparencia
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# --- Constantes --- 
CIDADES_PADRAO = [
    "MANAUS", "PIN - ITA", "MUR - CRI - HUM", "PORTO VELHO", 
    "JIP - ARI - COA", "RORAIMA", "AMAPA", "AM SAT E RIO BRANCO", "SUPERVISAO"
]
HORARIOS_PADRAO = [
    "06:00 - 12:00", "12:00 - 18:00", "18:00 - 00:00", "00:00 - 06:00", "Folguista"
]
HORARIOS_SUPERVISAO = [
    "06:00 - 14:00", "14:00 - 22:00", "16:00 - 02:00"
]
OPCOES_EDICAO = HORARIOS_PADRAO + ["Folga", "Folga (Banco de Horas)", "-", "Escrever..."]
ARQUIVO_DADOS = "pessoas_data.json"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_SCRIPT_HTML = os.path.join(BASE_DIR, "escala.js")

class EscalaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Gerador de Escalas v9.7 - Interface Compacta") 
        # Tamanho inicial mais proporcional ao conteudo, sem limitar o
        # aproveitamento de telas maiores.
        self.geometry("1180x760")
        self.minsize(980, 640)
        self.after_idle(self.centralizar_janela)
        
        self.pessoas = [] 
        self.carregar_pessoas()
        self.df_escala = None
        self.edit_widgets = {}
        
        self.configure(fg_color=("#f0f0f0", "#2b2b2b"))
        
        self.tabview = ctk.CTkTabview(self, fg_color=("#e0e0e0", "#333333"))
        self.tabview.pack(fill="both", expand=True, padx=8, pady=8)
        
        self.tab_config = self.tabview.add("Configuracoes")
        self.tab_result = self.tabview.add("Resultado")
        self.tabview.set("Configuracoes")
        
        self.setup_config_tab()
        self.setup_result_tab()
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def centralizar_janela(self):
        """Centraliza a janela usando o tamanho real calculado pelo Tk."""
        self.update_idletasks()
        largura = self.winfo_width()
        altura = self.winfo_height()
        pos_x = max(0, (self.winfo_screenwidth() - largura) // 2)
        pos_y = max(0, (self.winfo_screenheight() - altura) // 2)
        self.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
        
    def on_closing(self):
        self.salvar_pessoas()
        self.destroy()

    def setup_config_tab(self):
        config_frame = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        config_frame.pack(fill="both", expand=True)
        config_frame.grid_columnconfigure(0, weight=1)
        config_frame.grid_rowconfigure(3, weight=1, minsize=260)

        header_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=6, pady=(4, 8))
        ctk.CTkLabel(
            header_frame, 
            text="Configuracao da Escala", 
            font=("Segoe UI", 22, "bold"),
            text_color=("navy", "lightblue")
        ).pack(side="left", padx=8)
        
        period_frame = ctk.CTkFrame(config_frame, corner_radius=10)
        period_frame.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        period_frame.grid_columnconfigure(1, weight=1, uniform="datas")
        period_frame.grid_columnconfigure(3, weight=1, uniform="datas")
        
        ctk.CTkLabel(
            period_frame, 
            text="Periodo da Escala", 
            font=("Segoe UI", 16, "bold"),
            anchor="w"
        ).grid(row=0, column=0, columnspan=4, sticky="ew", padx=12, pady=(6, 2))
        
        ctk.CTkLabel(period_frame, text="Data Inicio:", font=("Segoe UI", 13)).grid(row=1, column=0, padx=(12, 2), pady=6, sticky="w")
        self.data_inicio_entry = DateEntry(
            period_frame, width=12, background="darkblue", foreground="white",
            borderwidth=2, date_pattern="dd/mm/yyyy", locale="pt_BR"
        )
        self.data_inicio_entry.grid(row=1, column=1, padx=(5, 18), pady=6, sticky="ew")
        
        ctk.CTkLabel(period_frame, text="Data Fim:", font=("Segoe UI", 13)).grid(row=1, column=2, padx=(12, 2), pady=6, sticky="w")
        self.data_fim_entry = DateEntry(
            period_frame, width=12, background="darkblue", foreground="white",
            borderwidth=2, date_pattern="dd/mm/yyyy", locale="pt_BR"
        )
        self.data_fim_entry.grid(row=1, column=3, padx=(5, 12), pady=6, sticky="ew")

        add_frame = ctk.CTkFrame(config_frame, corner_radius=10)
        add_frame.grid(row=2, column=0, sticky="ew", padx=6, pady=4)
        add_frame.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(
            add_frame, text="Adicionar Pessoa", font=("Segoe UI", 16, "bold"), anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=(12, 22), pady=8)
        
        ctk.CTkLabel(add_frame, text="Nome", font=("Segoe UI", 13)).grid(row=0, column=1, padx=(0, 6), pady=8, sticky="w")
        self.nome_entry = ctk.CTkEntry(add_frame, font=("Segoe UI", 13))
        self.nome_entry.grid(row=0, column=2, padx=5, pady=8, sticky="ew")
        self.nome_entry.bind("<Return>", lambda _event: self.adicionar_pessoa())
        
        add_button = ctk.CTkButton(
            add_frame, text="Adicionar", command=self.adicionar_pessoa, font=("Segoe UI", 13, "bold"),
            fg_color=("#3a7ebf", "#1f538d"), hover_color=("#2a6da9", "#14375e")
        )
        add_button.grid(row=0, column=3, padx=12, pady=8)

        list_section_frame = ctk.CTkFrame(config_frame, corner_radius=10)
        list_section_frame.grid(row=3, column=0, sticky="nsew", padx=6, pady=4)
        list_section_frame.grid_rowconfigure(1, weight=1)
        list_section_frame.grid_columnconfigure(0, weight=1)
        
        list_header = ctk.CTkFrame(list_section_frame, fg_color="transparent")
        list_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 3))
        list_header.grid_columnconfigure(2, weight=1)
        
        ctk.CTkLabel(
            list_header, text="Pessoas Cadastradas", font=("Segoe UI", 16, "bold"), anchor="w"
        ).grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(list_header, text="Filtrar por Cidade:", font=("Segoe UI", 13)).grid(row=0, column=1, padx=(30, 5), sticky="e")
        self.cidade_filtro_var = ctk.StringVar(value="Todas as Cidades")
        opcoes_filtro_cidade = ["Todas as Cidades"] + CIDADES_PADRAO
        self.cidade_filtro_menu = ctk.CTkOptionMenu(
            list_header, values=opcoes_filtro_cidade, variable=self.cidade_filtro_var, 
            command=self.atualizar_lista_pessoas, width=200, font=("Segoe UI", 13), dropdown_font=("Segoe UI", 13)
        ) 
        self.cidade_filtro_menu.grid(row=0, column=2, padx=5, sticky="e")
        
        list_table_frame = ctk.CTkFrame(list_section_frame)
        list_table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(3, 6))
        list_table_frame.grid_columnconfigure(0, weight=1)
        list_table_frame.grid_rowconfigure(1, weight=1)
        
        header_frame = ctk.CTkFrame(list_table_frame, fg_color=("#e0e0e0", "#333333"))
        header_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        header_frame.grid_columnconfigure(0, weight=2)
        header_frame.grid_columnconfigure(1, weight=2)
        header_frame.grid_columnconfigure(2, weight=2)
        header_frame.grid_columnconfigure(3, weight=1)
        header_frame.grid_columnconfigure(4, weight=1)
        header_frame.grid_columnconfigure(5, weight=1)
        
        ctk.CTkLabel(header_frame, text="Nome", anchor="w", font=("Segoe UI", 13, "bold")).grid(row=0, column=0, sticky="ew", padx=5)
        ctk.CTkLabel(header_frame, text="Cidade", anchor="w", font=("Segoe UI", 13, "bold")).grid(row=0, column=1, sticky="ew", padx=5) 
        ctk.CTkLabel(header_frame, text="Horario / Status", anchor="w", font=("Segoe UI", 13, "bold")).grid(row=0, column=2, sticky="ew", padx=5) 
        ctk.CTkLabel(header_frame, text="Folga Inicial", anchor="w", font=("Segoe UI", 13, "bold")).grid(row=0, column=3, sticky="ew", padx=5) 
        ctk.CTkLabel(header_frame, text="Banco Horas", anchor="w", font=("Segoe UI", 13, "bold")).grid(row=0, column=4, sticky="ew", padx=5)
        ctk.CTkLabel(header_frame, text="Acoes", anchor="w", font=("Segoe UI", 13, "bold")).grid(row=0, column=5, sticky="ew", padx=5)
        
        self.pessoas_list_frame = ctk.CTkScrollableFrame(list_table_frame)
        self.pessoas_list_frame.grid(row=1, column=0, sticky="nsew", padx=2, pady=2)
        self.pessoas_list_frame.grid_columnconfigure(0, weight=1)
        
        action_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        action_frame.grid(row=4, column=0, sticky="ew", padx=6, pady=(4, 5))
        action_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkButton(
            action_frame, text="Limpar Tudo", command=self.limpar_dados, font=("Segoe UI", 13),
            fg_color=("#d32f2f", "#b71c1c"), hover_color=("#c62828", "#7f0000"), width=130
        ).grid(row=0, column=0, padx=8, pady=2)
        
        self.gerar_button = ctk.CTkButton(
            action_frame, text="Gerar Escala", command=self.gerar_escala, font=("Segoe UI", 16, "bold"),
            fg_color=("#4caf50", "#2e7d32"), hover_color=("#43a047", "#1b5e20"), width=160, height=45
        )
        self.gerar_button.grid(row=0, column=2, padx=8, pady=2)
        
        self.atualizar_lista_pessoas()
    
    def setup_result_tab(self):
        result_frame = ctk.CTkFrame(self.tab_result)
        result_frame.pack(fill="both", expand=True, padx=4, pady=4)
        result_frame.grid_columnconfigure(0, weight=1)
        result_frame.grid_rowconfigure(2, weight=1, minsize=360)
        
        # Header compacto com filtro e alerta na mesma linha
        header_frame = ctk.CTkFrame(result_frame, fg_color="transparent", height=40)
        header_frame.grid(row=0, column=0, sticky="ew", padx=6, pady=(5, 6))
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_propagate(False)
        
        # Titulo
        ctk.CTkLabel(
            header_frame, text="Escala Gerada", font=("Segoe UI", 18, "bold"), text_color=("navy", "lightblue")
        ).grid(row=0, column=0, sticky="w", padx=10)
        
        # Filtro compacto no centro
        filter_compact_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        filter_compact_frame.grid(row=0, column=1, sticky="ew", padx=20)
        filter_compact_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            filter_compact_frame, text="Cidade:", font=("Segoe UI", 12)
        ).grid(row=0, column=0, padx=(0, 5), sticky="w")
        
        self.resultado_cidade_filtro_var = ctk.StringVar(value="Todas")
        opcoes_filtro_resultado = ["Todas"] + self.obter_cidades_com_pessoas()
        self.resultado_cidade_filtro_menu = ctk.CTkOptionMenu(
            filter_compact_frame, values=opcoes_filtro_resultado, variable=self.resultado_cidade_filtro_var, 
            command=self.filtrar_resultado_por_cidade, width=120, height=28, font=("Segoe UI", 11), dropdown_font=("Segoe UI", 11)
        ) 
        self.resultado_cidade_filtro_menu.grid(row=0, column=1, sticky="w")
        
        # Alerta de conflito - quadrado pequeno no canto direito
        self.conflict_indicator = ctk.CTkFrame(header_frame, width=35, height=35, corner_radius=8)
        self.conflict_indicator.grid(row=0, column=2, padx=10, sticky="e")
        self.conflict_indicator.grid_propagate(False)
        
        self.conflict_label = ctk.CTkLabel(
            self.conflict_indicator, text="âœ“", font=("Segoe UI", 16, "bold"), text_color="white"
        )
        self.conflict_label.pack(expand=True)
        
        # Inicialmente verde (sem conflitos)
        self.conflict_indicator.configure(fg_color=("#4caf50", "#2e7d32"))
        
        # Instrucao compacta
        instruction_frame = ctk.CTkFrame(result_frame, fg_color="transparent", height=25)
        instruction_frame.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 4))
        instruction_frame.grid_propagate(False)
        
        ctk.CTkLabel(
            instruction_frame, text="(Clique nas celulas para editar)", 
            font=("Segoe UI", 11, "italic"), text_color=("gray40", "gray60")
        ).pack(side="left", padx=10)
        
        # Tabela da escala
        table_frame = ctk.CTkFrame(result_frame, corner_radius=10)
        table_frame.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        self.result_table_frame = ctk.CTkScrollableFrame(table_frame)
        self.result_table_frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.result_table_frame.grid_columnconfigure(0, minsize=180, weight=1)
        for i in range(1, 8):
            self.result_table_frame.grid_columnconfigure(i, minsize=105, weight=1)
        
        # Botoes de acao compactos
        action_frame = ctk.CTkFrame(result_frame, fg_color="transparent", height=40)
        action_frame.grid(row=3, column=0, sticky="ew", padx=6, pady=(4, 5))
        action_frame.grid_columnconfigure(0, weight=1)
        action_frame.grid_propagate(False)
        
        buttons_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        buttons_frame.grid(row=0, column=0, sticky="e")
        
        ctk.CTkButton(
            buttons_frame, text="Detectar Conflitos", command=self.detectar_conflitos_manual, font=("Segoe UI", 11),
            fg_color=("#f44336", "#d32f2f"), hover_color=("#e53935", "#c62828"), width=130, height=32
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            buttons_frame, text="Excel", command=self.exportar_excel, font=("Segoe UI", 11),
            fg_color=("#3a7ebf", "#1f538d"), hover_color=("#2a6da9", "#14375e"), width=80, height=32
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            buttons_frame, text="HTML", command=self.exportar_html, font=("Segoe UI", 11),
            fg_color=("#ff9800", "#e65100"), hover_color=("#fb8c00", "#d84315"), width=80, height=32
        ).pack(side="left", padx=5)
    
    def salvar_pessoas(self):
        try:
            with open(ARQUIVO_DADOS, "w", encoding="utf-8") as f:
                json.dump(self.pessoas, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Erro ao salvar dados: {e}")

    def carregar_pessoas(self):
        if os.path.exists(ARQUIVO_DADOS):
            try:
                with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
                    self.pessoas = json.load(f)
                for pessoa in self.pessoas: 
                    if "banco_horas" not in pessoa:
                        pessoa["banco_horas"] = False
            except Exception as e:
                print(f"Erro ao carregar dados: {e}")
                self.pessoas = []
        else:
            self.pessoas = []

    def adicionar_pessoa(self):
        nome = self.nome_entry.get().strip()
        if not nome:
            self.mostrar_erro("O nome nao pode estar vazio.")
            return
        if any(p["nome"].lower() == nome.lower() for p in self.pessoas):
            self.mostrar_erro(f"O nome '{nome}' ja existe.")
            return
        
        self.pessoas.append({
            "nome": nome, 
            "cidade": CIDADES_PADRAO[0], 
            "horario": HORARIOS_PADRAO[0], 
            "folga_inicial": 1, 
            "banco_horas": False
        })
        self.nome_entry.delete(0, "end")
        self.atualizar_lista_pessoas()
        self.atualizar_filtro_resultado()
        self.salvar_pessoas()
        self.mostrar_info(f"Pessoa '{nome}' adicionada. Edite Cidade, Horario/Status e Folga na lista.")

    def atualizar_lista_pessoas(self, filtro_selecionado=None):
        for widget in self.pessoas_list_frame.winfo_children():
            widget.destroy()
        
        filtro_cidade = self.cidade_filtro_var.get()
        pessoas_filtradas = [p for p in self.pessoas if filtro_cidade == "Todas as Cidades" or p["cidade"] == filtro_cidade]
        cores_linhas = [("white", "#2d2d2d"), ("#f5f5f5", "#333333")]
        
        for i, pessoa in enumerate(pessoas_filtradas):
            try:
                original_index = self.pessoas.index(pessoa)
            except ValueError:
                continue
            
            cor_linha = cores_linhas[i % 2]
            pessoa_frame = ctk.CTkFrame(self.pessoas_list_frame, fg_color=cor_linha)
            pessoa_frame.grid(row=i, column=0, sticky="ew", padx=2, pady=1)
            pessoa_frame.grid_columnconfigure(0, weight=2)
            pessoa_frame.grid_columnconfigure(1, weight=2)
            pessoa_frame.grid_columnconfigure(2, weight=2)
            pessoa_frame.grid_columnconfigure(3, weight=1)
            pessoa_frame.grid_columnconfigure(4, weight=1)
            pessoa_frame.grid_columnconfigure(5, weight=1)
            
            ctk.CTkLabel(
                pessoa_frame, text=pessoa["nome"], anchor="w", font=("Segoe UI", 13)
            ).grid(row=0, column=0, sticky="ew", padx=5)
            
            cidade_var_inline = ctk.StringVar(value=pessoa["cidade"])
            cidade_dropdown_inline = ctk.CTkOptionMenu(
                pessoa_frame, values=CIDADES_PADRAO, variable=cidade_var_inline, 
                font=("Segoe UI", 13), dropdown_font=("Segoe UI", 13),
                command=lambda choice, idx=original_index: self.salvar_campo(idx, "cidade", choice)
            )
            cidade_dropdown_inline.grid(row=0, column=1, sticky="ew", padx=5)
            
            horarios_disponiveis = HORARIOS_PADRAO.copy()
            if pessoa["cidade"] == "SUPERVISAO":
                horarios_disponiveis = HORARIOS_SUPERVISAO + ["Folguista"]
            
            horario_var_inline = ctk.StringVar(value=pessoa["horario"])
            horario_dropdown_inline = ctk.CTkOptionMenu(
                pessoa_frame, values=horarios_disponiveis, variable=horario_var_inline, 
                font=("Segoe UI", 13), dropdown_font=("Segoe UI", 13),
                command=lambda choice, idx=original_index: self.salvar_campo_horario(idx, choice)
            )
            horario_dropdown_inline.grid(row=0, column=2, sticky="ew", padx=5)
            
            is_folguista = pessoa["horario"] == "Folguista"
            is_supervisao = pessoa["cidade"] == "SUPERVISAO"
            
            folga_entry = ctk.CTkEntry(
                pessoa_frame, width=50, font=("Segoe UI", 13),
                state="disabled" if (is_folguista or is_supervisao) else "normal"
            )
            if not is_folguista and not is_supervisao:
                folga_entry.insert(0, str(pessoa["folga_inicial"]))
                folga_entry.bind("<Return>", lambda event, idx=original_index, entry=folga_entry: self.salvar_folga_inicial(idx, entry))
            else:
                folga_entry.insert(0, "-")
            folga_entry.grid(row=0, column=3, sticky="ew", padx=5)
            
            banco_horas_var = ctk.BooleanVar(value=pessoa.get("banco_horas", False))
            banco_horas_checkbox = ctk.CTkCheckBox(
                pessoa_frame, text="", variable=banco_horas_var, width=20,
                state="disabled" if is_supervisao else "normal",
                command=lambda idx=original_index, var=banco_horas_var: self.salvar_campo(idx, "banco_horas", var.get())
            )
            banco_horas_checkbox.grid(row=0, column=4, sticky="ew", padx=5)
            
            ctk.CTkButton(
                pessoa_frame, text="Remover", width=80, font=("Segoe UI", 13),
                fg_color=("#d32f2f", "#b71c1c"), hover_color=("#c62828", "#7f0000"),
                command=lambda idx=original_index: self.remover_pessoa(idx)
            ).grid(row=0, column=5, sticky="ew", padx=5)

    def salvar_campo(self, index, campo, novo_valor):
        if 0 <= index < len(self.pessoas):
            valor_antigo = self.pessoas[index].get(campo)
            self.pessoas[index][campo] = novo_valor
            
            if campo == "cidade" and novo_valor == "SUPERVISAO":
                if self.pessoas[index]["horario"] not in HORARIOS_SUPERVISAO + ["Folguista"]:
                    self.pessoas[index]["horario"] = HORARIOS_SUPERVISAO[0]
                    self.pessoas[index]["folga_inicial"] = 0
                    self.pessoas[index]["banco_horas"] = False
                    self.mostrar_info(f"Funcionario {self.pessoas[index]['nome']} movido para supervisao. Horario e configuracoes ajustadas.")
            
            elif campo == "cidade" and valor_antigo == "SUPERVISAO" and novo_valor != "SUPERVISAO":
                if self.pessoas[index]["horario"] in HORARIOS_SUPERVISAO:
                    self.pessoas[index]["horario"] = HORARIOS_PADRAO[0]
                    self.pessoas[index]["folga_inicial"] = 1
                    self.mostrar_info(f"Funcionario {self.pessoas[index]['nome']} saiu da supervisao. Configuracoes restauradas.")
            
            self.salvar_pessoas()
            if campo == "cidade" and valor_antigo != novo_valor:
                self.atualizar_filtro_resultado()
            if campo == "cidade" and self.cidade_filtro_var.get() != "Todas as Cidades" and valor_antigo != novo_valor:
                self.atualizar_lista_pessoas()
        else:
            self.mostrar_erro("Erro ao salvar: Indice da pessoa invalido.")
    
    def salvar_campo_horario(self, index, novo_horario):
        if 0 <= index < len(self.pessoas):
            horario_antigo = self.pessoas[index]["horario"]
            self.pessoas[index]["horario"] = novo_horario
            
            if self.pessoas[index]["cidade"] == "SUPERVISAO":
                self.pessoas[index]["folga_inicial"] = 0
                self.pessoas[index]["banco_horas"] = False
            elif novo_horario == "Folguista" and horario_antigo != "Folguista":
                self.pessoas[index]["folga_inicial"] = 0
                self.mostrar_info(f"Folga inicial removida para o folguista {self.pessoas[index]['nome']}.")
            elif horario_antigo == "Folguista" and novo_horario != "Folguista":
                self.pessoas[index]["folga_inicial"] = 1
                self.mostrar_info(f"Folga inicial definida como 1 para {self.pessoas[index]['nome']}.")
            
            self.salvar_pessoas()
            self.atualizar_lista_pessoas()
        else:
            self.mostrar_erro("Erro ao salvar: Indice da pessoa invalido.")

    def salvar_folga_inicial(self, index, entry_widget):
        if self.pessoas[index]["horario"] == "Folguista":
            self.mostrar_erro("Folguistas nao possuem folga inicial definida.")
            entry_widget.delete(0, "end")
            entry_widget.insert(0, "-")
            return
        
        if self.pessoas[index]["cidade"] == "SUPERVISAO":
            self.mostrar_erro("Funcionarios da supervisao nao possuem folga inicial definida.")
            entry_widget.delete(0, "end")
            entry_widget.insert(0, "-")
            return
            
        novo_valor_str = entry_widget.get().strip()
        if not novo_valor_str.isdigit():
            self.mostrar_erro("A folga inicial deve ser um numero.")
            entry_widget.delete(0, "end")
            entry_widget.insert(0, str(self.pessoas[index]["folga_inicial"]))
            return
        
        novo_valor = int(novo_valor_str)
        
        try:
            data_inicio_date = self.data_inicio_entry.get_date()
            data_inicio = datetime.combine(data_inicio_date, time.min)
            
            if novo_valor < 1 or novo_valor > 31:
                self.mostrar_erro(f"A folga inicial deve ser um dia valido (1-31).")
                entry_widget.delete(0, "end")
                entry_widget.insert(0, str(self.pessoas[index]["folga_inicial"]))
                return
            
            try:
                primeiro_dia_mes_inicio = data_inicio.replace(day=1)
                ultimo_dia_mes_inicio = primeiro_dia_mes_inicio.replace(day=calendar.monthrange(data_inicio.year, data_inicio.month)[1])
                if novo_valor > ultimo_dia_mes_inicio.day:
                    self.mostrar_erro(f"O dia {novo_valor} nao existe no mes de inicio ({data_inicio.strftime('%m/%Y')}).")
                    entry_widget.delete(0, "end")
                    entry_widget.insert(0, str(self.pessoas[index]["folga_inicial"]))
                    return
            except Exception as date_val_err:
                print(f"Aviso: Nao foi possivel validar o dia da folga com o mes de inicio: {date_val_err}")

        except Exception as get_date_err: 
            print(f"Aviso: Nao foi possivel obter as datas para validar a folga inicial: {get_date_err}")

        self.salvar_campo(index, "folga_inicial", novo_valor)
        self.salvar_pessoas()
        self.mostrar_info(f"Folga inicial de {self.pessoas[index]['nome']} atualizada para {novo_valor}.")
        self.pessoas_list_frame.focus()

    def remover_pessoa(self, index):
        if 0 <= index < len(self.pessoas):
            nome_removido = self.pessoas[index]["nome"]
            del self.pessoas[index]
            self.atualizar_lista_pessoas()
            self.atualizar_filtro_resultado()
            self.salvar_pessoas()
            self.mostrar_info(f"Pessoa '{nome_removido}' removida.")
    
    def limpar_dados(self):
        self.pessoas = []
        self.cidade_filtro_var.set("Todas as Cidades") 
        self.atualizar_lista_pessoas()
        self.atualizar_filtro_resultado()
        self.nome_entry.delete(0, "end")
        self.salvar_pessoas()
        self.mostrar_info("Todos os dados de pessoas foram limpos.")
    
    def gerar_escala(self):
        try:
            data_inicio_date = self.data_inicio_entry.get_date()
            data_fim_date = self.data_fim_entry.get_date()
            data_inicio = datetime.combine(data_inicio_date, time.min)
            data_fim = datetime.combine(data_fim_date, time.min)
        except Exception as e:
            self.mostrar_erro(f"Erro ao obter datas: {e}. Verifique o formato (DD/MM/AAAA) e se as datas sao validas.")
            return

        if data_inicio > data_fim:
            self.mostrar_erro("A data de inicio nao pode ser posterior a data de fim.")
            return

        filtro_cidade = self.cidade_filtro_var.get()
        pessoas_para_escala = [p for p in self.pessoas if filtro_cidade == "Todas as Cidades" or p["cidade"] == filtro_cidade]

        if not pessoas_para_escala:
            self.mostrar_erro(f"Nenhuma pessoa encontrada para a cidade '{filtro_cidade}'.")
            return

        folguistas = [p for p in pessoas_para_escala if p["horario"] == "Folguista"]
        regulares = [p for p in pessoas_para_escala if p["horario"] != "Folguista"]

        if not regulares:
            self.mostrar_erro("Nenhum funcionario regular encontrado para gerar a escala base.")
            return
        if not folguistas:
            self.mostrar_info("Nenhum folguista encontrado. A escala sera gerada apenas para os regulares.")

        try:
            self.gerar_escala_final_daterange(data_inicio, data_fim, regulares, folguistas)
            self.tabview.set("Resultado")
        except Exception as e:
            self.mostrar_erro(f"Ocorreu um erro inesperado ao gerar a escala: {e}")
            traceback.print_exc()
    
    def gerar_escala_final_daterange(self, data_inicio, data_fim, regulares, folguistas):
        try:
            for widget in self.result_table_frame.winfo_children():
                widget.destroy()

            datas_escala = []
            data_atual = data_inicio
            while data_atual <= data_fim:
                datas_escala.append(data_atual)
                data_atual += timedelta(days=1)

            if not datas_escala:
                self.mostrar_erro("Intervalo de datas invalido resultou em 0 dias.")
                return

            cidades = {}
            for pessoa in regulares + folguistas:
                cidade = pessoa.get("cidade", "SEM_CIDADE")
                if cidade not in cidades:
                    cidades[cidade] = {"regulares": [], "folguistas": []}
                
                if pessoa in regulares:
                    cidades[cidade]["regulares"].append(pessoa)
                else:
                    cidades[cidade]["folguistas"].append(pessoa)

            todos_nomes = [p["nome"] for p in regulares + folguistas]
            df_escala_final = pd.DataFrame(index=todos_nomes, columns=datas_escala).fillna("")

            for cidade, funcionarios_cidade in cidades.items():
                regulares_cidade = funcionarios_cidade["regulares"]
                folguistas_cidade = funcionarios_cidade["folguistas"]
                
                if not regulares_cidade:
                    continue
                
                df_cidade = self.processar_escala_cidade(
                    data_inicio, data_fim, datas_escala, 
                    regulares_cidade, folguistas_cidade, cidade
                )
                
                for nome in df_cidade.index:
                    if nome in df_escala_final.index:
                        for data in datas_escala:
                            df_escala_final.at[nome, data] = df_cidade.at[nome, data]

            self.df_escala = df_escala_final
            self.exibir_escala_editavel_colunas_fixas(df_escala_final)

        except Exception as e:
            self.mostrar_erro(f"Erro interno ao calcular a escala: {e}")
            traceback.print_exc()

    def processar_escala_cidade(self, data_inicio, data_fim, datas_escala, regulares, folguistas, cidade):
        nomes_regulares = [p["nome"] for p in regulares]
        nomes_folguistas = [p["nome"] for p in folguistas]
        todos_nomes = nomes_regulares + nomes_folguistas
        
        df_escala = pd.DataFrame(index=todos_nomes, columns=datas_escala).fillna("")
        
        horario_por_nome = {p["nome"]: p["horario"] for p in regulares}
        folgas_por_data = {data: [] for data in datas_escala}
        folgas_regulares_map = {}
        
        if cidade == "SUPERVISAO":
            for pessoa in regulares:
                nome = pessoa["nome"]
                horario = pessoa["horario"]
                for data in datas_escala:
                    df_escala.at[nome, data] = horario
            
            for pessoa_f in folguistas:
                nome_f = pessoa_f["nome"]
                for data in datas_escala:
                    df_escala.at[nome_f, data] = "-"
            
            return df_escala
        
        turno_18_00_nomes = [p["nome"] for p in regulares if p["horario"] == "18:00 - 00:00"]
        turno_00_06_nomes = [p["nome"] for p in regulares if p["horario"] == "00:00 - 06:00"]
        folgas_apos_cobertura = {nome: set() for nome in turno_18_00_nomes}

        for pessoa in regulares:
            nome = pessoa["nome"]
            folga_inicial_dia = pessoa["folga_inicial"]
            horario = pessoa["horario"]
            tem_banco_horas = pessoa.get("banco_horas", False)
            
            if horario == "Folguista":
                continue
            
            try:
                data_teste = datetime(data_inicio.year, data_inicio.month, folga_inicial_dia)
            except ValueError:
                ultimo_dia_mes_inicio = calendar.monthrange(data_inicio.year, data_inicio.month)[1]
                data_teste = datetime(data_inicio.year, data_inicio.month, ultimo_dia_mes_inicio)
            
            while data_teste < data_inicio:
                if data_teste.month == 12:
                    try:
                        data_teste = datetime(data_teste.year + 1, 1, folga_inicial_dia)
                    except ValueError:
                        data_teste = datetime(data_teste.year + 1, 1, calendar.monthrange(data_teste.year + 1, 1)[1])
                else:
                    try:
                        data_teste = datetime(data_teste.year, data_teste.month + 1, folga_inicial_dia)
                    except ValueError:
                        data_teste = datetime(data_teste.year, data_teste.month + 1, calendar.monthrange(data_teste.year, data_teste.month + 1)[1])
            
            primeira_folga_data = data_teste
            folgas_finais = set()
            data_folga = primeira_folga_data
            
            while data_folga <= data_fim:
                if data_folga >= data_inicio:
                    folgas_finais.add(data_folga)
                    
                    if data_folga.weekday() == 5:  # Sabado
                        data_domingo = data_folga + timedelta(days=1)
                        if data_inicio <= data_domingo <= data_fim:
                            folgas_finais.add(data_domingo)
                        
                        if tem_banco_horas:
                            data_segunda = data_folga + timedelta(days=2)
                            if data_inicio <= data_segunda <= data_fim:
                                folgas_finais.add(data_segunda)
                        
                        data_folga += timedelta(days=8)
                        
                    elif data_folga.weekday() == 6:  # Domingo
                        if tem_banco_horas:
                            data_segunda = data_folga + timedelta(days=1)
                            if data_inicio <= data_segunda <= data_fim:
                                folgas_finais.add(data_segunda)
                        
                        data_folga += timedelta(days=8)
                        
                    else:
                        data_folga += timedelta(days=8)
                else:
                    data_folga += timedelta(days=8)
            
            folgas_regulares_map[nome] = folgas_finais
            for data in datas_escala:
                if data in folgas_finais:
                    df_escala.at[nome, data] = "Folga"
                    folgas_por_data[data].append(nome)
                else:
                    df_escala.at[nome, data] = horario

        for data in datas_escala:
            folgas_turno_00_06 = []
            
            for nome_00_06 in turno_00_06_nomes:
                if nome_00_06 in df_escala.index:
                    status_funcionario = df_escala.at[nome_00_06, data]
                    if status_funcionario == "Folga":
                        folgas_turno_00_06.append(nome_00_06)
            
            if folgas_turno_00_06:
                cobertores_disponiveis = []
                for nome_18_00 in turno_18_00_nomes:
                    if nome_18_00 not in folgas_por_data.get(data, []):
                        if data not in folgas_apos_cobertura.get(nome_18_00, set()):
                            cobertores_disponiveis.append(nome_18_00)
                
                if cobertores_disponiveis:
                    for i, funcionario_de_folga in enumerate(folgas_turno_00_06):
                        if i < len(cobertores_disponiveis):
                            cobertor = cobertores_disponiveis[i]
                            df_escala.at[cobertor, data] = "00:00 - 06:00 (Cob)"
                            
                            cobertor_horario = next((p["horario"] for p in regulares if p["nome"] == cobertor), "")
                            if cobertor_horario == "18:00 - 00:00":
                                folguista_disponivel = None
                                for folguista in folguistas:
                                    folguista_nome = folguista["nome"]
                                    folguista_status = df_escala.at[folguista_nome, data]
                                    if folguista_status == "-" or pd.isna(folguista_status) or folguista_status == "":
                                        folguista_disponivel = folguista_nome
                                        break
                                
                                if folguista_disponivel:
                                    df_escala.at[folguista_disponivel, data] = "18:00 - 00:00"
                            
                            cobertor_pessoa = next((p for p in regulares if p["nome"] == cobertor), None)
                            tem_banco_horas_cobertor = cobertor_pessoa.get("banco_horas", False) if cobertor_pessoa else False
                            
                            data_seguinte = data + timedelta(days=1)
                            if data_seguinte <= data_fim and data_seguinte not in folgas_regulares_map.get(cobertor, set()):
                                df_escala.at[cobertor, data_seguinte] = "Folga"
                                folgas_apos_cobertura.setdefault(cobertor, set()).add(data_seguinte)
                                folgas_por_data.setdefault(data_seguinte, []).append(cobertor)
                                
                                if tem_banco_horas_cobertor:
                                    data_segundo_dia = data_seguinte + timedelta(days=1)
                                    if data_segundo_dia <= data_fim and data_segundo_dia not in folgas_regulares_map.get(cobertor, set()):
                                        df_escala.at[cobertor, data_segundo_dia] = "Folga"
                                        folgas_apos_cobertura.setdefault(cobertor, set()).add(data_segundo_dia)
                                        folgas_por_data.setdefault(data_segundo_dia, []).append(cobertor)

        if folguistas:
            for data in datas_escala:
                funcionarios_de_folga = []
                for nome in df_escala.index:
                    nomes_folguistas = [f["nome"] for f in folguistas]
                    if nome not in nomes_folguistas:
                        valor_celula = df_escala.at[nome, data]
                        if valor_celula == "Folga":
                            horario_funcionario = horario_por_nome.get(nome, "")
                            if horario_funcionario != "00:00 - 06:00":
                                funcionarios_de_folga.append((nome, horario_funcionario))
                
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
                            df_escala.at[nome_folguista, data] = horario_funcionario

        for pessoa_f in folguistas:
            nome_f = pessoa_f["nome"]
            for data in datas_escala:
                if df_escala.at[nome_f, data] == "":
                    df_escala.at[nome_f, data] = "-"

        return df_escala

    def exibir_escala_editavel_colunas_fixas(self, df):
        try:
            for widget in self.result_table_frame.winfo_children():
                widget.destroy()
            
            self.edit_widgets = {}
            
            datas_escala = list(df.columns)
            if not datas_escala:
                return
            
            data_inicio = datas_escala[0]
            data_fim = datas_escala[-1]
            dias_semana_abrev = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex", 5: "Sab", 6: "Dom"}
            nomes_folguistas = [p["nome"] for p in self.pessoas if p["horario"] == "Folguista"]

            blocos_seg_dom = []
            dias_para_primeira_segunda = (0 - data_inicio.weekday() + 7) % 7
            primeira_segunda_data = data_inicio + timedelta(days=dias_para_primeira_segunda)
            
            dias_iniciais = [d for d in datas_escala if d < primeira_segunda_data]
            if dias_iniciais:
                blocos_seg_dom.append(dias_iniciais)
            
            data_bloco_atual = primeira_segunda_data
            while data_bloco_atual <= data_fim:
                fim_bloco_data = data_bloco_atual + timedelta(days=6)
                bloco_atual = [d for d in datas_escala if data_bloco_atual <= d <= fim_bloco_data]
                if bloco_atual:
                    blocos_seg_dom.append(bloco_atual)
                data_bloco_atual += timedelta(days=7)
            
            row_idx_global = 0
            for bloco_num, bloco_datas in enumerate(blocos_seg_dom):
                bloco_header_frame = ctk.CTkFrame(self.result_table_frame, fg_color=("#1a237e", "#0d47a1"), height=65)
                bloco_header_frame.grid(row=row_idx_global, column=0, columnspan=8, sticky="ew", padx=8, pady=(15, 5))
                bloco_header_frame.grid_columnconfigure(0, weight=1)
                bloco_header_frame.grid_propagate(False)
                
                primeira_data_bloco = bloco_datas[0]
                ultima_data_bloco = bloco_datas[-1]
                titulo_bloco = f"Semana {bloco_num + 1}: {primeira_data_bloco.strftime('%d/%m')} - {ultima_data_bloco.strftime('%d/%m/%Y')}"
                ctk.CTkLabel(
                    bloco_header_frame, text=titulo_bloco, font=("Segoe UI", 20, "bold"), text_color="white"
                ).grid(row=0, column=0, pady=18)
                row_idx_global += 1
                
                table_header_frame = ctk.CTkFrame(self.result_table_frame, fg_color=("#263238", "#37474f"), height=55)
                table_header_frame.grid(row=row_idx_global, column=0, columnspan=8, sticky="ew", padx=8, pady=3)
                table_header_frame.grid_propagate(False)
                
                table_header_frame.grid_columnconfigure(0, minsize=180, weight=1)
                for col_idx in range(len(bloco_datas)):
                    table_header_frame.grid_columnconfigure(col_idx + 1, minsize=105, weight=1)
                
                ctk.CTkLabel(
                    table_header_frame, text="Funcionario", font=("Segoe UI", 15, "bold"), 
                    text_color="white", anchor="w"
                ).grid(row=0, column=0, sticky="ew", padx=15, pady=10)
                
                for col_idx, data in enumerate(bloco_datas):
                    dia_str = f"{data.day}/{data.month}\n({dias_semana_abrev[data.weekday()]})"
                    ctk.CTkLabel(
                        table_header_frame, text=dia_str, font=("Segoe UI", 13, "bold"), 
                        text_color="white", anchor="center"
                    ).grid(row=0, column=col_idx + 1, sticky="ew", padx=3, pady=10)
                row_idx_global += 1
                
                for pessoa_idx, nome in enumerate(df.index):
                    is_folguista = nome in nomes_folguistas
                    pessoa_obj = next((p for p in self.pessoas if p["nome"] == nome), None)
                    is_supervisao = pessoa_obj and pessoa_obj.get("cidade") == "SUPERVISAO"
                    
                    cor_linha = [("white", "#2d2d2d"), ("#f8f9fa", "#3a3a3a")][pessoa_idx % 2]
                    
                    linha_frame = ctk.CTkFrame(self.result_table_frame, fg_color=cor_linha, height=50)
                    linha_frame.grid(row=row_idx_global + pessoa_idx, column=0, columnspan=8, sticky="ew", padx=8, pady=2)
                    linha_frame.grid_propagate(False)
                    
                    linha_frame.grid_columnconfigure(0, minsize=180, weight=1)
                    for col_idx in range(len(bloco_datas)):
                        linha_frame.grid_columnconfigure(col_idx + 1, minsize=105, weight=1)
                    
                    nome_display = f"{nome}"
                    if is_folguista:
                        nome_display += " (Folguista)"
                    elif is_supervisao:
                        nome_display += " (Supervisao)"
                    
                    nome_label = ctk.CTkLabel(
                        linha_frame, text=nome_display, 
                        font=("Segoe UI", 14, "bold" if not is_folguista else "normal"), 
                        anchor="w", 
                        fg_color=("#9c27b0", "#7b1fa2") if is_supervisao else ("#263238", "#37474f"), 
                        text_color="white"
                    )
                    nome_label.grid(row=0, column=0, sticky="nsew", padx=3, pady=2)
                    
                    for col_idx, data in enumerate(bloco_datas):
                        valor_celula = df.at[nome, data]
                        
                        opcoes_disponiveis = HORARIOS_SUPERVISAO + ["-", "Escrever..."] if is_supervisao else OPCOES_EDICAO.copy()
                        
                        valor_var = ctk.StringVar(value=str(valor_celula))
                        dropdown = ctk.CTkOptionMenu(
                            linha_frame, values=opcoes_disponiveis, variable=valor_var,
                            width=125, height=40, font=("Segoe UI", 13), dropdown_font=("Segoe UI", 13),
                            command=lambda choice, n=nome, d=data: self.atualizar_celula_escala(n, d, choice)
                        )
                        
                        self.aplicar_cor_celula(dropdown, valor_celula, is_supervisao, is_folguista)
                        
                        dropdown.grid(row=0, column=col_idx + 1, sticky="ew", padx=2, pady=3)
                        
                        self.edit_widgets[(nome, data)] = dropdown

                row_idx_global += len(df.index) + 3

        except Exception as e:
            self.mostrar_erro(f"Erro ao exibir a escala editavel: {e}")
            traceback.print_exc()

    def aplicar_cor_celula(self, dropdown, valor_celula, is_supervisao, is_folguista):
        cor_padrao = ("#3a7ebf", "#1f538d")
        cor_folga = ("#1f538d", "#1976d2")
        cor_cobertura_00_06 = ("#4fc3f7", "#03a9f4")
        cor_supervisao_fg = ("#9c27b0", "#7b1fa2")
        
        cor_final = cor_padrao
        
        if valor_celula == "Folga" or "Folga" in str(valor_celula):
            cor_final = cor_folga
        elif "(Cob)" in str(valor_celula):
            cor_final = cor_cobertura_00_06
        elif is_supervisao and valor_celula != "-":
            cor_final = cor_supervisao_fg
        
        dropdown.configure(fg_color=cor_final)

    def atualizar_celula_escala(self, nome, data, novo_valor):
        if novo_valor == "Escrever...":
            self.abrir_janela_edicao_livre(nome, data)
            return
        
        if self.df_escala is not None:
            self.df_escala.at[nome, data] = novo_valor
            
            if (nome, data) in self.edit_widgets:
                dropdown = self.edit_widgets[(nome, data)]
                pessoa_obj = next((p for p in self.pessoas if p["nome"] == nome), None)
                is_supervisao = pessoa_obj and pessoa_obj.get("cidade") == "SUPERVISAO"
                is_folguista = pessoa_obj and pessoa_obj.get("horario") == "Folguista"
                
                self.aplicar_cor_celula(dropdown, novo_valor, is_supervisao, is_folguista)

    def abrir_janela_edicao_livre(self, nome, data):
        dialog = ctk.CTkInputDialog(
            text=f"Digite o valor para {nome} em {data.strftime('%d/%m')}:", 
            title="Edicao Livre"
        )
        novo_valor = dialog.get_input()
        
        if novo_valor is not None and self.df_escala is not None:
            self.df_escala.at[nome, data] = novo_valor
            
            if (nome, data) in self.edit_widgets:
                dropdown = self.edit_widgets[(nome, data)]
                dropdown.set(novo_valor)
                
                pessoa_obj = next((p for p in self.pessoas if p["nome"] == nome), None)
                is_supervisao = pessoa_obj and pessoa_obj.get("cidade") == "SUPERVISAO"
                is_folguista = pessoa_obj and pessoa_obj.get("horario") == "Folguista"
                
                self.aplicar_cor_celula(dropdown, novo_valor, is_supervisao, is_folguista)

    def obter_cidades_com_pessoas(self):
        """Retorna somente cidades utilizadas, na ordem padrao da interface."""
        cidades_utilizadas = {
            pessoa.get("cidade") for pessoa in self.pessoas if pessoa.get("cidade")
        }
        cidades_ordenadas = [
            cidade for cidade in CIDADES_PADRAO if cidade in cidades_utilizadas
        ]
        # Mantem compativel qualquer cidade antiga que nao esteja mais na lista padrao.
        cidades_extras = sorted(cidades_utilizadas.difference(CIDADES_PADRAO))
        return cidades_ordenadas + cidades_extras

    def atualizar_filtro_resultado(self):
        if not hasattr(self, "resultado_cidade_filtro_menu"):
            return

        opcoes = ["Todas"] + self.obter_cidades_com_pessoas()
        cidade_atual = self.resultado_cidade_filtro_var.get()
        self.resultado_cidade_filtro_menu.configure(values=opcoes)

        if cidade_atual not in opcoes:
            self.resultado_cidade_filtro_var.set("Todas")
            if self.df_escala is not None and not self.df_escala.empty:
                self.filtrar_resultado_por_cidade()

    def filtrar_resultado_por_cidade(self, cidade_selecionada=None):
        if self.df_escala is None or self.df_escala.empty:
            return
        
        cidade_filtro = self.resultado_cidade_filtro_var.get()
        
        if cidade_filtro == "Todas":
            funcionarios_filtrados = [p["nome"] for p in self.pessoas]
        else:
            funcionarios_filtrados = [
                p["nome"] for p in self.pessoas if p["cidade"] == cidade_filtro
            ]
        
        df_filtrado = self.df_escala.loc[self.df_escala.index.intersection(funcionarios_filtrados)]
        self.exibir_escala_editavel_colunas_fixas(df_filtrado)

    def detectar_conflitos_manual(self):
        if self.df_escala is None or self.df_escala.empty:
            self.mostrar_info("Nenhuma escala gerada para detectar conflitos.")
            return
        
        try:
            conflitos = []
            datas_escala = list(self.df_escala.columns)
            
            funcionarios_por_cidade = {}
            for nome in self.df_escala.index:
                pessoa = next((p for p in self.pessoas if p["nome"] == nome), None)
                if pessoa:
                    cidade = pessoa.get("cidade", "SEM_CIDADE")
                    if cidade not in funcionarios_por_cidade:
                        funcionarios_por_cidade[cidade] = []
                    funcionarios_por_cidade[cidade].append(nome)
            
            # Reset cores
            for (nome, data), widget in self.edit_widgets.items():
                pessoa_obj = next((p for p in self.pessoas if p["nome"] == nome), None)
                is_supervisao = pessoa_obj and pessoa_obj.get("cidade") == "SUPERVISAO"
                is_folguista = pessoa_obj and pessoa_obj.get("horario") == "Folguista"
                self.aplicar_cor_celula(widget, self.df_escala.at[nome, data], is_supervisao, is_folguista)
            
            for cidade, funcionarios in funcionarios_por_cidade.items():
                if cidade == "SUPERVISAO":
                    continue
                
                for data in datas_escala:
                    horarios_ocupados = {}
                    pessoas_de_folga = []
                    horarios_presentes = set()
                    
                    for nome in funcionarios:
                        valor_celula = self.df_escala.at[nome, data]
                        
                        if valor_celula == "Folga" or "Folga" in str(valor_celula):
                            pessoas_de_folga.append(nome)
                        elif valor_celula not in ["-", ""] and ":" in str(valor_celula):
                            horario_base = str(valor_celula).replace("(Cob)", "").strip()
                            
                            if horario_base not in horarios_ocupados:
                                horarios_ocupados[horario_base] = []
                            horarios_ocupados[horario_base].append(nome)
                            horarios_presentes.add(horario_base)
                    
                    # Conflitos de horÃ¡rio
                    for horario, pessoas in horarios_ocupados.items():
                        if len(pessoas) > 1:
                            conflitos.append({
                                "tipo": "horario",
                                "data": data,
                                "horario": horario,
                                "cidade": cidade,
                                "funcionarios": pessoas
                            })
                            
                            for nome_conflito in pessoas:
                                if (nome_conflito, data) in self.edit_widgets:
                                    widget = self.edit_widgets[(nome_conflito, data)]
                                    widget.configure(fg_color=("#f44336", "#d32f2f"))
                    
                    # Conflitos de folga
                    if len(pessoas_de_folga) > 1:
                        conflitos.append({
                            "tipo": "folga",
                            "data": data,
                            "horario": "Folga",
                            "cidade": cidade,
                            "funcionarios": pessoas_de_folga
                        })
                        
                        for nome_folga in pessoas_de_folga:
                            if (nome_folga, data) in self.edit_widgets:
                                widget = self.edit_widgets[(nome_folga, data)]
                                widget.configure(fg_color=("#f44336", "#d32f2f"))
                    
                    # HorÃ¡rios faltantes
                    horarios_necessarios = set(h for h in HORARIOS_PADRAO if h != "Folguista")
                    horarios_faltantes = horarios_necessarios - horarios_presentes
                    if horarios_faltantes:
                        conflitos.append({
                            "tipo": "faltante",
                            "data": data,
                            "horario": ", ".join(horarios_faltantes),
                            "cidade": cidade,
                            "funcionarios": []
                        })
            
            self.atualizar_indicador_conflitos(conflitos)
            
        except Exception as e:
            print(f"Erro na deteccao de conflitos: {e}")
            self.mostrar_erro(f"Erro ao detectar conflitos: {e}")

    def atualizar_indicador_conflitos(self, conflitos):
        """Atualiza o pequeno quadrado indicador de conflitos"""
        try:
            if conflitos:
                # Vermelho com nÃºmero de conflitos
                total_conflitos = len(conflitos)
                self.conflict_indicator.configure(fg_color=("#f44336", "#d32f2f"))
                self.conflict_label.configure(text=str(total_conflitos), text_color="white")
                
                # Tooltip com detalhes (opcional)
                conflitos_horario = len([c for c in conflitos if c["tipo"] == "horario"])
                conflitos_folga = len([c for c in conflitos if c["tipo"] == "folga"])
                conflitos_faltante = len([c for c in conflitos if c["tipo"] == "faltante"])
                
                tooltip_text = f"Conflitos: {conflitos_horario} horÃ¡rio, {conflitos_folga} folga, {conflitos_faltante} faltante"
                
                self.mostrar_info(f"Detectados {total_conflitos} conflito(s): {conflitos_horario} de horario, {conflitos_folga} de folga e {conflitos_faltante} de horarios faltantes.")
                
            else:
                # Verde - sem conflitos
                self.conflict_indicator.configure(fg_color=("#4caf50", "#2e7d32"))
                self.conflict_label.configure(text="âœ“", text_color="white")
                self.mostrar_info("Nenhum conflito detectado na escala atual.")
                
        except Exception as e:
            print(f"Erro ao atualizar indicador de conflitos: {e}")

    def exportar_excel(self):
        if self.df_escala is None or self.df_escala.empty:
            self.mostrar_erro("Nenhuma escala gerada para exportar.")
            return
        try:
            filtro_cidade = self.cidade_filtro_var.get().replace(" ", "_").replace("-", "_")
            data_inicio_str = self.df_escala.columns[0].strftime("%d_%m_%Y")
            data_fim_str = self.df_escala.columns[-1].strftime("%d_%m_%Y")
            filename = f"escala_{filtro_cidade}_{data_inicio_str}_a_{data_fim_str}.xlsx"
            
            writer = pd.ExcelWriter(filename, engine="openpyxl")
            
            df_export = self.df_escala.copy()
            df_export.columns = [data.strftime("%d/%m/%Y") for data in df_export.columns]
            
            df_export.to_excel(writer, sheet_name="Escala", index=True)
            
            worksheet = writer.sheets["Escala"]
            for i, col in enumerate(df_export.columns):
                max_len = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(66 + i)].width = max_len
            
            max_nome_len = max(len(str(nome)) for nome in df_export.index) + 5
            worksheet.column_dimensions["A"].width = max_nome_len
            
            writer.close()
            
            self.mostrar_info(f"Escala exportada para {filename}")
            if os.path.exists(filename):
                webbrowser.open(f"file://{os.path.realpath(filename)}")
        except Exception as e:
            self.mostrar_erro(f"Erro ao exportar para Excel: {e}")
            traceback.print_exc()

    def exportar_html(self):
        if self.df_escala is None or self.df_escala.empty:
            self.mostrar_erro("Nenhuma escala gerada para exportar.")
            return
        try:
            filtro_cidade = self.cidade_filtro_var.get().replace(" ", "_").replace("-", "_")
            data_inicio_str = self.df_escala.columns[0].strftime("%d_%m_%Y")
            data_fim_str = self.df_escala.columns[-1].strftime("%d_%m_%Y")
            filename = f"visualizacao_{filtro_cidade}_{data_inicio_str}_a_{data_fim_str}.html"
            
            html_content = self.gerar_html_somente_visualizacao(self.df_escala)
            
            with open(filename, "w", encoding="utf-8") as f:
                f.write(html_content)
            self.mostrar_info(f"Visualizacao HTML gerada: {filename}")
            webbrowser.open(f"file://{os.path.realpath(filename)}")
        except Exception as e:
            self.mostrar_erro(f"Erro ao gerar visualizacao HTML: {e}")
            traceback.print_exc()

    def gerar_html_somente_visualizacao(self, df):
        datas_escala = list(df.columns)
        data_inicio = datas_escala[0]
        data_fim = datas_escala[-1]
        dias_semana_abrev = {0: "Seg", 1: "Ter", 2: "Qua", 3: "Qui", 4: "Sex", 5: "Sab", 6: "Dom"}
        nomes_folguistas = [p["nome"] for p in self.pessoas if p["horario"] == "Folguista"]

        cidades_na_escala = set()
        for nome in df.index:
            pessoa = next((p for p in self.pessoas if p["nome"] == nome), None)
            if pessoa:
                cidades_na_escala.add(pessoa["cidade"])
        cidades_na_escala = sorted(list(cidades_na_escala))

        nome_para_cidade = {}
        for nome in df.index:
            pessoa = next((p for p in self.pessoas if p["nome"] == nome), None)
            if pessoa:
                nome_para_cidade[nome] = pessoa["cidade"]

        blocos_seg_dom = []
        dias_para_primeira_segunda = (0 - data_inicio.weekday() + 7) % 7
        primeira_segunda_data = data_inicio + timedelta(days=dias_para_primeira_segunda)
        dias_iniciais = [d for d in datas_escala if d < primeira_segunda_data]
        if dias_iniciais:
            blocos_seg_dom.append(dias_iniciais)
        data_bloco_atual = primeira_segunda_data
        while data_bloco_atual <= data_fim:
            fim_bloco_data = data_bloco_atual + timedelta(days=6)
            bloco_atual = [d for d in datas_escala if data_bloco_atual <= d <= fim_bloco_data]
            if bloco_atual:
                blocos_seg_dom.append(bloco_atual)
            data_bloco_atual += timedelta(days=7)

        css = """
        <style>
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f0f0f0; }
            h1 { color: #0d47a1; text-align: center; }
            .filter-container { 
                text-align: center; 
                margin: 20px 0; 
                padding: 15px; 
                background-color: #ffffff; 
                border-radius: 8px; 
                box-shadow: 0 2px 5px rgba(0,0,0,0.1); 
            }
            .filter-button { 
                display: inline-block; 
                margin: 5px; 
                padding: 8px 16px; 
                background-color: #e0e0e0; 
                color: #333; 
                border: none; 
                border-radius: 5px; 
                cursor: pointer; 
                font-size: 12px; 
                transition: all 0.3s ease; 
            }
            .filter-button:hover { 
                background-color: #d0d0d0; 
            }
            .filter-button.active { 
                background-color: #0d47a1; 
                color: white; 
            }
            .escala-container { margin-bottom: 25px; background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); overflow: hidden; }
            table { width: 100%; border-collapse: collapse; }
            th, td { border: 1px solid #ccc; padding: 8px 10px; text-align: center; font-size: 11px; min-width: 80px; height: 35px; }
            th { background-color: #0d47a1; color: white; font-weight: bold; font-size: 10px; white-space: pre-wrap; }
            .nome-col { background-color: #37474f; color: white; text-align: left; font-weight: bold; min-width: 150px; }
            .folguista-nome { font-weight: normal; }
            .supervisao-nome { background-color: #9c27b0; color: white; }
            .folga { background-color: #1f538d; color: white; }
            .folga-banco-horas { background-color: #1f538d; color: white; }
            .cobertura-00-06 { background-color: #4fc3f7; color: black; }
            .folguista-trabalhando { background-color: #f5f5f5; color: black; }
            .supervisao-trabalhando { background-color: #9c27b0; color: white; }
            .funcionario-row { transition: opacity 0.3s ease; }
            .funcionario-row.hidden { display: none; }
            .print-button { display: block; margin: 20px auto; padding: 10px 20px; font-size: 14px; cursor: pointer; background-color: #4caf50; color: white; border: none; border-radius: 5px; }
            
            .conflict-alert {
                background-color: #f44336;
                color: white;
                padding: 15px;
                margin: 20px 0;
                border-radius: 8px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.2);
                display: none;
            }
            .conflict-alert h3 {
                margin: 0 0 10px 0;
                font-size: 16px;
            }
            .conflict-section {
                margin: 10px 0;
                padding: 8px;
                background-color: rgba(255,255,255,0.1);
                border-radius: 4px;
            }
            .conflict-section h4 {
                margin: 0 0 5px 0;
                font-size: 14px;
                font-weight: bold;
            }
            .conflict-item {
                background-color: rgba(255,255,255,0.05);
                padding: 6px;
                margin: 3px 0;
                border-radius: 3px;
                font-size: 11px;
            }
            .conflict-cell {
                background-color: #f44336 !important;
                color: white !important;
                animation: pulse 1.5s infinite;
                border: 2px solid #d32f2f !important;
            }
            @keyframes pulse {
                0% { box-shadow: 0 0 0 0 rgba(244, 67, 54, 0.7); }
                70% { box-shadow: 0 0 0 10px rgba(244, 67, 54, 0); }
                100% { box-shadow: 0 0 0 0 rgba(244, 67, 54, 0); }
            }
            @media print {
                body { margin: 0; background-color: white; }
                .escala-container { box-shadow: none; border: 1px solid #ccc; margin-bottom: 15px; }
                .print-button, .filter-container { display: none; }
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
    <title>Visualizacao da Escala</title>
    {css}
</head>
<body>
    <h1>Visualizacao da Escala ({data_inicio.strftime('%d/%m/%Y')} - {data_fim.strftime('%d/%m/%Y')})</h1>
    
    <div class="filter-container">
        <h3 style="margin: 0 0 15px 0; color: #0d47a1;">Filtrar por Cidade:</h3>
        <button class="filter-button active" data-cidade-filtro="todas">Todas as Cidades</button>"""
        
        for cidade in cidades_na_escala:
            cidade_id = cidade.replace(' ', '_').replace('-', '_')
            html += f"""
        <button class="filter-button" data-cidade-filtro="{cidade_id}">{cidade}</button>"""
        
        html += """
    </div>
    
    <div id="conflictAlert" class="conflict-alert">
        <h3>âš ï¸ CONFLITOS DETECTADOS</h3>
        <div id="conflictList"></div>
    </div>
    
    <button id="printButton" class="print-button">Imprimir Escala</button>
"""

        for bloco_num, bloco_datas in enumerate(blocos_seg_dom):
            html += f"""<div class="escala-container">
        <table>
            <thead>
                <tr>
                    <th class="nome-col">Funcionario</th>
"""
            for data in bloco_datas:
                dia_str_html = f"{data.day}/{data.month}<br>({dias_semana_abrev[data.weekday()]})"
                html += f"<th>{dia_str_html}</th>"
            html += "</tr></thead><tbody>"

            for nome in df.index:
                is_folguista = nome in nomes_folguistas
                pessoa_obj = next((p for p in self.pessoas if p["nome"] == nome), None)
                is_supervisao = pessoa_obj and pessoa_obj.get("cidade") == "SUPERVISAO"
                
                nome_class = "folguista-nome" if is_folguista else ("supervisao-nome" if is_supervisao else "")
                cidade_funcionario = nome_para_cidade.get(nome, "")
                cidade_id = cidade_funcionario.replace(' ', '_').replace('-', '_')
                
                nome_display = nome
                if is_folguista:
                    nome_display += " (Folguista)"
                elif is_supervisao:
                    nome_display += " (Supervisao)"
                
                html += f"""<tr class="funcionario-row" data-cidade="{cidade_id}">
                    <td class="nome-col {nome_class}">{nome_display}</td>
"""
                for data in bloco_datas:
                    valor_celula = df.at[nome, data]
                    cell_class = ""
                    if valor_celula == "Folga": 
                        cell_class += " folga"
                    elif valor_celula == "Folga (Banco de Horas)": 
                        cell_class += " folga-banco-horas"
                    elif "(Cob)" in str(valor_celula): 
                        cell_class += " cobertura-00-06"
                    elif is_supervisao and valor_celula != "-":
                        cell_class += " supervisao-trabalhando"
                    elif is_folguista and valor_celula != "-": 
                        cell_class += " folguista-trabalhando"
                    
                    html += f"""<td class="{cell_class}">{str(valor_celula)}</td>"""
                html += "</tr>"
            html += "</tbody></table></div>"

        try:
            with open(ARQUIVO_SCRIPT_HTML, "r", encoding="utf-8") as arquivo_script:
                script_compilado = arquivo_script.read()
        except OSError as erro:
            raise RuntimeError(
                f"Nao foi possivel carregar o JavaScript compilado em {ARQUIVO_SCRIPT_HTML}. "
                "Execute 'tsc -p tsconfig.json' antes de exportar o HTML."
            ) from erro

        html += f"<script>\n{script_compilado}\n</script>"
        html += "</body></html>"
        return html

    def mostrar_erro(self, mensagem):
        from tkinter import messagebox
        messagebox.showerror("Erro", mensagem)

    def mostrar_info(self, mensagem):
        from tkinter import messagebox
        messagebox.showinfo("Informacao", mensagem)

if __name__ == "__main__":
    app = EscalaApp()
    app.mainloop()

