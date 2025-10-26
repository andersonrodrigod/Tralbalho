import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox
from controllers.excluir_linhas_controller import ExcluirLinhasController

class ExcluirLinhas(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.controller = ExcluirLinhasController()
        
        # Configuração de cores
        self.cor_primaria = "#238AD9"
        self.cor_sucesso = "#28A745"
        self.cor_perigo = "#DC3545"
        self.cor_alerta = "#FFC107"

        # Variáveis
        self.arquivo_selecionado = None
        self.abas_disponiveis = []
        self.abas_selecionadas = []
        self.criterios_exclusao = []

        # === NOVO: FRAME PRINCIPAL SCROLLABLE ===
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self,
            scrollbar_button_color=self.cor_primaria,
            scrollbar_button_hover_color="#1A6BA6"
        )
        self.scrollable_frame.pack(fill="both", expand=True)

        # Label principal (agora dentro do scrollable)
        self.label = ctk.CTkLabel(
            self.scrollable_frame,
            text="🗑️ Excluir Linhas de Planilhas",
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color=self.cor_primaria
        )
        self.label.pack(pady=(50, 30))

        # Botão Voltar (fora do scrollable para ficar fixo)
        self.btn_back = ctk.CTkButton(
            self, 
            text="← Voltar", 
            command=lambda: master.show_frame(master.menu_frame), 
            width=100,
            height=35,
            fg_color="gray",
            hover_color="#5A6268"
        )
        self.btn_back.place(x=20, y=20)

        # Frame principal (agora dentro do scrollable)
        self.frame_principal = ctk.CTkFrame(self.scrollable_frame)
        self.frame_principal.pack(pady=10, padx=40, fill="both", expand=True)

        # --- Frame para Seleção de Arquivo ---
        frame_arquivo = ctk.CTkFrame(self.frame_principal)
        frame_arquivo.pack(fill="x", pady=(0, 20), padx=20)

        label_titulo_arquivo = ctk.CTkLabel(
            frame_arquivo,
            text="📁 Selecionar Arquivo Excel",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label_titulo_arquivo.pack(anchor="w", pady=(15, 10))

        subframe_arquivo = ctk.CTkFrame(frame_arquivo, fg_color="transparent")
        subframe_arquivo.pack(fill="x", padx=10, pady=(0, 15))

        self.btn_selecionar_arquivo = ctk.CTkButton(
            subframe_arquivo,
            text="Procurar Arquivo",
            command=self.selecionar_arquivo,
            width=140,
            height=35,
            fg_color=self.cor_primaria,
            hover_color="#1A6BA6"
        )
        self.btn_selecionar_arquivo.pack(side="left", padx=(0, 10))

        self.entry_arquivo = ctk.CTkEntry(
            subframe_arquivo,
            width=400,
            height=35,
            placeholder_text="Nenhum arquivo selecionado",
            state="readonly"
        )
        self.entry_arquivo.pack(side="left", fill="x", expand=True)

        # --- Frame para Seleção de Abas ---
        self.frame_abas = ctk.CTkFrame(self.frame_principal)
        self.frame_abas.pack(fill="x", pady=(0, 20), padx=20)

        label_titulo_abas = ctk.CTkLabel(
            self.frame_abas,
            text="📊 Selecionar Abas para Processar",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label_titulo_abas.pack(anchor="w", pady=(15, 10))

        # Frame para controles de seleção de abas
        frame_controles_abas = ctk.CTkFrame(self.frame_abas, fg_color="transparent")
        frame_controles_abas.pack(fill="x", padx=10, pady=(0, 10))

        self.btn_selecionar_todas = ctk.CTkButton(
            frame_controles_abas,
            text="✅ Selecionar Todas",
            command=self.selecionar_todas_abas,
            width=140,
            height=30,
            fg_color=self.cor_sucesso,
            hover_color="#218838",
            state="disabled"
        )
        self.btn_selecionar_todas.pack(side="left", padx=(0, 10))

        self.btn_limpar_selecao = ctk.CTkButton(
            frame_controles_abas,
            text="🗑️ Limpar Seleção",
            command=self.limpar_selecao_abas,
            width=140,
            height=30,
            fg_color=self.cor_perigo,
            hover_color="#C82333",
            state="disabled"
        )
        self.btn_limpar_selecao.pack(side="left")

        # Label para mostrar seleção
        self.label_selecao_abas = ctk.CTkLabel(
            frame_controles_abas,
            text="Nenhuma aba selecionada",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.label_selecao_abas.pack(side="right")

        # Frame para a lista de abas (scrollable) - REDUZIDO
        self.frame_lista_abas = ctk.CTkScrollableFrame(
            self.frame_abas,
            height=80,  # Reduzido
            fg_color="#f0f0f0"
        )
        self.frame_lista_abas.pack(fill="x", padx=10, pady=(0, 10))

        # --- Frame para Critérios de Exclusão ---
        self.frame_criterios = ctk.CTkFrame(self.frame_principal)
        self.frame_criterios.pack(fill="x", pady=(0, 20), padx=20)

        label_titulo_criterios = ctk.CTkLabel(
            self.frame_criterios,
            text="🎯 Critérios de Exclusão de Linhas",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label_titulo_criterios.pack(anchor="w", pady=(15, 10))

        # Frame para adicionar novos critérios
        frame_novo_criterio = ctk.CTkFrame(self.frame_criterios, fg_color="transparent")
        frame_novo_criterio.pack(fill="x", padx=10, pady=(0, 10))

        # Label de instrução
        label_instrucao = ctk.CTkLabel(
            frame_novo_criterio,
            text="Adicione valores para excluir linhas. Deixe a coluna em branco para buscar em todas as colunas.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            wraplength=600
        )
        label_instrucao.pack(anchor="w", pady=(0, 10))

        # Subframe para inputs do critério
        subframe_inputs = ctk.CTkFrame(frame_novo_criterio, fg_color="transparent")
        subframe_inputs.pack(fill="x", pady=5)

        # Input do valor a ser excluído
        self.label_valor = ctk.CTkLabel(
            subframe_inputs,
            text="Valor a excluir:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.label_valor.pack(side="left", padx=(0, 10))

        self.entry_valor = ctk.CTkEntry(
            subframe_inputs,
            placeholder_text="Ex: _Julho, teste, 2023...",
            width=200,
            height=35
        )
        self.entry_valor.pack(side="left", padx=(0, 20))

        # Input da coluna específica (opcional)
        self.label_coluna = ctk.CTkLabel(
            subframe_inputs,
            text="Coluna específica (opcional):",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.label_coluna.pack(side="left", padx=(0, 10))

        self.entry_coluna = ctk.CTkEntry(
            subframe_inputs,
            placeholder_text="Ex: Coluna_A, Nome, ID...",
            width=200,
            height=35
        )
        self.entry_coluna.pack(side="left", padx=(0, 20))

        # Botão para adicionar critério
        self.btn_adicionar_criterio = ctk.CTkButton(
            subframe_inputs,
            text="➕ Adicionar Critério",
            command=self.adicionar_criterio,
            width=140,
            height=35,
            fg_color=self.cor_sucesso,
            hover_color="#218838"
        )
        self.btn_adicionar_criterio.pack(side="left")

        # Frame para lista de critérios adicionados - REDUZIDO
        self.frame_lista_criterios = ctk.CTkScrollableFrame(
            self.frame_criterios,
            height=60,  # Reduzido
            label_text="Critérios de Exclusão Adicionados",
            fg_color="#f8f9fa"
        )
        self.frame_lista_criterios.pack(fill="x", padx=10, pady=(0, 10))

        # --- Frame para Configurações Avançadas ---
        self.frame_config = ctk.CTkFrame(self.frame_principal)
        self.frame_config.pack(fill="x", pady=(0, 20), padx=20)

        label_titulo_config = ctk.CTkLabel(
            self.frame_config,
            text="⚙️ Configurações Avançadas",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label_titulo_config.pack(anchor="w", pady=(15, 10))

        # Frame para opções avançadas
        frame_opcoes = ctk.CTkFrame(self.frame_config, fg_color="transparent")
        frame_opcoes.pack(fill="x", padx=10, pady=(0, 15))

        # Coluna 1 - Opções básicas
        frame_col1 = ctk.CTkFrame(frame_opcoes, fg_color="transparent")
        frame_col1.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.check_var_vazias = ctk.BooleanVar(value=True)
        self.check_duplicatas = ctk.BooleanVar(value=False)

        self.check_vazias = ctk.CTkCheckBox(
            frame_col1,
            text="Excluir linhas completamente vazias",
            variable=self.check_var_vazias,
            font=ctk.CTkFont(size=13)
        )
        self.check_vazias.pack(anchor="w", pady=5)

        self.check_duplicatas = ctk.CTkCheckBox(
            frame_col1,
            text="Excluir linhas duplicadas",
            variable=self.check_duplicatas,
            font=ctk.CTkFont(size=13)
        )
        self.check_duplicatas.pack(anchor="w", pady=5)

        # Coluna 2 - Modo de busca
        frame_col2 = ctk.CTkFrame(frame_opcoes, fg_color="transparent")
        frame_col2.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.radio_var = ctk.StringVar(value="contem")
        
        self.radio_contem = ctk.CTkRadioButton(
            frame_col2,
            text="Valor contém o texto",
            variable=self.radio_var,
            value="contem",
            font=ctk.CTkFont(size=13)
        )
        self.radio_contem.pack(anchor="w", pady=5)

        self.radio_exato = ctk.CTkRadioButton(
            frame_col2,
            text="Valor exato",
            variable=self.radio_var,
            value="exato",
            font=ctk.CTkFont(size=13)
        )
        self.radio_exato.pack(anchor="w", pady=5)

        # === IMPORTANTE: Botão Executar e Status FORA do frame_principal ===
        # Mas DENTRO do scrollable_frame para aparecer no scroll

        # --- Botão Executar ---
        self.btn_executar = ctk.CTkButton(
            self.scrollable_frame,  # Agora no scrollable_frame, não no self
            text="🚀 Executar Exclusão de Linhas",
            command=self.executar_exclusao,
            height=45,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=self.cor_sucesso,
            hover_color="#218838",
            state="disabled"
        )
        self.btn_executar.pack(pady=(20, 20), padx=40, fill="x")

        # Frame para status
        self.frame_status = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")  # No scrollable_frame
        self.frame_status.pack(pady=(0, 30), padx=40, fill="x")  # Pady aumentado na parte inferior

        self.label_status = ctk.CTkLabel(
            self.frame_status,
            text="Selecione um arquivo Excel para começar",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.label_status.pack()

        # Barra de progresso (inicialmente oculta)
        self.progress_bar = ctk.CTkProgressBar(self.frame_status)
        self.progress_bar.set(0)

        # Inicialmente esconder frames que dependem de arquivo
        self.frame_abas.pack_forget()
        self.frame_criterios.pack_forget()
        self.frame_config.pack_forget()

    # Os demais métodos permanecem EXATAMENTE os mesmos...
    def adicionar_criterio(self):
        """Adiciona um novo critério à lista"""
        valor = self.entry_valor.get().strip()
        coluna = self.entry_coluna.get().strip()

        if not valor:
            messagebox.showwarning("Aviso", "Digite um valor para excluir!")
            return

        # Criar critério
        criterio = {
            'valor': valor,
            'coluna': coluna if coluna else None,  # None significa todas as colunas
            'tipo_busca': self.radio_var.get()
        }

        # Adicionar à lista
        self.criterios_exclusao.append(criterio)

        # Atualizar interface
        self.atualizar_lista_criterios()

        # Limpar campos
        self.entry_valor.delete(0, 'end')
        self.entry_coluna.delete(0, 'end')

        self.label_status.configure(text=f"Critério adicionado! Total: {len(self.criterios_exclusao)}", text_color="blue")

    def atualizar_lista_criterios(self):
        """Atualiza a lista visual de critérios"""
        # Limpar lista atual
        for widget in self.frame_lista_criterios.winfo_children():
            widget.destroy()

        # Adicionar cada critério
        for i, criterio in enumerate(self.criterios_exclusao):
            frame_criterio = ctk.CTkFrame(self.frame_lista_criterios, fg_color="white")
            frame_criterio.pack(fill="x", pady=2, padx=5)

            # Texto do critério
            if criterio['coluna']:
                texto = f"🔍 {criterio['valor']} → Coluna: {criterio['coluna']} ({criterio['tipo_busca']})"
            else:
                texto = f"🔍 {criterio['valor']} → Todas as colunas ({criterio['tipo_busca']})"

            label_criterio = ctk.CTkLabel(
                frame_criterio,
                text=texto,
                font=ctk.CTkFont(size=11),
                anchor="w"
            )
            label_criterio.pack(side="left", padx=10, pady=5, fill="x", expand=True)

            # Botão para remover critério
            btn_remover = ctk.CTkButton(
                frame_criterio,
                text="✕",
                width=30,
                height=25,
                fg_color="red",
                hover_color="darkred",
                command=lambda idx=i: self.remover_criterio(idx)
            )
            btn_remover.pack(side="right", padx=5)

    def remover_criterio(self, index):
        """Remove um critério da lista"""
        if 0 <= index < len(self.criterios_exclusao):
            self.criterios_exclusao.pop(index)
            self.atualizar_lista_criterios()
            self.label_status.configure(text=f"Critério removido! Total: {len(self.criterios_exclusao)}", text_color="orange")

    def selecionar_arquivo(self):
        """Seleciona um arquivo Excel e carrega suas abas"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        
        if arquivo:
            self.arquivo_selecionado = arquivo
            self.entry_arquivo.configure(state="normal")
            self.entry_arquivo.delete(0, "end")
            self.entry_arquivo.insert(0, arquivo.split("/")[-1])
            self.entry_arquivo.configure(state="readonly")
            
            # Carregar abas do arquivo
            self.carregar_abas(arquivo)
            
            # Mostrar frames adicionais
            self.frame_abas.pack(fill="x", pady=(0, 20), padx=20)
            self.frame_criterios.pack(fill="x", pady=(0, 20), padx=20)
            self.frame_config.pack(fill="x", pady=(0, 20), padx=20)
            
            self.label_status.configure(text="Arquivo carregado com sucesso! Selecione as abas e adicione critérios.", text_color="green")
            self.btn_selecionar_todas.configure(state="normal")
            self.btn_limpar_selecao.configure(state="normal")

    def carregar_abas(self, arquivo):
        """Carrega todas as abas do arquivo Excel"""
        try:
            # Limpar lista anterior
            for widget in self.frame_lista_abas.winfo_children():
                widget.destroy()
            
            # Ler abas do Excel
            excel_file = pd.ExcelFile(arquivo)
            self.abas_disponiveis = excel_file.sheet_names
            self.abas_selecionadas = []
            
            # Criar checkboxes para cada aba
            for aba in self.abas_disponiveis:
                self.criar_checkbox_aba(aba)
            
            self.atualizar_label_selecao()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel:\n{str(e)}")

    def criar_checkbox_aba(self, nome_aba):
        """Cria um checkbox para uma aba específica"""
        var = ctk.BooleanVar(value=False)
        
        frame_aba = ctk.CTkFrame(self.frame_lista_abas, fg_color="white")
        frame_aba.pack(fill="x", pady=2, padx=5)
        
        checkbox = ctk.CTkCheckBox(
            frame_aba,
            text=nome_aba,
            variable=var,
            font=ctk.CTkFont(size=12),
            command=lambda: self.atualizar_selecao_aba(nome_aba, var.get())
        )
        checkbox.pack(side="left", padx=10, pady=5)
        
        # Armazenar referência
        checkbox.var = var
        checkbox.nome_aba = nome_aba

    def atualizar_selecao_aba(self, nome_aba, selecionada):
        """Atualiza a lista de abas selecionadas"""
        if selecionada and nome_aba not in self.abas_selecionadas:
            self.abas_selecionadas.append(nome_aba)
        elif not selecionada and nome_aba in self.abas_selecionadas:
            self.abas_selecionadas.remove(nome_aba)
        
        self.atualizar_label_selecao()
        self.verificar_pronto_execucao()

    def selecionar_todas_abas(self):
        """Seleciona todas as abas disponíveis"""
        for widget in self.frame_lista_abas.winfo_children():
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    if isinstance(child, ctk.CTkCheckBox):
                        child.var.set(True)
                        if child.nome_aba not in self.abas_selecionadas:
                            self.abas_selecionadas.append(child.nome_aba)
        
        self.atualizar_label_selecao()
        self.verificar_pronto_execucao()

    def limpar_selecao_abas(self):
        """Limpa a seleção de todas as abas"""
        for widget in self.frame_lista_abas.winfo_children():
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    if isinstance(child, ctk.CTkCheckBox):
                        child.var.set(False)
        
        self.abas_selecionadas = []
        self.atualizar_label_selecao()
        self.verificar_pronto_execucao()

    def atualizar_label_selecao(self):
        """Atualiza o label que mostra quantas abas estão selecionadas"""
        total = len(self.abas_disponiveis)
        selecionadas = len(self.abas_selecionadas)
        
        if selecionadas == 0:
            texto = "Nenhuma aba selecionada"
            cor = "gray"
        elif selecionadas == total:
            texto = f"Todas as {total} abas selecionadas"
            cor = "green"
        else:
            texto = f"{selecionadas} de {total} abas selecionadas"
            cor = "blue"
        
        self.label_selecao_abas.configure(text=texto, text_color=cor)

    def verificar_pronto_execucao(self):
        """Verifica se pode habilitar o botão executar"""
        if (self.arquivo_selecionado and 
            self.abas_selecionadas and 
            (self.criterios_exclusao or self.check_var_vazias.get() or self.check_duplicatas.get())):
            
            self.btn_executar.configure(state="normal")
            
            if self.criterios_exclusao:
                status_text = f"Pronto! {len(self.abas_selecionadas)} aba(s) e {len(self.criterios_exclusao)} critério(s)"
            else:
                status_text = f"Pronto! {len(self.abas_selecionadas)} aba(s) com filtros básicos"
                
            self.label_status.configure(text=status_text, text_color="green")
        else:
            self.btn_executar.configure(state="disabled")
            if self.arquivo_selecionado and self.abas_selecionadas:
                self.label_status.configure(text="Adicione pelo menos um critério de exclusão", text_color="orange")

    def executar_exclusao(self):
        """Executa o processo de exclusão de linhas"""
        if not self.arquivo_selecionado or not self.abas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione um arquivo e pelo menos uma aba!")
            return

        if not self.criterios_exclusao and not self.check_var_vazias.get() and not self.check_duplicatas.get():
            messagebox.showwarning("Aviso", "Adicione pelo menos um critério de exclusão!")
            return

        # Mostrar barra de progresso
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_bar.set(0.1)
        
        # Atualizar status
        self.label_status.configure(text="Iniciando processo de exclusão...", text_color="orange")
        self.btn_executar.configure(state="disabled")

        try:
            # Coletar configurações
            config = {
                'excluir_vazias': self.check_var_vazias.get(),
                'excluir_duplicatas': self.check_duplicatas.get(),
                'criterios_personalizados': self.criterios_exclusao,
                'tipo_busca': self.radio_var.get(),
                'abas_selecionadas': self.abas_selecionadas
            }

            # Processar arquivo
            self.progress_bar.set(0.3)
            
            # Chamar o controller para processar
            resultado = self.controller.processar_arquivo(
                self.arquivo_selecionado, 
                config
            )

            self.progress_bar.set(0.8)

            if resultado:
                # Salvar arquivo processado
                self.salvar_arquivo_processado(resultado)
            else:
                raise Exception("Nenhum dado processado")

        except Exception as e:
            self.label_status.configure(text="❌ Erro no processamento!", text_color="red")
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")
            self.progress_bar.pack_forget()
            self.btn_executar.configure(state="normal")

    def salvar_arquivo_processado(self, dados_processados):
        """Salva o arquivo com as linhas excluídas"""
        caminho = filedialog.asksaveasfilename(
            title="Salvar arquivo processado", 
            defaultextension=".xlsx", 
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )

        if not caminho:
            self.label_status.configure(text="Operação cancelada pelo usuário", text_color="gray")
            self.progress_bar.pack_forget()
            self.btn_executar.configure(state="normal")
            return

        try:
            # Salvar o arquivo processado
            with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
                for aba, df in dados_processados.items():
                    df.to_excel(writer, sheet_name=aba, index=False)
            
            self.progress_bar.set(1.0)
            
            # Mensagem de resumo
            total_criterios = len(self.criterios_exclusao)
            total_abas = len(self.abas_selecionadas)
            
            self.label_status.configure(
                text=f"✅ Sucesso! {total_abas} aba(s) processada(s) com {total_criterios} critério(s)", 
                text_color="green"
            )
            
            messagebox.showinfo(
                "Concluído", 
                f"Arquivo processado com sucesso!\n\n"
                f"• Abas processadas: {total_abas}\n"
                f"• Critérios aplicados: {total_criterios}\n"
                f"• Local: {caminho}"
            )
            
        except Exception as e:
            self.label_status.configure(text="❌ Erro ao salvar arquivo!", text_color="red")
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo:\n{str(e)}")
        
        finally:
            # Esconder barra de progresso após 3 segundos
            self.after(3000, self.progress_bar.pack_forget)
            self.btn_executar.configure(state="normal")