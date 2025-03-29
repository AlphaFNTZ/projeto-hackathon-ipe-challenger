import customtkinter as ctk
from tkinter import filedialog, messagebox
import shutil
import os
import pandas as pd

# Configuração da janela principal
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Ordenador de Tarefas")
        self.geometry("500x350")
        self.iconbitmap("icon.ico") 
        
        self.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)

        self.label_descricao = ctk.CTkLabel(
            self, 
            text="Sistema de Ordenação e Alocação de Tarefas\n\n"
                 "Selecione um arquivo Excel com:\n"
                 "- Planilha1: Tarefas (Status, Prioridades, Time)\n"
                 "- Planilha2: Profissionais (Desenvolvedor, Time)",
            wraplength=450, 
            justify="center"
        )
        self.label_descricao.pack(pady=15)

        self.botao_upload = ctk.CTkButton(self, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        self.botao_upload.pack(pady=10)

        self.label_arquivo = ctk.CTkLabel(self, text="Nenhum arquivo selecionado")
        self.label_arquivo.pack(pady=5)

        self.botao_enviar = ctk.CTkButton(
            self, 
            text="Processar e Alocar", 
            command=self.processar_arquivo, 
            state="disabled"
        )
        self.botao_enviar.pack(pady=10)

        self.arquivo_path = None

    def selecionar_arquivo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            self.arquivo_path = file_path
            self.label_arquivo.configure(text=os.path.basename(file_path))
            self.botao_enviar.configure(state="normal")

    def ordenar_e_alocar_tarefas(self, arquivo_excel):
        try:
            # Carrega as planilhas
            df_tarefas = pd.read_excel(arquivo_excel, sheet_name="Planilha1")
            df_profissionais = pd.read_excel(arquivo_excel, sheet_name="Planilha2")
            
            # Verifica colunas obrigatórias
            colunas_tarefas = ['Status', 'Prioridade por cliente', 'Prioridade PM', 'Complexidade', 'Tamanho', 'Time']
            colunas_profissionais = ['Desenvolvedor ', 'Time']
            
            for col in colunas_tarefas:
                if col not in df_tarefas.columns:
                    raise ValueError(f"Coluna '{col}' não encontrada na Planilha1")
            
            for col in colunas_profissionais:
                if col not in df_profissionais.columns:
                    raise ValueError(f"Coluna '{col}' não encontrada na Planilha2")
            
            # Filtra apenas tarefas aprovadas e com PM != 0
            df_tarefas = df_tarefas[
                (df_tarefas['Status'].astype(str).str.strip().str.lower() == 'aprovada') &
                (df_tarefas['Prioridade PM'] != 0)
            ].copy()
            
            if df_tarefas.empty:
                raise ValueError("Nenhuma tarefa válida encontrada (Status 'Aprovada' e PM diferente de 0)")
            
            # Limpeza e normalização dos dados
            df_tarefas['Prioridade por cliente'] = df_tarefas['Prioridade por cliente'].astype(str).str.strip().str.lower()
            df_tarefas['Complexidade'] = df_tarefas['Complexidade'].astype(str).str.strip().str.lower()
            df_tarefas['Tamanho'] = df_tarefas['Tamanho'].astype(str).str.strip().str.upper()
            
            # Mapeamento para ordenação
            complexidade_map = {'alta': 3, 'média': 2, 'baixa': 1}
            tamanho_map = {'PP': 4, 'P': 3, 'M': 2, 'G': 1}
            
            df_tarefas['_complexidade_ord'] = df_tarefas['Complexidade'].map(complexidade_map)
            df_tarefas['_tamanho_ord'] = df_tarefas['Tamanho'].map(tamanho_map)
            
            # Separa tarefas por prioridade do cliente
            df_sim = df_tarefas[df_tarefas['Prioridade por cliente'] == 'sim'].copy()
            df_nao = df_tarefas[df_tarefas['Prioridade por cliente'] == 'não'].copy()
            
            # Ordenação para tarefas com "sim"
            df_sim_ordenado = df_sim.sort_values(
                by=['Prioridade PM'],
                ascending=True
            )
            
            # Ordenação para tarefas com "não"
            # Separa em PM <= 5 e PM > 5 (já excluímos PM == 0)
            df_nao_pm_baixo = df_nao[df_nao['Prioridade PM'] <= 5].copy()
            df_nao_pm_alto = df_nao[df_nao['Prioridade PM'] > 5].copy()
            
            # Ordena PM <= 5 por Tamanho (maior primeiro) e Complexidade (maior primeiro)
            df_nao_pm_baixo_ordenado = df_nao_pm_baixo.sort_values(
                by=['_tamanho_ord', '_complexidade_ord'],
                ascending=[False, False]
            )
            
            # Ordena PM > 5 por Prioridade PM (menor primeiro)
            df_nao_pm_alto_ordenado = df_nao_pm_alto.sort_values(
                by=['Prioridade PM'],
                ascending=True
            )
            
            # Combina todas as tarefas ordenadas
            df_ordenado = pd.concat([
                df_sim_ordenado,
                df_nao_pm_baixo_ordenado,
                df_nao_pm_alto_ordenado
            ])
            
            # Remove colunas temporárias
            df_ordenado = df_ordenado.drop(columns=['_complexidade_ord', '_tamanho_ord'])
            
            # Alocação por área de profissionais
            profissionais_por_time = df_profissionais['Time'].value_counts().to_dict()
            alocacao = {time: 0 for time in profissionais_por_time.keys()}
            
            # Filtra apenas as tarefas que podem ser alocadas
            tarefas_alocadas = []
            
            for _, tarefa in df_ordenado.iterrows():
                time_tarefa = tarefa['Time']
                
                if time_tarefa in alocacao and alocacao[time_tarefa] < profissionais_por_time.get(time_tarefa, 0):
                    tarefas_alocadas.append(tarefa)
                    alocacao[time_tarefa] += 1
            
            if not tarefas_alocadas:
                raise ValueError("Nenhuma tarefa pôde ser alocada com os profissionais disponíveis")
            
            # Cria DataFrame final com as tarefas alocadas
            df_final = pd.DataFrame(tarefas_alocadas)
            
            # Adiciona coluna com a quantidade de profissionais disponíveis por time
            df_final['Profissionais no Time'] = df_final['Time'].map(profissionais_por_time)
            
            return df_final
        
        except Exception as e:
            raise Exception(f"Erro no processamento: {str(e)}")

    def processar_arquivo(self):
        if not self.arquivo_path:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo primeiro.")
            return

        try:
            self.label_descricao.configure(text="Processando e alocando tarefas...")
            self.update()
            
            df_final = self.ordenar_e_alocar_tarefas(self.arquivo_path)
            
            nome_arquivo_original = os.path.basename(self.arquivo_path)
            nome_arquivo_processado = f"ALOCADAS_{nome_arquivo_original}"
            df_final.to_excel(nome_arquivo_processado, index=False)
            
            # Mostra resumo da alocação
            resumo = df_final['Time'].value_counts().to_string()
            messagebox.showinfo(
                "Alocação concluída",
                f"Tarefas alocadas por time:\n\n{resumo}\n\n"
                f"Arquivo salvo como: {nome_arquivo_processado}"
            )
            
            self.withdraw()  
            JanelaDownload(self, nome_arquivo_processado)
            
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.label_descricao.configure(text="Erro no processamento.\nVerifique os dados e tente novamente.")

    def fechar_aplicacao(self):
        self.destroy()
        os._exit(0)

class JanelaDownload(ctk.CTkToplevel):
    def __init__(self, parent, arquivo_processado):
        super().__init__(parent)

        self.title("Download do Arquivo Processado")
        self.geometry("500x250")
        self.wm_iconbitmap("icon.ico")
        self.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)

        self.label_mensagem = ctk.CTkLabel(
            self, 
            text="Tarefas ordenadas e alocadas com sucesso!\n\n"
                 "O arquivo contém apenas tarefas 'Aprovadas'\n"
                 "com PM ≠ 0, alocadas conforme disponibilidade.",
            wraplength=450,
            justify="center"
        )
        self.label_mensagem.pack(pady=15)

        self.arquivo_processado = arquivo_processado
        self.parent = parent

        self.botao_download = ctk.CTkButton(
            self, 
            text="Salvar Arquivo", 
            command=self.baixar_arquivo
        )
        self.botao_download.pack(pady=10)

        self.botao_novo_arquivo = ctk.CTkButton(
            self, 
            text="Processar Outro Arquivo", 
            command=self.voltar_para_tela_principal
        )
        self.botao_novo_arquivo.pack(pady=5)

    def baixar_arquivo(self):
        pasta_destino = filedialog.askdirectory()
        if pasta_destino:
            try:
                destino = os.path.join(pasta_destino, os.path.basename(self.arquivo_processado))
                shutil.copy2(self.arquivo_processado, destino)
                messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{destino}")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar arquivo: {str(e)}")
            finally:
                if os.path.exists(self.arquivo_processado):
                    os.remove(self.arquivo_processado)

    def voltar_para_tela_principal(self):
        if os.path.exists(self.arquivo_processado):
            os.remove(self.arquivo_processado)
        self.destroy()
        self.parent.deiconify()

    def fechar_aplicacao(self):
        if os.path.exists(self.arquivo_processado):
            os.remove(self.arquivo_processado)
        self.parent.destroy()
        os._exit(0)

if __name__ == "__main__":
    app = App()
    app.mainloop()