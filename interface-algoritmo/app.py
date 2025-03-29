import customtkinter as ctk
from tkinter import filedialog, messagebox
import shutil
import os

# Configuração da janela principal
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Upload de Arquivo")
        self.geometry("500x300")
        self.iconbitmap("icon.ico") 
        
        # Fechar toda a aplicação ao fechar a janela principal
        self.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)

        # Descrição do sistema
        self.label_descricao = ctk.CTkLabel(self, text="Bem-vindo ao sistema de processamento de arquivos .xlsx!\n\nSelecione um arquivo e clique em enviar.", wraplength=400, justify="center")
        self.label_descricao.pack(pady=20)

        # Botão para selecionar arquivo
        self.botao_upload = ctk.CTkButton(self, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        self.botao_upload.pack(pady=10)

        # Label para mostrar o arquivo selecionado
        self.label_arquivo = ctk.CTkLabel(self, text="Nenhum arquivo selecionado")
        self.label_arquivo.pack(pady=5)

        # Botão de enviar
        self.botao_enviar = ctk.CTkButton(self, text="Enviar", command=self.processar_arquivo, state="disabled")
        self.botao_enviar.pack(pady=10)

        self.arquivo_path = None  # Variável para armazenar o caminho do arquivo

    def selecionar_arquivo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            self.arquivo_path = file_path
            self.label_arquivo.configure(text=os.path.basename(file_path))  # Exibe o nome do arquivo
            self.botao_enviar.configure(state="normal")  # Ativa o botão de enviar

    def processar_arquivo(self):
        if not self.arquivo_path:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo primeiro.")
            return

        # Simulando o processamento e salvando o arquivo modificado
        novo_arquivo = "arquivo_processado.xlsx"
        shutil.copy(self.arquivo_path, novo_arquivo)  # Apenas copia o arquivo original para simular o processamento
        
        # Oculta a janela principal em vez de fechá-la
        self.withdraw()  
        
        # Abrir nova janela para baixar o arquivo
        JanelaDownload(self, novo_arquivo)

    def fechar_aplicacao(self):
        """Fecha toda a aplicação"""
        self.destroy()
        os._exit(0)  # Garante que todos os processos do Tkinter são encerrados


class JanelaDownload(ctk.CTkToplevel):
    def __init__(self, parent, arquivo_processado):
        super().__init__(parent)

        self.title("Download do Arquivo")
        self.geometry("400x200")

        # Definir ícone corretamente
        self.wm_iconbitmap("icon.ico")

        # Configurar para fechar toda a aplicação ao fechar a janela
        self.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)

        self.label_mensagem = ctk.CTkLabel(self, text="Seu arquivo foi processado com sucesso!\nClique abaixo para baixá-lo.")
        self.label_mensagem.pack(pady=20)

        self.arquivo_processado = arquivo_processado  # Caminho do arquivo processado

        self.botao_download = ctk.CTkButton(self, text="Baixar Arquivo", command=self.baixar_arquivo)
        self.botao_download.pack(pady=10)

        self.botao_voltar = ctk.CTkButton(self, text="Voltar", command=self.voltar_para_tela_principal)
        self.botao_voltar.pack(pady=10)

        self.parent = parent  # Referência à janela principal

    def baixar_arquivo(self):
        pasta_destino = filedialog.askdirectory()
        if pasta_destino:
            shutil.move(self.arquivo_processado, os.path.join(pasta_destino, self.arquivo_processado))
            messagebox.showinfo("Sucesso", "Arquivo baixado com sucesso!")

    def voltar_para_tela_principal(self):
        self.destroy()  # Fecha a janela de download
        self.parent.deiconify()  # Reexibe a janela principal

    def fechar_aplicacao(self):
        """Fecha toda a aplicação ao fechar a janela de download"""
        self.parent.destroy()
        os._exit(0)

if __name__ == "__main__":
    app = App()
    app.mainloop()