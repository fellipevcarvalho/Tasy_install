import tkinter as tk
from tkinter import messagebox
import subprocess
import winreg as reg
import shutil
import os
import time
import sys

def centralizar_janela(janela):
    janela.update_idletasks()
    largura = 300
    altura = 150
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry('{}x{}+{}+{}'.format(largura, altura, x, y))

def mostrar_janela_informativa():
    try:
        # Instalação do BDE
        resposta = messagebox.askyesno("Pergunta", "Deseja instalar o BDE ?")

        if resposta:
            caminho_executavel_bde = r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\Borland_win64\Setup_BDE52_PTB'
            subprocess.run([caminho_executavel_bde], check=True)
            messagebox.showinfo("Instalação BDE Concluída", "A instalação do BDE foi bem-sucedida!")

            # Configuração após a instalação do BDE
            caminho_origem_config_bde = r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\IDAPI32.cfg'
            caminho_destino_bde = r'C:\Program Files (x86)\Common Files\Borland Shared\BDE'
            shutil.copy(caminho_origem_config_bde, caminho_destino_bde)
            
            caminho_arquivo_sql8 =r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\SQLORA8.DLL'
            caminho_destino_sql8 =r'C:\Program Files (x86)\Common Files\Borland Shared\BDE'
            shutil.copy(caminho_arquivo_sql8,caminho_destino_sql8)
            
            caminho_arquivo_sql32 =r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\SQLORA32.DLL'
            caminho_destino_sql32 =r'C:\Program Files (x86)\Common Files\Borland Shared\BDE'
            shutil.copy(caminho_arquivo_sql32,caminho_destino_sql32)
            
            caminho_arquivo_reg_bde = r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\Borland_win64\Borland.reg'
            subprocess.run(['regedit', '/s', caminho_arquivo_reg_bde], check=True)
            messagebox.showinfo("Configuração BDE Concluída", "A configuração do BDE foi realizada com sucesso!")

        # Instalação Oracle Client 11
        resposta_oracle = messagebox.askyesno("Pergunta", "Deseja instalar o Oracle ?")
        if resposta_oracle:
            caminho_executavel_oracle = r'\\ti14\SOFTWARE\TASY - instalação completa\3º - Oracle Client 11\setup.exe'
            subprocess.run([caminho_executavel_oracle], check=True)
            time.sleep(60*2)
            while True:
                resposta_instalacao_oracle = messagebox.askyesno("Pergunta", "A instalação do Oracle foi finalizada?")
                if resposta_instalacao_oracle:
                    messagebox.showinfo("Instalação Oracle Client Concluída", "A instalação do Oracle Client 11 foi bem-sucedida!")
                    # Execução do Oracle.reg
                    caminho_arquivo_reg_oracle = r'\\ti14\SOFTWARE\TASY - instalação completa\3º - Oracle Client 11\Oracle.reg'
                    subprocess.run(['regedit', '/s', caminho_arquivo_reg_oracle], check=True)
                    messagebox.showinfo("Configuração Oracle Concluída", "A configuração do Oracle Client foi realizada com sucesso!")
                    break
                else:
                    time.sleep(30)

        # Registros Fim - Registros_win64
        caminho_registros_win64 = r'\\ti14\SOFTWARE\TASY - instalação completa\4º - Registros Fim\Registros_win64.reg'
        subprocess.run(['regedit', '/s', caminho_registros_win64], check=True)        
        messagebox.showinfo("Configuração Registros Concluída", "A configuração dos registros foi realizada com sucesso!")

        # Copiar arquivos tnsnames.ora e sqlnet.ora para a pasta C:\oracle\product\11.2.0\client_1\network\admin
        caminho_origem_tnsnames = r'\\ti14\SOFTWARE\TASY - instalação completa\5º - TNS local - colar no network\tnsnames.ora'
        caminho_origem_sqlnet = r'\\ti14\SOFTWARE\TASY - instalação completa\5º - TNS local - colar no network\sqlnet.ora'
        caminho_destino_oracle = r'C:\oracle\product\11.2.0\client_1\network\admin'

        shutil.copy(caminho_origem_tnsnames, caminho_destino_oracle)
        shutil.copy(caminho_origem_sqlnet, caminho_destino_oracle)

        messagebox.showinfo("Configuração Tasy", "O Tasy foi configurado!")
        
        # Pergunta se deseja instalar certificado JAVA  
        resposta_certificado_java = messagebox.askyesno("Pergunta", "Deseja instalar o certificado ?")

        if resposta_certificado_java:
            caminho_java_certificado = r'\\ti14\SOFTWARE\TASY - instalação completa\Java Certificado\Java_Runtime_Environment_(32bit)_v8_Update_241.exe'
            subprocess.run([caminho_java_certificado], check=True)
            # Obtém a chave do Registro do Sistema
            key_path = r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
            with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, key_path, 0, reg.KEY_SET_VALUE) as key:
                # Obtém o valor atual do PATH
                current_path = os.environ.get('PATH', '')
                # Adiciona %JAVA_HOME% e C:\Program Files (x86)\Java\jre1.8.0_241\bin ao PATH
                new_path = f'%JAVA_HOME%;C:\\Program Files (x86)\\Java\\jre1.8.0_241\\bin;{current_path}'
                reg.SetValueEx(key, 'Path', 0, reg.REG_EXPAND_SZ, new_path)

            # Atualiza as variáveis de ambiente na sessão atual
            os.environ['PATH'] = new_path

        janela_principal.destroy()
        sys.exit()

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro durante a instalação", f"Ocorreu um erro durante a instalação: {e}")
        time.sleep(5)
        janela_principal.destroy()
        sys.exit()
def reparar_software():
    try:
        #Reparação
        caminho_executavel_bde = r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\Borland_win64\Setup_BDE52_PTB'
        subprocess.run([caminho_executavel_bde], check=True)
        messagebox.showinfo("Instalação BDE Concluída", "A instalação do BDE foi bem-sucedida!")

        # Configuração após a instalação do BDE
        caminho_origem_config_bde = r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\IDAPI32.cfg'
        caminho_destino_bde = r'C:\Program Files (x86)\Common Files\Borland Shared\BDE'
        shutil.copy(caminho_origem_config_bde, caminho_destino_bde)

        caminho_arquivo_reg_bde = r'\\ti14\SOFTWARE\TASY - instalação completa\2º - BDE\Borland_win64\Borland.reg'
        subprocess.run(['regedit', '/s', caminho_arquivo_reg_bde], check=True)
        messagebox.showinfo("Configuração BDE Concluída", "A configuração do BDE foi realizada com sucesso!")
        
        pasta_oracle = r'C:\Oracle'
        try:
            shutil.rmtree(pasta_oracle)
            messagebox.showinfo("Exclusão Concluída")
        except PermissionError as permission_error:
            messagebox.showerror("Erro de permissão")
        
        pasta_program_files_oracle = r'C:\Program Files (x86)\Oracle'
        try:
            shutil.rmtree(pasta_program_files_oracle)
        except PermissionError as permission_error:
            messagebox.showerror("Erro de Permissão")

            caminho_executavel_oracle = r'\\ti14\SOFTWARE\TASY - instalação completa\3º - Oracle Client 11\setup.exe'
            subprocess.run([caminho_executavel_oracle], check=True)
            time.sleep(60*2)

        while True:
                resposta_instalacao_oracle = messagebox.askyesno("Pergunta", "A instalação do Oracle foi finalizada?")
        if resposta_instalacao_oracle:
                messagebox.showinfo("Instalação Oracle Client Concluída", "A instalação do Oracle Client 11 foi bem-sucedida!")
                # Execução do Oracle.reg
                caminho_arquivo_reg_oracle = r'\\ti14\SOFTWARE\TASY - instalação completa\3º - Oracle Client 11\Oracle.reg'
                subprocess.run(['regedit', '/s', caminho_arquivo_reg_oracle], check=True)
                messagebox.showinfo("Configuração Oracle Concluída", "A configuração do Oracle Client foi realizada com sucesso!")
        else:
                time.sleep(30)
                messagebox.showinfo("Reparo Concluído com sucesso")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro durante o reparo", f"Ocorreu um erro durante as operações de reparo: {e}")
# Iniciar o loop principal da interface gráfica
janela_principal = tk.Tk()
janela_principal.title("Instalação de Software")

# Botão que aciona a janela informativa e inicia a instalação
botao_instalar = tk.Button(janela_principal, text="Instalação completa", command=mostrar_janela_informativa)
botao_instalar.pack(pady=20)
# Botão para acionar as operações de reparo
botao_reparar = tk.Button(janela_principal, text="Reparar Oracle", command=reparar_software)
botao_reparar.pack(pady=20)

#Centralizar Janela
centralizar_janela(janela_principal)
# Iniciar o loop principal da interface gráfica
janela_principal.mainloop()
