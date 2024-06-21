import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

def merge_sheets():
    # Abrir diálogo para selecionar o arquivo Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    # Perguntar a quantidade de abas
    try:
        num_sheets = int(simpledialog.askstring("Input", "Quantas abas deseja fazer o merge?"))
    except (ValueError, TypeError):
        messagebox.showerror("Error", "Por favor, insira um número válido.")
        return

    if num_sheets < 2:
        messagebox.showerror("Error", "Por favor, insira pelo menos 2 abas para mesclar.")
        return

    # Perguntar os nomes das abas
    sheet_names = []
    for i in range(num_sheets):
        sheet_name = simpledialog.askstring("Input", f"Nome da aba {i + 1}:")
        if not sheet_name:
            messagebox.showerror("Error", "Nome da aba não pode estar vazio.")
            return
        sheet_names.append(sheet_name)

    # Carregar o arquivo Excel
    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Não foi possível ler o arquivo Excel: {e}")
        return

    # Ler as abas em dataframes
    dataframes = []
    try:
        for sheet_name in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            dataframes.append(df)
    except Exception as e:
        messagebox.showerror("Error", f"Não foi possível ler uma ou mais abas: {e}")
        return

    # Concatenar os dataframes
    combined_df = pd.concat(dataframes)

    # Remover duplicatas
    combined_df = combined_df.drop_duplicates()

    # Nome da nova aba
    new_sheet_name = simpledialog.askstring("Input", "Nome da nova aba:")
    if not new_sheet_name:
        messagebox.showerror("Error", "Nome da nova aba não pode estar vazio.")
        return

    # Salvar o resultado em uma nova aba
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            combined_df.to_excel(writer, sheet_name=new_sheet_name, index=False)  # <- index=False para evitar índice
        messagebox.showinfo("Success", "Dados combinados e duplicatas removidas com sucesso!")
    except Exception as e:
        messagebox.showerror("Error", f"Não foi possível salvar os dados: {e}")

# Configurar a interface gráfica
root = tk.Tk()
root.title("Mesclar Abas do Excel")

# Botão para iniciar o processo
merge_button = tk.Button(root, text="Selecionar Arquivo e Mesclar Abas", command=merge_sheets)
merge_button.pack(pady=20)

root.mainloop()
