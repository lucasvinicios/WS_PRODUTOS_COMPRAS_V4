import tkinter as tk
from tkinter import messagebox
import time
import threading
import bot

# Listas de produtos e supermercados
produtos = ["Arroz 5kg",
            "Feijão Carioca 1kg", 
            "Macarrão Espaguete 500g",
            "Óleo de Soja 900ml",
            "Açúcar Cristal 5kg", 
            "Leite Integral 1L", 
            "Pão de Forma", 
            "Café Tradicional 500g", 
            "Detergente Líquido 500ml", 
            "Sabão em Pó 800g", 
            "Papel Higiênico",
            "Creme Dental 70g",
            "Água Sanitária 2l",
            "Sabonete 85g",
            "Fio Dental",
            "Molho de Tomate 300g",
            "Azeite Extra Virgem 500ml",
            "Farinha de Trigo 1kg",
            "Queijo Mussarela 150g",
            "Creme de Leite 200g",
            ]

supermercados = ["Tauste", "Barbosa", "Confiança", "Mercado Livre", "Coop Supermercado", "Tenda Atacado", "Boa Supermercado"]

# Criando janela principal
root = tk.Tk()
root.title("Lista de Compras")
root.geometry("500x800")
root.configure(bg="#ffffff")

# Cores e estilos
title_font = ("Arial", 14, "bold")
button_font = ("Arial", 10, "bold")
primary_color = "#0078D7"
secondary_color = "#E0E0E0"
text_color = "#333333"

# Criando os frames principais
frame_top = tk.Frame(root, bg=primary_color, height=60)
frame_top.pack(fill=tk.X)

frame_main = tk.Frame(root, bg="#ffffff", padx=20, pady=10)
frame_main.pack(fill=tk.BOTH, expand=True)

# Título
label_title = tk.Label(frame_top, text="Lista de Compras", font=title_font, fg="white", bg=primary_color)
label_title.pack(pady=15)

# Criando os frames para alinhar as colunas
frame_produtos = tk.LabelFrame(frame_main, text="Produtos", font=title_font, fg=text_color, bg="#ffffff", padx=10, pady=10)
frame_produtos.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

frame_supermercados = tk.LabelFrame(frame_main, text="Supermercados", font=title_font, fg=text_color, bg="#ffffff", padx=10, pady=10)
frame_supermercados.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

# Dicionários para armazenar os valores dos checkboxes
produtos_var = {produto: tk.BooleanVar() for produto in produtos}
supermercados_var = {mercado: tk.BooleanVar() for mercado in supermercados}

# Função para selecionar/deselecionar todos os produtos
def selecionar_todos_produtos():
    novo_estado = not all(var.get() for var in produtos_var.values())
    for var in produtos_var.values():
        var.set(novo_estado)

# Função para selecionar/deselecionar todos os supermercados
def selecionar_todos_supermercados():
    novo_estado = not all(var.get() for var in supermercados_var.values())
    for var in supermercados_var.values():
        var.set(novo_estado)

# Botão para selecionar todos os produtos
btn_select_all_produtos = tk.Button(frame_produtos, text="Selecionar Todos", command=selecionar_todos_produtos, font=button_font, bg=primary_color, fg="white")
btn_select_all_produtos.pack(pady=5)

# Criando os checkboxes para produtos
for produto, var in produtos_var.items():
    tk.Checkbutton(frame_produtos, text=produto, variable=var, font=("Arial", 10), bg="#ffffff").pack(anchor="w")

# Botão para selecionar todos os supermercados
btn_select_all_supermercados = tk.Button(frame_supermercados, text="Selecionar Todos", command=selecionar_todos_supermercados, font=button_font, bg=primary_color, fg="white")
btn_select_all_supermercados.pack(pady=5)

# Criando os checkboxes para supermercados
for mercado, var in supermercados_var.items():
    tk.Checkbutton(frame_supermercados, text=mercado, variable=var, font=("Arial", 10), bg="#ffffff").pack(anchor="w")

# Label para o loading
loading_label = tk.Label(root, text="", font=("Arial", 10, "bold"), fg="blue", bg="#ffffff")
loading_label.pack(pady=5)

# Função para exibir os itens selecionados após o loading
def mostrar_selecionados():
    btn_selecionar.config(text="Carregando...", state=tk.DISABLED, bg=secondary_color)
    loading_label.config(text="⏳ Carregando...")

    def processar():
        time.sleep(5)  # Simula o tempo de carregamento

        produtos_selecionados = [p for p, v in produtos_var.items() if v.get()]
        supermercados_selecionados = [s for s, v in supermercados_var.items() if v.get()]

        supermarket = bot.SuperMarket(supermarkets=supermercados_selecionados, products=produtos_selecionados)
        supermarket.extract_data()
        # supermarket.extract_data()

        msg = f"Produtos selecionados: {', '.join(produtos_selecionados) if produtos_selecionados else 'Nenhum'}\n"
        msg += f"Supermercados selecionados: {', '.join(supermercados_selecionados) if supermercados_selecionados else 'Nenhum'}"

        root.after(0, lambda: [
            loading_label.config(text=""),
            btn_selecionar.config(text="Selecionar", state=tk.NORMAL, bg=primary_color),
            messagebox.showinfo("Itens Selecionados", msg)
        ])

    threading.Thread(target=processar).start()

# Botão de seleção
btn_selecionar = tk.Button(root, text="Selecionar", command=mostrar_selecionados, font=button_font, bg=primary_color, fg="white", padx=20, pady=5)
btn_selecionar.pack(pady=20)

# Rodar o aplicativo
root.mainloop()
