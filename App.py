import random
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
import os

class BancaTCC:
    def __init__(self, professor1, professor2, data, horario):
        self.professor1 = professor1
        self.professor2 = professor2
        self.data = data
        self.horario = horario
        self.fitness = 0


def criar_populacao_inicial_do_excel(file_path):
    df = pd.read_excel(file_path)
    populacao = []
    professores = df['Professor'].unique().tolist()

    for _, row in df.iterrows():
        dias_banca = str(row['Dia']).split(';')
        horarios_banca = row['Horario'].split(';')

        for dia, horario in zip(dias_banca, horarios_banca):
            professor1 = random.choice(professores)
            professor2 = random.choice(professores)
            while professor1 == professor2:
                professor2 = random.choice(professores)

            banca = BancaTCC(professor1, professor2,
                             dia.strip(), horario.strip())
            populacao.append(banca)

    return populacao


def calcular_fitness(banca, horarios_desejados, dias_desejados):
    fitness = 0

    # Penaliza se professor1 for igual a professor2
    if banca.professor1 == banca.professor2:
        fitness -= 10

    # Recompensar bancas em horários desejados
    if banca.horario in horarios_desejados[banca.professor1]:
        fitness += 10  # Recompensa por horário desejado

    # Recompensar bancas em dias desejados
    if banca.data in dias_desejados[banca.professor1]:
        fitness += 5  # Recompensa por dia desejado

    # Recompensar bancas em horários desejados
    if banca.horario in horarios_desejados[banca.professor2]:
        fitness += 10  # Recompensa por horário desejado

    # Recompensar bancas em dias desejados
    if banca.data in dias_desejados[banca.professor2]:
        fitness += 5  # Recompensa por dia desejado

    # Atualizar o atributo de fitness da banca
    banca.fitness = fitness


def selecionar_pais(populacao, tamanho_torneio):
    torneio = random.sample(populacao, tamanho_torneio)
    torneio.sort(key=lambda x: x.fitness, reverse=True)
    melhor_participante = torneio[0]
    return melhor_participante


def realizar_crossover(pai1, pai2):
    while True:
        professor1_filho1 = pai1.professor1
        professor2_filho1 = pai2.professor2

        professor1_filho2 = pai2.professor1
        professor2_filho2 = pai1.professor2

        filho1 = BancaTCC(professor1_filho1, professor2_filho1,
                          pai1.data, pai1.horario)
        filho2 = BancaTCC(professor1_filho2, professor2_filho2,
                          pai2.data, pai2.horario)

        # Verificar se os filhos já existem na população
        if filho1 not in populacao and filho2 not in populacao:
            return filho1, filho2


def realizar_mutacao(banca, professores, horarios_desejados_unicos, dias_desejados_unicos, taxa_mutacao):
    if random.random() < taxa_mutacao:
        novo_professor1 = random.choice(professores)
        novo_professor2 = random.choice(professores)

        while novo_professor1 == novo_professor2:
            novo_professor2 = random.choice(professores)

        novo_dia = random.choice(dias_desejados_unicos)
        novo_horario = random.choice(horarios_desejados_unicos)

        banca.professor1 = novo_professor1
        banca.professor2 = novo_professor2
        banca.data = novo_dia
        banca.horario = novo_horario


def exibir_resultados(resultados):
    root = tk.Tk()
    root.title("Resultados da População Final")

    largura_janela = 1024
    altura_janela = 720
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    x_pos = (largura_tela - largura_janela) // 2
    y_pos = (altura_tela - altura_janela) // 2
    root.geometry(f"{largura_janela}x{altura_janela}+{x_pos}+{y_pos}")

    tree = ttk.Treeview(root, columns=("Banca", "Resultado"),
                        show="headings", height=len(resultados))
    tree.column("Banca", width=100, anchor=tk.CENTER)
    tree.column("Resultado", width=900, anchor=tk.W)
    tree.heading("Banca", text="Banca")
    tree.heading("Resultado", text="Resultado")

    for i, resultado in enumerate(resultados):
        tree.insert("", tk.END, values=(f"Banca {i+1}", resultado))

    tree.pack(pady=10)

    root.mainloop()


def abrir_janela():
    def obter_caminho_arquivo():
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo .xlsx")
        return file_path

    def selecionar_arquivo():
        file_path = obter_caminho_arquivo()
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                tk.messagebox.showwarning(
                    "Aviso", "Por favor, selecione um arquivo .xlsx válido.")
                return
            excel_path.set(file_path)
            caminho_selecionado.config(state=tk.NORMAL)
            caminho_selecionado.delete(0, tk.END)
            caminho_selecionado.insert(0, file_path)
            caminho_selecionado.config(state=tk.DISABLED)

    root = tk.Tk()
    root.title("Gerador Inteligente de Bancas de TCC")

    largura_janela = 400
    altura_janela = 300
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    x_pos = (largura_tela - largura_janela) // 2
    y_pos = (altura_tela - altura_janela) // 2
    root.geometry(f"{largura_janela}x{altura_janela}+{x_pos}+{y_pos}")

    excel_path = tk.StringVar()

    titulo_label = tk.Label(
        root, text="Gerador Inteligente de Bancas de TCC", font=("Helvetica", 16))
    titulo_label.pack(pady=10)

    descricao_label = tk.Label(root, text="Para gerar as bancas de TCC, informe uma planilha .xlsx com as colunas: Professor, Dia, Horario", font=(
        "Helvetica", 12), wraplength=300)
    descricao_label.pack(pady=10)

    caminho_selecionado = tk.Entry(
        root, width=40, state=tk.DISABLED, font=("Helvetica", 10))
    caminho_selecionado.pack(pady=10)

    btn_selecionar = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo, bg="#4CAF50",
                               fg="white", padx=10, pady=5, borderwidth=0, font=("Helvetica", 12), cursor="hand2")
    btn_selecionar.pack(pady=10)

    def gerar():
        arquivo_selecionado = excel_path.get()
        if not arquivo_selecionado:
            tk.messagebox.showwarning(
                "Aviso", "Por favor, selecione um arquivo.")
            return

        if not os.path.exists(arquivo_selecionado):
            tk.messagebox.showwarning(
                "Aviso", "O arquivo selecionado não existe.")
            return

        if not arquivo_selecionado.lower().endswith('.xlsx'):
            tk.messagebox.showwarning(
                "Aviso", "Por favor, selecione um arquivo .xlsx válido.")
            return

        root.destroy()

    btn_gerar = tk.Button(root, text="Gerar", command=gerar, bg="#3498db", fg="white",
                          padx=10, pady=5, borderwidth=0, font=("Helvetica", 12), cursor="hand2")
    btn_gerar.pack(pady=20)

    root.mainloop()

    return excel_path.get()


# Caminho do arquivo Excel
excel_path = abrir_janela()

# Hiperparâmetros
tamanho_populacao = 200
tamanho_torneio = 3
taxa_mutacao = 0.1
numero_geracoes = 100

# Inicializa a população
populacao = criar_populacao_inicial_do_excel(excel_path)

# Definir as variáveis globais para professores
df = pd.read_excel(excel_path)
professores = df.Professor

# Cria um dicionário para armazenar os horários e dias desejados para cada professor
horarios_desejados = {}
dias_desejados = {}

# Itera sobre as linhas da planilha
for _, row in df.iterrows():
    professor = row['Professor']
    horarios_desejados[professor] = []
    dias_desejados[professor] = []

    # Separando a coluna 'Horario' que contém os horários desejados no formato "10:00-12:00;14:00-16:00"
    horarios_desejados_str = row['Horario']
    dias_desejados_str = row['Dia']

    # Divide a string em horários individuais
    str(row['Dia']).split(';')
    horarios = horarios_desejados_str.split(';')
    dias = str(dias_desejados_str).split(';')

    # Adiciona os horários desejados à lista para o professor
    for horario in horarios:
        horarios_desejados[professor].append(horario.strip())

    # Adiciona os dias desejados à lista para o professor
    for dia in dias:
        dias_desejados[professor].append(dia.strip())

# Criar lista de horários desejados únicos
horarios_desejados_unicos = []

# Iterar sobre os horários desejados de cada professor
for horarios_professor in horarios_desejados.values():
    # Adicionar os horários à lista sem repetição
    for horario in horarios_professor:
        if horario not in horarios_desejados_unicos:
            horarios_desejados_unicos.append(horario)

# Ordenar a lista de horários
horarios_desejados_unicos.sort()

# Criar lista de dias desejados únicos
dias_desejados_unicos = []

# Iterar sobre os dias desejados de cada professor
for dias_professor in dias_desejados.values():
    # Adicionar os dias à lista sem repetição
    for dia in dias_professor:
        if dia not in dias_desejados_unicos:
            dias_desejados_unicos.append(dia)

# Ordenar a lista de dias
dias_desejados_unicos.sort()

melhores_bancas_globais = []

# Avalia o fitness da população
for banca in populacao:
    calcular_fitness(banca, horarios_desejados, dias_desejados)

nova_populacao = populacao

# Loop principal do algoritmo genético
for geracao in range(numero_geracoes):

    for _ in range(tamanho_populacao):
        pai1 = selecionar_pais(populacao, tamanho_torneio)
        pai2 = selecionar_pais(populacao, tamanho_torneio)

        # Realiza crossover para criar dois filhos
        filho1, filho2 = realizar_crossover(pai1, pai2)

        # Realiza mutação nos filhos
        realizar_mutacao(filho1, professores, horarios_desejados_unicos,
                         dias_desejados_unicos, taxa_mutacao)
        realizar_mutacao(filho2, professores, horarios_desejados_unicos,
                         dias_desejados_unicos, taxa_mutacao)

        # Adiciona os filhos à nova população
        nova_populacao.extend([filho1, filho2])

    # Substitui a antiga população pela nova população
    populacao = nova_populacao

    # Avalia o fitness da população
    for banca in populacao:
        calcular_fitness(banca, horarios_desejados, dias_desejados)

        # Atualiza a lista de melhores bancas globais
        if not any(banca.fitness == melhor.fitness and banca.professor1 == melhor.professor1 and banca.professor2 == melhor.professor2 and banca.data == melhor.data and banca.horario == melhor.horario for melhor in melhores_bancas_globais):
            melhores_bancas_globais.append(banca)

    # Ordena a lista de melhores bancas globais
    melhores_bancas_globais.sort(key=lambda x: x.fitness, reverse=True)

    # Mantém apenas as melhores bancas na lista
    melhores_bancas_globais = melhores_bancas_globais[:tamanho_populacao]

    # Adiciona a melhor banca global à nova população
    nova_populacao = melhores_bancas_globais.copy()

# Cria uma matriz para armazenar os resultados da população final
resultados_populacao_final = []

# Preenche a matriz com os resultados da população final
for banca in nova_populacao:
    resultados_populacao_final.append([
        banca.professor1,
        banca.professor2,
        banca.data,
        banca.horario,
        banca.fitness
    ])

resultados_populacao_final = []
for banca in nova_populacao:
    resultados_populacao_final.append(
        f"Professor 1: {banca.professor1}, Professor 2: {banca.professor2}, Data da Banca: {
            banca.data}, Horario da Banca: {banca.horario}, Fitness da Banca: {banca.fitness}"
    )

exibir_resultados(resultados_populacao_final)
