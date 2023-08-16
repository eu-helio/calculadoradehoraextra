import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import openpyxl

class HorasExtrasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Calculadora de Horas Extras")
        self.root.geometry("600x400")

        self.dados = []
        self.limite_diario = 9 * 60 + 20  # Limite de 9 horas e 20 minutos em minutos

        self.label_titulo = tk.Label(root, text="Calculadora de Horas Extras", font=("Helvetica", 16))
        self.label_criador = tk.Label(root, text="Criado por: Seu Nome", font=("Helvetica", 10))
        
        self.label_entrada = tk.Label(root, text="Hora de Entrada (HH:MM):")
        self.label_saida = tk.Label(root, text="Hora de Saída (HH:MM):")
        self.label_almoco = tk.Label(root, text="Duração do Almoço (HH:MM):")

        self.entrada_entrada = tk.Entry(root)
        self.entrada_saida = tk.Entry(root)
        self.entrada_almoco = tk.Entry(root)

        self.botao_calcular = tk.Button(root, text="Calcular", command=self.calcular)
        self.botao_salvar = tk.Button(root, text="Salvar", command=self.salvar_planilha)

        self.tree = ttk.Treeview(root, columns=("Data", "Entrada", "Saída", "Almoço", "Horas Trabalhadas", "Horas Extras"), show="headings")
        self.tree.heading("Data", text="Data")
        self.tree.heading("Entrada", text="Entrada")
        self.tree.heading("Saída", text="Saída")
        self.tree.heading("Almoço", text="Almoço")
        self.tree.heading("Horas Trabalhadas", text="Horas Trabalhadas")
        self.tree.heading("Horas Extras", text="Horas Extras")
        self.tree.bind("<ButtonRelease-1>", self.selecionar_item)

        self.label_total_mes = tk.Label(root, text="Total de Horas Extras no Mês:")
        self.label_total = tk.Label(root, text="0:00")

        self.botao_excluir = tk.Button(root, text="Excluir", command=self.excluir_item)
        self.botao_editar = tk.Button(root, text="Editar", command=self.editar_item)

        self.label_aviso = tk.Label(root, text="", font=("Helvetica", 10), fg="green")
        
        self.label_titulo.pack()
        self.label_criador.pack()
        self.label_entrada.pack()
        self.entrada_entrada.pack()
        self.label_saida.pack()
        self.entrada_saida.pack()
        self.label_almoco.pack()
        self.entrada_almoco.pack()
        self.botao_calcular.pack()
        self.botao_salvar.pack()
        self.tree.pack()
        self.label_total_mes.pack()
        self.label_total.pack()
        self.botao_excluir.pack()
        self.botao_editar.pack()
        self.label_aviso.pack()

        self.item_selecionado = None

    def calcular(self):
        entrada = self.entrada_entrada.get()
        saida = self.entrada_saida.get()
        almoco = self.entrada_almoco.get()

        if entrada and saida and almoco:
            try:
                entrada = datetime.strptime(entrada, '%H:%M')
                saida = datetime.strptime(saida, '%H:%M')
                almoco = datetime.strptime(almoco, '%H:%M')

                minutos_almoco = almoco.hour * 60 + almoco.minute
                minutos_trabalhados = (saida - entrada - timedelta(minutes=minutos_almoco)).total_seconds() / 60

                if minutos_trabalhados > self.limite_diario:
                    minutos_extras = minutos_trabalhados - self.limite_diario
                else:
                    minutos_extras = 0

                data_atual = datetime.now().strftime('%Y-%m-%d')
                self.dados.append((data_atual, entrada.strftime('%H:%M'), saida.strftime('%H:%M'), almoco.strftime('%H:%M'), minutos_trabalhados, minutos_extras))
                self.atualizar_tabela()
                self.atualizar_total_mes()
                self.label_aviso.config(text="Dados calculados com sucesso.", fg="green")
            except ValueError:
                tk.messagebox.showerror("Erro", "Formato de hora inválido. Use HH:MM.")
                self.label_aviso.config(text="Erro ao calcular. Formato de hora inválido.", fg="red")

    def formatar_horas_minutos(self, minutos):
        horas = int(minutos // 60)
        minutos = int(minutos % 60)
        return f"{horas:02d}:{minutos:02d}"

    def atualizar_tabela(self):
        self.tree.delete(*self.tree.get_children())
        for data, entrada, saida, almoco, minutos_trabalhados, minutos_extras in self.dados:
            self.tree.insert("", "end", values=(data, entrada, saida, almoco, self.formatar_horas_minutos(minutos_trabalhados), self.formatar_horas_minutos(minutos_extras)))

    def atualizar_total_mes(self):
        total_mes = sum(minutos_extras for _, _, _, _, _, minutos_extras in self.dados)
        self.label_total.config(text=self.formatar_horas_minutos(total_mes))

    def salvar_planilha(self):
        planilha = openpyxl.Workbook()
        aba = planilha.active
        aba.title = 'Horas Extras'

        cabecalho = ['Data', 'Hora de Entrada', 'Hora de Saída', 'Almoço', 'Horas Trabalhadas', 'Horas Extras']
        aba.append(cabecalho)

        for data, entrada, saida, almoco, minutos_trabalhados, minutos_extras in self.dados:
            aba.append([data, entrada, saida, almoco, self.formatar_horas_minutos(minutos_trabalhados), self.formatar_horas_minutos(minutos_extras)])

        nome_arquivo = 'horas_extras.xlsx'
        planilha.save(nome_arquivo)
        self.label_aviso.config(text=f"Planilha '{nome_arquivo}' salva com sucesso.", fg="green")

    def selecionar_item(self, event):
        item_selecionado = self.tree.focus()
        if item_selecionado:
            self.item_selecionado = item_selecionado

    def excluir_item(self):
        if self.item_selecionado:
            indice = self.tree.index(self.item_selecionado)
            del self.dados[indice]
            self.item_selecionado = None
            self.atualizar_tabela()
            self.atualizar_total_mes()

    def editar_item(self):
        if self.item_selecionado:
            indice = self.tree.index(self.item_selecionado)
            item = self.dados[indice]
            self.entrada_entrada.delete(0, tk.END)
            self.entrada_saida.delete(0, tk.END)
            self.entrada_almoco.delete(0, tk.END)
            self.entrada_entrada.insert(0, item[1])
            self.entrada_saida.insert(0, item[2])
            self.entrada_almoco.insert(0, item[3])
            self.excluir_item()

if __name__ == "__main__":
    root = tk.Tk()
    app = HorasExtrasApp(root)
    root.mainloop()
