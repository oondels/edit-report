from tkinter import filedialog
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from docx import Document
from docx.shared import Inches
from datetime import date
from tkinter import messagebox

relatorio = Document("./data/relatorio-legendado.docx")

def editDoc(doc, padrao, novo):
    for paragrafo in doc.paragraphs:
        if padrao in paragrafo.text:
            paragrafo.text = paragrafo.text.replace(padrao, novo)

def addDefects(array):
    defects_table = relatorio.tables[0]
    for defect in array:       
        newDefect = defects_table.cell(0,0)
        newDefect.add_paragraph().add_run().add_text(defect)
    print("Defeitos adicionados...")
    
def addServices(array):
    services_table = relatorio.tables[1]
    for service in array:       
        newService = services_table.cell(0,0)
        newService.add_paragraph().add_run().add_text(service)
    print("Defeitos adicionados...")

def addMeasures(array):
    measures_table = relatorio.tables[2]
    
    for measure in array:
        new_meausre = measures_table.cell(0,0)
        new_meausre.add_paragraph().add_run().add_text(measure)

def sendImages():
    filenames = filedialog.askopenfilenames(title="Selecione as imagens", filetypes=[("Image files", "*.jpg *.png *.jpeg *.bmp")])
    pictures_table = relatorio.tables[3]

    for i in range(0, len(filenames), 2):
        sub_array = filenames[i : i+2]
        row = pictures_table.add_row()
        
        for j in range(len(sub_array)):
            cell = row.cells[j]
            cell.add_paragraph().add_run().add_picture(sub_array[j], width=Inches(1.0))
            
def success_popup():
    # Exibe uma caixa de diálogo de informação
    messagebox.showinfo("Sucesso!", "O relatório foi salvo com sucesso!")

def show_error():
    # Exibe uma caixa de diálogo de erro
    messagebox.showerror("Título do Erro", "Ocorreu um erro!")

def saveRelatorio():
    data_atual = date.today()
    print(data_atual)
    equipament.delete(0, END)
    osNumber.delete(0, END)
    defects.delete("1.0",END)
    services.delete("1.0",END)
    measure.delete("1.0",END)
    
    editDoc(relatorio, "nome-equipamento", equipament.get())
    editDoc(relatorio, "numero-os", osNumber.get())
    editDoc(relatorio, "data-entrada", date_entrada.entry.get())
    editDoc(relatorio, "data-execucao", execution.entry.get())
    
    array_defects = defects.get("1.0",END).split(",")
    addDefects(array_defects)
    
    array_services = services.get("1.0",END).split(",")
    addServices(array_services)
    
    array_measures = measure.get("1.0",END).split(",")
    addMeasures(array_measures)

    success_popup()
    relatorio.save("./Relatorios-Editados/relatorio_editado.docx")
    print("Relatorio salvo")

if __name__ == "__main__":
    app = ttk.Window(themename="superhero", minsize = (900,600))
    
    app.title("Relatórios Engenharia")
    app.geometry("900x700") 

    # Frame principal
    main_frame = ttk.Frame(app, padding="20")
    main_frame.pack(expand=True)
    
    # Title Frame
    info_frame = ttk.Frame(main_frame, padding="10", relief="solid", borderwidth=1)
    info_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
    
    labelTitle = ttk.Label(info_frame, text="Relatórios - Diesel Hidráulica", font=("Helvetica", 16, "bold"))
    labelTitle.pack(pady=10)
    
    label_select_number = ttk.Label(info_frame, text="Selecione quantidade de relatórios:")
    label_select_number.pack(pady=5)
    select_number = ttk.Spinbox(info_frame, width=5, bootstyle="info")
    select_number.pack(pady=5)
    
    # Frame para a primeira coluna
    div_one = ttk.Frame(main_frame, padding="10", relief=RAISED, borderwidth=2)
    div_one.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
    
    labelEquipament = ttk.Label(div_one, text="Nome do equipamento", font=("Helvetica", 12, "bold"))
    labelEquipament.grid(row=0, column=0, padx=5, pady=5, sticky=W)
    equipament = ttk.Entry(div_one, bootstyle="primary", width=30)
    equipament.grid(row=1, column=0, padx=5, pady=5)
    
    labelOs = ttk.Label(div_one, text="Número de O.S", font=("Helvetica", 12, "bold"))
    labelOs.grid(row=2, column=0, padx=5, pady=5, sticky=W)
    osNumber = ttk.Entry(div_one, bootstyle="primary", width=30)
    osNumber.grid(row=3, column=0, padx=5, pady=5)
    
    dateLabel = ttk.Label(div_one, text="Data de entrada", font=("Helvetica", 12, "bold"))
    dateLabel.grid(row=4, column=0, padx=5, pady=5, sticky=W)
    date_entrada = ttk.DateEntry(div_one, bootstyle="primary", width=30)
    date_entrada.grid(row=5, column=0, padx=5, pady=5)
    
    executionLabel = ttk.Label(div_one, text="Data de execução", font=("Helvetica", 12, "bold"))
    executionLabel.grid(row=6, column=0, padx=5, pady=5, sticky=W)
    execution = ttk.DateEntry(div_one, bootstyle="success", width=30)
    execution.grid(row=7, column=0, padx=5, pady=5)
    
    # Frame para a segunda coluna
    div_two = ttk.Frame(main_frame, padding="10", relief=RAISED, borderwidth=2)
    div_two.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
    
    defectsLabel = ttk.Label(div_two, text="Defeitos", font=("Helvetica", 12, "bold"))
    defectsLabel.grid(row=0, column=0, padx=5, pady=5, sticky=W)
    defects = tk.Text(div_two, wrap=tk.WORD, height=3, width=30)
    defects.grid(row=1, column=0, padx=5, pady=5)
    
    servicesLabel = ttk.Label(div_two, text="Serviços", font=("Helvetica", 12, "bold"))
    servicesLabel.grid(row=2, column=0, padx=5, pady=5, sticky=W)
    services = tk.Text(div_two, wrap=tk.WORD, height=3, width=30)
    services.grid(row=3, column=0, padx=5, pady=5)
    
    measureLabel = ttk.Label(div_two, text="Medidas", font=("Helvetica", 12, "bold"))
    measureLabel.grid(row=4, column=0, padx=5, pady=5, sticky=W)
    measure = tk.Text(div_two, wrap=tk.WORD, height=3, width=30)
    measure.grid(row=5, column=0, padx=5, pady=5)
    
    picturesLabel = ttk.Label(div_two, text="Fotos", font=("Helvetica", 12, "bold"))
    picturesLabel.grid(row=6, column=0, padx=5, pady=5, sticky=W)
    pictures = ttk.Button(div_two, text="Selecionar fotos", bootstyle="success", command=sendImages)
    pictures.grid(row=7, column=0, padx=5, pady=5)
    
    # Submit Frame
    submitFrame = ttk.Frame(main_frame, padding="10", relief=RAISED, borderwidth=2)
    submitFrame.grid(row=2, column=0, columnspan=2, sticky="nsew")

    buttonEnviar = ttk.Button(submitFrame, text="Salvar", command=saveRelatorio)
    buttonEnviar.grid(row=0, column=0, columnspan=2, sticky="nsew")
    
    # Ajuste a proporção das colunas e linhas
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_rowconfigure(1, weight=1)
    
    app.mainloop()
