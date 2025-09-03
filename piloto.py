import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Frame, Label
import pandas as pd
import PyPDF2
import xml.etree.ElementTree as ET

# Variável global para armazenar itens
itens = []

# Funções para extrair dados das NFe
def extrair_valores_nfe_xml(file_path):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        emitente = root.find('.//nfe:emit/nfe:xNome', namespaces=ns).text if root.find('.//nfe:emit/nfe:xNome', namespaces=ns) is not None else 'Emitente não disponível'
        numero_nf = root.find('.//nfe:ide/nfe:nNF', namespaces=ns).text if root.find('.//nfe:ide/nfe:nNF', namespaces=ns) is not None else 'Número não disponível'
        serie_nf = root.find('.//nfe:ide/nfe:serie', namespaces=ns).text if root.find('.//nfe:ide/nfe:serie', namespaces=ns) is not None else 'Série não disponível'

        global itens
        itens = []

        for det in root.findall('.//nfe:det', namespaces=ns):
            item = {}
            item['codigo'] = det.find('.//nfe:cProd', namespaces=ns).text if det.find('.//nfe:cProd', namespaces=ns) is not None else 'Código não disponível'
            item['descricao'] = det.find('.//nfe:xProd', namespaces=ns).text if det.find('.//nfe:xProd', namespaces=ns) is not None else 'Descrição não disponível'
            item['quantidade'] = float(det.find('.//nfe:qCom', namespaces=ns).text) if det.find('.//nfe:qCom', namespaces=ns) is not None else 0.0
            item['valor_unitario'] = float(det.find('.//nfe:vUnCom', namespaces=ns).text) if det.find('.//nfe:vUnCom', namespaces=ns) is not None else 0.0
            item['valor_total'] = float(det.find('.//nfe:vProd', namespaces=ns).text) if det.find('.//nfe:vProd', namespaces=ns) is not None else 0.0
            item['icms'] = float(det.find('.//nfe:vICMS', namespaces=ns).text) if det.find('.//nfe:vICMS', namespaces=ns) is not None else 0.0
            item['ipi'] = float(det.find('.//nfe:vIPI', namespaces=ns).text) if det.find('.//nfe:vIPI', namespaces=ns) is not None else 0.0
            item['pis'] = float(det.find('.//nfe:vPIS', namespaces=ns).text) if det.find('.//nfe:vPIS', namespaces=ns) is not None else 0.0
            item['cofins'] = float(det.find('.//nfe:vCOFINS', namespaces=ns).text) if det.find('.//nfe:vCOFINS', namespaces=ns) is not None else 0.0

            item['valor_liquido'] = item['valor_total'] - (item['icms'] + item['ipi'] + item['pis'] + item['cofins'])
            item['valor_liquido_unitario'] = item['valor_liquido'] / item['quantidade'] if item['quantidade'] != 0 else 0.0
            itens.append(item)

        return emitente, numero_nf, serie_nf

    except ET.ParseError:
        messagebox.showerror("Erro", "Erro ao processar o XML. Verifique se o arquivo está correto.")
        return None, None, None

def extrair_valores_nfe_pdf(file_path):
    try:
        emitente = "Emitente do PDF"
        numero_nf = "Número do PDF"
        serie_nf = "Série do PDF"
        itens = []
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                # Implementar a lógica para extrair os valores da NFe do texto do PDF
        return emitente, numero_nf, serie_nf, itens

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o PDF: {e}")
        return None, None, None

def extrair_valores_nfe(file_path):
    if file_path.endswith('.xml'):
        return extrair_valores_nfe_xml(file_path)
    elif file_path.endswith('.pdf'):
        return extrair_valores_nfe_pdf(file_path)
    else:
        messagebox.showerror("Erro", "Formato de arquivo não suportado. Selecione um arquivo XML ou PDF.")
        return None, None, None

def selecionar_arquivo():
    file_path = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=[("Arquivos XML", "*.xml"), ("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
    )

    if file_path:
        emitente, numero_nf, serie_nf = extrair_valores_nfe(file_path)
        if emitente and numero_nf and serie_nf and itens:
            resultado = f"Emitente: {emitente}\nNúmero da NF: {numero_nf}\nSérie da NF: {serie_nf}\n"
            resultado += "-" * 40 + "\n"
            resultado += f"Total de itens encontrados na NFe: {len(itens)}\n"
            for idx, item in enumerate(itens, start=1):
                resultado += f"Item {idx}:\n"
                resultado += f"Código: {item.get('codigo', 'Código não disponível')}\n"
                resultado += f"Descrição: {item.get('descricao', 'Descrição não disponível')}\n"
                resultado += f"Quantidade: {item['quantidade']}\n"
                resultado += f"Valor unitário: R${item['valor_unitario']:.2f}\n"
                resultado += f"Valor total: R${item['valor_total']:.2f}\n"
                resultado += f"Icms: R${item['icms']:.2f}\n"
                resultado += f"Ipi: R${item['ipi']:.2f}\n"
                resultado += f"Pis: R${item['pis']:.2f}\n"
                resultado += f"Cofins: R${item['cofins']:.2f}\n"
                resultado += f"Valor líquido unitário do item: R${item['valor_liquido_unitario']:.2f}\n"
                resultado += f"Valor líquido do item: R${item['valor_liquido']:.2f}\n"
                resultado += "-" * 20 + "\n"
            exibir_resultado(resultado)
        else:
            messagebox.showwarning("Aviso", "Nenhuma informação encontrada no arquivo.")
    else:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")

def exibir_resultado(resultado):
    resultado_text.config(state=tk.NORMAL)  # Habilita a edição temporariamente
    resultado_text.delete(1.0, tk.END)  # Limpa o texto anterior
    resultado_text.insert(tk.END, resultado)  # Insere o novo resultado
    resultado_text.config(state=tk.DISABLED)  # Desabilita a edição

def gerar_relatorio():
    if not itens:
        messagebox.showwarning("Aviso", "Nenhum item encontrado para gerar relatório.")
        return

    # Cria um DataFrame do pandas
    df = pd.DataFrame(itens)
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso: {save_path}")

# Criando a interface gráfica com estilo moderno
root = tk.Tk()
root.title("Calculadora de NFe")
root.geometry("800x600")
root.configure(bg="#f0f0f0")

# Estilo do cabeçalho
header_frame = Frame(root, bg="#2196F3")  # Azul
header_frame.pack(fill=tk.X)

header_label = Label(header_frame, text="Calculadora de NFe", font=("Helvetica", 20), bg="#2196F3", fg="white")
header_label.pack(pady=10)

# Botão para selecionar arquivo
btn_selecionar = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo, fg="black", font=("Helvetica", 12), relief=tk.FLAT)
btn_selecionar.pack(pady=10)

# Botão para gerar relatório
btn_relatorio = tk.Button(root, text="Gerar Relatório Excel", command=gerar_relatorio, fg="black", font=("Helvetica", 12), relief=tk.FLAT)
btn_relatorio.pack(pady=10)

# Área de texto para mostrar os resultados
resultado_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=90, height=30, bg="#ffffff", fg="#000000", font=("Helvetica", 12))
resultado_text.pack(padx=10, pady=10)
resultado_text.config(state=tk.DISABLED)  # Inicia como não editável

# Rodar a interface gráfica
root.mainloop()