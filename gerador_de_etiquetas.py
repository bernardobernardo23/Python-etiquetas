import tkinter as tk
from tkinter import messagebox, filedialog
from fpdf import FPDF
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials
import io
import os
import subprocess
import sys
import re  # para validação regex


def recurso_caminho(relativo):
    """Retorna o caminho absoluto do recurso, lidando com o ambiente do PyInstaller."""
    if getattr(sys, 'frozen', False):  # Executável gerado
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relativo)
# --- Configuração de conexão com Google Sheets ---
credenciais_path = recurso_caminho("credenciais.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
CREDS = Credentials.from_service_account_file(credenciais_path, scopes=SCOPES)

gc = gspread.authorize(CREDS)
SHEET_ID = "id da planilha"
SHEET_NAME = "PRODUTOS"
planilha = gc.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# --- Funções auxiliares ---
def buscar_dados_por_id(id_produto):
    dados = planilha.get_all_records()
    for linha in dados:
        if str(linha["Codigo"]) == str(id_produto):
            return linha
    return None

def open_file(filepath):
    try:
        if os.path.exists(filepath):
            if os.name == 'nt':
                os.startfile(filepath)
            elif sys.platform == 'darwin':
                subprocess.call(('open', filepath))
            else:
                subprocess.call(('xdg-open', filepath))
        else:
            messagebox.showerror("Erro", f"Arquivo não encontrado: {filepath}")
    except Exception as e:
        messagebox.showerror("Erro ao abrir arquivo", f"Não foi possível abrir o arquivo:\n{e}")

if getattr(sys, 'frozen', False):
    dir_path = os.path.dirname(sys.executable)
else:
    dir_path = os.path.dirname(os.path.abspath(__file__))

pasta_etiquetas = os.path.join(dir_path, "etiquetas")
os.makedirs(pasta_etiquetas, exist_ok=True)

# --- Validação de data ---
def validar_data(texto):
    # Permite apenas números e '/' no formato dd/mm/aaaa ou parciais válidas
    # Limita tamanho máximo a 10 caracteres (dd/mm/aaaa)
    if len(texto) > 10:
        return False
    if re.fullmatch(r"(\d{0,2})(/)?(\d{0,2})?(/)?(\d{0,4})?", texto) or texto == "":
        return True
    return False

class EtiquetaPDF(FPDF):
    def __init__(self, width, height):
        super().__init__(unit="mm", format=(width, height))
        self.set_auto_page_break(auto=False)

    def gerar_etiqueta(self, tipo, descricao, codigo_barras, id_digitado, quantidade, unidade, lote, data_chegada, data_validade):
        self.add_page()

        logo_path = recurso_caminho("logo.png")
        logo_width_mm = 20
        logo_x = 0
        logo_y = 0

        if os.path.exists(logo_path):
            self.image(logo_path, x=logo_x, y=logo_y, w=logo_width_mm)
        else:
            print(f"⚠️ Logo não encontrada: {logo_path}")

        self.set_xy(logo_x + logo_width_mm + 10, logo_y)
        self.set_font("Arial", "B", 20)
        self.cell(0, 15, "CHESIQUÍMICA", ln=True, align="L")

        y_current = max(logo_y + 15, self.get_y() + 4)

        self.set_font("Arial", "B", 12)
        self.set_xy(0, y_current)
        self.cell(0, 10, f"Tipo: {tipo}", ln=True)
        y_current = self.get_y()

        if len(descricao) > 30:
            self.set_font("Arial", "B", 9)  # sem negrito
        else:
            self.set_font("Arial", "B", 12)  # negrito
        self.set_xy(0, y_current)
        self.multi_cell(self.w - 10, 6, f"Descrição: {descricao}", align="L")
        y_current = self.get_y() + 2

        self.set_font("Arial", "B", 12)
        self.set_xy(0, y_current)
        self.cell(0, 6, f"Quantidade: {quantidade} {unidade}", ln=True)
        y_current = self.get_y()

        # Reduz espaçamento se descrição for longa
        espacamento = 0
        
            
        if lote:
            self.set_xy(0, y_current)
            self.cell(0, 6, f"Lote: {lote}", ln=True)
            y_current = self.get_y() + espacamento

        if data_chegada:
            self.set_xy(0, y_current)
            self.cell(0, 6, f"Data de fabricação: {data_chegada}", ln=True)
            y_current = self.get_y() + espacamento

        if data_validade:
            self.set_xy(0, y_current)
            self.cell(0, 6, f"Data de Vencimento: {data_validade}", ln=True)
            y_current = self.get_y() + espacamento

        barcode_width = 60
        barcode_height = 20
        barcode_x = (self.w - barcode_width) / 2
        barcode_y = y_current + 0  # ajuste para manter o código visível

        barcode_img = Code128(str(codigo_barras), writer=ImageWriter())
        buffer = io.BytesIO()
        barcode_img.write(buffer, options={
            "module_width": 0.3,
            "module_height": barcode_height,
            "quiet_zone": 1,
            "write_text": False
        })
        buffer.seek(0)

        temp_img_path = os.path.join(dir_path, "temp_barcode.png")
        with Image.open(buffer) as img:
            img.save(temp_img_path)

        self.image(temp_img_path, x=barcode_x-4, y=barcode_y, w=barcode_width, h=barcode_height)

        self.set_xy(barcode_x, barcode_y + barcode_height + 2)
        self.set_font("Arial", "B", 12)
        self.cell(barcode_width, 0, f"ID do Produto: {id_digitado}", align="C")

        if os.path.exists(temp_img_path):
            os.remove(temp_img_path)

# [O restante do código permanece igual — inclusive as f

def gerar_pdf_action():
    id_produto_digitado = entry_id.get().strip()
    quantidade = entry_quantidade.get().strip()
    lote = entry_lote.get().strip()
    data_chegada = entry_data_chegada.get().strip()
    data_validade = entry_data_validade.get().strip()

    if not id_produto_digitado or not quantidade:
        messagebox.showerror("Erro", "Preencha o Código do produto e a quantidade")
        return

    dados = buscar_dados_por_id(id_produto_digitado)
    if not dados:
        messagebox.showerror("Erro", "Produto não encontrado na planilha.")
        return

    tipo = dados.get("TIPO_INSUMO", "")
    descricao = dados.get("Descricao", "")
    codigo_barras = str(dados.get("Codigo", ""))
    unidade = dados.get("Unidade", "").upper()

    tamanho_selecionado = tamanho_var.get()
    if tamanho_selecionado == "100x96":
        largura, altura = 100, 96
        tamanho_str = "100x96"
    elif tamanho_selecionado == "120x96":
        largura, altura = 120, 96
        tamanho_str = "120x96"
    elif tamanho_selecionado == "personalizado":
        try:
            largura = int(entry_largura.get())
            altura = int(entry_altura.get())
            tamanho_str = f"{largura}x{altura}"
        except ValueError:
            messagebox.showerror("Erro", "Informe valores numéricos válidos para largura e altura.")
            return
    else:
        messagebox.showerror("Erro", "Selecione um tamanho de etiqueta.")
        return

    pdf = EtiquetaPDF(largura, altura)
    pdf.gerar_etiqueta(tipo, descricao, codigo_barras, id_produto_digitado, quantidade, unidade, lote, data_chegada, data_validade)

    descricao_limpa = re.sub(r"[^\w]", "", descricao)
    salvar_caminho = os.path.join(pasta_etiquetas, f"{id_produto_digitado}_{descricao_limpa}_{tamanho_str}.pdf")

    try:
        pdf.output(salvar_caminho)
        if messagebox.showinfo("Sucesso", f"Etiqueta PDF gerada e salva em:\n{salvar_caminho}"):
            open_file(salvar_caminho)
    except Exception as e:
        messagebox.showerror("Erro ao salvar PDF", f"Não foi possível salvar o PDF:\n{e}")

# --- Interface gráfica ---
root = tk.Tk()
root.title("Gerador de Etiquetas Chesiquimica")

tk.Label(root, text="Código do Produto:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
entry_id = tk.Entry(root, width=30)
entry_id.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Quantidade (texto na etiqueta):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
entry_quantidade = tk.Entry(root, width=30)
entry_quantidade.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Lote:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
entry_lote = tk.Entry(root, width=30)
entry_lote.grid(row=2, column=1, padx=5, pady=5)

# Validação para campos de data (só números e '/')
vcmd = (root.register(validar_data), "%P")

tk.Label(root, text="Data de fabricação:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
entry_data_chegada = tk.Entry(root, width=30, validate="key", validatecommand=vcmd)
entry_data_chegada.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Data de Vencimento:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
entry_data_validade = tk.Entry(root, width=30, validate="key", validatecommand=vcmd)
entry_data_validade.grid(row=4, column=1, padx=5, pady=5)

tamanho_var = tk.StringVar(value="100x96")

frame_tamanhos = tk.Frame(root)
frame_tamanhos.grid(row=5, column=0, columnspan=2, pady=5)

rb1 = tk.Radiobutton(frame_tamanhos, text="100 x 96 mm", variable=tamanho_var, value="100x96")
rb2 = tk.Radiobutton(frame_tamanhos, text="120 x 96 mm", variable=tamanho_var, value="120x96")
rb3 = tk.Radiobutton(frame_tamanhos, text="Personalizado", variable=tamanho_var, value="personalizado")

rb1.grid(row=0, column=1, padx=5)
rb2.grid(row=0, column=0, padx=5)
rb3.grid(row=0, column=2, padx=5)

tk.Label(root, text="Largura (mm):").grid(row=6, column=0, sticky="e", padx=5, pady=5)
entry_largura = tk.Entry(root, width=10, state="disabled")
entry_largura.grid(row=6, column=1, sticky="w", padx=5, pady=5)

tk.Label(root, text="Altura (mm):").grid(row=7, column=0, sticky="e", padx=5, pady=5)
entry_altura = tk.Entry(root, width=10, state="disabled")
entry_altura.grid(row=7, column=1, sticky="w", padx=5, pady=5)

def toggle_custom_size():
    if tamanho_var.get() == "personalizado":
        entry_largura.config(state="normal")
        entry_altura.config(state="normal")
    else:
        entry_largura.config(state="disabled")
        entry_altura.config(state="disabled")

tamanho_var.trace_add("write", lambda *args: toggle_custom_size())

btn_gerar_pdf = tk.Button(root, text="Gerar Etiqueta PDF", command=gerar_pdf_action, width=20)
btn_gerar_pdf.grid(row=8, column=0, columnspan=2, pady=10)

root.mainloop()
