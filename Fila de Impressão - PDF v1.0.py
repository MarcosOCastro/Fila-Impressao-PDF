import os
import time  
import win32print
import win32ui
import win32con
from PIL import Image, ImageWin
import fitz  # PyMuPDF
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
from screeninfo import get_monitors

class FilaImpressaoApp:
    def __init__(self, root):
        self.root = root
        self.root.overrideredirect(True)
        
        # --- CORES ---
        self.cor_base = "#1e1e1e"      
        self.cor_lista = "#181818"     
        self.branco = "#ffffff"
        self.vermelho_hover = "#cc3333"
        
        self.largura_janela = 750
        self.altura_janela = 650
        self.centralizar_no_monitor_do_mouse()
        self.root.configure(bg=self.cor_base)

        # --- BARRA DE TÍTULO ---
        self.barra_titulo = tk.Frame(root, bg=self.cor_base, height=35)
        self.barra_titulo.pack(fill="x", side="top")
        self.barra_titulo.bind("<ButtonPress-1>", self.start_move)
        self.barra_titulo.bind("<B1-Motion>", self.do_move)

        lbl_titulo = tk.Label(self.barra_titulo, text="Fila de Impressão - PDF v1.0", 
                              bg=self.cor_base, fg="#aaaaaa", font=("Arial", 9, "bold"))
        lbl_titulo.pack(side="left", padx=15)

        self.btn_fechar = tk.Button(self.barra_titulo, text="✕", bg=self.cor_base, fg="#888888", 
                               relief="flat", command=root.quit, font=("Arial", 10),
                               activebackground=self.vermelho_hover, activeforeground="white", bd=0, padx=15)
        self.btn_fechar.pack(side="right", fill="y")
        self.btn_fechar.bind("<Enter>", lambda e: self.btn_fechar.configure(bg=self.vermelho_hover, fg="white"))
        self.btn_fechar.bind("<Leave>", lambda e: self.btn_fechar.configure(bg=self.cor_base, fg="#888888"))

        # --- ESTILOS ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=self.cor_lista, foreground=self.branco, 
                        fieldbackground=self.cor_lista, borderwidth=0, rowheight=30)
        
        style.configure("Treeview.Heading", background="#2d2d2d", foreground=self.branco, borderwidth=0, font=("Arial", 9, "bold"))
        style.map("Treeview.Heading", background=[('active', '#2d2d2d')], foreground=[('active', self.branco)])

        frame_topo = tk.Frame(root, bg=self.cor_base, padx=20, pady=20)
        frame_topo.pack(fill="x")
        
        tk.Label(frame_topo, text="Selecione a Impressora:", bg=self.cor_base, fg=self.branco, 
                 font=("Arial", 11, "bold")).pack(side="left")
        
        self.combo_impressoras = ttk.Combobox(frame_topo, values=self.get_printers(), width=40, state="readonly")
        self.combo_impressoras.pack(side="left", padx=15)
        try: self.combo_impressoras.set(win32print.GetDefaultPrinter())
        except: pass

        self.frame_central = tk.Frame(root, bg=self.cor_base, padx=20)
        self.frame_central.pack(fill="both", expand=True)

        self.lista_box = ttk.Treeview(self.frame_central, columns=("Arquivo", "Caminho"), show="headings")
        self.lista_box.heading("Arquivo", text="Nome do Arquivo")
        self.lista_box.heading("Caminho", text="Localização")
        self.lista_box.column("Arquivo", width=400)
        self.lista_box.column("Caminho", width=250)
        self.lista_box.pack(side="left", fill="both", expand=True)

        self.lista_box.tag_configure('sucesso', background='#1b4332', foreground='white')
        self.lista_box.tag_configure('erro', background='#800f2f', foreground='white')

        self.label_instrucao = tk.Label(self.lista_box, text="Arraste seus PDF's aqui\nou selecione os arquivos clicando no botão Adicionar abaixo", 
                                       fg="#777777", bg=self.cor_lista, font=("Arial", 10, "bold italic"))
        
        self.lista_box.drop_target_register(DND_FILES)
        self.lista_box.dnd_bind('<<Drop>>', self.processar_drop)

        # --- CONTADOR DE ARQUIVOS ---
        self.lbl_contador = tk.Label(root, text="Arquivos na fila: 0", bg=self.cor_base, fg="#28a745", font=("Arial", 10, "bold"))
        self.lbl_contador.pack(pady=(5, 0))

        frame_ctrl = tk.Frame(root, bg=self.cor_base, pady=10)
        frame_ctrl.pack()
        btn_estilo = {"fg": "white", "relief": "flat", "font": ("Arial", 9, "bold"), "padx": 15, "pady": 6, "cursor": "hand2"}

        tk.Button(frame_ctrl, text="+ Adicionar", bg="#0056b3", command=self.selecionar_manual, **btn_estilo).pack(side="left", padx=3)
        tk.Button(frame_ctrl, text="Remover", bg="#8b0000", command=self.remover_item, **btn_estilo).pack(side="left", padx=3)
        tk.Button(frame_ctrl, text="Limpar", bg="#444444", command=self.limpar_lista, **btn_estilo).pack(side="left", padx=3)
        tk.Button(frame_ctrl, text="↑", bg="#333333", command=lambda: self.mover(-1), **btn_estilo).pack(side="left", padx=2)
        tk.Button(frame_ctrl, text="↓", bg="#333333", command=lambda: self.mover(1), **btn_estilo).pack(side="left", padx=2)

        self.btn_print = tk.Button(root, text="INICIAR IMPRESSÃO", bg="#28a745", fg="white", 
                                   font=("Arial", 13, "bold"), relief="flat", height=2, command=self.imprimir)
        self.btn_print.pack(pady=(0, 15), padx=20, fill="x")

        self.atualizar_placeholder() # Chama pela primeira vez

        footer = tk.Label(root, text="Desenvolvido por: Marcos Castro | GitHub: marcosocastro", 
                          bg=self.cor_base, fg="#777777", font=("Arial", 9), cursor="hand2")
        footer.pack(side="bottom", pady=8)
        footer.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/marcosocastro"))

    def verificar_erro_detalhado(self, printer_name):
        try:
            phandle = win32print.OpenPrinter(printer_name)
            time.sleep(2.0) 
            info = win32print.GetPrinter(phandle, 2)
            p_status = info['Status']
            jobs = win32print.EnumJobs(phandle, 0, 50, 1)
            win32print.ClosePrinter(phandle)

            msg_erro = ""
            if p_status & win32print.PRINTER_STATUS_OFFLINE:
                msg_erro = "Impressora DESCONECTADA ou OFFLINE."
            elif p_status & win32print.PRINTER_STATUS_PAPER_OUT:
                msg_erro = "Impressora está SEM PAPEL."
            elif p_status & win32print.PRINTER_STATUS_PAPER_JAM:
                msg_erro = "Detectado ATOLAMENTO de papel."
            elif p_status & win32print.PRINTER_STATUS_ERROR:
                msg_erro = "Erro Geral de Hardware (verifique cabos/tampa)."

            if not msg_erro:
                for job in jobs:
                    status = job.get('Status', 0)
                    if status & win32print.JOB_STATUS_PAPEROUT:
                        msg_erro = "Aviso: Sem Papel no meio da impressão."
                        break
                    elif status & win32print.JOB_STATUS_OFFLINE:
                        msg_erro = "Aviso: Impressora caiu durante o processo."
                        break
                    elif status & win32print.JOB_STATUS_ERROR:
                        msg_erro = "Falha crítica no processamento do arquivo."
                        break
            return msg_erro if msg_erro else None
        except: return "Não foi possível comunicar com a impressora."

    def imprimir(self):
        printer = self.combo_impressoras.get(); items = self.lista_box.get_children()
        if not printer or not items: return
        success = 0
        teve_erro = False 

        for item in items:
            path = self.lista_box.item(item)['values'][1]
            try:
                doc = fitz.open(path); hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(printer); hdc.StartDoc(os.path.basename(path))
                for page in doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    pw, ph = hdc.GetDeviceCaps(win32con.HORZRES), hdc.GetDeviceCaps(win32con.VERTRES)
                    ir, pr = img.size[0]/img.size[1], pw/ph
                    nw, nh = (pw, int(pw/ir)) if ir > pr else (int(ph*ir), ph)
                    hdc.StartPage(); dib = ImageWin.Dib(img)
                    xo, yo = (pw-nw)//2, (ph-nh)//2
                    dib.draw(hdc.GetHandleOutput(), (xo, yo, xo+nw, yo+nh))
                    hdc.EndPage()
                hdc.EndDoc(); hdc.DeleteDC(); doc.close()

                detalhe_erro = self.verificar_erro_detalhado(printer)
                if detalhe_erro:
                    self.lista_box.item(item, tags=('erro',))
                    self.root.update()
                    messagebox.showerror("Interrupção Crítica", 
                                         f"ERRO: {detalhe_erro}\n\n"
                                         f"Arquivo que parou: {os.path.basename(path)}")
                    teve_erro = True
                    break
                else:
                    self.lista_box.item(item, tags=('sucesso',))
                    self.root.update()
                    success += 1
            except Exception:
                self.lista_box.item(item, tags=('erro',))
                self.root.update()
                messagebox.showerror("Erro de Arquivo", f"Não foi possível processar o PDF:\n{os.path.basename(path)}")
                teve_erro = True
                break
        
        if not teve_erro:
            messagebox.showinfo("Relatório", f"Processo finalizado com sucesso!\nTotal enviados: {success}")

    def centralizar_no_monitor_do_mouse(self):
        mx, my = self.root.winfo_pointerx(), self.root.winfo_pointery()
        mon = next((m for m in get_monitors() if m.x <= mx <= m.x+m.width and m.y <= my <= m.y+m.height), get_monitors()[0])
        px, py = mon.x + (mon.width//2) - (self.largura_janela//2), mon.y + (mon.height//2) - (self.altura_janela//2)
        self.root.geometry(f"{self.largura_janela}x{self.altura_janela}+{px}+{py}")

    def start_move(self, event): self.x = event.x; self.y = event.y
    def do_move(self, event):
        x, y = self.root.winfo_x() + (event.x - self.x), self.root.winfo_y() + (event.y - self.y)
        self.root.geometry(f"+{x}+{y}")

    def get_printers(self): return [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    
    def atualizar_placeholder(self):
        # Contagem de arquivos
        qtd = len(self.lista_box.get_children())
        self.lbl_contador.config(text=f"Arquivos na fila: {qtd}")
        
        if qtd == 0: 
            self.label_instrucao.place(relx=0.5, rely=0.5, anchor="center")
        else: 
            self.label_instrucao.place_forget()

    def selecionar_manual(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.add_to_list(files)
    def processar_drop(self, event):
        files = self.root.tk.splitlist(event.data); self.add_to_list(files)
    def add_to_list(self, paths):
        for p in paths:
            clean_p = os.path.normpath(p)
            if clean_p.lower().endswith(".pdf"):
                if not any(self.lista_box.item(i)['values'][1] == clean_p for i in self.lista_box.get_children()):
                    self.lista_box.insert("", "end", values=(os.path.basename(clean_p), clean_p))
        self.atualizar_placeholder()
    def remover_item(self):
        for i in self.lista_box.selection(): self.lista_box.delete(i)
        self.atualizar_placeholder()
    def limpar_lista(self):
        for i in self.lista_box.get_children(): self.lista_box.delete(i)
        self.atualizar_placeholder()
    def mover(self, d):
        for i in self.lista_box.selection():
            idx = self.lista_box.index(i); self.lista_box.move(i, '', idx + d)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = FilaImpressaoApp(root)
    root.mainloop()