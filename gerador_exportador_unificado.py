
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from openpyxl import Workbook

def gerar_excel():
    csv_path = "C:/temp/exportado.csv"
    xlsx_path = "C:/temp/exportado.xlsx"

    if not os.path.exists(csv_path):
        messagebox.showerror("Erro", "Arquivo exportado.csv não encontrado em C:/temp.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Exportado"

    with open(csv_path, "r", encoding="utf-8") as f:
        for line in f:
            row = [cell.strip() for cell in line.strip().split(",")]
            ws.append(row)

    wb.save(xlsx_path)
    return xlsx_path

def gerar_lisp(dwg_path):
    exportar_lsp = """(defun c:EXPORTARTUDO ()
  (vl-load-com)
  (setq arquivo (open "C:\\temp\\exportado.csv" "w"))
  (write-line "Tipo,Subtipo,Valor,Layer,X,Y,InfoExtra" arquivo)

  ;; Textos
  (setq ss (ssget '((0 . "TEXT,MTEXT"))))
  (if ss (progn (setq i 0)
    (while (< i (sslength ss))
      (setq ent (ssname ss i))
      (setq dados (entget ent))
      (setq valor (cdr (assoc 1 dados)))
      (setq layer (cdr (assoc 8 dados)))
      (setq ponto (cdr (assoc 10 dados)))
      (write-line (strcat "Texto,," valor "," layer "," 
                  (rtos (car ponto) 2 2) "," 
                  (rtos (cadr ponto) 2 2) ",") arquivo)
      (setq i (1+ i)))))

  ;; Blocos e atributos
  (setq ss (ssget '((0 . "INSERT"))))
  (if ss (progn (setq i 0)
    (while (< i (sslength ss))
      (setq ent (ssname ss i))
      (setq dados (entget ent))
      (setq nome (cdr (assoc 2 dados)))
      (setq layer (cdr (assoc 8 dados)))
      (setq ponto (cdr (assoc 10 dados)))
      (write-line (strcat "Bloco," nome ",," layer "," 
                  (rtos (car ponto) 2 2) "," 
                  (rtos (cadr ponto) 2 2) ",") arquivo)
      (setq ent2 ent)
      (while (setq ent2 (entnext ent2))
        (setq dados2 (entget ent2))
        (if (= (cdr (assoc 0 dados2)) "ATTRIB")
          (progn
            (setq valor (cdr (assoc 1 dados2)))
            (setq tag (cdr (assoc 2 dados2)))
            (write-line (strcat "Atributo," tag "," valor "," layer "," 
                        (rtos (car ponto) 2 2) "," 
                        (rtos (cadr ponto) 2 2) ",") arquivo))))
      (setq i (1+ i)))))

  ;; Outras entidades geométricas
  (setq tipos '("LINE" "CIRCLE" "ARC" "ELLIPSE" "SPLINE" "POLYLINE" "LWPOLYLINE" "HATCH" "DIMENSION" "REGION" "SOLID"))
  (foreach tipo tipos
    (setq ss (ssget (list (cons 0 tipo))))
    (if ss (progn (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq dados (entget ent))
        (setq layer (cdr (assoc 8 dados)))
        (setq ponto (cdr (assoc 10 dados)))
        (setq infoextra "")
        (cond
          ((= tipo "CIRCLE") (setq infoextra (strcat "Raio=" (rtos (cdr (assoc 40 dados)) 2 2))))
          ((= tipo "ARC") (setq infoextra (strcat "Raio=" (rtos (cdr (assoc 40 dados)) 2 2))))
          ((= tipo "ELLIPSE") (setq infoextra (strcat "RazãoEixos=" (rtos (cdr (assoc 40 dados)) 2 2))))
          ((= tipo "LWPOLYLINE")
            (setq area (if (= (logand (cdr (assoc 70 dados)) 1) 1)
                           (rtos (vlax-curve-getarea ent) 2 2) "Aberta"))
            (setq infoextra (strcat "Área=" area)))
          ((= tipo "HATCH") (setq infoextra (strcat "Padrão=" (cdr (assoc 2 dados)))))
          ((= tipo "DIMENSION") (setq infoextra (strcat "Medida=" (rtos (cdr (assoc 42 dados)) 2 2))))
        )
        (write-line (strcat tipo ",,," layer "," 
                    (rtos (car ponto) 2 2) "," 
                    (rtos (cadr ponto) 2 2) "," infoextra) arquivo)
        (setq i (1+ i))))))

  (close arquivo)
  (princ "\nExportação COMPLETA para C:\\temp\\exportado.csv")
  (princ))"""

    with open("exportar.lsp", "w", encoding="utf-8") as f:
        f.write(exportar_lsp)

    return gerar_excel()

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(filetypes=[("DWG Files", "*.dwg")])
    if caminho:
        gerar_lisp(caminho)
        messagebox.showinfo("Sucesso", "Arquivo exportar.lsp criado e exportado.xlsx gerado com sucesso em C:/temp")

root = tk.Tk()
root.title("Lispro Unificado")
root.geometry("400x200")
root.resizable(False, False)

label = tk.Label(root, text="Clique no botão ou arraste um arquivo .dwg aqui", pady=20)
label.pack()

botao = tk.Button(root, text="Abrir arquivo .DWG", command=selecionar_arquivo)
botao.pack()

root.mainloop()
