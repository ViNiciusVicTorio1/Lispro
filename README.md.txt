# 🧩 Lispro - Exportador de Dados DWG para Excel

Aplicativo desktop que automatiza a criação de arquivos `.lsp` para AutoCAD e gera uma planilha Excel com textos, atributos, áreas e entidades geométricas de arquivos `.dwg`.

## ✅ Funcionalidades
- Geração automática do `exportar.lsp`
- Exportação de textos, blocos, atributos, cotas, hachuras, linhas e áreas
- Criação do `.csv` via AutoCAD (comando `EXPORTARTUDO`)
- Conversão automática para `.xlsx` com formatação limpa

## 🖥️ Requisitos
- Python 3.8 ou superior
- Bibliotecas: `openpyxl`, `tkinter`
- AutoCAD 2007 ou superior

## ▶️ Como usar

```bash
pip install openpyxl pyinstaller
pyinstaller --onefile --noconsole --icon=lispro_icon.ico --name=Lispro gerador_exportador_unificado.py
