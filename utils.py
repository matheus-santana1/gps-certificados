from tkinter import messagebox
from PyPDF2 import PdfMerger
import comtypes.client
import os
import pythoncom
import config
import win32process

def unificar_pdfs_da_pasta(pasta, nome_pessoa, frente_verso):
    pasta_destino = os.path.dirname(pasta)
    merger_geral = PdfMerger()
    merger_fv = PdfMerger()
    tem_geral = False
    tem_fv = False
    try:
        arquivos = [f for f in os.listdir(pasta) if f.endswith(".pdf") and not f.startswith("DOC_UNIFICADO")]
        arquivos.sort()

        for nome_arquivo in arquivos:
            caminho_completo = os.path.join(pasta, nome_arquivo)
            e_frente_verso = any(termo.upper() in nome_arquivo.upper() for termo in frente_verso)
            if e_frente_verso:
                merger_fv.append(caminho_completo)
                tem_fv = True
            else:
                merger_geral.append(caminho_completo)
                tem_geral = True

        if tem_geral:
            caminho_final_geral = os.path.join(pasta_destino, f"DOC_UNIFICADO {nome_pessoa}.pdf")
            merger_geral.write(caminho_final_geral)
            merger_geral.close()

        if tem_fv:
            caminho_final_fv = os.path.join(pasta_destino, f"FRENTE_VERSO {nome_pessoa}.pdf")
            merger_fv.write(caminho_final_fv)
            merger_fv.close()
    
        extensoes_apagar = (".xlsx", ".docx", ".pptx")
        for arquivo in os.listdir(pasta_destino):
            if arquivo.lower().endswith(extensoes_apagar):
                caminho_para_apagar = os.path.join(pasta_destino, arquivo)
                try:
                    os.remove(caminho_para_apagar)
                except Exception as e_del:
                    print(f"Não foi possível apagar o arquivo {arquivo}: {e_del}")
        return True

    except Exception as e:
        print(f"Erro ao mesclar PDF para {nome_pessoa}: {e}")
        return False

def get_word_pid(word_app):
    import time
    import win32gui
    """
    Finds the process ID (PID) of a given Word Application COM object.
    """
    # 1. Set a unique, temporary caption to identify the window
    unique_caption = f"Python_COM_Identifier_{int(time.time())}"
    word_app.Visible = True # Must be visible for a main window handle to exist and be detectable
    word_app.Caption = unique_caption
    
    # 2. Find the window handle (hWnd) by its unique caption
    hWnd = win32gui.FindWindow(None, unique_caption)
    
    if hWnd:
        # 3. Get the PID from the window handle
        _, pid = win32process.GetWindowThreadProcessId(hWnd)
        
        # Optionally, hide the window again and reset the caption
        word_app.Visible = False
        word_app.Caption = ""
        return pid
    else:
        word_app.Visible = False
        word_app.Caption = ""
        return None

def converter_pasta_pdf(diretorio_raiz, app_interface, config_excel={}):
    """
    diretorio_raiz: Geralmente "saida"
    Percorre: saida -> nome_pessoa -> arquivos
    Salva em: saida -> nome_pessoa -> pdf -> arquivo.pdf
    """
    pythoncom.CoInitialize()
    
    app_ppt = None
    app_excel = None
    app_word = None

    try:
        # 1. Lista todas as pastas de pessoas dentro de 'saida'
        if not os.path.exists(diretorio_raiz):
            app_interface.log(f"⚠️ Pasta '{diretorio_raiz}' não encontrada.")
            return

        pastas_pessoas = [p for p in os.listdir(diretorio_raiz) 
                          if os.path.isdir(os.path.join(diretorio_raiz, p))]

        if not pastas_pessoas:
            app_interface.log("⚠️ Nenhuma pasta de pessoa encontrada dentro de 'saida'.")
            return

        for pessoa in pastas_pessoas:
            caminho_pessoa = os.path.join(os.path.abspath(diretorio_raiz), pessoa)
            pasta_pdf_destino = os.path.join(caminho_pessoa, "pdf")

            # Cria a pasta 'pdf' interna se não existir
            if not os.path.exists(pasta_pdf_destino):
                os.makedirs(pasta_pdf_destino)

            # 2. Lista arquivos da pessoa
            arquivos = [f for f in os.listdir(caminho_pessoa) if f.endswith(('.pptx', '.xlsx', '.docx'))]
            
            if arquivos:
                app_interface.log(f"📂 Processando pasta: {pessoa}")
            
            for nome_arquivo in arquivos:
                caminho_entrada = os.path.join(caminho_pessoa, nome_arquivo)
                nome_base = os.path.splitext(nome_arquivo)[0]
                caminho_saida_pdf = os.path.join(pasta_pdf_destino, f"{nome_base}.pdf")

                # --- LÓGICA PARA WORD ---
                if nome_arquivo.endswith(".docx"):
                    if not app_word:
                        app_word = comtypes.client.CreateObject("Word.Application")
                        try:
                            pid = get_word_pid(app_word)
                            if pid:
                                config.pids_automacao.append(pid)
                            else:
                                raise Exception("Não foi possível capturar o PID do Word.")
                        except Exception as e:
                            print(e)
                    doc = app_word.Documents.Open(caminho_entrada, ReadOnly=True)
                    doc.ExportAsFixedFormat(caminho_saida_pdf, 17)
                    doc.Close(SaveChanges=False)
                    app_interface.log(f"   ✔️ DOCX -> PDF: {nome_arquivo}")

                # --- LÓGICA PARA POWERPOINT ---
                elif nome_arquivo.endswith(".pptx"):
                    if not app_ppt:
                        app_ppt = comtypes.client.CreateObject("Powerpoint.Application")
                        _, pid = win32process.GetWindowThreadProcessId(app_ppt.Hwnd)
                        config.pids_automacao.append(pid)

                    pres = app_ppt.Presentations.Open(caminho_entrada, WithWindow=False)
                    pres.SaveAs(caminho_saida_pdf, 32)
                    pres.Close()
                    app_interface.log(f"   ✔️ PPTX -> PDF: {nome_arquivo}")

                # --- LÓGICA PARA EXCEL ---
                elif nome_arquivo.endswith(".xlsx"):
                    if not app_excel:
                        app_excel = comtypes.client.CreateObject("Excel.Application")
                        app_excel.Visible = False
                        app_excel.DisplayAlerts = False
                        _, pid = win32process.GetWindowThreadProcessId(app_excel.Hwnd)
                        config.pids_automacao.append(pid)
                    
                    wb = app_excel.Workbooks.Open(caminho_entrada)
                    
                    config_pessoa = config_excel.get(pessoa, {}) if config_excel else {}
                    abas_permitidas = config_pessoa.get(nome_arquivo, [])
                    if abas_permitidas:
                        for sheet in wb.Sheets:
                            sheet.Visible = -1 if sheet.Name in abas_permitidas else 0
                    
                    wb.ExportAsFixedFormat(0, caminho_saida_pdf, Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False)
                    wb.Close(SaveChanges=False)
                    app_interface.log(f"   ✔️ XLSX -> PDF: {nome_arquivo}")

            if os.path.exists(pasta_pdf_destino):
                sucesso = unificar_pdfs_da_pasta(pasta_pdf_destino, pessoa, config_excel.get("frente_verso", []))
            if sucesso:
                app_interface.log(f"   ✅ PDF Unificado gerado para {pessoa}")

        app_interface.log(f"\n✅ Todos os arquivos foram convertidos!")
        messagebox.showinfo("Sucesso", "Conversão das pastas concluída!")

    except Exception as e:
        app_interface.log(f"❌ Erro na conversão: {str(e)}")
    
    finally:
        if app_word: 
            try: app_word.Quit()
            except: pass
        if app_ppt: 
            try: app_ppt.Quit()
            except: pass
        if app_excel: 
            try: app_excel.Quit()
            except: pass
        pythoncom.CoUninitialize()