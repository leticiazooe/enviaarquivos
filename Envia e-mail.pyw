import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime

def send_emails():
    try:
        df = pd.read_excel(file_path.get(), engine='openpyxl')

        smtp_server = 'smtp.office365.com'
        smtp_port = 587
        smtp_user = 'seuemail@dominio.com'
        smtp_password = 'suasenha'

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)

        log_entries = []
        total_emails = len(df)
        progress["maximum"] = total_emails
        progress["value"] = 0
        root.update_idletasks()

        assunto_template = assunto_entry.get()
        mensagem_template = email_body_text.get("1.0", tk.END)

        for index, row in df.iterrows():
            try:
                nome = row['Nome do Fornecedor']
                email = row['E-mail']

                assunto = assunto_template.format(nome=nome)
                mensagem = mensagem_template.format(nome=nome)

                msg = MIMEMultipart()
                msg['From'] = smtp_user
                msg['To'] = email
                msg['Subject'] = assunto
                msg.attach(MIMEText(mensagem, 'plain'))

                for path in attachment_paths:
                    with open(path, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(path)}"')
                        msg.attach(part)

                server.sendmail(smtp_user, email, msg.as_string())
                status_text.insert(tk.END, f"✅ E-mail enviado com sucesso para {nome} ({email})\n")
                log_entries.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Sucesso - {nome} ({email})")

            except Exception as e:
                status_text.insert(tk.END, f"❌ Erro ao enviar e-mail para {nome} ({email}): {e}\n")
                log_entries.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Erro - {nome} ({email}) - {e}")

            progress["value"] += 1
            root.update_idletasks()
            status_text.see(tk.END)

        server.quit()
        status_text.insert(tk.END, "\n✅ Todos os e-mails foram processados.\n")

        log_folder = log_path.get()
        if not log_folder:
            log_folder = os.path.join(os.path.expanduser("~"), "Documents", "envio automatico de e-mails")
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
        log_file = os.path.join(log_folder, f"log_{datetime.now().strftime('%Y%m%d')}.txt")
        with open(log_file, 'w') as f:
            f.write("\n".join(log_entries))
        status_text.insert(tk.END, f"\n📄 Log salvo em: {log_file}\n")
        status_text.see(tk.END)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro geral no processo: {e}")
        status_text.insert(tk.END, f"\n❌ Erro geral no processo: {e}\n")
        status_text.see(tk.END)

def select_file():
    file_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))

def clear_supplier_file():
    file_path.set("")
    status_text.insert(tk.END, "📄 Lista de destinatários removida.\n")
    status_text.see(tk.END)

def add_attachment():
    paths = filedialog.askopenfilenames(filetypes=[
        ("Documentos permitidos", "*.doc *.docx *.pdf *.xls *.xlsx *.ppt *.pptx *.pbix")
    ])
    for path in paths:
        if path not in attachment_paths:
            attachment_paths.append(path)
            attachments_list.insert(tk.END, path)

def remove_selected_attachment():
    selected = attachments_list.curselection()
    for index in reversed(selected):
        attachment_paths.pop(index)
        attachments_list.delete(index)

def select_log_folder():
    log_path.set(filedialog.askdirectory())
    status_text.insert(tk.END, f"📂 Pasta de log selecionada: {log_path.get()}\n")
    status_text.see(tk.END)

def view_log_folder():
    log_folder = log_path.get()
    if not log_folder:
        log_folder = os.path.join(os.path.expanduser("~"), "Documents", "envio automatico de e-mails")
    os.startfile(log_folder)

# Interface
root = tk.Tk()
root.title("Envio de E-mails para destinatários")
root.geometry("1000x950")
root.configure(bg="#f0f0f0")

style = ttk.Style()
style.theme_use("default")
style.configure("Rounded.TButton",
                padding=6,
                relief="flat",
                background="#000080",
                foreground="white",
                font=('Arial', 10, 'bold'))
style.map("Rounded.TButton", background=[('active', '#0000a0')])

file_path = tk.StringVar()
attachment_paths = []
log_path = tk.StringVar()

frame = tk.Frame(root, bg="#f0f0f0")
frame.pack(fill="both", expand=True, padx=20, pady=10)

tk.Label(frame, text="Selecione a planilha Excel com os dados dos destinatários:", bg="#f0f0f0", font=('Arial', 11)).pack(anchor="w")
tk.Entry(frame, textvariable=file_path, font=('Arial', 10), width=120).pack(pady=5, fill="x")

file_btn_frame = tk.Frame(frame, bg="#f0f0f0")
file_btn_frame.pack(pady=5, anchor="w")
ttk.Button(file_btn_frame, text="Selecionar Arquivo", command=select_file, style="Rounded.TButton").pack(side="left", padx=5)
ttk.Button(file_btn_frame, text="🧹 Limpar Lista de destinatários", command=clear_supplier_file, style="Rounded.TButton").pack(side="left", padx=5)

# Anexos
tk.Label(frame, text="Anexos (PDF/Word/Excel/PPT/PBIX):", bg="#f0f0f0", font=('Arial', 11)).pack(anchor="w", pady=(20, 0))
tk.Label(frame, text="Tipos aceitos: .doc, .docx, .pdf, .xls, .xlsx, .ppt, .pptx, .pbix", bg="#f0f0f0", font=('Arial', 9, 'italic')).pack(anchor="w")
attachments_list = tk.Listbox(frame, height=6, font=('Arial', 9), width=120)
attachments_list.pack(pady=2, fill="x")
btn_frame = tk.Frame(frame, bg="#f0f0f0")
btn_frame.pack(anchor="w", pady=5)
ttk.Button(btn_frame, text="+ Adicionar Anexo", command=add_attachment, style="Rounded.TButton").pack(side="left", padx=5)
ttk.Button(btn_frame, text="🗑 Remover Selecionado", command=remove_selected_attachment, style="Rounded.TButton").pack(side="left", padx=5)

tk.Label(frame, text="Assunto do E-mail (use {nome} para personalização):", bg="#f0f0f0", font=('Arial', 11)).pack(anchor="w", pady=(20, 0))
assunto_entry = tk.Entry(frame, font=('Arial', 11), width=120)
assunto_entry.insert(0, "Email teste para {nome} - por favor, desconsiderar")
assunto_entry.pack(pady=5, fill="x")

tk.Label(frame, text="Mensagem do E-mail (use {nome} para personalização):", bg="#f0f0f0", font=('Arial', 11)).pack(anchor="w", pady=(10, 0))
email_body_text = tk.Text(frame, height=12, font=('Arial', 10), wrap="word")
email_body_text.insert(tk.END, """Prezado(a), {nome},

Espero que esteja bem!

Queria compartilhar uma novidade: estou estudando Python! 
Estou me aprofundando nessa linguagem incrível e já comecei a criar alguns scripts e pequenas automações. 
É fascinante ver como algumas linhas de código podem fazer tanta coisa acontecer!

Atenciosamente,

Letícia SILVA
Analista de Desenvolvimento
""")
email_body_text.pack(fill="x", pady=5)

tk.Label(frame, text="Pasta de Log:", bg="#f0f0f0", font=('Arial', 11)).pack(anchor="w", pady=(20, 0))
tk.Entry(frame, textvariable=log_path, font=('Arial', 10), width=120).pack(pady=5, fill="x")
log_btn_frame = tk.Frame(frame, bg="#f0f0f0")
log_btn_frame.pack(pady=5, anchor="w")
ttk.Button(log_btn_frame, text="Selecionar Pasta de Log", command=select_log_folder, style="Rounded.TButton").pack(side="left", padx=5)
ttk.Button(log_btn_frame, text="📂 Visualizar Log", command=view_log_folder, style="Rounded.TButton").pack(side="left", padx=5)

# Barra de progresso
progress = ttk.Progressbar(frame, orient="horizontal", length=800, mode="determinate")
progress.pack(pady=10)

# Botão de envio
ttk.Button(frame, text="Iniciar Envio de E-mails", command=send_emails, style="Rounded.TButton").pack(pady=20)

# Área de status
status_text = tk.Text(root, height=15, font=('Courier New', 9), wrap="word")
status_text.pack(fill="both", expand=True, padx=20, pady=10)

# Rodapé
tk.Label(root, text="Programa desenvolvido por Letícia Silva.",
         bg="#f0f0f0", font=('Arial', 9, 'italic')).pack(pady=10)

root.mainloop()
