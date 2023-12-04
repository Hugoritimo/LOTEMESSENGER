import win32com.client as win32
import pandas as pd
import os
import PySimpleGUI as sg

# Reading the database from a CSV file
db = pd.read_csv("destinos_anexo_dbteste.csv", delimiter=";", encoding='utf8')

# Extracting the necessary data from the database
emails = db['email']
nomes = db['nome']
arquivos = db['arquivo']

# DataFrame for the error log
log_df = pd.DataFrame(columns=['Email', 'Arquivo', 'Erro'])

# Mensagem pré-definida
mensagem_predefinida = """
<p style="font-size: 16px; color: #333;">Boa tarde, {nome}.</p>
<br>
<p style="font-size: 12px; color: #999;">Assinatura e informações adicionais aqui...</p>
"""

def enviar_email(email_destino, nome_destino, mensagem_editada=''):
    try:
        # Se uma mensagem editada foi fornecida, use-a em vez da padrão
        email_content = mensagem_editada if mensagem_editada else mensagem_predefinida.format(nome=nome_destino)

        # Creating and sending the email with the attached file
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.Subject = "BOLETIM DE MEDIÇÃO DO FORNECEDOR"
        email.HTMLBody = email_content
        
        # Adicionando os destinatários
        email.To = ';'.join(email_destino)

        email.Send()

        print(f"Email Sent to {email_destino}")

    except Exception as e:
        # Adding a log entry if an exception occurs
        log_df.loc[len(log_df)] = [email_destino, '', str(e)]
        print(f"Error sending email to {email_destino}: {str(e)}")

# Creating the GUI
layout = [
    [sg.Text("Select the recipients:")],
    [sg.Listbox(emails, size=(30, len(emails)), key='-LIST-', enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED)],
    [sg.Text("Edit Email Body:"), sg.Multiline(default_text=mensagem_predefinida, size=(60, 10), key='-EMAIL_BODY-')],
    [sg.Button('Send Email'), sg.Button('Exit')]
]

window = sg.Window('Send Email').Layout(layout)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    elif event == 'Send Email':
        selected_indices = values.get('-LIST-', [])
        for selected_index in selected_indices:
            if selected_index in range(len(emails)):
                selected_index = selected_index
                email_destino = emails.iloc[selected_index].split(';')
                nome_destino = nomes.iloc[selected_index]
                mensagem_editada = values['-EMAIL_BODY-']
                enviar_email(email_destino, nome_destino, mensagem_editada)

# Exporting the error log file
log_df.to_excel('error_log.xlsx', index=False)

window.close()
