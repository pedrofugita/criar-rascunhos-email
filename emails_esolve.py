import win32com.client as win32
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox

data_atual = datetime.today()
dias_ate_segunda = (0-data_atual.weekday())%7
proxima_segunda = data_atual + timedelta(days=dias_ate_segunda)

lista_destinatarios = ["michelle.santos@embraer.com.br",
                       ["evandro.creste@embraer.com.br", "joao.zanatto@embraer.com.br"],
                       "she_engenharia@embraer.onmicrosoft.com",
                       "EPM-Bot@embraer.onmicrosoft.com"]

for destinatario in lista_destinatarios:
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    
    if isinstance(destinatario, list):
        destinatarios_grupo = ";".join(destinatario)
        email.To = destinatarios_grupo
    else:
        email.To = destinatario

    email.Subject = f"Ações e Documentos - Atrasados e Pendentes - {data_atual.strftime('%d/%m/%Y')}"
    
    email.HTMLBody = f"""
    <p>Bom dia!</p>
    <br>
    <p>Constam itens de E-solve com status <strong>ATRASADO</strong> no dia {data_atual.strftime('%d/%m/%Y')}.</p>
    <br>
    <p><strong>Ações:</strong></p>
    <ul>
    <li>8D - <strong>Zero atraso</strong></li>
    <li>CA não conformidade de Produto ou Processo - <strong>Zero atraso</strong></li>
    <li>PV oportunidade de melhoria - <strong>Zero atraso</strong></li>
    </ul>
    <p><strong>Documentos:</strong></p>
    <ul>
    <li>8D - <strong>Zero atraso</strong></li>
    <li>CA não conformidade de Produto ou Processo - <strong>Zero atraso</strong></li>
    <li>PV oportunidade de melhoria - <strong>Zero atraso</strong></li>
    </ul>
    <p>Ações com <strong>prazo até {proxima_segunda.strftime('%d/%m/%Y')}</strong>:</p>
    <br>
    <br>
    <p>Atenciosamente,</p>
    <br>
    <p style="font-family: Trebuchet MS; color: #002060"><strong>Pedro Henrique Fugita Bóis</strong></p>
    <p style="font-family: Trebuchet MS;">Time de Engenharia de Manufatura</p>
    <br>
    <p style="font-family: Trebuchet MS;">+55 17 99635-5383</p>
    <p style="font-family: Trebuchet MS;">pedro.bois@embraer.com.br</p>
    <p style="font-family: Trebuchet MS;">Embraer / Botucatu</p>
    <p style="font-family: Trebuchet MS;">embraer.com</p>
    """

    email.Save()

messagebox.showinfo("Concluido", "Os rascunhos foram criados no seu Outlook.")