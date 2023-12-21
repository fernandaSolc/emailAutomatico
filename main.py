import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)
email.To = "fernandaoliv.8272@gmail.com"
email.Subject = "Olá, mundo!"
email.HTMLBody = """<p>Teste para email automatico em python
</p>
<p> Abs, </p>
<p>Fernanda</p>
"""
email.Send()
print("Email enviado")