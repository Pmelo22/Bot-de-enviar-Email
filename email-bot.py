import win32com.client as win32

try:
    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # lista de emails
    destinatarios = [
        
    ]

    # configurar as informações do seu e-mail
    email.To = ""  # Deixe o campo 'To' vazio
    email.Subject = "escreva aqui o Assunto do email"
    email.HTMLBody = f"""
    Aqui 
    nesse campo
    escreva o que deseja em HTML
    """

    # Adiciona os destinatários no campo BCC
    email.BCC = "; ".join(destinatarios)

    # caminho para a imagem(caso o email enviado não vá usar imagem, remova essas linhas abaixo)
    imagem = r"C:\Seu\caminho\para\imagem\e o\pix.jpg"
    email.Attachments.Add(imagem)

    email.Send()
    print("Email Enviado")
except Exception as e:
    print(f"Erro: {e}")
