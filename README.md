# Bot de enviar Email
 Bot que enviar email através do outlook e lança toda vez que o código é executado


# Esse comando faz a portabilidade com o outlook
outlook = win32.Dispatch('outlook.application')

# Adicione os emails que quer enviar nesse modelo acima
destinatarios = ["emailaserenviado@gmail.com"]

# Caso não vá enviar imagens no email, remova essas linhas
imagem = r"C:\Seu\caminho\para\imagem\e o\pix.jpg"
email.Attachments.Add(imagem)

