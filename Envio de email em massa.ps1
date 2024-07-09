# Configurações do email
$EmailFrom = "quemuelrodriguesdesousa@hotmail.com"
$Subject = "Curriculum"
$Body = "Analista de TI"
$SMTPServer = "smtp.outlook.com"
$SMTPPort = 587
$SMTPUsername = "quemuelrodriguesdesousa@hotmail.com"
$SMTPPassword = "Raiane@123*"

# Lista de destinatários (emails em massa)
$EmailTo = @(
    "curriculos@hospitalpremium.com.br",
    "gotpagany@gmail.com",
    "gotpagany@gmail.com",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br"
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br",
    "curriculos@hospitalpremium.com.br"

    # Adicione mais emails conforme necessário
)

# Caminho para o arquivo que será anexado
$AttachmentPath = "C:\Users\FbwTech\Downloads\Curriculo.pdf"

# Configuração do cliente SMTP
$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUsername, $SMTPPassword)

# Loop para enviar email para cada destinatário na lista
foreach ($recipient in $EmailTo) {
    # Criando o objeto de mensagem
    $Message = New-Object System.Net.Mail.MailMessage
    $Message.From = $EmailFrom
    $Message.To.Add($recipient)
    $Message.Subject = $Subject
    $Message.Body = $Body

    # Anexando o arquivo
    $Attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
    $Message.Attachments.Add($Attachment)

    # Enviando o email
    $SMTPClient.Send($Message)
    
    Write-Host "Email enviado para: $recipient"

    # Limpando o objeto de mensagem para o próximo loop
    $Message.Dispose()
}
