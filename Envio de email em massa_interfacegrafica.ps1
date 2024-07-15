Add-Type -AssemblyName System.Windows.Forms

# Cria o formulário
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enviar Email"
$form.Size = New-Object System.Drawing.Size(400, 300)
$form.StartPosition = "CenterScreen"

# Labels e Textboxes para os parâmetros do email
$emailFromLabel = New-Object System.Windows.Forms.Label
$emailFromLabel.Location = New-Object System.Drawing.Point(10, 20)
$emailFromLabel.Size = New-Object System.Drawing.Size(120, 20)
$emailFromLabel.Text = "Email Remetente:"
$form.Controls.Add($emailFromLabel)

$emailFromTextBox = New-Object System.Windows.Forms.TextBox
$emailFromTextBox.Location = New-Object System.Drawing.Point(140, 20)
$emailFromTextBox.Size = New-Object System.Drawing.Size(240, 20)
$emailFromTextBox.Text = "quemuelrodriguesdesousa@hotmail.com"
$form.Controls.Add($emailFromTextBox)

$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Location = New-Object System.Drawing.Point(10, 50)
$subjectLabel.Size = New-Object System.Drawing.Size(120, 20)
$subjectLabel.Text = "Assunto:"
$form.Controls.Add($subjectLabel)

$subjectTextBox = New-Object System.Windows.Forms.TextBox
$subjectTextBox.Location = New-Object System.Drawing.Point(140, 50)
$subjectTextBox.Size = New-Object System.Drawing.Size(240, 20)
$subjectTextBox.Text = "Curriculum"
$form.Controls.Add($subjectTextBox)

$bodyLabel = New-Object System.Windows.Forms.Label
$bodyLabel.Location = New-Object System.Drawing.Point(10, 80)
$bodyLabel.Size = New-Object System.Drawing.Size(120, 20)
$bodyLabel.Text = "Corpo do Email:"
$form.Controls.Add($bodyLabel)

$bodyTextBox = New-Object System.Windows.Forms.TextBox
$bodyTextBox.Location = New-Object System.Drawing.Point(140, 80)
$bodyTextBox.Size = New-Object System.Drawing.Size(240, 100)
$bodyTextBox.Multiline = $true
$bodyTextBox.Text = "Analista de TI"
$form.Controls.Add($bodyTextBox)

$attachmentLabel = New-Object System.Windows.Forms.Label
$attachmentLabel.Location = New-Object System.Drawing.Point(10, 200)
$attachmentLabel.Size = New-Object System.Drawing.Size(120, 20)
$attachmentLabel.Text = "Caminho do Anexo:"
$form.Controls.Add($attachmentLabel)

$attachmentTextBox = New-Object System.Windows.Forms.TextBox
$attachmentTextBox.Location = New-Object System.Drawing.Point(140, 200)
$attachmentTextBox.Size = New-Object System.Drawing.Size(240, 20)
$attachmentTextBox.Text = "C:\Caminho\Para\Seu\Anexo.pdf"
$form.Controls.Add($attachmentTextBox)

$sendButton = New-Object System.Windows.Forms.Button
$sendButton.Location = New-Object System.Drawing.Point(140, 230)
$sendButton.Size = New-Object System.Drawing.Size(100, 30)
$sendButton.Text = "Enviar"
$sendButton.Add_Click({
    # Configurações do email
    $EmailFrom = $emailFromTextBox.Text
    $Subject = $subjectTextBox.Text
    $Body = $bodyTextBox.Text

    # Lista de destinatários (emails em massa)
    $EmailTo = @(
        "curriculos@hospitalpremium.com.br",
        "gotpagany@gmail.com",
        "gotpagany@gmail.com",
        "gotpagany@gmail.com"
    )

    # Caminho para o arquivo que será anexado
    $AttachmentPath = $attachmentTextBox.Text

    # Configuração do cliente SMTP
    $SMTPServer = "smtp.outlook.com"
    $SMTPPort = 587
    $SMTPUsername = "quemuelrodriguesdesousa@hotmail.com"
    $SMTPPassword = "senha*"

    $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUsername, $SMTPPassword)

    foreach ($recipient in $EmailTo) {
        # Criando o objeto de mensagem
        $Message = New-Object System.Net.Mail.MailMessage
        $Message.From = $EmailFrom
        $Message.To.Add($recipient)
        $Message.Subject = $Subject
        $Message.Body = $Body

        # Anexando o arquivo, se houver
        if (Test-Path $AttachmentPath) {
            $Attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
            $Message.Attachments.Add($Attachment)
        }

        # Enviando o email
        $SMTPClient.Send($Message)

        Write-Host "Email enviado para: $recipient"

        # Limpando o objeto de mensagem para o próximo loop
        $Message.Dispose()
    }
})
$form.Controls.Add($sendButton)

# Exibe o formulário
$form.ShowDialog()
