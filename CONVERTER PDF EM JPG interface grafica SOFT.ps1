Add-Type -AssemblyName System.Windows.Forms

# Variáveis globais para armazenar caminhos
$script:pdfPath = ""
$script:outputFolder = ""

# Função para selecionar o arquivo PDF
function Select-PDFFile {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Selecione o arquivo PDF"
    $fileDialog.Filter = "Arquivos PDF (*.pdf)|*.pdf"

    $result = $fileDialog.ShowDialog()

    if ($result -eq 'OK') {
        $script:pdfPath = $fileDialog.FileName
        $pdfTextBox.Text = $script:pdfPath
    }
}

# Função para selecionar o diretório de saída das imagens JPEG
function Select-OutputFolder {
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Selecione a pasta para salvar as imagens JPEG"
    $folderDialog.RootFolder = [System.Environment+SpecialFolder]::Desktop

    $result = $folderDialog.ShowDialog()

    if ($result -eq 'OK') {
        $script:outputFolder = $folderDialog.SelectedPath
        $outputFolderTextBox.Text = $script:outputFolder
    }
}

# Função para iniciar a conversão
function Convert-PDFToJPEG {
    # Caminho para o Ghostscript
    $ghostscriptPath = "C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe"

    # Verifica se o caminho do PDF foi selecionado
    if ([string]::IsNullOrEmpty($script:pdfPath)) {
        Write-Host "Selecione um arquivo PDF primeiro."
        return
    }

    # Verifica se o diretório de saída foi selecionado
    if ([string]::IsNullOrEmpty($script:outputFolder)) {
        Write-Host "Selecione um diretório de saída para as imagens JPEG."
        return
    }

    # Define o nome base para as imagens JPEG de saída
    $outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($script:pdfPath)

    # Comando para executar a conversão usando Ghostscript
    & $ghostscriptPath -sDEVICE=jpeg -r300 -o "$script:outputFolder\$outputBaseName-%03d.jpg" $script:pdfPath

    Write-Host "Conversão concluída. Imagens JPEG geradas em: $script:outputFolder"
}

# Criar janela
$form = New-Object System.Windows.Forms.Form
$form.Text = "Conversor PDF para JPEG"
$form.Width = 600
$form.Height = 250
$form.StartPosition = "CenterScreen"

# Botão para selecionar arquivo PDF
$selectPDFButton = New-Object System.Windows.Forms.Button
$selectPDFButton.Location = New-Object System.Drawing.Point(50, 30)
$selectPDFButton.Size = New-Object System.Drawing.Size(150, 30)
$selectPDFButton.Text = "Selecionar PDF"
$selectPDFButton.Add_Click({ Select-PDFFile })
$form.Controls.Add($selectPDFButton)

# Textbox para exibir o caminho do arquivo PDF selecionado
$pdfTextBox = New-Object System.Windows.Forms.TextBox
$pdfTextBox.Location = New-Object System.Drawing.Point(210, 30)
$pdfTextBox.Size = New-Object System.Drawing.Size(300, 30)
$pdfTextBox.ReadOnly = $true
$form.Controls.Add($pdfTextBox)

# Botão para selecionar diretório de saída
$selectOutputButton = New-Object System.Windows.Forms.Button
$selectOutputButton.Location = New-Object System.Drawing.Point(50, 70)
$selectOutputButton.Size = New-Object System.Drawing.Size(200, 30)
$selectOutputButton.Text = "Selecionar Diretório de Saída"
$selectOutputButton.Add_Click({ Select-OutputFolder })
$form.Controls.Add($selectOutputButton)

# Textbox para exibir o diretório de saída selecionado
$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Location = New-Object System.Drawing.Point(260, 70)
$outputFolderTextBox.Size = New-Object System.Drawing.Size(250, 30)
$outputFolderTextBox.ReadOnly = $true
$form.Controls.Add($outputFolderTextBox)

# Botão para iniciar a conversão
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Location = New-Object System.Drawing.Point(50, 110)
$convertButton.Size = New-Object System.Drawing.Size(200, 30)
$convertButton.Text = "Converter PDF para JPEG"
$convertButton.Add_Click({ Convert-PDFToJPEG })
$form.Controls.Add($convertButton)

# Mostrar a janela
$form.ShowDialog()
$form.Dispose()
# Função para iniciar a conversão
function Convert-PDFToJPEG {
    # Caminho para o Ghostscript
    $ghostscriptPath = "C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe"

    # Verifica se o caminho do PDF foi selecionado
    if ([string]::IsNullOrEmpty($script:pdfPath)) {
        Write-Host "Selecione um arquivo PDF primeiro."
        return
    }

    # Verifica se o diretório de saída foi selecionado
    if ([string]::IsNullOrEmpty($script:outputFolder)) {
        Write-Host "Selecione um diretório de saída para as imagens JPEG."
        return
    }

    # Define o nome base para as imagens JPEG de saída
    $outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($script:pdfPath)

    # Comando para executar a conversão usando Ghostscript
    & $ghostscriptPath -sDEVICE=jpeg -r300 -o "$script:outputFolder\$outputBaseName-%03d.jpg" $script:pdfPath

    # Mensagem de confirmação após a conversão
    [System.Windows.Forms.MessageBox]::Show("Conversão concluída. Imagens JPEG geradas em: $script:outputFolder", "Conversão Concluída", "OK", [System.Windows.Forms.MessageBoxIcon]::Information)
}
