Add-Type -AssemblyName System.Windows.Forms

# Cria um formulário
$form = New-Object System.Windows.Forms.Form
$form.Text = "Conversor PDF para DOCX"
$form.Size = New-Object System.Drawing.Size(500, 250)
$form.MaximizeBox = $false  # Desabilita a maximização do formulário

# Adiciona um botão para selecionar o arquivo PDF
$buttonPDF = New-Object System.Windows.Forms.Button
$buttonPDF.Location = New-Object System.Drawing.Point(50, 50)
$buttonPDF.Size = New-Object System.Drawing.Size(400, 30)
$buttonPDF.Text = "Selecionar arquivo PDF"
$buttonPDF.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Arquivos PDF (*.pdf)|*.pdf"
    $openFileDialog.Title = "Selecione um arquivo PDF"

    if ($openFileDialog.ShowDialog() -eq "OK") {
        $global:pdfPath = $openFileDialog.FileName
        $form.Controls["labelPDF"].Text = $global:pdfPath
    }
})
$form.Controls.Add($buttonPDF)

# Adiciona um rótulo para mostrar o caminho do arquivo PDF selecionado
$labelPDF = New-Object System.Windows.Forms.Label
$labelPDF.Location = New-Object System.Drawing.Point(50, 90)
$labelPDF.Size = New-Object System.Drawing.Size(400, 30)
$labelPDF.Name = "labelPDF"
$labelPDF.Text = ""
$form.Controls.Add($labelPDF)

# Adiciona um botão para iniciar a conversão
$buttonConvert = New-Object System.Windows.Forms.Button
$buttonConvert.Location = New-Object System.Drawing.Point(200, 130)
$buttonConvert.Size = New-Object System.Drawing.Size(100, 30)
$buttonConvert.Text = "Converter"
$buttonConvert.Add_Click({
    if (-not [string]::IsNullOrEmpty($global:pdfPath)) {
        $docxPath = [System.IO.Path]::ChangeExtension($global:pdfPath, ".docx")
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($global:pdfPath)
        $doc.SaveAs([ref] $docxPath, [ref] 16)
        $doc.Close()
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        Remove-Variable word
        [System.Windows.Forms.MessageBox]::Show("Conversão concluída!")

        # Atualiza o rótulo com o caminho do arquivo DOCX
        $form.Controls["labelDOCX"].Text = $docxPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Selecione um arquivo PDF primeiro.")
    }
})
$form.Controls.Add($buttonConvert)

# Adiciona um rótulo para mostrar o caminho do arquivo DOCX convertido
$labelDOCX = New-Object System.Windows.Forms.Label
$labelDOCX.Location = New-Object System.Drawing.Point(50, 170)
$labelDOCX.Size = New-Object System.Drawing.Size(400, 30)
$labelDOCX.Name = "labelDOCX"
$labelDOCX.Text = ""
$form.Controls.Add($labelDOCX)

# Mostra o formulário
$form.ShowDialog()
