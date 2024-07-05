# Carregar a assembly do Microsoft Office Interop
Add-Type -AssemblyName Microsoft.Office.Interop.Word

# Caminho para o arquivo PDF que você deseja converter
$pdfPath = "C:\Users\FbwTech\Downloads\RFPMITRWtg.pdf"

# Caminho para salvar o arquivo DOCX convertido
$docxPath = "C:\Users\FbwTech\Desktop\arquivo.docx"

# Criar uma instância do Word
$word = New-Object -ComObject Word.Application

# Especificar que não queremos exibir o Word durante o processo
$word.Visible = $false

# Abrir o PDF
$doc = $word.Documents.Open($pdfPath)

# Salvar como DOCX
$doc.SaveAs([ref] $docxPath, [ref] 16)  # 16 para formato DOCX

# Fechar o documento PDF e o aplicativo Word
$doc.Close()
$word.Quit()

# Limpar a instância do Word da memória
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable word
