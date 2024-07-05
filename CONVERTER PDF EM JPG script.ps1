# Caminho para o Ghostscript
$ghostscriptPath = "C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe"

# Caminho para o arquivo PDF de entrada
$pdfPath = "C:\Users\FbwTech\Downloads\diagramaCOZxONS.pdf"

# Pasta de saída para as imagens JPEG
$outputFolder = "C:\Users\FbwTech\Desktop\q.jpg"

# Verifica se a pasta de saída existe, se não, cria-a
if (!(Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# Define o nome base para as imagens JPEG de saída
$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)

# Comando para executar a conversão usando Ghostscript
& $ghostscriptPath -sDEVICE=jpeg -r300 -o "$outputFolder\$outputBaseName-%03d.jpg" $pdfPath

Write-Host "Conversão concluída. Imagens JPEG geradas em: $outputFolder"
