# Ricrea "Convertitore PUB.exe" dopo ogni modifica a pub2sla.py, pub2idml.py o Convertitore PUB.pyw.
# Serve un Python con requirements.txt installato: .\crea_eseguibile.ps1 -python <percorso di python.exe>
param([string]$python = "python")

$sorgenti = $PSScriptRoot
$lavoro = Join-Path $env:TEMP "convertitore-pub-build"
$copia = Join-Path $lavoro "src"  # si compila da una copia: i percorsi locali non finiscono nell'exe

Remove-Item $lavoro -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $copia | Out-Null
foreach ($file in "pub2sla.py", "pub2idml.py", "raccogli_font.py", "Convertitore PUB.pyw", "font_base.json", "icona.ico") {
    Copy-Item -LiteralPath (Join-Path $sorgenti $file) $copia
}

& $python -m PyInstaller --noconfirm --clean --onefile --windowed --name "Convertitore PUB" `
    --paths "$copia" --add-data "$copia\font_base.json;." `
    --icon "$copia\icona.ico" --add-data "$copia\icona.ico;." `
    --exclude-module onnxruntime --exclude-module numpy --exclude-module networkx `
    --exclude-module pymupdf_layout --exclude-module pymupdf4llm --exclude-module PIL `
    --exclude-module fontTools --exclude-module yaml --exclude-module psutil --exclude-module tabulate `
    --workpath "$lavoro\work" --distpath "$sorgenti" --specpath "$lavoro" "$copia\Convertitore PUB.pyw"
