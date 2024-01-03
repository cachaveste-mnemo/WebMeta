 # Solicitar al usuario la ruta del directorio a filtrar
param (
    [string]$dom,
    [string]$ld, # Inicia la búsqueda a partir de una lsita de dominios
    [string]$dir, #  Por default obtiene archivos y resultados en la carpeta webmeta en Documents
    [string]$rc, # Carpeta en la que se entregaran los resultados, distinta de Documents
    [switch]$help = $false #Para ayuda
    #,[switch]$r para recursividad
)



function MostrarAyuda {
    Write-Host "
    ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⢀⣰⡋⠉⢆⠀⠀⠀⡀⠤⠒⠒⠀⠐⠒⠤⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⣼⣃⠀⠉⠳⡀⠈⢆⡴⠋⠀⠀⠀⠀⠀⠀⠀⠀⠀⠷⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣠⠊⠀⢸⠘⠳⣄⠀⠙⣤⠋⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⡄⡠⡤⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⠜⢁⣤⡴⠏⣢⣀⢈⡶⠤⠧⠀⣀⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣏⣴⠷⣌⠢⣀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⣠⠖⠙⣤⠋⢀⣇⣀⣠⠼⣟⡀⠀⠀⠀⠀⠈⠳⡄⠀⠀⠀⠀⠀⡠⠤⡄⡠⠋⢸⠤⡈⠻⣎⣳⡄⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⢀⡤⠒⠉⢀⣤⠞⢁⠴⡏⠀⡍⢡⣀⣧⢿⣫⡇⠺⠏⠀⠀⠈⡆⡜⠉⣇⢰⡇⠀⢿⠁⣠⠃⠀⠈⠳⢌⠢⣱⣄⠀⠀⠀⠀
⠀⠀⠀⢀⡴⢁⣴⡶⠟⠋⢀⡰⠋⡼⠀⡼⣁⡼⠋⢱⠖⠓⣄⠀⠀⡰⢲⠾⡁⡇⠀⢸⣿⢱⠀⢸⠖⠁⠀⠀⠀⠀⠀⠳⣿⣯⣑⣢⡄⠀
⠀⠀⠶⠉⢴⡿⠋⡠⠜⠂⠉⢀⠴⢉⠞⢠⠎⠀⢠⠋⠀⢀⡞⠦⢴⠁⢸⠴⠋⢳⠀⠸⡟⠉⡇⠈⡆⠀⠀⠀⠀⠀⠀⠀⠈⢦⠣⡀⠀⠀
⠀⠀⣠⠔⢉⡤⠊⠀⠀⠀⠀⠈⠀⠉⠀⠘⠦⠠⣇⠀⢀⡾⠀⢀⠎⢰⠇⠀⠀⠈⡄⠀⡇⠀⢡⠄⢳⠀⠀⠀⠀⠀⠀⠀⠀⠀⠳⡙⢦⡀
⣔⣊⡡⠖⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⠉⠀⢀⠎⠀⡎⠀⠀⠀⠀⡗⠒⠃⠀⠸⠀⢸⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠓⠒
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠘⠒⠚⠁⠀⠀⠀⠀⢹⠀⢸⠀⠀⡇⢸⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢇⢸⠀⠀⡗⠚⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢸⠚⡇⠀⢹⠀⢧⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠘⡆⢰⠀⠈⠓⠚⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣸⣀⣸⠃⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
        __    __     _                 _                    
       / / /\ \ \___| |__   /\/\   ___| |_ __ _             
       \ \/  \/ / _ \ '_ \ /    \ / _ \ __/ _` |            
        \  /\  /  __/ |_) / /\/\ \  __/ || (_| |            
         \/  \/ \___|_.__/\/    \/\___|\__\__,_|            
              _                          _       _          
      ___ ___| |__   __ ___   _____  ___| |_ ___| |__       
     / __/ __| '_ \ / _` \ \ / / _ \/ __| __/ _ \ '_ \      
    | (_| (__| | | | (_| |\ V /  __/\__ \ ||  __/ |_) |     
     \___\___|_| |_|\__,_| \_/ \___||___/\__\___|_.__/     

"
    Write-Host "Para la utilización del script se necesita de dos parámetros, `nuno para indicar el dominio o lista de dominios a los que se `nles realizará la extracción.`nEl otro es para indicar si se pondrá el resultado bajo el `ndirectorio documentos o bien en una ruta absoluta.`n  "
    Write-Host "Opciones referentes al dominino:"
    Write-Host "-dom   : El dominio que se va a buscar."
    Write-Host "-ld    : La ruta absoluta de un archivo .txt, el separador es un salto de línea.`n"
    Write-Host "Opciones referentes al directorio destino:"
    Write-Host "-dir   : Nombre del directorio bajo Documentos en donde aparecera un nuevo directorio. Debe ser una palabra solamente."
    Write-Host "-rc    : Es la ruta completa."
    Write-Host "-h, --help    : Muestra esta ayuda.`n`n"
    # Agregar más información sobre otras opciones si es necesario
}

$currentTime = Get-Date

function WM{
param (
    [string]$dom,
    [string]$dir, #  Por default obtiene archivos y resultados en la carpeta webmeta en Documents
    [string]$rc # Carpeta en la que se entregaran los resultados, distinta de Documents
    #,[switch]$r para recursividad
)
    #Comprobación de banderas
    $dominioAdquisicion = $dom #Read-Host "Ingrese el dominio del que se obtendrán archivos de ofimática"
    $nombreCarpeta = $dir #Read-Host "Ingrese el nombre de la carpeta en la que se van a alojar los documentos extraídos"
    if ( -not $dir -and -not $rc -or -not $dom) { 
        Write-Host "Especifique banderas"
        exit
    }
    #Si se recibe dir
    elseif (-not $rc){
        $nombreCarpeta = $dir
        $rutaDestino = Join-Path -Path $rutaDocumentos -ChildPath $nombreCarpeta
        }
    else{$rutaDestino = $rc} #Si se rcibe rc 


    $nuevaLineaComando = "& '$rutaRelativaWC'"+" "+ $dominioAdquisicion + " /recursive /NOLOGO /o  " +"'$rutaDestino"+"'"
    if (-not (Test-Path $rutaDestino -PathType Container)) {
        # Si no existe, se crea el directorio
        New-Item -Path $rutaDestino -ItemType Directory
        Write-Host "Directory created at: $rutaDestino"
        } else {
            Write-Host "Directory already exists at: $rutaDestino"
        }

    $nuevaLineaComando = "& '$rutaRelativaWC'"+" "+ $dominioAdquisicion + " /recursive /NOLOGO /o  " +"'$rutaDestino"+"'"
    # Write-Host "Nueva línea de comando: $nuevaLineaComando"
    # Ejecutar la nueva línea de comando
    powershell -command $nuevaLineaComando

    # Filtrar y mover los archivos
    $rutaArchivos = Join-Path $rutaDestino "Archivos obtenidos"
    if (-not (Test-Path -Path $rutaArchivos)) {
        New-Item -ItemType Directory -Path $rutaArchivos | Out-Null
    }
    $extensiones = @(".docx", ".doc", ".odt", ".xlsx", ".xls", ".csv", ".ods", ".pptx", ".ppt", ".odp", ".pdf")

    #Recorrer toda la carpeta para encontrar los archivos, obtener la metadata y moverlos a "Archivos obtenidos"
    Get-ChildItem -Path $rutaDestino -Recurse |
        Where-Object { $_.Extension -in $extensiones -and $_.Name -ne  "compilado$($currentTime.ToString('yyyy-MM-dd')).csv"} |
        ForEach-Object {
            $RutaCompleta= $_.FullName
            Write-Host "La ruta del archivo es:" $RutaCompleta
            $exiftoolOutput = powershell -command "& '$rutaRelativaET' '$RutaCompleta' -charset utf8"     
            [string]$nombreArchivo = ($exifToolOutput | Select-String "File Name").Line -replace "File Name\s*:\s*", ""  
            [string]$Title =  ($exifToolOutput | Select-String "Title").Line -replace "Title\s*:\s*", ""
            [string]$FileType = ($exifToolOutput | Select-String "File Type Extension").Line -replace "File Type Extension\s*:\s*", ""
            [string]$Software = ($exifToolOutput | Select-String "Software").Line -replace "Software\s*:\s*", ""
            [string]$Company = ($exifToolOutput | Select-String "Company").Line -replace "Company\s*:\s*", ""
            [string]$Subject = ($exifToolOutput | Select-String "Subject").Line -replace "subject\s*:\s*", ""
            [string]$Keywords = ($exifToolOutput | Select-String "Keywords").Line -replace "Keywords\s*:\s*", ""
            [string]$createDate = ($exifToolOutput | Select-String "Create Date").Line -replace "Create Date\s*:\s*", ""
            [string]$modifyDate = ($exifToolOutput | Select-String "Modify Date").Line -replace "Modify Date\s*:\s*", ""
            [string]$Producer = ($exifToolOutput | Select-String "Producer").Line -replace "Producer\s*:\s*", ""
            [string]$paginas = ($exifToolOutput | Select-String "Page Count").Line -replace "Page Count\s*:\s*", ""
            [string]$creator = ($exifToolOutput | Select-String "Creator" | Select-Object -First 1).Line -replace "Creator\s*:\s*", ""
            [string]$creatorTool = ($exifToolOutput | Select-String "Creator Tool").Line -replace "Creator Tool\s*:\s*", ""
            [string]$author = ($exifToolOutput | Select-String "Author").Line -replace "Author\s*:\s*", ""
            [string]$language = ($exifToolOutput | Select-String "Language").Line -replace "Language\s*:\s*", ""        
            [string]$Parrafos = ($exifToolOutput | Select-String "Paragraphs").Line -replace "Paragraphs\s*:\s*", ""
            [string]$ObjectType = [regex]::Match($exifToolOutput -join "`n", 'Comp Obj User Type\s*:\s*(.+)$').Groups[1].Value
            $dominio = $dominioAdquisicion -replace '^https?://([^/]+).*$', '$1'
            #$URL = $RutaCompleta -replace ".*\\$nombreCarpeta\\", "$dominio/"
            #$URL= $URL -replace '\\', '/'
            $rutaArchivoWebCopy=$rutaDestino+'\webcopy-origin.txt'
            $contenido = Get-Content $rutaArchivoWebCopy
            $cont = 0
            $URL = $contenido | ForEach-Object {
            if ($_ -match "File Name: (.*)") {
                $rutaArchivo = $matches[1].Trim()
                if ($rutaArchivo -eq $rutaCompleta) {
                    $contenido[--$cont] -replace "Uri\s*:\s*", "" -as [string]
                    }
                }
                $cont++
            }
                    
            $resultados = [PSCustomObject]@{
            'Fecha'= $currentTime.ToString('yyyy-MM-dd')
            'URI' = $URL
            'Dominio' = $dominio
            'Nombre del Archivo' = $nombreArchivo
            'Tipo de Archivo' = $FileType
            'Titulo' = $Title
            'Autor'= $author            
            'Asunto' = $Subject
            'Palabra clave' = $Keywords
            'Creado' = $createDate
            'Modificado' = $modifyDate
            'Software'= $Software
            'Compañia' = $Company           
            'Proveedor' = $Producer
            'Numero de paginas' =$paginas
            'Creador'           = $creator
            'Herramienta del Creador' = $creatorTool
            'Idioma'            = $language
            'Parrafos'          = $Parrafos
            'Tipo de Objeto'    = $ObjectType
            }
       
            $destinoArchivo = Join-Path $rutaArchivos $_.Name
            Write-Host "Se ha obtenido la metadata de $destinoArchivo"
            Move-Item $_.FullName -Destination $destinoArchivo -Force
            Write-Host "Se ha movido $($_.Name) a $destinoArchivo"
            $rutaCsv = Join-Path $rutaDestino "compilado$($currentTime.ToString('yyyy-MM-dd')).csv"
            Write-Host "rutaCSV = "$rutaCSV
            if (Test-Path "$rutaCsv") {
            $resultados | Export-Csv -Path "$rutaCsv" -Append -NoTypeInformation -Encoding UTF8} 
            else {
            $resultados | Export-Csv -Path "$rutaCsv" -NoTypeInformation -Encoding UTF8}
        
        }
    Write-Host "Proceso completado."
}

# Verificar si la opción de ayuda está presente
if ($help) {
    MostrarAyuda
    # Salir del script después de mostrar la ayuda si es necesario
    return
}

#Obtiene la ruta de Documents
$rutaDocumentos = Join-Path -Path $env:USERPROFILE -ChildPath "Documents" 

# Obtiene la ruta del directorio actual
$directorioWM = Get-Location

# Concatena la ruta relativa al directorio actual
$rutaRelativaWC = Join-Path -Path $directorioWM -ChildPath "Ejecutables\wcopy.exe"
$rutaRelativaET = Join-Path -Path $directorioWM -ChildPath ".\Ejecutables\exiftool.exe"

# Navega al directorio que contiene el ejecutable
Set-Location -Path $directorioWM

# Verificar si la bandera $ld tiene contenido
if (-not ($ld -eq $null -or $ld -eq '')) {
    Write-Host "La bandera -ld tiene contenido: $ld, se procederá a verificar que la ruta dirija a un archivo .txt válido"
    if (-not (Test-Path $ld -PathType Leaf)){
    Write-Host "La ruta especificada no contiene un archivo .txt válido."
    }
    else{
        $listaDeDominios = Get-Content -Path $ld
        foreach ($linea in $listaDeDominios){
            $dom=$linea
            WM -dom $dom -dir $dir -rc $rc
        }
    }
} else {
    Write-Host "La bandera -ld no tiene contenido. Se verificará el contenido de -dom"
    Write-Host $dom $dir $rc
    WM -dom $dom -dir $dir -rc $rc
}
