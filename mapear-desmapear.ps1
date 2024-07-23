function Get-DriveLetterMap {
    $allLetters = [System.Collections.ArrayList]@("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

    $psdrives = Get-PSDrive -PSProvider FileSystem

    while ($true) {
        Write-Host "`nUnidades sendo utilizadas:"

        foreach ($psdrive in $psdrives) {

            if ([string]::IsNullOrEmpty($psdrive.DisplayRoot)) {
                Write-Host "$($psdrive.Name)"
            }
            else {
                Write-Host "$($psdrive.Name) --- $($psdrive.DisplayRoot)"
            }

            if ($psdrive.Name -in $allLetters) {
                $allLetters.Remove($psdrive.Name)
            }
        }
        Write-Host "`nUnidades disponiveis:"

        foreach ($letter in $allLetters) {
            Write-Host "$letter"
        }
        $driveLetter = Read-Host "`nDigite a letra de uma unidade disponivel"
        $driveLetter = $driveLetter.ToUpper()

        $driveNames = $psdrives | Select-Object -ExpandProperty Name

        if (-not ($driveLetter.Length -eq 1 -and $driveLetter -match '^[a-zA-Z]$')) {
            Write-Host "Entrada '$driveLetter' inválida."
        }
        elseif (-not($driveNames -contains $driveLetter)) {
            Write-Host "`nA unidade '$driveLetter' esta disponivel."
            break
        }
        else {
            Write-Host "`nA unidade '$driveLetter' esta sendo utilizada, tente outra."
        }
    }
    return $driveLetter
}

function Get-PSPath {
    $pspath = Read-Host "`nDigite o caminho de rede"

    while ($true) {
        if (Test-Path -PSPath $pspath) {
            break
        }
        else {
            try {
                Get-Item -PSPath $pspath -ErrorAction Stop
            }
            catch {
                Write-Host "`nDetalhes do erro: $_"
            }
            $pspath = Read-Host "`nCaminho nao encontrado, digite outro caminho"
        }
    }
    return $pspath
}

function Set-PSDriveValues {
    param(
        [string] $driveLetter, [string] $pspath, [string] $description
    )

    New-PSDrive -Name $driveLetter -Description $description -PSProvider FileSystem -Root $pspath -Persist -Scope Global

    $psdrives = Get-PSDrive -PSProvider FileSystem
    $driveNames = $psdrives | Select-Object -ExpandProperty Name

    if ($driveNames -contains $driveLetter) {
        Write-Host "`nO caminho '$pspath' foi mapeado na unidade '$driveLetter'.`n"
    }
    else {
        Write-Host "`nO caminho '$pspath' nao foi mapeado.`n"
    }

    Start-Sleep -Seconds 1
}

function New-Shortcut {
    param (
        [string] $driveLetter, [string] $description
    )

    while ($true) {
        $inputValue = Read-Host "`nDeseja criar um atalho na area de trabalho? [S] Sim [N] Nao"
        $value = $inputValue.ToLower()

        if ($value -eq "s") {
            $shortcutPath = "$env:USERPROFILE\Desktop\$description.lnk"
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath)
            $Shortcut.TargetPath = "$driveLetter`:\"
            $Shortcut.WorkingDirectory = "$driveLetter`:\"
            $Shortcut.IconLocation = "explorer.exe,0"
            $Shortcut.Save()
            Write-Host "`nAtalho criado com sucesso.`n"
            break
        }
        elseif ($value -eq "n") {
            break
        }
        else {
            Write-Host "`nTecla invalida."
        }
    }
}

function Get-DriveLetterDesmap {
    param(
        [System.Object] $psdrives
    )

    $allLetters = [System.Collections.ArrayList]@("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

    while ($true) {
        Write-Host "`nUnidades sendo utilizadas:"

        foreach ($psdrive in $psdrives) {

            if ([string]::IsNullOrEmpty($psdrive.DisplayRoot)) {
                Write-Host "$($psdrive.Name)"
            }
            else {
                Write-Host "$($psdrive.Name) --- $($psdrive.DisplayRoot)"
            }

            if ($psdrive.Name -in $allLetters) {
                $allLetters.Remove($psdrive.Name)
            }
        }
        Write-Host "* --- Selecionar Todos"

        Write-Host "`nUnidades disponiveis:"

        foreach ($letter in $allLetters) {
            Write-Host "$letter"
        }
        $driveLetter = Read-Host "`nDigite a letra da unidade utilizada que deseja desmapear"
        $driveLetter = $driveLetter.ToUpper()

        $driveNames = $psdrives | Select-Object -ExpandProperty Name
        $driveNames += "*"

        if (-not ($driveLetter.Length -eq 1 -and $driveLetter -match '^[a-zA-Z*]$')) {
            Write-Host "Entrada '$driveLetter' inválida."
        }
        elseif ($driveNames -contains $driveLetter) {
            Write-Host "`nA unidade '$driveLetter' foi selecionada."
            break
        }
        else {
            Write-Host "`nA unidade '$driveLetter' nao esta sendo utilizada, tente outra."
        }
    }
    return $driveLetter
}

function Remove-PSDriveValues {
    param (
        [string] $driveLetter
    )
    while ($true) {
        if ($driveLetter -ne "*") {

            $inputValue = Read-Host "`nDeseja remover a unidade '$driveLetter'? [S] Sim [N] Nao"
            $value = $inputValue.ToUpper()

            if ($value -eq "S") {
                $driveLetter = "$driveLetter`:"
                net use $driveLetter /del
                break
            }
            elseif ($value -eq "N") {
                break
            }
            else {
                Write-Host "`nTecla invalida."
            }
        }
        else {
            net use $driveLetter /del
            break
        }
    }
}

function Remove-Shortcut {
    param(
        [string] $driveLetter, [System.Object] $psdrives
    )

    $desktopPath = "$env:USERPROFILE\Desktop\"
    $shortcutsDesktop = Get-ChildItem -Path $desktopPath -Filter *.lnk

    if ($driveLetter -eq "*") {
        foreach ($psdrive in $psdrives) {

            foreach ($shortcutDesktop in $shortcutsDesktop) {
                $WshShell = New-Object -ComObject WScript.Shell
                $shortcutObject = $WshShell.CreateShortcut($shortcutDesktop.FullName)

                if ($shortcutObject.TargetPath -eq $psdrive.Root) {
                    Remove-Item -Path $shortcutObject.FullName -Force
                }
            } 
        }
    }
    else {
        $driveLetter = "$driveLetter`:\"
        foreach ($shortcutDesktop in $shortcutsDesktop) {
            $WshShell = New-Object -ComObject WScript.Shell
            $shortcutObject = $WshShell.CreateShortcut($shortcutDesktop.FullName)
    
            if ($shortcutObject.TargetPath -eq $driveLetter) {
                Remove-Item -Path $shortcutObject.FullName -Force
                Write-Host "Atalho "$shortcutObject.FullName" eliminado com sucesso.`n"
            }
        }
    }
}

try {
    while ($true) {
        Write-Host "Digite 'mapear' para mapear um caminho de rede."
        Write-Host "Digite 'desmapear' para desmapear um caminho de rede."
        Write-Host "Digite 'exit' para sair."

        $inputValue = Read-Host
        $value = $inputValue.ToLower()
        
        if ($value -eq "mapear") {
            $driveLetter = Get-DriveLetterMap
            $pspath = Get-PSPath
            $description = Read-Host "`nDigite uma descricao"
            Set-PSDriveValues $driveLetter $pspath $description
            New-Shortcut $driveLetter $description
        }
        elseif ($value -eq "desmapear") {
            $psdrives = Get-PSDrive -PSProvider FileSystem
            $driveLetter = Get-DriveLetterDesmap $psdrives
            Remove-PSDriveValues $driveLetter
            Remove-Shortcut $driveLetter $psdrives
        }
        elseif ($value -eq "exit") {
            break
        }
        else {
            Write-Host "Texto invalido.`n"
        }
    }
}
catch {
    Write-Error "`nOcorreu um erro: $_"
}