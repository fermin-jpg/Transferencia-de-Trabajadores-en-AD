@echo off
setlocal

net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs -WindowStyle Hidden"
    exit /b
)

set "BATPATH=%~f0"
set "PS1=%TEMP%\ADTransfer_%RANDOM%.ps1"
powershell -Command "$f=Get-Content $env:BATPATH -Encoding Default; $n=($f | Select-String '^:PS1START').LineNumber; $f | Select-Object -Skip $n | Set-Content $env:PS1 -Encoding UTF8"
powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%PS1%"
if exist "%PS1%" del /f /q "%PS1%"
exit /b

:PS1START
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Verificar modulo ActiveDirectory ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "No se encontro el modulo 'ActiveDirectory'.`n`nNecesitas instalar las RSAT Tools:`nConfiguracion > Aplicaciones > Caracteristicas opcionales`n> RSAT: Active Directory Domain Services y Lightweight Directory Tools",
        "Modulo no encontrado",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit
}

# --- Variables globales ---
$global:managerOrigen     = $null
$global:managerDestino    = $null
$global:subordinados      = @()
$global:resultadosOrigen  = @()
$global:resultadosDestino = @()
$global:animStep = 0
$global:animDir  = 1

# ================================================================
# DIMENSIONES
# ================================================================
$formW      = 1150
$margen     = 12
$anchoTotal = $formW - ($margen * 2) - 12   # ~1114

$colW  = [int](($anchoTotal - 8) / 2)        # ~553
$xDer  = $margen + $colW + 8

$alturaTop = 170   # PASO 1 y PASO 3 (busqueda)
$alturaMid = 400   # PASO 2 y lista trabajadores destino
$alturaBtn = 68    # PASO 4 ejecutar

# ================================================================
# FORMULARIO PRINCIPAL
# ================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Transferencia de Trabajadores v1.0.0 by Fermin32"
$form.StartPosition   = "CenterScreen"
$form.BackColor       = [System.Drawing.Color]::FromArgb(228, 228, 228)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ================================================================
# CABECERA OSCURA
# ================================================================
$panelHeader = New-Object System.Windows.Forms.Panel
$panelHeader.Location  = New-Object System.Drawing.Point(0, 0)
$panelHeader.Size      = New-Object System.Drawing.Size($formW, 62)
$panelHeader.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 35)
$form.Controls.Add($panelHeader)

$lblAD = New-Object System.Windows.Forms.Label
$lblAD.Text      = "AD"
$lblAD.Font      = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblAD.ForeColor = [System.Drawing.Color]::FromArgb(0, 210, 120)
$lblAD.Location  = New-Object System.Drawing.Point(16, 10)
$lblAD.AutoSize  = $true
$panelHeader.Controls.Add($lblAD)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "Transferencia de Trabajadores"
$lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Location  = New-Object System.Drawing.Point(65, 10)
$lblTitle.AutoSize  = $true
$panelHeader.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text      = "Active Directory  -  Reasignacion de manager entre trabajadores"
$lblSub.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(155, 155, 175)
$lblSub.Location  = New-Object System.Drawing.Point(67, 36)
$lblSub.AutoSize  = $true
$panelHeader.Controls.Add($lblSub)

# ================================================================
# FUNCIONES
# ================================================================
function Write-Log {
    param([string]$Msg, [string]$Tipo = "Info")
    $ts     = Get-Date -Format "HH:mm:ss"
    $prefix = switch ($Tipo) {
        "OK"    { "[  OK  ] " }
        "Error" { "[ ERROR ] " }
        "Warn"  { "[ AVISO ] " }
        "Info"  { "[ INFO  ] " }
        "Step"  { "[       ] " }
        default { "          " }
    }
    $color = switch ($Tipo) {
        "OK"    { [System.Drawing.Color]::Lime }
        "Error" { [System.Drawing.Color]::Red }
        "Warn"  { [System.Drawing.Color]::Yellow }
        "Info"  { [System.Drawing.Color]::Cyan }
        "Step"  { [System.Drawing.Color]::White }
        default { [System.Drawing.Color]::White }
    }
    $txtLog.SelectionStart  = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor  = $color
    $txtLog.AppendText("[$ts] $prefix$Msg`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-SubordinadosSeleccionados {
    $sel = @()
    for ($i = 0; $i -lt $lstSubs.Items.Count; $i++) {
        if ($lstSubs.Items[$i].Checked) { $sel += $global:subordinados[$i] }
    }
    return $sel
}

function Actualizar-Boton {
    $sel = Get-SubordinadosSeleccionados
    $ok  = ($global:managerDestino -ne $null) -and
           ($sel.Count -gt 0)
    $btnEjecutar.Enabled = $ok
    if ($ok) {
        $n = $sel.Count
        $btnEjecutar.Text = " PULSA PARA EJECUTAR TRANSFERENCIA   ---   $n trabajador(es) seran reasignados"
    } else {
        $btnEjecutar.BackColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
        $btnEjecutar.Text = "PULSA PARA EJECUTAR TRANSFERENCIA"
    }
}

function Actualizar-ContadorSubs {
    $sel   = 0
    $total = $lstSubs.Items.Count
    for ($i = 0; $i -lt $total; $i++) {
        if ($lstSubs.Items[$i].Checked) { $sel++ }
    }
    if ($total -gt 0) {
        $lblSubsStatus.Text = "$sel de $total seleccionados  -  listos para ser transferidos"
        $lblSubsStatus.ForeColor = if ($sel -eq $total) {
            [System.Drawing.Color]::FromArgb(0, 120, 0)
        } else {
            [System.Drawing.Color]::FromArgb(160, 100, 0)
        }
    }
    Actualizar-Boton
}

function Cargar-Subordinados {
    $lstSubs.Items.Clear()
    $global:subordinados = @()
    if ($global:managerOrigen -eq $null) { return }

    $lblSubsStatus.Text      = "Cargando trabajadores en Active Directory..."
    $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $dn   = $global:managerOrigen.DistinguishedName
        $subs = Get-ADUser -Filter "Manager -eq '$dn'" `
                    -Properties DisplayName, SamAccountName, Department, Title |
                    Sort-Object DisplayName

        $global:subordinados = @($subs)

        foreach ($s in $subs) {
            $item = New-Object System.Windows.Forms.ListViewItem("")
            $item.Checked = $true
            $item.SubItems.Add($s.DisplayName) | Out-Null
            $item.SubItems.Add($s.SamAccountName) | Out-Null
            $item.SubItems.Add($(if ($s.Department) { $s.Department } else { "---" })) | Out-Null
            $item.SubItems.Add($(if ($s.Title)      { $s.Title }      else { "---" })) | Out-Null
            $lstSubs.Items.Add($item) | Out-Null
        }

        if ($subs.Count -eq 0) {
            $lblSubsStatus.Text      = "Este manager no tiene trabajadores directos registrados en AD."
            $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(160, 100, 0)
            Write-Log "El manager origen no tiene trabajadores directos en AD." "Warn"
        } else {
            $lblSubsStatus.Text      = "$($subs.Count) de $($subs.Count) seleccionados  -  listos para ser transferidos"
            $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
            Write-Log "$($subs.Count) trabajador(es) cargados para $($global:managerOrigen.DisplayName)." "OK"
        }
    } catch {
        $lblSubsStatus.Text      = "Error al cargar trabajadores: $_"
        $lblSubsStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "Error al cargar trabajadores: $_" "Error"
    }

    Actualizar-Boton
}

function Cargar-SubordinadosDestino {
    $lstSubsDestino.Items.Clear()
    if ($global:managerDestino -eq $null) { return }

    $lblSubsDestinoStatus.Text      = "Cargando trabajadores actuales del manager destino..."
    $lblSubsDestinoStatus.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $dn   = $global:managerDestino.DistinguishedName
        $subs = Get-ADUser -Filter "Manager -eq '$dn'" `
                    -Properties DisplayName, SamAccountName, Department, Title |
                    Sort-Object DisplayName

        foreach ($s in $subs) {
            $item = New-Object System.Windows.Forms.ListViewItem($s.DisplayName)
            $item.SubItems.Add($s.SamAccountName) | Out-Null
            $item.SubItems.Add($(if ($s.Department) { $s.Department } else { "---" })) | Out-Null
            $item.SubItems.Add($(if ($s.Title)      { $s.Title }      else { "---" })) | Out-Null
            $lstSubsDestino.Items.Add($item) | Out-Null
        }

        if ($subs.Count -eq 0) {
            $lblSubsDestinoStatus.Text      = "El manager destino no tiene trabajadores directos actualmente."
            $lblSubsDestinoStatus.ForeColor = [System.Drawing.Color]::FromArgb(160, 100, 0)
        } else {
            $lblSubsDestinoStatus.Text      = "$($subs.Count) trabajador(es) actuales del manager destino"
            $lblSubsDestinoStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
        }
    } catch {
        $lblSubsDestinoStatus.Text      = "Error al cargar trabajadores destino: $_"
        $lblSubsDestinoStatus.ForeColor = [System.Drawing.Color]::Red
    }
}

# ================================================================
# FILA 1: PASO 1 (izquierda) | PASO 3 (derecha)  <-- ambas busquedas arriba
# ================================================================
$y = 72

# --- PASO 1 - MANAGER ORIGEN ---
$gbOrigen = New-Object System.Windows.Forms.GroupBox
$gbOrigen.Text      = "  PASO 1  -  Manager ORIGEN  (el que se va o cambia de departamento)"
$gbOrigen.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$gbOrigen.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
$gbOrigen.Location  = New-Object System.Drawing.Point($margen, $y)
$gbOrigen.Size      = New-Object System.Drawing.Size($colW, $alturaTop)
$gbOrigen.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$form.Controls.Add($gbOrigen)

    $lbO1 = New-Object System.Windows.Forms.Label
    $lbO1.Text     = "Escribe el nombre o usuario del manager que se va:"
    $lbO1.Location = New-Object System.Drawing.Point(12, 26)
    $lbO1.AutoSize = $true
    $gbOrigen.Controls.Add($lbO1)

    $txtBuscaOrigen = New-Object System.Windows.Forms.TextBox
    $txtBuscaOrigen.Location  = New-Object System.Drawing.Point(12, 46)
    $txtBuscaOrigen.Size      = New-Object System.Drawing.Size(($colW - 210), 26)
    $txtBuscaOrigen.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $gbOrigen.Controls.Add($txtBuscaOrigen)

    $btnBuscaOrigen = New-Object System.Windows.Forms.Button
    $btnBuscaOrigen.Text      = "  Buscar en AD"
    $btnBuscaOrigen.Location  = New-Object System.Drawing.Point(($colW - 195), 44)
    $btnBuscaOrigen.Size      = New-Object System.Drawing.Size(177, 30)
    $btnBuscaOrigen.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $btnBuscaOrigen.ForeColor = [System.Drawing.Color]::White
    $btnBuscaOrigen.FlatStyle = "Flat"
    $btnBuscaOrigen.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBuscaOrigen.Cursor    = "Hand"
    $gbOrigen.Controls.Add($btnBuscaOrigen)

    $lstOrigenRes = New-Object System.Windows.Forms.ListBox
    $lstOrigenRes.Location            = New-Object System.Drawing.Point(12, 82)
    $lstOrigenRes.Size                = New-Object System.Drawing.Size(($colW - 26), 52)
    $lstOrigenRes.Font                = New-Object System.Drawing.Font("Consolas", 9)
    $lstOrigenRes.ScrollAlwaysVisible = $true
    $gbOrigen.Controls.Add($lstOrigenRes)

    $lblOrigenSel = New-Object System.Windows.Forms.Label
    $lblOrigenSel.Text      = "Ningun manager origen seleccionado"
    $lblOrigenSel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblOrigenSel.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblOrigenSel.Location  = New-Object System.Drawing.Point(12, 142)
    $lblOrigenSel.AutoSize  = $true
    $gbOrigen.Controls.Add($lblOrigenSel)

# --- PASO 3 - MANAGER DESTINO (arriba derecha, al lado del PASO 1) ---
$gbDestino = New-Object System.Windows.Forms.GroupBox
$gbDestino.Text      = "  PASO 3  -  Manager DESTINO  (el que recibe a los trabajadores)"
$gbDestino.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$gbDestino.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 50)
$gbDestino.Location  = New-Object System.Drawing.Point($xDer, $y)
$gbDestino.Size      = New-Object System.Drawing.Size($colW, $alturaTop)
$gbDestino.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$form.Controls.Add($gbDestino)

    $lbD1 = New-Object System.Windows.Forms.Label
    $lbD1.Text     = "Escribe el nombre o usuario del nuevo manager:"
    $lbD1.Location = New-Object System.Drawing.Point(12, 26)
    $lbD1.AutoSize = $true
    $gbDestino.Controls.Add($lbD1)

    $txtBuscaDestino = New-Object System.Windows.Forms.TextBox
    $txtBuscaDestino.Location  = New-Object System.Drawing.Point(12, 46)
    $txtBuscaDestino.Size      = New-Object System.Drawing.Size(($colW - 210), 26)
    $txtBuscaDestino.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $gbDestino.Controls.Add($txtBuscaDestino)

    $btnBuscaDestino = New-Object System.Windows.Forms.Button
    $btnBuscaDestino.Text      = "  Buscar en AD"
    $btnBuscaDestino.Location  = New-Object System.Drawing.Point(($colW - 195), 44)
    $btnBuscaDestino.Size      = New-Object System.Drawing.Size(177, 30)
    $btnBuscaDestino.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $btnBuscaDestino.ForeColor = [System.Drawing.Color]::White
    $btnBuscaDestino.FlatStyle = "Flat"
    $btnBuscaDestino.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBuscaDestino.Cursor    = "Hand"
    $gbDestino.Controls.Add($btnBuscaDestino)

    $lstDestinoRes = New-Object System.Windows.Forms.ListBox
    $lstDestinoRes.Location            = New-Object System.Drawing.Point(12, 82)
    $lstDestinoRes.Size                = New-Object System.Drawing.Size(($colW - 26), 52)
    $lstDestinoRes.Font                = New-Object System.Drawing.Font("Consolas", 9)
    $lstDestinoRes.ScrollAlwaysVisible = $true
    $gbDestino.Controls.Add($lstDestinoRes)

    $lblDestinoSel = New-Object System.Windows.Forms.Label
    $lblDestinoSel.Text      = "Ningun manager destino seleccionado"
    $lblDestinoSel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblDestinoSel.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblDestinoSel.Location  = New-Object System.Drawing.Point(12, 142)
    $lblDestinoSel.AutoSize  = $true
    $gbDestino.Controls.Add($lblDestinoSel)

$y += $alturaTop + 6

# ================================================================
# FILA 2: PASO 2 (izquierda) | Trabajadores destino (derecha)
# ================================================================

# --- PASO 2 - LISTA ORIGEN CON CHECKBOXES ---
$gbSubs = New-Object System.Windows.Forms.GroupBox
$gbSubs.Text      = "  PASO 2  -  Trabajadores del manager ORIGEN  (marca los que quieres transferir)"
$gbSubs.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$gbSubs.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 0)
$gbSubs.Location  = New-Object System.Drawing.Point($margen, $y)
$gbSubs.Size      = New-Object System.Drawing.Size($colW, $alturaMid)
$gbSubs.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$form.Controls.Add($gbSubs)

    $btnMarcarTodos = New-Object System.Windows.Forms.Button
    $btnMarcarTodos.Text      = "Marcar todos"
    $btnMarcarTodos.Location  = New-Object System.Drawing.Point(12, 22)
    $btnMarcarTodos.Size      = New-Object System.Drawing.Size(110, 24)
    $btnMarcarTodos.BackColor = [System.Drawing.Color]::FromArgb(60, 120, 60)
    $btnMarcarTodos.ForeColor = [System.Drawing.Color]::White
    $btnMarcarTodos.FlatStyle = "Flat"
    $btnMarcarTodos.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnMarcarTodos.Cursor    = "Hand"
    $gbSubs.Controls.Add($btnMarcarTodos)

    $btnDesmarcarTodos = New-Object System.Windows.Forms.Button
    $btnDesmarcarTodos.Text      = "Desmarcar todos"
    $btnDesmarcarTodos.Location  = New-Object System.Drawing.Point(128, 22)
    $btnDesmarcarTodos.Size      = New-Object System.Drawing.Size(120, 24)
    $btnDesmarcarTodos.BackColor = [System.Drawing.Color]::FromArgb(160, 60, 60)
    $btnDesmarcarTodos.ForeColor = [System.Drawing.Color]::White
    $btnDesmarcarTodos.FlatStyle = "Flat"
    $btnDesmarcarTodos.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnDesmarcarTodos.Cursor    = "Hand"
    $gbSubs.Controls.Add($btnDesmarcarTodos)

    $btnImportar = New-Object System.Windows.Forms.Button
    $btnImportar.Text      = "  Importar Excel/CSV"
    $btnImportar.Location  = New-Object System.Drawing.Point(254, 22)
    $btnImportar.Size      = New-Object System.Drawing.Size(150, 24)
    $btnImportar.BackColor = [System.Drawing.Color]::FromArgb(0, 130, 80)
    $btnImportar.ForeColor = [System.Drawing.Color]::White
    $btnImportar.FlatStyle = "Flat"
    $btnImportar.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnImportar.Cursor    = "Hand"
    $gbSubs.Controls.Add($btnImportar)

    $lstSubs = New-Object System.Windows.Forms.ListView
    $lstSubs.Location      = New-Object System.Drawing.Point(12, 52)
    $lstSubs.Size          = New-Object System.Drawing.Size(($colW - 26), ($alturaMid - 90))
    $lstSubs.View          = "Details"
    $lstSubs.FullRowSelect = $true
    $lstSubs.GridLines     = $true
    $lstSubs.CheckBoxes    = $true
    $lstSubs.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

    $cChk = New-Object System.Windows.Forms.ColumnHeader; $cChk.Text = "";               $cChk.Width = 24
    $c1   = New-Object System.Windows.Forms.ColumnHeader; $c1.Text   = "Nombre completo"; $c1.Width = 185
    $c2   = New-Object System.Windows.Forms.ColumnHeader; $c2.Text   = "Login";           $c2.Width = 90
    $c3   = New-Object System.Windows.Forms.ColumnHeader; $c3.Text   = "Departamento";    $c3.Width = 155
    $c4   = New-Object System.Windows.Forms.ColumnHeader; $c4.Text   = "Cargo";           $c4.Width = 80
    $lstSubs.Columns.AddRange(@($cChk, $c1, $c2, $c3, $c4))
    $gbSubs.Controls.Add($lstSubs)

    $lblSubsStatus = New-Object System.Windows.Forms.Label
    $lblSubsStatus.Text      = "Selecciona un manager origen en el Paso 1 para ver aqui sus Trabajadores."
    $lblSubsStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lblSubsStatus.Location  = New-Object System.Drawing.Point(12, ($alturaMid - 32))
    $lblSubsStatus.AutoSize  = $true
    $gbSubs.Controls.Add($lblSubsStatus)

# --- LISTA TRABAJADORES DEL DESTINO (derecha, fila 2) ---
$gbSubsDestino = New-Object System.Windows.Forms.GroupBox
$gbSubsDestino.Text      = "  Trabajadores actuales del manager DESTINO"
$gbSubsDestino.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$gbSubsDestino.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 50)
$gbSubsDestino.Location  = New-Object System.Drawing.Point($xDer, $y)
$gbSubsDestino.Size      = New-Object System.Drawing.Size($colW, $alturaMid)
$gbSubsDestino.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$form.Controls.Add($gbSubsDestino)

    $lstSubsDestino = New-Object System.Windows.Forms.ListView
    $lstSubsDestino.Location      = New-Object System.Drawing.Point(12, 22)
    $lstSubsDestino.Size          = New-Object System.Drawing.Size(($colW - 26), ($alturaMid - 58))
    $lstSubsDestino.View          = "Details"
    $lstSubsDestino.FullRowSelect = $true
    $lstSubsDestino.GridLines     = $true
    $lstSubsDestino.BackColor     = [System.Drawing.Color]::FromArgb(240, 255, 248)
    $lstSubsDestino.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

    $dC1 = New-Object System.Windows.Forms.ColumnHeader; $dC1.Text = "Nombre completo"; $dC1.Width = 185
    $dC2 = New-Object System.Windows.Forms.ColumnHeader; $dC2.Text = "Login";           $dC2.Width = 90
    $dC3 = New-Object System.Windows.Forms.ColumnHeader; $dC3.Text = "Departamento";    $dC3.Width = 155
    $dC4 = New-Object System.Windows.Forms.ColumnHeader; $dC4.Text = "Cargo";           $dC4.Width = 80
    $lstSubsDestino.Columns.AddRange(@($dC1, $dC2, $dC3, $dC4))
    $gbSubsDestino.Controls.Add($lstSubsDestino)

    $lblSubsDestinoStatus = New-Object System.Windows.Forms.Label
    $lblSubsDestinoStatus.Text      = "Selecciona manager destino para ver sus Trabajadores actuales."
    $lblSubsDestinoStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblSubsDestinoStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $lblSubsDestinoStatus.Location  = New-Object System.Drawing.Point(12, ($alturaMid - 32))
    $lblSubsDestinoStatus.AutoSize  = $true
    $gbSubsDestino.Controls.Add($lblSubsDestinoStatus)

$y += $alturaMid + 6

# ================================================================
# PASO 4 - BOTON EJECUTAR (ancho completo)
# ================================================================
$gbEjecutar = New-Object System.Windows.Forms.GroupBox
$gbEjecutar.Text      = "  PASO 4  -  Ejecutar la transferencia"
$gbEjecutar.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$gbEjecutar.ForeColor = [System.Drawing.Color]::FromArgb(0, 80, 160)
$gbEjecutar.Location  = New-Object System.Drawing.Point($margen, $y)
$gbEjecutar.Size      = New-Object System.Drawing.Size($anchoTotal, $alturaBtn)
$gbEjecutar.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$form.Controls.Add($gbEjecutar)

    $btnEjecutar = New-Object System.Windows.Forms.Button
    $btnEjecutar.Text      = "PULSA PARA EJECUTAR TRANSFERENCIA"
    $btnEjecutar.Location  = New-Object System.Drawing.Point(12, 20)
    $btnEjecutar.Size      = New-Object System.Drawing.Size(($anchoTotal - 26), 40)
    $btnEjecutar.BackColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
    $btnEjecutar.ForeColor = [System.Drawing.Color]::White
    $btnEjecutar.FlatStyle = "Flat"
    $btnEjecutar.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnEjecutar.Enabled   = $false
    $btnEjecutar.Cursor    = "Hand"
    $gbEjecutar.Controls.Add($btnEjecutar)

$y += $alturaBtn + 4

# Barra de progreso
$progBar = New-Object System.Windows.Forms.ProgressBar
$progBar.Location = New-Object System.Drawing.Point($margen, $y)
$progBar.Size     = New-Object System.Drawing.Size($anchoTotal, 10)
$progBar.Minimum  = 0
$progBar.Maximum  = 100
$progBar.Value    = 0
$progBar.Style    = "Continuous"
$form.Controls.Add($progBar)

$lblProgresoTxt = New-Object System.Windows.Forms.Label
$lblProgresoTxt.Text      = "Progreso: en espera"
$lblProgresoTxt.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblProgresoTxt.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$lblProgresoTxt.Location  = New-Object System.Drawing.Point($margen, ($y + 12))
$lblProgresoTxt.AutoSize  = $true
$form.Controls.Add($lblProgresoTxt)
$y += 28

# ================================================================
# LOG
# ================================================================
$lblLogHdr = New-Object System.Windows.Forms.Label
$lblLogHdr.Text      = "Registro de actividad:"
$lblLogHdr.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblLogHdr.ForeColor = [System.Drawing.Color]::Black
$lblLogHdr.Location  = New-Object System.Drawing.Point($margen, $y)
$lblLogHdr.AutoSize  = $true
$form.Controls.Add($lblLogHdr)
$y += 16

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location    = New-Object System.Drawing.Point($margen, $y)
$txtLog.Size        = New-Object System.Drawing.Size($anchoTotal, 110)
$txtLog.BackColor   = [System.Drawing.Color]::Black
$txtLog.ForeColor   = [System.Drawing.Color]::Lime
$txtLog.Font        = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.ReadOnly    = $true
$txtLog.BorderStyle = "None"
$txtLog.ScrollBars  = "Vertical"
$form.Controls.Add($txtLog)
$y += 115

$form.ClientSize = New-Object System.Drawing.Size($formW, $y)

# ================================================================
# TIMER: ANIMACION DE PULSO EN BOTON EJECUTAR
# ================================================================
$timerAnim = New-Object System.Windows.Forms.Timer
$timerAnim.Interval = 55
$timerAnim.Add_Tick({
    if (-not $btnEjecutar.Enabled) { return }
    $global:animStep += $global:animDir
    if ($global:animStep -ge 30) { $global:animDir = -1 }
    if ($global:animStep -le 0)  { $global:animDir = 1 }
    # Pulsa entre verde oscuro (0, 130, 58) y verde brillante (30, 210, 95)
    $t = $global:animStep / 30.0
    $r = [int](  0 + $t * 30)
    $g = [int](130 + $t * 80)
    $b = [int]( 58 + $t * 37)
    $btnEjecutar.BackColor = [System.Drawing.Color]::FromArgb($r, $g, $b)
})
$timerAnim.Start()

# ================================================================
# EVENTOS
# ================================================================

# --- MARCAR / DESMARCAR TODOS ---
$btnMarcarTodos.Add_Click({
    for ($i = 0; $i -lt $lstSubs.Items.Count; $i++) { $lstSubs.Items[$i].Checked = $true }
    Actualizar-ContadorSubs
})

$btnDesmarcarTodos.Add_Click({
    for ($i = 0; $i -lt $lstSubs.Items.Count; $i++) { $lstSubs.Items[$i].Checked = $false }
    Actualizar-ContadorSubs
})

$lstSubs.Add_ItemChecked({ Actualizar-ContadorSubs })

# --- IMPORTAR DESDE ARCHIVO ---
$btnImportar.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Seleccionar archivo de usuarios (Excel, CSV o Texto)"
    $dialog.Filter = "Archivos de usuarios (*.xlsx;*.xls;*.csv;*.txt)|*.xlsx;*.xls;*.csv;*.txt|Archivos de Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Archivos CSV (*.csv)|*.csv|Archivos de Texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    
    if ($dialog.ShowDialog() -ne "OK") { return }
    $filePath = $dialog.FileName
    
    Write-Log "Importando usuarios desde archivo: '$($dialog.SafeFileName)'..." "Info"
    $usernames = @()
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    
    if ($ext -eq ".txt") {
        try {
            $usernames = Get-Content $filePath | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" -and -not $_.StartsWith("#") }
        } catch {
            Write-Log "Error al leer archivo TXT: $_" "Error"
            return
        }
    } elseif ($ext -eq ".csv") {
        try {
            $lines = Get-Content $filePath
            if ($lines.Count -gt 0) {
                $firstLine = $lines[0]
                $delimiter = ","
                if ($firstLine -match ";") { $delimiter = ";" }
                
                $csvData = Import-Csv -Path $filePath -Delimiter $delimiter
                if ($csvData.Count -gt 0) {
                    $firstRow = $csvData[0]
                    $properties = $firstRow.PSObject.Properties | Select-Object -ExpandProperty Name
                    
                    # Escaneo avanzado de cabeceras para CSV
                    $bestScore = -1
                    $userColumn = $null
                    foreach ($prop in $properties) {
                        $propLower = $prop.ToLower().Trim()
                        if ($propLower -match "^(usuario\s+adm|adm)$") { continue }
                        $score = -1
                        if ($propLower -match "^(usuario\s+nominal|nominal|n[oó]minal)$") { $score = 10 }
                        elseif ($propLower -match "^(samaccountname|login|usuario|user|username)$") { $score = 9 }
                        elseif ($propLower -match "^(email|correo|userprincipalname|upn)$") { $score = 7 }
                        elseif ($propLower -match "^(nombre|role|displayname|nombre\s+completo)$") { $score = 5 }
                        
                        if ($score -gt $bestScore) {
                            $bestScore = $score
                            $userColumn = $prop
                        }
                    }
                    if ($userColumn -eq $null) {
                        $userColumn = $properties[0]
                    }
                    $usernames = $csvData | ForEach-Object { $_.$userColumn } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                } else {
                    $usernames = $lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                }
            }
        } catch {
            Write-Log "Error al leer archivo CSV: $_" "Error"
            return
        }
    } elseif ($ext -eq ".xlsx" -or $ext -eq ".xls") {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($filePath)
            $sheet = $workbook.Sheets.Item(1)
            
            # Obtener ultima fila y columna absolutas usando SpecialCells(11) [xlCellTypeLastCell]
            $lastCell = $sheet.Cells.SpecialCells(11)
            $rowsCount = $lastCell.Row
            $colsCount = $lastCell.Column

            # Escaneo avanzado de cabeceras en matriz de 10x10
            $bestScore = -1
            $bestRow = 1
            $bestCol = 1

            for ($r = 1; $r -le [Math]::Min($rowsCount, 10); $r++) {
                for ($c = 1; $c -le [Math]::Min($colsCount, 10); $c++) {
                    $val = [string]($sheet.Cells.Item($r, $c).Value2)
                    if ([string]::IsNullOrWhiteSpace($val)) { continue }
                    $valLower = $val.ToLower().Trim()
                    if ($valLower -match "^(usuario\s+adm|adm)$") { continue }
                    
                    $score = -1
                    if ($valLower -match "^(usuario\s+nominal|nominal|n[oó]minal)$") { $score = 10 }
                    elseif ($valLower -match "^(samaccountname|login|usuario|user|username)$") { $score = 9 }
                    elseif ($valLower -match "^(email|correo|userprincipalname|upn)$") { $score = 7 }
                    elseif ($valLower -match "^(nombre|role|displayname|nombre\s+completo)$") { $score = 5 }
                    
                    if ($score -gt $bestScore) {
                        $bestScore = $score
                        $bestRow = $r
                        $bestCol = $c
                    }
                }
            }

            # Si no se encontro cabecera, buscar primera columna con datos
            if ($bestScore -eq -1) {
                for ($c = 1; $c -le [Math]::Min($colsCount, 10); $c++) {
                    $nonEmptyCount = 0
                    for ($r = 1; $r -le [Math]::Min($rowsCount, 20); $r++) {
                        $val = [string]($sheet.Cells.Item($r, $c).Value2)
                        if (-not [string]::IsNullOrWhiteSpace($val)) { $nonEmptyCount++ }
                    }
                    if ($nonEmptyCount -gt 3) {
                        $bestCol = $c
                        $bestRow = 0
                        break
                    }
                }
            }

            # Leer filas a partir de la fila siguiente al encabezado (o desde 1 si no hay encabezado)
            $startRow = $bestRow + 1
            for ($r = $startRow; $r -le $rowsCount; $r++) {
                $val = [string]($sheet.Cells.Item($r, $bestCol).Value2)
                if (-not [string]::IsNullOrWhiteSpace($val)) {
                    $val = $val.Trim()
                    if ($val -match "^\d+[\.,]0+$") { $val = $val -replace "[\.,]0+$", "" }
                    $usernames += $val
                }
            }

            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($lastCell) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        } catch {
            Write-Log "Error al leer archivo Excel: $_" "Error"
            Write-Log "Asegurate de que Excel este instalado en tu equipo, o guarda el archivo como CSV/Texto." "Warn"
            return
        }
    }

    if ($usernames.Count -eq 0) {
        Write-Log "No se encontraron nombres de usuario en el archivo." "Warn"
        [System.Windows.Forms.MessageBox]::Show("No se encontraron usuarios en el archivo.", "Archivo vacio", "OK", "Warning") | Out-Null
        return
    }

    Write-Log "Se leyeron $($usernames.Count) entradas. Buscando en Active Directory..." "Info"
    if ($usernames.Count -gt 0) {
        $first5 = $usernames | Select-Object -First 5
        Write-Log "Primeras entradas leidas del archivo: $($first5 -join ', ')" "Info"
    }
    $lstSubs.Items.Clear()
    $global:subordinados = @()
    $global:managerOrigen = $null
    $lstOrigenRes.Items.Clear()
    $global:resultadosOrigen = @()
    
    $lblOrigenSel.Text      = "LISTA IMPORTADA DESDE ARCHIVO"
    $lblOrigenSel.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 50)
    $lblOrigenSel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $foundUsers = @()
    $notFound = @()
    
    # Mostrar estado de carga en Paso 2
    $lblSubsStatus.Text      = "Buscando usuarios importados en AD..."
    $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    [System.Windows.Forms.Application]::DoEvents()

    # Deshabilitar el formulario principal temporalmente
    $form.Enabled = $false

    # Crear ventana de carga (Loading Popup)
    $loadingForm = New-Object System.Windows.Forms.Form
    $loadingForm.Text            = "Cargando Usuarios"
    $loadingForm.Size            = New-Object System.Drawing.Size(420, 130)
    $loadingForm.StartPosition   = "Manual"
    $loadingForm.FormBorderStyle = "FixedToolWindow"
    $loadingForm.BackColor       = [System.Drawing.Color]::FromArgb(28, 28, 35)
    $loadingForm.ControlBox      = $false
    $loadingForm.ShowInTaskbar   = $false
    
    # Centrar respecto al formulario principal
    $loadingForm.Location = New-Object System.Drawing.Point(
        ($form.Location.X + ($form.Width - $loadingForm.Width) / 2),
        ($form.Location.Y + ($form.Height - $loadingForm.Height) / 2)
    )
    
    $lblLoadingTitle = New-Object System.Windows.Forms.Label
    $lblLoadingTitle.Text      = "Buscando usuarios en Active Directory..."
    $lblLoadingTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblLoadingTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 210, 120)
    $lblLoadingTitle.Location  = New-Object System.Drawing.Point(18, 15)
    $lblLoadingTitle.Size      = New-Object System.Drawing.Size(384, 22)
    $loadingForm.Controls.Add($lblLoadingTitle)
    
    $lblLoadingUser = New-Object System.Windows.Forms.Label
    $lblLoadingUser.Text      = "Iniciando busqueda..."
    $lblLoadingUser.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblLoadingUser.ForeColor = [System.Drawing.Color]::White
    $lblLoadingUser.Location  = New-Object System.Drawing.Point(18, 40)
    $lblLoadingUser.Size      = New-Object System.Drawing.Size(384, 20)
    $loadingForm.Controls.Add($lblLoadingUser)
    
    $loadingBar = New-Object System.Windows.Forms.ProgressBar
    $loadingBar.Location = New-Object System.Drawing.Point(18, 65)
    $loadingBar.Size     = New-Object System.Drawing.Size(368, 14)
    $loadingBar.Minimum  = 0
    $loadingBar.Maximum  = 100
    $loadingBar.Value    = 0
    $loadingForm.Controls.Add($loadingBar)
    
    $loadingForm.Show()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $idxUser = 0
        $totalUsers = $usernames.Count
        
        foreach ($userRaw in $usernames) {
            $idxUser++
            $uName = $userRaw.Trim()
            if ([string]::IsNullOrEmpty($uName)) { continue }
            
            $lblLoadingUser.Text = "Procesando $idxUser de $totalUsers : $uName"
            $loadingBar.Value = [int](($idxUser / $totalUsers) * 100)
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                $adUser = @()
                
                # 1. Intentar busqueda exacta con el original y con el relleno de ceros (si es numerico < 8 digitos)
                $termsExact = @($uName)
                if ($uName -match "^\d+$" -and $uName.Length -lt 8) {
                    $termsExact += $uName.PadLeft(8, '0')
                }
                
                foreach ($term in $termsExact) {
                    $adUser = @(Get-ADUser -Filter "SamAccountName -eq '$term' -or DisplayName -eq '$term' -or UserPrincipalName -eq '$term'" -Properties DisplayName, SamAccountName, Department, Title)
                    if ($adUser.Count -gt 0) { break }
                }
                
                # 2. Si no se encontro exacto, intentar busqueda parcial
                if ($adUser.Count -eq 0) {
                    foreach ($term in $termsExact) {
                        $adUser = @(Get-ADUser -Filter "DisplayName -like '*$term*' -or SamAccountName -like '*$term*'" -Properties DisplayName, SamAccountName, Department, Title)
                        if ($adUser.Count -gt 0) { break }
                    }
                }
                
                if ($adUser.Count -eq 1) {
                    $foundUsers += $adUser[0]
                } elseif ($adUser.Count -gt 1) {
                    $foundUsers += $adUser[0]
                    Write-Log "Multiples coincidencias para '$uName', seleccionado '$($adUser[0].DisplayName)'." "Warn"
                } else {
                    $notFound += $uName
                }
            } catch {
                $notFound += $uName
                Write-Log "Error buscando '$uName' en AD: $_" "Error"
            }
        }
    } finally {
        # Garantizar el cierre y restauracion de la UI
        $loadingForm.Close()
        $form.Enabled = $true
        $form.Activate()
    }

    # Cargar en el Paso 2
    $global:subordinados = @($foundUsers)
    foreach ($s in $foundUsers) {
        $item = New-Object System.Windows.Forms.ListViewItem("")
        $item.Checked = $true
        $item.SubItems.Add($s.DisplayName) | Out-Null
        $item.SubItems.Add($s.SamAccountName) | Out-Null
        $item.SubItems.Add($(if ($s.Department) { $s.Department } else { "---" })) | Out-Null
        $item.SubItems.Add($(if ($s.Title)      { $s.Title }      else { "---" })) | Out-Null
        $lstSubs.Items.Add($item) | Out-Null
    }

    if ($foundUsers.Count -eq 0) {
        $lblSubsStatus.Text      = "No se identifico ningun usuario en Active Directory."
        $lblSubsStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "No se identifico ningun usuario del archivo en Active Directory." "Error"
    } else {
        $lblSubsStatus.Text      = "$($foundUsers.Count) de $($foundUsers.Count) seleccionados  -  listos para ser transferidos"
        $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
        Write-Log "Importacion finalizada. $($foundUsers.Count) usuario(s) identificado(s) de AD." "OK"
    }

    if ($notFound.Count -gt 0) {
        $msgAviso = "$($notFound.Count) Usuario/s no fueron encontrados en AD"
        Write-Log "$($msgAviso): $($notFound -join ', ')" "Warn"
        
        $avisoForm = New-Object System.Windows.Forms.Form
        $avisoForm.Text            = "Aviso - Usuarios no encontrados"
        $avisoForm.Size            = New-Object System.Drawing.Size(550, 320)
        $avisoForm.StartPosition   = "CenterParent"
        $avisoForm.FormBorderStyle = "FixedDialog"
        $avisoForm.MaximizeBox     = $false
        $avisoForm.MinimizeBox     = $false
        $avisoForm.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 240)
        
        $avisoHeader = New-Object System.Windows.Forms.Panel
        $avisoHeader.Size      = New-Object System.Drawing.Size(550, 52)
        $avisoHeader.BackColor = [System.Drawing.Color]::FromArgb(215, 120, 0)
        $avisoHeader.Location  = New-Object System.Drawing.Point(0, 0)
        $avisoForm.Controls.Add($avisoHeader)
        
        $lblAvisoTitle = New-Object System.Windows.Forms.Label
        $lblAvisoTitle.Text      = "Aviso: Usuario/s no encontradas en AD"
        $lblAvisoTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $lblAvisoTitle.ForeColor = [System.Drawing.Color]::White
        $lblAvisoTitle.Location  = New-Object System.Drawing.Point(16, 14)
        $lblAvisoTitle.AutoSize  = $true
        $avisoHeader.Controls.Add($lblAvisoTitle)
        
        $txtAvisoReport = New-Object System.Windows.Forms.RichTextBox
        $txtAvisoReport.Location    = New-Object System.Drawing.Point(16, 68)
        $txtAvisoReport.Size        = New-Object System.Drawing.Size(502, 145)
        $txtAvisoReport.Font        = New-Object System.Drawing.Font("Consolas", 10)
        $txtAvisoReport.ReadOnly    = $true
        $txtAvisoReport.BackColor   = [System.Drawing.Color]::FromArgb(255, 252, 240)
        $txtAvisoReport.ForeColor   = [System.Drawing.Color]::FromArgb(120, 70, 0)
        $txtAvisoReport.BorderStyle = "FixedSingle"
        $txtAvisoReport.ScrollBars  = "Vertical"
        $avisoForm.Controls.Add($txtAvisoReport)
        
        $txtAvisoReport.Text = "$($msgAviso):`n`n$($notFound -join ', ')"
        
        $btnCerrarAviso = New-Object System.Windows.Forms.Button
        $btnCerrarAviso.Text      = "Entendido"
        $btnCerrarAviso.Location  = New-Object System.Drawing.Point(195, 230)
        $btnCerrarAviso.Size      = New-Object System.Drawing.Size(145, 34)
        $btnCerrarAviso.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
        $btnCerrarAviso.ForeColor = [System.Drawing.Color]::White
        $btnCerrarAviso.FlatStyle = "Flat"
        $btnCerrarAviso.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnCerrarAviso.Cursor    = "Hand"
        $btnCerrarAviso.Add_Click({ $avisoForm.Close() })
        $avisoForm.Controls.Add($btnCerrarAviso)
        
        $avisoForm.ShowDialog($form) | Out-Null
    }
    
    Actualizar-Boton
})

# --- BUSCAR ORIGEN ---
$btnBuscaOrigen.Add_Click({
    $lstOrigenRes.Items.Clear()
    $global:resultadosOrigen = @()
    $global:managerOrigen    = $null
    $lblOrigenSel.Text       = "Ningun manager origen seleccionado"
    $lblOrigenSel.Font       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblOrigenSel.ForeColor  = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lstSubs.Items.Clear()
    $global:subordinados     = @()
    $lblSubsStatus.Text      = "Selecciona un manager origen en el Paso 1 para ver aqui sus Trabajadores."
    $lblSubsStatus.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    Actualizar-Boton

    $t = $txtBuscaOrigen.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) {
        [System.Windows.Forms.MessageBox]::Show("Escribe un nombre o usuario para buscar.", "Campo vacio", "OK", "Warning") | Out-Null
        return
    }

    Write-Log "Buscando manager origen: '$t'..." "Info"
    try {
        $res = Get-ADUser -Filter "DisplayName -like '*$t*' -or SamAccountName -like '*$t*'" `
                   -Properties DisplayName, SamAccountName, Department, Title |
                   Sort-Object DisplayName
        $global:resultadosOrigen = @($res)

        foreach ($u in $res) {
            $d = if ($u.Department) { "  |  $($u.Department)" } else { "" }
            $lstOrigenRes.Items.Add("$($u.DisplayName)  [$($u.SamAccountName)]$d") | Out-Null
        }

        if ($res.Count -eq 0) {
            Write-Log "Sin resultados para '$t'. Prueba con otro termino." "Warn"
        } else {
            # Seleccionar siempre el primer resultado automaticamente
            $lstOrigenRes.SelectedIndex = 0
            if ($res.Count -eq 1) {
                Write-Log "1 usuario encontrado y seleccionado automaticamente." "OK"
            } else {
                Write-Log "$($res.Count) usuario(s) encontrado(s). Primer resultado seleccionado automaticamente." "OK"
            }
        }
    } catch {
        Write-Log "Error en busqueda origen: $_" "Error"
    }
})

$txtBuscaOrigen.Add_KeyDown({ param($s,$e); if ($e.KeyCode -eq "Return") { $btnBuscaOrigen.PerformClick() } })

$lstOrigenRes.Add_SelectedIndexChanged({
    $i = $lstOrigenRes.SelectedIndex
    if ($i -lt 0 -or $i -ge $global:resultadosOrigen.Count) { return }
    $global:managerOrigen   = $global:resultadosOrigen[$i]
    $dept = if ($global:managerOrigen.Department) { "  |  $($global:managerOrigen.Department)" } else { "" }
    $lblOrigenSel.Text      = "SELECCIONADO:  $($global:managerOrigen.DisplayName)  [$($global:managerOrigen.SamAccountName)]$dept"
    $lblOrigenSel.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
    $lblOrigenSel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    Write-Log "Manager origen seleccionado: $($global:managerOrigen.DisplayName) [$($global:managerOrigen.SamAccountName)]" "Step"
    Cargar-Subordinados
})

# --- BUSCAR DESTINO ---
$btnBuscaDestino.Add_Click({
    $lstDestinoRes.Items.Clear()
    $global:resultadosDestino = @()
    $global:managerDestino    = $null
    $lblDestinoSel.Text       = "Ningun manager destino seleccionado"
    $lblDestinoSel.Font       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblDestinoSel.ForeColor  = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $lstSubsDestino.Items.Clear()
    $lblSubsDestinoStatus.Text      = "Selecciona manager destino para ver sus Trabajadores actuales."
    $lblSubsDestinoStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    Actualizar-Boton

    $t = $txtBuscaDestino.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($t)) {
        [System.Windows.Forms.MessageBox]::Show("Escribe un nombre o usuario para buscar.", "Campo vacio", "OK", "Warning") | Out-Null
        return
    }

    Write-Log "Buscando manager destino: '$t'..." "Info"
    try {
        $res = Get-ADUser -Filter "DisplayName -like '*$t*' -or SamAccountName -like '*$t*'" `
                   -Properties DisplayName, SamAccountName, Department, Title |
                   Sort-Object DisplayName
        $global:resultadosDestino = @($res)

        foreach ($u in $res) {
            $d = if ($u.Department) { "  |  $($u.Department)" } else { "" }
            $lstDestinoRes.Items.Add("$($u.DisplayName)  [$($u.SamAccountName)]$d") | Out-Null
        }

        if ($res.Count -eq 0) {
            Write-Log "Sin resultados para '$t'. Prueba con otro termino." "Warn"
        } else {
            # Seleccionar siempre el primer resultado automaticamente
            $lstDestinoRes.SelectedIndex = 0
            if ($res.Count -eq 1) {
                Write-Log "1 usuario encontrado y seleccionado automaticamente." "OK"
            } else {
                Write-Log "$($res.Count) usuario(s) encontrado(s). Primer resultado seleccionado automaticamente." "OK"
            }
        }
    } catch {
        Write-Log "Error en busqueda destino: $_" "Error"
    }
})

$txtBuscaDestino.Add_KeyDown({ param($s,$e); if ($e.KeyCode -eq "Return") { $btnBuscaDestino.PerformClick() } })

$lstDestinoRes.Add_SelectedIndexChanged({
    $i = $lstDestinoRes.SelectedIndex
    if ($i -lt 0 -or $i -ge $global:resultadosDestino.Count) { return }
    $global:managerDestino   = $global:resultadosDestino[$i]
    $dept = if ($global:managerDestino.Department) { "  |  $($global:managerDestino.Department)" } else { "" }
    $lblDestinoSel.Text      = "SELECCIONADO:  $($global:managerDestino.DisplayName)  [$($global:managerDestino.SamAccountName)]$dept"
    $lblDestinoSel.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 50)
    $lblDestinoSel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    Write-Log "Manager destino seleccionado: $($global:managerDestino.DisplayName) [$($global:managerDestino.SamAccountName)]" "Step"
    Cargar-SubordinadosDestino
    Actualizar-Boton
})

# --- EJECUTAR TRANSFERENCIA ---
$btnEjecutar.Add_Click({
    $subsSeleccionados = Get-SubordinadosSeleccionados
    if ($global:managerDestino -eq $null -or $subsSeleccionados.Count -eq 0) { return }

    $n   = $subsSeleccionados.Count
    $msg = "Estas a punto de transferir $n trabajador(es) en Active Directory:`n`n"
    if ($global:managerOrigen -ne $null) {
        $msg += "  DE:   $($global:managerOrigen.DisplayName)  [$($global:managerOrigen.SamAccountName)]`n"
    } else {
        $msg += "  DE:   [Lista Importada de Excel/CSV/TXT]`n"
    }
    $msg += "  A:    $($global:managerDestino.DisplayName)  [$($global:managerDestino.SamAccountName)]`n`n"
    $msg += "Esta accion modificara el campo 'Manager' de cada trabajador en AD.`n`nDeseas continuar?"

    $resp = [System.Windows.Forms.MessageBox]::Show(
        $msg, "Confirmar transferencia",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($resp -ne "Yes") {
        Write-Log "Operacion cancelada por el usuario." "Warn"
        return
    }

    # Bloquear UI
    $btnEjecutar.Enabled     = $false
    $btnBuscaOrigen.Enabled  = $false
    $btnBuscaDestino.Enabled = $false
    $progBar.Value = 0

    # Deshabilitar el formulario principal temporalmente
    $form.Enabled = $false

    # Crear ventana de carga (Loading Popup) para la transferencia
    $loadingForm = New-Object System.Windows.Forms.Form
    $loadingForm.Text            = "Ejecutando Transferencia"
    $loadingForm.Size            = New-Object System.Drawing.Size(420, 130)
    $loadingForm.StartPosition   = "Manual"
    $loadingForm.FormBorderStyle = "FixedToolWindow"
    $loadingForm.BackColor       = [System.Drawing.Color]::FromArgb(28, 28, 35)
    $loadingForm.ControlBox      = $false
    $loadingForm.ShowInTaskbar   = $false
    
    # Centrar respecto al formulario principal
    $loadingForm.Location = New-Object System.Drawing.Point(
        ($form.Location.X + ($form.Width - $loadingForm.Width) / 2),
        ($form.Location.Y + ($form.Height - $loadingForm.Height) / 2)
    )
    
    $lblLoadingTitle = New-Object System.Windows.Forms.Label
    $lblLoadingTitle.Text      = "Transfiriendo trabajadores en Active Directory..."
    $lblLoadingTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblLoadingTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 210, 120)
    $lblLoadingTitle.Location  = New-Object System.Drawing.Point(18, 15)
    $lblLoadingTitle.Size      = New-Object System.Drawing.Size(384, 22)
    $loadingForm.Controls.Add($lblLoadingTitle)
    
    $lblLoadingUser = New-Object System.Windows.Forms.Label
    $lblLoadingUser.Text      = "Iniciando transferencia..."
    $lblLoadingUser.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblLoadingUser.ForeColor = [System.Drawing.Color]::White
    $lblLoadingUser.Location  = New-Object System.Drawing.Point(18, 40)
    $lblLoadingUser.Size      = New-Object System.Drawing.Size(384, 20)
    $loadingForm.Controls.Add($lblLoadingUser)
    
    $loadingBar = New-Object System.Windows.Forms.ProgressBar
    $loadingBar.Location = New-Object System.Drawing.Point(18, 65)
    $loadingBar.Size     = New-Object System.Drawing.Size(368, 14)
    $loadingBar.Minimum  = 0
    $loadingBar.Maximum  = 100
    $loadingBar.Value    = 0
    $loadingForm.Controls.Add($loadingBar)
    
    $loadingForm.Show()
    [System.Windows.Forms.Application]::DoEvents()

    Write-Log "=================================" "Step"
    Write-Log "INICIO DE TRANSFERENCIA" "Step"
    if ($global:managerOrigen -ne $null) {
        Write-Log "DE:    $($global:managerOrigen.DisplayName) [$($global:managerOrigen.SamAccountName)]" "Step"
    } else {
        Write-Log "DE:    [Lista Importada de Excel/CSV/TXT]" "Step"
    }
    Write-Log "HACIA: $($global:managerDestino.DisplayName) [$($global:managerDestino.SamAccountName)]" "Step"
    Write-Log "Total: $n trabajador(es) a transferir" "Step"
    Write-Log "---------------------------------" "Step"

    $ok = 0; $ko = 0; $idx = 0
    $dnDestino = $global:managerDestino.DistinguishedName
    $erroresLista = @()

    try {
        foreach ($sub in $subsSeleccionados) {
            $idx++
            $progVal = [int](($idx / $n) * 100)
            $progBar.Value = $progVal
            $loadingBar.Value = $progVal
            
            $statusTxt = "Procesando $idx de $n : $($sub.DisplayName)"
            $lblProgresoTxt.Text = $statusTxt
            $lblLoadingUser.Text = $statusTxt
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # 1. Comprobar si existe el usuario o está deshabilitado
                $adCheck = Get-ADUser -Identity $sub.DistinguishedName -Properties Enabled
                if (-not $adCheck.Enabled) {
                    throw "DESHABILITADA: La cuenta del usuario está deshabilitada en Active Directory."
                }
                
                Set-ADUser -Identity $sub.DistinguishedName -Manager $dnDestino
                Write-Log "[$idx/$n] OK  ---  $($sub.DisplayName)  [$($sub.SamAccountName)]" "OK"
                $ok++
            } catch {
                $exMessage = $_.Exception.Message
                $errReason = ""
                
                if ($exMessage -match "DESHABILITADA") {
                    $errReason = "La cuenta del usuario está DESHABILITADA en Active Directory."
                } elseif ($exMessage -match "Access is denied" -or $exMessage -match "acceso denegado" -or $exMessage -match "privilegios") {
                    $errReason = "Permisos insuficientes en Active Directory (Acceso Denegado)."
                } elseif ($exMessage -match "Cannot find an object" -or $exMessage -match "No se encuentra" -or $exMessage -match "no existe") {
                    $errReason = "El usuario no existe o su DistinguishedName es incorrecto."
                } else {
                    $errReason = ($exMessage -replace '^.*?:\s*', '').Trim()
                    if ([string]::IsNullOrWhiteSpace($errReason)) {
                        $errReason = "Error desconocido al asignar manager en AD."
                    }
                }
                
                Write-Log "[$idx/$n] ERROR  ---  $($sub.DisplayName): $errReason" "Error"
                $erroresLista += [PSCustomObject]@{
                    Nombre = $sub.DisplayName
                    Login  = $sub.SamAccountName
                    Motivo = $errReason
                }
                $ko++
            }
            Start-Sleep -Milliseconds 80
        }
    } finally {
        $loadingForm.Close()
        $form.Enabled = $true
        $form.Activate()
    }

    Write-Log "---------------------------------" "Step"
    $tipoFin = if ($ko -eq 0) { "OK" } else { "Warn" }
    Write-Log "COMPLETADO: $ok transferidos correctamente,  $ko con error." $tipoFin
    Write-Log "=================================" "Step"
    $lblProgresoTxt.Text = "Completado: $ok OK  /  $ko errores"

    $msgFin = "Transferencia finalizada.`n`n  Correctos: $ok`n  Errores:   $ko"
    $icoFin = if ($ko -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
    [System.Windows.Forms.MessageBox]::Show($msgFin, "Resultado final", "OK", $icoFin) | Out-Null

    if ($ko -gt 0) {
        $errForm = New-Object System.Windows.Forms.Form
        $errForm.Text            = "Reporte de Errores - Transferencia AD"
        $errForm.Size            = New-Object System.Drawing.Size(600, 440)
        $errForm.StartPosition   = "CenterParent"
        $errForm.FormBorderStyle = "FixedDialog"
        $errForm.MaximizeBox     = $false
        $errForm.MinimizeBox     = $false
        $errForm.BackColor       = [System.Drawing.Color]::FromArgb(240, 240, 245)
        
        $errHeader = New-Object System.Windows.Forms.Panel
        $errHeader.Size      = New-Object System.Drawing.Size(600, 52)
        $errHeader.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $errHeader.Location  = New-Object System.Drawing.Point(0, 0)
        $errForm.Controls.Add($errHeader)
        
        $lblErrTitle = New-Object System.Windows.Forms.Label
        $lblErrTitle.Text      = "Detalle de Errores detectados en la transferencia"
        $lblErrTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $lblErrTitle.ForeColor = [System.Drawing.Color]::White
        $lblErrTitle.Location  = New-Object System.Drawing.Point(16, 14)
        $lblErrTitle.AutoSize  = $true
        $errHeader.Controls.Add($lblErrTitle)
        
        $txtReport = New-Object System.Windows.Forms.RichTextBox
        $txtReport.Location    = New-Object System.Drawing.Point(16, 68)
        $txtReport.Size        = New-Object System.Drawing.Size(552, 265)
        $txtReport.Font        = New-Object System.Drawing.Font("Consolas", 9.5)
        $txtReport.ReadOnly    = $true
        $txtReport.BackColor   = [System.Drawing.Color]::FromArgb(253, 243, 243)
        $txtReport.ForeColor   = [System.Drawing.Color]::FromArgb(120, 20, 20)
        $txtReport.BorderStyle = "FixedSingle"
        $txtReport.ScrollBars  = "Vertical"
        $errForm.Controls.Add($txtReport)
        
        # Construir reporte
        $reportText = "DETALLE DE ERRORES POR TRABAJADOR:`r`n"
        $reportText += "========================================================`r`n`r`n"
        foreach ($e in $erroresLista) {
            $reportText += "Trabajador:  $($e.Nombre)`r`n"
            $reportText += "Usuario:     $($e.Login)`r`n"
            $reportText += "Motivo:      $($e.Motivo)`r`n"
            $reportText += "--------------------------------------------------------`r`n`r`n"
        }
        $txtReport.Text = $reportText
        
        $btnCerrarErr = New-Object System.Windows.Forms.Button
        $btnCerrarErr.Text      = "Aceptar y Cerrar"
        $btnCerrarErr.Location  = New-Object System.Drawing.Point(220, 348)
        $btnCerrarErr.Size      = New-Object System.Drawing.Size(145, 34)
        $btnCerrarErr.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
        $btnCerrarErr.ForeColor = [System.Drawing.Color]::White
        $btnCerrarErr.FlatStyle = "Flat"
        $btnCerrarErr.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnCerrarErr.Cursor    = "Hand"
        $btnCerrarErr.Add_Click({ $errForm.Close() })
        $errForm.Controls.Add($btnCerrarErr)
        
        $errForm.ShowDialog($form) | Out-Null
    }

    if ($ko -eq 0) { Cargar-Subordinados }
    Cargar-SubordinadosDestino

    $btnBuscaOrigen.Enabled  = $true
    $btnBuscaDestino.Enabled = $true
    Actualizar-Boton
})

# ================================================================
# INICIO
# ================================================================
Write-Log "Script iniciado con permisos de Administrador." "OK"
Write-Log "Modulo ActiveDirectory cargado correctamente." "OK"
Write-Log "Usa el PASO 1 para buscar el manager de origen." "Info"

[System.Windows.Forms.Application]::Run($form)
