# ====================================================================
# ANALIZADOR AVAYA ONEX AGENT V22
# NUEVO: ContactLog.xml como fuente primaria de llamadas y estados
# ====================================================================
Clear-Host
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- PARCHE ANTI-ANSI (Símbolos Seguros) ---
$symPlay  = [char]9658
$symStop  = [char]9632
$symUp    = [char]9650
$symRes   = [char]9654
$symMusic = [char]9834
$symArr   = [char]8594
$symUser  = [char]9679
$symOK    = [char]10003   # ✓  confirmación de éxito

# --- VARIABLES GLOBALES DE SESIÓN ---
$Script:Creds          = $null
$Script:DriveName      = "UnidadAuditoria"
$Script:CurrentIP      = ""
$Script:RutaManual     = ""
$Script:DirFinalGlobal = ""
$Script:ModoContestacion = @{}   # VI_ID -> modo de contestación
$Script:AlertingHoras    = @{}   # UUID(36) -> HH:mm:ss del alerting por llamada

# --- ESTILOS ---
$ColorFondo = [System.Drawing.Color]::FromArgb(30, 30, 30)
$ColorPanel = [System.Drawing.Color]::FromArgb(45, 45, 48)
$ColorTexto = [System.Drawing.Color]::White

# --- FORMULARIO PRINCIPAL ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Analizador OneX V22"
$Form.Size = New-Object System.Drawing.Size(1550, 750)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = $ColorFondo
$Form.ForeColor = $ColorTexto

# --- CONTROLES FILA 1 ---
$lblIP = New-Object System.Windows.Forms.Label; $lblIP.Text = "IP del Equipo:"; $lblIP.Location = New-Object System.Drawing.Point(20, 20); $lblIP.AutoSize = $true
$txtIP = New-Object System.Windows.Forms.TextBox; $txtIP.Location = New-Object System.Drawing.Point(105, 18); $txtIP.Size = New-Object System.Drawing.Size(100, 25); $txtIP.BackColor = $ColorPanel; $txtIP.ForeColor = $ColorTexto

$btnBuscarUsr = New-Object System.Windows.Forms.Button; $btnBuscarUsr.Text = "1. BUSCAR USUARIOS"; $btnBuscarUsr.Location = New-Object System.Drawing.Point(215, 16); $btnBuscarUsr.Size = New-Object System.Drawing.Size(140, 28); $btnBuscarUsr.BackColor = [System.Drawing.Color]::DarkGoldenrod; $btnBuscarUsr.FlatStyle = "Flat"; $btnBuscarUsr.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$lblUsuario = New-Object System.Windows.Forms.Label; $lblUsuario.Text = "Usuario Windows:"; $lblUsuario.Location = New-Object System.Drawing.Point(365, 20); $lblUsuario.AutoSize = $true
$cmbUsuarios = New-Object System.Windows.Forms.ComboBox; $cmbUsuarios.Location = New-Object System.Drawing.Point(475, 18); $cmbUsuarios.Size = New-Object System.Drawing.Size(130, 25); $cmbUsuarios.BackColor = $ColorPanel; $cmbUsuarios.ForeColor = $ColorTexto; $cmbUsuarios.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$lblFecha = New-Object System.Windows.Forms.Label; $lblFecha.Text = "Fecha:"; $lblFecha.Location = New-Object System.Drawing.Point(615, 20); $lblFecha.AutoSize = $true
$dtpFecha = New-Object System.Windows.Forms.DateTimePicker; $dtpFecha.Location = New-Object System.Drawing.Point(660, 18); $dtpFecha.Size = New-Object System.Drawing.Size(105, 25)
$dtpFecha.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom; $dtpFecha.CustomFormat = "dd/MM/yyyy"
$dtpFecha.MinDate = [datetime]::new(2020, 1, 1); $dtpFecha.MaxDate = [datetime]::new(2035, 12, 31)

$btnAnalizar = New-Object System.Windows.Forms.Button; $btnAnalizar.Text = "2. AUDITAR"; $btnAnalizar.Location = New-Object System.Drawing.Point(775, 16); $btnAnalizar.Size = New-Object System.Drawing.Size(100, 28); $btnAnalizar.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnAnalizar.FlatStyle = "Flat"; $btnAnalizar.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnAnalizar.Enabled = $false

$btnRutaManual = New-Object System.Windows.Forms.Button; $btnRutaManual.Text = "3. RUTA MANUAL"; $btnRutaManual.Location = New-Object System.Drawing.Point(885, 16); $btnRutaManual.Size = New-Object System.Drawing.Size(120, 28); $btnRutaManual.BackColor = [System.Drawing.Color]::Teal; $btnRutaManual.FlatStyle = "Flat"; $btnRutaManual.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$btnInfoPC = New-Object System.Windows.Forms.Button; $btnInfoPC.Text = "4. INFO PC"; $btnInfoPC.Location = New-Object System.Drawing.Point(1015, 16); $btnInfoPC.Size = New-Object System.Drawing.Size(80, 28); $btnInfoPC.BackColor = [System.Drawing.Color]::SlateGray; $btnInfoPC.FlatStyle = "Flat"; $btnInfoPC.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnInfoPC.Visible = $false

$btnExtraccion = New-Object System.Windows.Forms.Button; $btnExtraccion.Text = "5. EXTRAER LOG"; $btnExtraccion.Location = New-Object System.Drawing.Point(1100, 16); $btnExtraccion.Size = New-Object System.Drawing.Size(135, 28); $btnExtraccion.BackColor = [System.Drawing.Color]::Maroon; $btnExtraccion.FlatStyle = "Flat"; $btnExtraccion.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnExtraccion.Enabled = $false
$btnExtraccion.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$btnBusqueda = New-Object System.Windows.Forms.Button; $btnBusqueda.Text = "6. BÚSQUEDA"; $btnBusqueda.Location = New-Object System.Drawing.Point(1240, 16); $btnBusqueda.Size = New-Object System.Drawing.Size(95, 28); $btnBusqueda.BackColor = [System.Drawing.Color]::Navy; $btnBusqueda.FlatStyle = "Flat"; $btnBusqueda.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnBusqueda.Enabled = $false
$btnBusqueda.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$btnReinicios = New-Object System.Windows.Forms.Button; $btnReinicios.Text = "7. REINICIOS"; $btnReinicios.Location = New-Object System.Drawing.Point(1340, 16); $btnReinicios.Size = New-Object System.Drawing.Size(90, 28); $btnReinicios.BackColor = [System.Drawing.Color]::Chocolate; $btnReinicios.FlatStyle = "Flat"; $btnReinicios.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnReinicios.Visible = $false
$btnReinicios.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$btnExportarCSV = New-Object System.Windows.Forms.Button; $btnExportarCSV.Text = "8. EXPORT CSV"; $btnExportarCSV.Location = New-Object System.Drawing.Point(1345, 16); $btnExportarCSV.Size = New-Object System.Drawing.Size(85, 28); $btnExportarCSV.BackColor = [System.Drawing.Color]::DarkGreen; $btnExportarCSV.FlatStyle = "Flat"; $btnExportarCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnExportarCSV.Enabled = $false
$btnExportarCSV.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$chkVerCalidad = New-Object System.Windows.Forms.CheckBox; $chkVerCalidad.Text = "Ver Calidad de Red"; $chkVerCalidad.Location = New-Object System.Drawing.Point(1435, 20); $chkVerCalidad.ForeColor = [System.Drawing.Color]::Cyan; $chkVerCalidad.AutoSize = $true; $chkVerCalidad.Checked = $false
$chkVerCalidad.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# --- FILA 2-3: ESTADO E INFO ---
$lblStatus = New-Object System.Windows.Forms.Label; $lblStatus.Location = New-Object System.Drawing.Point(20, 55); $lblStatus.Size = New-Object System.Drawing.Size(1400, 20); $lblStatus.ForeColor = [System.Drawing.Color]::Yellow; $lblStatus.Text = "Ingresa la IP y haz clic en '1. BUSCAR USUARIOS', o usa '3. RUTA MANUAL' si tienes los logs guardados en tu equipo."
$lblExtension = New-Object System.Windows.Forms.Label; $lblExtension.Text = "Ids Detectados: Pendiente..."; $lblExtension.Location = New-Object System.Drawing.Point(20, 80); $lblExtension.AutoSize = $true; $lblExtension.ForeColor = [System.Drawing.Color]::Cyan; $lblExtension.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# --- GRID DE RESULTADOS ---
$GridResultados = New-Object System.Windows.Forms.DataGridView
$GridResultados.Size = New-Object System.Drawing.Size(1500, 580)
$GridResultados.Location = New-Object System.Drawing.Point(20, 110)
$GridResultados.BackgroundColor = $ColorPanel
$GridResultados.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$GridResultados.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10.5)
$GridResultados.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
$GridResultados.AllowUserToAddRows = $false; $GridResultados.RowHeadersVisible = $false; $GridResultados.ReadOnly = $true; $GridResultados.AutoSizeColumnsMode = "Fill"
$GridResultados.ShowCellToolTips = $false
$GridResultados.DefaultCellStyle.BackColor = $ColorPanel; $GridResultados.DefaultCellStyle.ForeColor = $ColorTexto
$GridResultados.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28); $GridResultados.ColumnHeadersDefaultCellStyle.ForeColor = $ColorTexto
$GridResultados.EnableHeadersVisualStyles = $false

$GridResultados.Columns.Add("Hora", "Hora") | Out-Null; $GridResultados.Columns["Hora"].FillWeight = 6
$GridResultados.Columns.Add("Sesion", "Sesión") | Out-Null; $GridResultados.Columns["Sesion"].FillWeight = 4
$GridResultados.Columns.Add("Telefono", "Teléfono") | Out-Null; $GridResultados.Columns["Telefono"].FillWeight = 10
$GridResultados.Columns.Add("Interpretacion", "Actividad del Agente") | Out-Null; $GridResultados.Columns["Interpretacion"].FillWeight = 20
$GridResultados.Columns.Add("EvAgente", "Endpoint.log") | Out-Null; $GridResultados.Columns["EvAgente"].FillWeight = 15
$GridResultados.Columns.Add("EvAudio", "Audio.log") | Out-Null; $GridResultados.Columns["EvAudio"].FillWeight = 15
$GridResultados.Columns.Add("EvAux", "AvayaOneX.log") | Out-Null; $GridResultados.Columns["EvAux"].FillWeight = 10
$GridResultados.Columns.Add("EvIspeac", "Calidad Red/Voz") | Out-Null; $GridResultados.Columns["EvIspeac"].FillWeight = 12; $GridResultados.Columns["EvIspeac"].Visible = $false
$GridResultados.Columns.Add("EvSysLog", "Log de Sistema") | Out-Null; $GridResultados.Columns["EvSysLog"].FillWeight = 10
$GridResultados.Columns.Add("EvAppLog", "Log de Aplicación") | Out-Null; $GridResultados.Columns["EvAppLog"].FillWeight = 10
foreach ($col in $GridResultados.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }

$Form.Controls.AddRange(@($lblIP, $txtIP, $btnBuscarUsr, $lblUsuario, $cmbUsuarios, $btnAnalizar, $lblFecha, $dtpFecha, $btnRutaManual, $btnInfoPC, $btnReinicios, $btnExtraccion, $btnBusqueda, $btnExportarCSV, $chkVerCalidad, $lblExtension, $lblStatus, $GridResultados))

# ====================================================================
# MENÚ CONTEXTUAL DEL GRID: DIAGNÓSTICO DE LLAMADA (botón derecho)
# ====================================================================
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip

$mnuDiag = New-Object System.Windows.Forms.ToolStripMenuItem
$mnuDiag.Text = "Diagnosticar: Por que no hay INICIO DE LLAMADA?"
$mnuDiag.Enabled = $false
[void]$ctxMenu.Items.Add($mnuDiag)

$mnuDiagFin = New-Object System.Windows.Forms.ToolStripMenuItem
$mnuDiagFin.Text = "Diagnosticar: Por que no hay FIN DE LLAMADA?"
$mnuDiagFin.Enabled = $false
[void]$ctxMenu.Items.Add($mnuDiagFin)

$mnuDiagExplicaFin = New-Object System.Windows.Forms.ToolStripMenuItem
$mnuDiagExplicaFin.Text = "Diagnosticar: Por que finalizo esta llamada?"
$mnuDiagExplicaFin.Enabled = $false
[void]$ctxMenu.Items.Add($mnuDiagExplicaFin)

$mnuDiagDir = New-Object System.Windows.Forms.ToolStripMenuItem
$mnuDiagDir.Text = "Diagnosticar: Por que es Entrante o Saliente?"
$mnuDiagDir.Enabled = $false
[void]$ctxMenu.Items.Add($mnuDiagDir)

$GridResultados.ContextMenuStrip = $ctxMenu

# Cancelar apertura del menú si ningún ítem está habilitado
$ctxMenu.Add_Opening({
    param($sender, $e)
    if (-not $mnuDiag.Enabled -and -not $mnuDiagFin.Enabled -and -not $mnuDiagExplicaFin.Enabled -and -not $mnuDiagDir.Enabled) { $e.Cancel = $true }
})

$GridResultados.Add_MouseDown({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hit = $GridResultados.HitTest($e.X, $e.Y)
        if ($hit.RowIndex -ge 0) {
            $GridResultados.ClearSelection()
            $GridResultados.Rows[$hit.RowIndex].Selected = $true
            $GridResultados.CurrentCell = $GridResultados.Rows[$hit.RowIndex].Cells[0]
            $CeldaInterp = $GridResultados.Rows[$hit.RowIndex].Cells["Interpretacion"].Value
            $mnuDiag.Enabled           = ($CeldaInterp -match "señal de llamada") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDiagFin.Enabled        = ($CeldaInterp -match "INICIO DE LLAMADA|LÍNEA ABIERTA") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDiagExplicaFin.Enabled = ($CeldaInterp -match "FIN DE LLAMADA|CUELGUE MANUAL") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDiagDir.Enabled        = ($CeldaInterp -match "INICIO DE LLAMADA") -and ($null -ne $Script:SnapEventosTiempo)
        } else { $mnuDiag.Enabled = $false; $mnuDiagFin.Enabled = $false; $mnuDiagExplicaFin.Enabled = $false; $mnuDiagDir.Enabled = $false }
    }
})

# ====================================================================
# FUNCIÓN GLOBAL PARA RESETEAR LA INTERFAZ
# ====================================================================
function Reset-Entorno {
    $GridResultados.Rows.Clear()
    $lblExtension.Text = "Ids Detectados: Pendiente..."
    $btnAnalizar.Enabled = $false
    $btnExtraccion.Enabled = $false
    $btnBusqueda.Enabled = $false
    $btnExportarCSV.Enabled = $false
    $btnInfoPC.Visible = $false
    $Script:DirFinalGlobal = ""
    $dtpFecha.MinDate = [datetime]::Parse("01/01/1753")
    $dtpFecha.MaxDate = [datetime]::Parse("31/12/9998")
    $dtpFecha.Value = Get-Date
}

# ====================================================================
# EVENTO DEL CHECKBOX DE CALIDAD
# ====================================================================
$chkVerCalidad.Add_CheckedChanged({
    $GridResultados.SuspendLayout()
    $GridResultados.CurrentCell = $null
    $GridResultados.Columns["EvIspeac"].Visible = $chkVerCalidad.Checked
    foreach ($row in $GridResultados.Rows) {
        if ($row.Cells["Interpretacion"].Value -eq "Llamada en curso (Analizando Calidad)") { $row.Visible = $chkVerCalidad.Checked }
    }
    $GridResultados.ResumeLayout()
})

# ====================================================================
# ACCIÓN 4: INFO PC (BOTÓN OCULTO)
# ====================================================================
$btnInfoPC.Add_Click({
    $TargetIP = $txtIP.Text.Trim()
    if (-not $TargetIP) { return }
    if (-not $Script:Creds) { try { $Script:Creds = Get-Credential -UserName "local\soporte" -Message "Credenciales para $TargetIP" -ErrorAction Stop } catch { return } }
    $lblStatus.Text = "Recopilando info del hardware de $TargetIP..."; $Form.Refresh()
    try {
        $CPU    = Get-WmiObject Win32_Processor -ComputerName $TargetIP -Credential $Script:Creds -EA Stop | Select-Object -First 1
        $RAM    = Get-WmiObject Win32_PhysicalMemory -ComputerName $TargetIP -Credential $Script:Creds -EA SilentlyContinue | Measure-Object -Property Capacity -Sum
        $OS     = Get-WmiObject Win32_OperatingSystem -ComputerName $TargetIP -Credential $Script:Creds -EA SilentlyContinue
        $NIC    = Get-WmiObject Win32_NetworkAdapter -ComputerName $TargetIP -Credential $Script:Creds -EA SilentlyContinue | Where-Object { $_.NetEnabled -eq $true -and $_.Speed -ne $null } | Select-Object -First 1
        $Users  = Get-WmiObject Win32_ComputerSystem -ComputerName $TargetIP -Credential $Script:Creds -EA SilentlyContinue

        $RamGB    = if ($RAM.Sum) { [math]::Round($RAM.Sum / 1GB, 1) } else { "N/A" }
        $OSVer    = if ($OS) { "$($OS.Caption) ($($OS.OSArchitecture))" } else { "N/A" }
        $NICSpeed = if ($NIC -and $NIC.Speed) { "$([math]::Round($NIC.Speed/1MB)) Mbps" } else { "N/A" }
        $Uptime   = if ($OS) { $OS.ConvertToDateTime($OS.LastBootUpTime) } else { $null }
        $UptimeStr = if ($Uptime) { $td = (Get-Date) - $Uptime; "$($td.Days)d $($td.Hours)h $($td.Minutes)m" } else { "N/A" }
        $UserOn   = if ($Users) { $Users.UserName } else { "N/A" }

        $Msg  = "=== REPORTE DE HARDWARE: $TargetIP ===`n`n"
        $Msg += "CPU    : $($CPU.Name)`n"
        $Msg += "RAM    : $RamGB GB`n"
        $Msg += "S.O.   : $OSVer`n"
        $Msg += "Red    : $NICSpeed`n"
        $Msg += "Uptime : $UptimeStr`n"
        $Msg += "Usuario: $UserOn`n"
        [System.Windows.Forms.MessageBox]::Show($Msg, "Info PC - $TargetIP", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch { [System.Windows.Forms.MessageBox]::Show("Error al conectar: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    $lblStatus.Text = "Listo."
})

# ====================================================================
# ACCIÓN 1: BUSCAR USUARIOS
# ====================================================================
$btnBuscarUsr.Add_Click({
    $TargetIP = $txtIP.Text.Trim()
    if (-not $TargetIP) { [System.Windows.Forms.MessageBox]::Show("Ingresa una IP.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }

    # --- Ping rapido: abortar si el equipo no responde (evita timeout de conexion) ---
    $lblStatus.Text = "Verificando disponibilidad de $TargetIP..."; $Form.Refresh()
    if (-not (Test-Connection -ComputerName $TargetIP -Count 1 -EA SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("El equipo $TargetIP no responde al ping.`nPuede estar apagado o fuera de la red.", "Equipo no disponible", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $lblStatus.Text = "Ping fallido — equipo no disponible."
        return
    }

    # --- Credenciales: reutilizar las guardadas; pedir solo si no hay ninguna ---
    if ($null -eq $Script:Creds) {
        try { $Script:Creds = Get-Credential -UserName "local\soporte" -Message "Credenciales para \\$TargetIP\c$" -ErrorAction Stop } catch { return }
    }

    $Script:RutaManual = ""   # limpiar ruta manual al cambiar a modo red
    $lblStatus.Text = "Conectando a $TargetIP..."; $Form.Refresh()
    Reset-Entorno
    try {
        if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null }
        New-PSDrive -Name $Script:DriveName -PSProvider FileSystem -Root "\\$TargetIP\c$" -Credential $Script:Creds -EA Stop | Out-Null
        $Excluir = @("Default", "Public", "All Users", "Default User", "desktop.ini")
        $Usuarios = Get-ChildItem "$($Script:DriveName):\Users" -EA SilentlyContinue | Where-Object { $_.PSIsContainer -and $Excluir -notcontains $_.Name } | Select-Object -ExpandProperty Name
        $cmbUsuarios.Items.Clear()
        if ($Usuarios) { foreach ($u in $Usuarios) { $cmbUsuarios.Items.Add($u) | Out-Null }; $cmbUsuarios.SelectedIndex = 0 }
        $btnAnalizar.Enabled = $true
        $btnInfoPC.Visible = $true
        $Script:CurrentIP = $TargetIP
        $lblStatus.Text = "Conexion exitosa. Selecciona un usuario y fecha, luego clic en AUDITAR."
    } catch {
        # Si el error es de autenticacion, borrar credenciales para que el proximo intento las pida de nuevo
        if ($_.Exception.Message -match "(?i)acceso denegado|access.denied|logon|credencial|1326") { $Script:Creds = $null }
        [System.Windows.Forms.MessageBox]::Show("No se pudo conectar a $TargetIP.`n$($_.Exception.Message)", "Error de Conexion", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $lblStatus.Text = "Error al conectar."
    } finally {
        if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null }
    }
})

# ====================================================================
# ACCIÓN 3: RUTA MANUAL
# ====================================================================
$btnRutaManual.Add_Click({
    $FolderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDlg.Description = "Selecciona la carpeta que contiene los logs de Avaya (Log Files)"
    if ($FolderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:RutaManual = $FolderDlg.SelectedPath
        $lblStatus.Text = "Ruta manual cargada: $($Script:RutaManual)"
        Reset-Entorno
        # Auto-ajuste del DateTimePicker según fechas disponibles en los logs
        $LogsDisponibles = Get-ChildItem -Path $Script:RutaManual -Filter "EndpointLog.txt*" -EA SilentlyContinue
        if ($LogsDisponibles) {
            $FechasLog = @()
            foreach ($lf in $LogsDisponibles) {
                $contenido = Get-Content $lf.FullName -TotalCount 50 -EA SilentlyContinue
                foreach ($linea in $contenido) {
                    if ($linea -match "(\d{2}/\d{2}/\d{4})") { try { $FechasLog += [datetime]::ParseExact($matches[1],"dd/MM/yyyy",$null) } catch {} }
                    if ($linea -match "(\d{4}-\d{2}-\d{2})") { try { $FechasLog += [datetime]::Parse($matches[1]) } catch {} }
                }
            }
            if ($FechasLog.Count -gt 0) {
                $FMin = ($FechasLog | Sort-Object)[0]; $FMax = ($FechasLog | Sort-Object)[-1]
                # Solo auto-seleccionamos la fecha más reciente — NO bloqueamos MinDate/MaxDate
                # (si MinDate=MaxDate el picker queda bloqueado y el usuario no puede cambiar la fecha)
                if ($FMax -ge $dtpFecha.MinDate -and $FMax -le $dtpFecha.MaxDate) { $dtpFecha.Value = $FMax }
                $lblStatus.Text = "Ruta manual cargada. Rango detectado: $($FMin.ToString('dd/MM/yyyy')) - $($FMax.ToString('dd/MM/yyyy')). Puedes cambiar la fecha libremente."
            }
        }
        $btnAnalizar.Enabled = $true
    }
})

# ====================================================================
# ACCIÓN 2: AUDITAR — MÓDULO 5 V22 (MOTOR PRINCIPAL)
# ====================================================================
$btnAnalizar.Add_Click({
    $TargetIP = $txtIP.Text.Trim()

    if ($Script:RutaManual -eq "") {
        if ($cmbUsuarios.SelectedItem -eq $null) { $lblStatus.Text = "Selecciona un usuario o carga una ruta manual."; return }
        $TargetUser = $cmbUsuarios.SelectedItem.ToString()
        & net use "\\$TargetIP\c$" /delete /y 2>&1 | Out-Null
        & net use "\\$TargetIP\IPC$" /delete /y 2>&1 | Out-Null
        if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null }
    } else {
        $TargetUser = "Usuario Local"
    }

    $FechaSeleccionada = $dtpFecha.Value.Date
    $FechaVisualStr    = $dtpFecha.Value.ToString("dd/MM/yyyy")
    $FechaFin          = $FechaSeleccionada.AddDays(1)
    $F1 = $dtpFecha.Value.ToString("MM/dd/yyyy"); $F2 = $dtpFecha.Value.ToString("dd/MM/yyyy")
    $F3 = $dtpFecha.Value.ToString("yyyy-MM-dd"); $F4 = $dtpFecha.Value.ToString("M/d/yyyy"); $F5 = $dtpFecha.Value.ToString("d/M/yyyy")
    $FechaOmni = "(?:$([regex]::Escape($F1))|$([regex]::Escape($F2))|$([regex]::Escape($F3))|$([regex]::Escape($F4))|$([regex]::Escape($F5)))"

    $GridResultados.Rows.Clear(); $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $EventosTiempo       = @{}
    $ListaExtensiones    = @()
    $ListaLogins         = @()
    $Script:ModoContestacion = @{}
    $Script:AlertingHoras    = @{}
    $DirLlamada          = @{}
    $ListaMakeCall       = @()
    $ConexionesYaIniciadas = @{}
    $PrimaryConnectedSeconds = @{}   # O(1) guard para SECONDARY CONNECTED
    $AlertingPorTel = @{}            # telNorm → horaLimpia: unifica dedup entre XML, WI.ADD y fallback
    $PhoneYaEnInicio = @{}           # telNorm → $true: bloquea SECONDARY para teléfono ya conectado por PRIMARY
    $SesionHoraInicio = @{}          # sesionId → horaSlot: para actualizar LÍNEA ABIERTA cuando llega tel real
    $CallStateDisconnected = @{}     # sesionId → horaSlot: fallback FIN cuando ProcessSessionEndedEvent no aparece
    $Script:UltimaHoraAlerting = $null
    $Script:UltimoTopic        = $null
    $XMLCargado = $false

    function Init-Hora ($hora) {
        if (-not $EventosTiempo.ContainsKey($hora)) {
            $EventosTiempo[$hora] = @{
                Interpretacion=""; Agente=""; Audio=""; Aux=""; Ispeac=""; SysLog=""; AppLog=""
                RawInterpretacion=""; RawAgente=""; RawAudio=""; RawAux=""; RawIspeac=""; RawSysLog=""; RawAppLog=""
                Sesion="-"; Tel="-"; ViId=""; Topic=""
                ColorInterpretacion=[System.Drawing.Color]::White; ColorAgente=[System.Drawing.Color]::White
                ColorAudio=[System.Drawing.Color]::White; ColorAux=[System.Drawing.Color]::White
                ColorIspeac=[System.Drawing.Color]::White; ColorSys=[System.Drawing.Color]::White; ColorApp=[System.Drawing.Color]::White
            }
        }
    }

    # --- WINDOWS EVENT LOG (solo modo remoto) ---
    if ($Script:RutaManual -eq "") {
        $lblStatus.Text = "Analizando Event Log de Aplicación para el $FechaVisualStr..."; $Form.Refresh()
        try {
            $EvApp = Get-WinEvent -ComputerName $TargetIP -Credential $Script:Creds -FilterHashtable @{LogName='Application'; Id=1000,1002; StartTime=$FechaSeleccionada; EndTime=$FechaFin} -EA SilentlyContinue
            if ($EvApp) {
                foreach ($Ev in $EvApp) {
                    $HoraLimpia = $Ev.TimeCreated.ToString("HH:mm:ss"); Init-Hora $HoraLimpia
                    $EventosTiempo[$HoraLimpia].AppLog = "APP CRASH/HANG: $($Ev.ProviderName)"
                    $EventosTiempo[$HoraLimpia].ColorApp = [System.Drawing.Color]::OrangeRed
                    $EventosTiempo[$HoraLimpia].Tel = "WINDOWS APP"
                    $EventosTiempo[$HoraLimpia].RawAppLog += "[$HoraLimpia] $($Ev.ProviderName) - $($Ev.Message)`n"
                }
            }
        } catch {}
        $lblStatus.Text = "Analizando Event Log de Sistema..."; $Form.Refresh()
        try {
            $EvSys = Get-WinEvent -ComputerName $TargetIP -Credential $Script:Creds -FilterHashtable @{LogName='System'; Id=41,1074,6008,1001,129,153,2004,10400; StartTime=$FechaSeleccionada; EndTime=$FechaFin} -EA SilentlyContinue
            if ($EvSys) {
                foreach ($Ev in $EvSys) {
                    $HoraLimpia = $Ev.TimeCreated.ToString("HH:mm:ss"); Init-Hora $HoraLimpia
                    $Desc = ""; $Omitir = $false
                    switch ($Ev.Id) {
                        41    { $Desc = "K-Power (Apagado Sucio/Corte energía)" }
                        1074  { $Desc = "Reinicio provocado por Usuario" }
                        6008  { $Desc = "Cierre Inesperado Previo" }
                        1001  { $Desc = "PANTALLAZO AZUL (BSOD BugCheck)" }
                        129   { if ($Ev.ProviderName -match "(?i)storahci|iaStor|disk|VDS") { $Desc = "Disco Duro: Reset Port (Lentitud)" } else { $Omitir = $true } }
                        153   { if ($Ev.ProviderName -match "(?i)storahci|iaStor|disk|VDS") { $Desc = "Disco Duro: Reintento IO (Lentitud)" } else { $Omitir = $true } }
                        2004  { $Desc = "¡ALERTA RAM!: Memoria Virtual Agotada" }
                        10400 { $Desc = "¡ALERTA NDIS!: Tarjeta de Red desconectada" }
                        default { $Desc = "Evento Sistema: $($Ev.Id)" }
                    }
                    if ($Omitir) { continue }
                    $EventosTiempo[$HoraLimpia].SysLog = $Desc; $EventosTiempo[$HoraLimpia].ColorSys = [System.Drawing.Color]::Red
                    $EventosTiempo[$HoraLimpia].Tel = "WINDOWS SYS"; $EventosTiempo[$HoraLimpia].RawSysLog += "[$HoraLimpia] ID $($Ev.Id): $($Ev.Message)`n"
                }
            }
        } catch {}
    } else { $lblStatus.Text = "Modo Local: Omitiendo visor de eventos de Windows..."; $Form.Refresh() }

    $lblStatus.Text = "Analizando logs de Endpoint, Audio, Ispeac y One-X..."; $Form.Refresh()

    try {
        if ($Script:RutaManual -eq "") {
            New-PSDrive -Name $Script:DriveName -PSProvider FileSystem -Root "\\$TargetIP\c$" -Credential $Script:Creds -EA Stop | Out-Null
            $RutaDir1 = "$($Script:DriveName):\Users\$TargetUser\AppData\Roaming\Avaya\Avaya one-X Agent\out"
            $RutaDir2 = "$($Script:DriveName):\Users\$TargetUser\AppData\Roaming\Avaya\one-X Agent\2.5\Log Files"
            $RutaDir3 = "$($Script:DriveName):\Users\$TargetUser\AppData\Roaming\Avaya\one-X Agent\2.5"
            $DirFinal = $null
            if (Test-Path $RutaDir1) { $DirFinal = $RutaDir1 } elseif (Test-Path $RutaDir2) { $DirFinal = $RutaDir2 } elseif (Test-Path $RutaDir3) { $DirFinal = $RutaDir3 }
        } else {
            $DirFinal = $Script:RutaManual
        }
        $Script:DirFinalGlobal = $DirFinal

        if ($DirFinal) {

            # ================================================================
            # PASO 1: CONTACTLOG.XML — FUENTE PRIMARIA
            # ================================================================
            $lblStatus.Text = "Leyendo ContactLog.xml (fuente primaria)..."; $Form.Refresh()
            $XMLPath = Get-ChildItem -Path $DirFinal -Filter "ContactLog.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if (-not $XMLPath) {
                $ParentDir = Split-Path $DirFinal -Parent
                $XMLPath = Get-ChildItem -Path $ParentDir -Filter "ContactLog.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            }

            if ($XMLPath -and (Test-Path $XMLPath)) {
                try {
                    $xml = [xml](Get-Content $XMLPath -Encoding UTF8)
                    $FechaFiltro = $FechaSeleccionada.Date

                    foreach ($log in $xml.ContactLogs.ContactLog) {
                        $LogTimeLocal = [datetimeoffset]::Parse($log.CreateTime).LocalDateTime
                        if ($LogTimeLocal.Date -ne $FechaFiltro) { continue }
                        $Topic = if ($log.Topic) { $log.Topic.Trim() } else { "" }

                        $Items = $log.ContactLogItem
                        if ($Items -eq $null) { continue }
                        if ($Items -isnot [System.Array]) { $Items = @($Items) }

                        foreach ($item in $Items) {
                            if ($item -eq $null) { continue }
                            $ItemTimeLocal = [datetimeoffset]::Parse($item.CreateTime).LocalDateTime
                            if ($ItemTimeLocal.Date -ne $FechaFiltro) { continue }
                            $HoraLimpia = $ItemTimeLocal.ToString("HH:mm:ss")
                            Init-Hora $HoraLimpia

                            $IsOutbound = ($item.Outbound -eq "true")
                            $IsMissed   = ($item.Missed   -eq "true")
                            $ViId       = if ($item.Id)   { $item.Id.Trim() }   else { "" }
                            $UriRaw     = if ($item.Uri)  { $item.Uri.Trim() }  else { "" }
                            $Uri        = ($UriRaw -split ',')[0] -replace '\s+.*$',''  # primer token, sin espacios
                            $NumLimpio  = $Uri -replace '\D',''
                            $NombreXML  = if ($item.Name) { $item.Name.Trim() } else { "" }
                            $DurSeg     = try { [int]$item.Duration } catch { -1 }
                            $TelDisplay = if ($NumLimpio) { $NumLimpio } elseif ($NombreXML) { $NombreXML } else { "Desconocido" }
                            $EsInterna  = ($NumLimpio.Length -gt 0 -and $NumLimpio.Length -le 6)

                            switch ($item.Type) {

                                "StateChange" {
                                    switch ($item.State) {
                                        "LoggedIn"  {
                                            $EventosTiempo[$HoraLimpia].Aux += "ESTADO: LOGGEDIN|"
                                            $EventosTiempo[$HoraLimpia].RawAux += "[XML] LoggedIn Ext:$($item.Uri)`n"
                                        }
                                        "LoggedOut" {
                                            $EventosTiempo[$HoraLimpia].Aux += "ESTADO: LOGGEDOUT|"
                                            $EventosTiempo[$HoraLimpia].RawAux += "[XML] LoggedOut`n"
                                        }
                                        "Ready" {
                                            $DurStr = if ($DurSeg -ge 0) { " ($DurSeg s)" } else { "" }
                                            $EventosTiempo[$HoraLimpia].Aux += "ESTADO: READY|"
                                            $EventosTiempo[$HoraLimpia].RawAux += "[XML] Ready$DurStr`n"
                                        }
                                        "Aux" {
                                            $RC = if ($item.ReasonCode) { $item.ReasonCode.Trim() } else { "default" }
                                            $DurStr = if ($DurSeg -ge 0) { " ($DurSeg s)" } else { "" }
                                            $EventosTiempo[$HoraLimpia].Aux += "NUEVO_RC_NOMBRE:$RC|"
                                            $EventosTiempo[$HoraLimpia].RawAux += "[XML] Aux[$RC]$DurStr`n"
                                        }
                                    }
                                }

                                "Voice" {
                                    $EventosTiempo[$HoraLimpia].ViId  = $ViId
                                    $EventosTiempo[$HoraLimpia].Topic = $Topic
                                    if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $TelDisplay }

                                    if ($IsOutbound) {
                                        # --- LLAMADA SALIENTE ---
                                        if ($TelDisplay -match "Desconocido|^$") {
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR (Posible evasión)"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        } else {
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE LLAMADA (Saliente)"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        }
                                    } else {
                                        # --- LLAMADA ENTRANTE — momento de alerting ---
                                        $PerdidaTag = if ($IsMissed) { " [LLAMADA PERDIDA]" } else { "" }
                                        if ($EsInterna) {
                                            $MsgAlert = "Agente con señal de llamada (Entrante - INTERNA - $TelDisplay)$PerdidaTag"
                                        } elseif ($Topic -ne "") {
                                            $MsgAlert = "Agente con señal de llamada (Entrante - ACD - $Topic)$PerdidaTag"
                                        } else {
                                            $MsgAlert = "Agente con señal de llamada (Entrante - EXTERNA - $TelDisplay)$PerdidaTag"
                                        }
                                        $EventosTiempo[$HoraLimpia].Interpretacion      = $MsgAlert
                                        $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        $Script:UltimaHoraAlerting = $HoraLimpia
                                        $Script:UltimoTopic        = $Topic
                                        # Guardar UUID para cruzar con CONNECTED en OneXAgent
                                        if ($ViId -match "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
                                            $Script:AlertingHoras[$matches[1]] = $HoraLimpia
                                        }
                                        # Registrar teléfono normalizado para dedup cruzada con WI.ADD/fallback
                                        $TelNormS = ($TelDisplay -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                        $AlertingPorTel[$TelNormS] = $HoraLimpia
                                    }
                                    $DirXML = if ($IsOutbound) { "Saliente" } else { "Entrante" }
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[XML] $ViId | Tel:$TelDisplay | $DirXML | Dur:${DurSeg}s | Topic:$Topic`n"
                                }

                                "VoiceConference" {
                                    $DurStr = if ($DurSeg -ge 0) { " (Dur: ${DurSeg}s)" } else { "" }
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[XML] Conference$DurStr`n"
                                    if ($EventosTiempo[$HoraLimpia].Interpretacion -ne "") {
                                        $EventosTiempo[$HoraLimpia].Interpretacion += " + CONFERENCIA"
                                    }
                                }
                            }
                        }
                    }
                    $XMLCargado = $true
                    $lblStatus.Text = "ContactLog.xml cargado correctamente. Continuando con logs..."; $Form.Refresh()
                } catch {
                    $lblStatus.Text = "ContactLog.xml no pudo leerse — usando modo fallback de logs..."; $Form.Refresh()
                }
            }

            # ================================================================
            # PASO 2: AUDIOLOGS
            # ================================================================
            $lblStatus.Text = "PASO 2/6: Leyendo AudioLogs..."; $Form.Refresh()
            $ArchivosAudio = Get-ChildItem -Path $DirFinal -Filter "AudioLog.txt*" | Sort-Object { if ($_.Name -match "\.(\d+)$") { [int]$matches[1] } else { -1 } } -Descending
            if ($ArchivosAudio) {
                foreach ($ArchivoA in $ArchivosAudio) {
                    $LineasA = Get-Content -Path $ArchivoA.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                    if (-not $LineasA) { continue }
                    foreach ($linea in $LineasA) {
                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2})") {
                            $HoraLimpia = $matches[1]; Init-Hora $HoraLimpia
                            if ($linea -match "(?i)sActiveWave(In|Out)Device\s+'([^']+)'") {
                                $Dispositivo = $matches[2].Trim()
                                $EventosTiempo[$HoraLimpia].Interpretacion = "Extensión en línea y conectada a dispositivos de audio"
                                $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Cyan
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                                if ($EventosTiempo[$HoraLimpia].Audio -eq "") { $EventosTiempo[$HoraLimpia].Audio = "$symMusic HW: $Dispositivo" } elseif ($EventosTiempo[$HoraLimpia].Audio -notmatch [regex]::Escape($Dispositivo)) { $EventosTiempo[$HoraLimpia].Audio += " / $Dispositivo" }
                                $EventosTiempo[$HoraLimpia].ColorAudio = [System.Drawing.Color]::DeepSkyBlue; $EventosTiempo[$HoraLimpia].RawAudio += "$linea`n"
                            }
                            elseif ($linea -match "StartSession: Start ISPEAC audio session") { $EventosTiempo[$HoraLimpia].Audio = "$symMusic Línea abierta"; $EventosTiempo[$HoraLimpia].ColorAudio = [System.Drawing.Color]::LimeGreen; $EventosTiempo[$HoraLimpia].RawAudio += "$linea`n" }
                            elseif ($linea -match "EndCall: Session ended") { $EventosTiempo[$HoraLimpia].Audio = "$symStop Línea cerrada"; $EventosTiempo[$HoraLimpia].ColorAudio = [System.Drawing.Color]::Gray; $EventosTiempo[$HoraLimpia].RawAudio += "$linea`n" }
                        }
                    }
                }
            }

            # ================================================================
            # PASO 3: ISPEACLOGS (Calidad de Red)
            # ================================================================
            $lblStatus.Text = "PASO 3/6: Leyendo IspeacLogs (calidad de red)..."; $Form.Refresh()
            $ArchivosIspeac = Get-ChildItem -Path $DirFinal -Filter "IspeacLog.txt*" | Sort-Object { if ($_.Name -match "\.(\d+)$") { [int]$matches[1] } else { -1 } } -Descending
            if ($ArchivosIspeac) {
                foreach ($ArchivoI in $ArchivosIspeac) {
                    $LineasI = Get-Content -Path $ArchivoI.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                    if (-not $LineasI) { continue }
                    foreach ($linea in $LineasI) {
                        if ($linea -match "^\[\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2}):\d{3}\]") {
                            $HoraLimpia = $matches[1]; Init-Hora $HoraLimpia
                            if ($linea -match "FRACTION DROPPED = (0\.[1-9]\d+|[1-9]\.\d+)") { $EventosTiempo[$HoraLimpia].Ispeac = "¡ALERTA RED! Pérdida de paquetes ($($matches[1]))"; $EventosTiempo[$HoraLimpia].ColorIspeac = [System.Drawing.Color]::OrangeRed; $EventosTiempo[$HoraLimpia].RawIspeac += "$linea`n" }
                            elseif ($linea -match "RECEIVED RTCP ROUND TRIP DELAY = (\d{3,}\.\d+)") { $EventosTiempo[$HoraLimpia].Ispeac = "¡ALERTA RED! Lag crítico ($($matches[1]) ms)"; $EventosTiempo[$HoraLimpia].ColorIspeac = [System.Drawing.Color]::OrangeRed; $EventosTiempo[$HoraLimpia].RawIspeac += "$linea`n" }
                        }
                    }
                }
            }

            # ================================================================
            # PASO 4: AVAYA ONEX LOG
            # ================================================================
            $lblStatus.Text = "PASO 4/6: Leyendo OneXAgent.log (puede tomar tiempo si el archivo es grande)..."; $Form.Refresh()
            if ($Script:RutaManual -eq "") {
                $ArchivosOneX = Get-ChildItem -Path "$($Script:DriveName):\Users\$TargetUser\AppData\Roaming\Avaya" -Recurse -EA SilentlyContinue | Where-Object { -not $_.PSIsContainer -and ($_.Name -match "(?i)one-?x.*\.log" -or $_.Name -match "(?i)one-?x.*\.txt") } | Sort-Object { if ($_.Name -match "\.(\d+)$") { [int]$matches[1] } else { -1 } } -Descending
            } else {
                $ArchivosOneX = Get-ChildItem -Path $DirFinal -Recurse -EA SilentlyContinue | Where-Object { -not $_.PSIsContainer -and ($_.Name -match "(?i)one-?x.*\.log" -or $_.Name -match "(?i)one-?x.*\.txt") } | Sort-Object { if ($_.Name -match "\.(\d+)$") { [int]$matches[1] } else { -1 } } -Descending
            }

            if ($ArchivosOneX) {
                foreach ($ArchivoX in $ArchivosOneX) {
                    $LineasX = Get-Content -Path $ArchivoX.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                    if (-not $LineasX) { continue }
                    $HoraLimpia = ""; $MsLimpio = "000"
                    foreach ($linea in $LineasX) {
                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2})(?:,(\d{3}))?") {
                            $HoraLimpia = $matches[1]; $MsLimpio = if ($matches[2]) { $matches[2] } else { "000" }

                            # --- Diadema desconectada ---
                            if ($linea -match "(?i)(DeviceRemoved|Audio device removed|Hardware removed|RemoveDevice)") {
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].Audio = "¡ALERTA! Dispositivo de Audio Desconectado"
                                $EventosTiempo[$HoraLimpia].ColorAudio = [System.Drawing.Color]::Red
                                $EventosTiempo[$HoraLimpia].RawAudio += "$linea`n"
                            }
                            # --- Intento de firma (para correlación Amnesia V21) ---
                            elseif ($linea -match "(?i)Attempting to Login \[stationId=(\d+)") {
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].Aux += "INTENTO_FIRMA_ONEX_EXT:$($matches[1])|"
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- Auto-Answer vs Manual (ModoContestacion) ---
                            elseif ($linea -match "(?i)(AnswerVoiceInteraction|Auto Accepting).*?(VI\d+:[^\s,\]]+)") {
                                $AccionDetectada = $matches[1]; $IdVoice = $matches[2]
                                $Script:ModoContestacion[$IdVoice] = if ($AccionDetectada -like "*Answer*") { "MANUAL (El asesor le dio clic a Contestar)" } else { "AUTOMATICO (Entró directo por Auto-Answer)" }
                            }
                            # --- GUI Ready/Aux confirmado ---
                            elseif ($linea -match "GUI Method (STARTED|ENDED): EnterReadyHandler") {
                                Init-Hora $HoraLimpia
                                if ($matches[1] -eq "ENDED") { $EventosTiempo[$HoraLimpia].Aux += "GUI_READY_CONFIRMADO|" }
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            elseif ($linea -match "GUI Method (STARTED|ENDED): EnterAuxWithReasonCodeHandler") {
                                Init-Hora $HoraLimpia
                                if ($matches[1] -eq "ENDED") { $EventosTiempo[$HoraLimpia].Aux += "GUI_AUX_CONFIRMADO|" }
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- ReasonCode y estados del agente (respaldo) ---
                            elseif ($linea -match "(?i)WorkServiceImpl EnterAux:session=.*?;code=(\d+)") {
                                Init-Hora $HoraLimpia; $Rc = $matches[1]
                                $EventosTiempo[$HoraLimpia].Aux += "NUEVO_RC:$Rc|"; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            elseif ($linea -match "(?i)ReasonCode[=\[>:\s]*(\d+)") {
                                Init-Hora $HoraLimpia; $Rc = $matches[1]
                                if ($Rc -eq "0") { $EventosTiempo[$HoraLimpia].Aux += "DEFAULT|" } elseif ($Rc -eq "10") { $EventosTiempo[$HoraLimpia].Aux += "SISTEMA_LOGOUT|" } else { $EventosTiempo[$HoraLimpia].Aux += "RC: $Rc|" }
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            elseif ($linea -match "(?i)(?:UpdateSessionState:\s*Agent State\s*|AgentState:\s*|State\s*[:=]\s*|Enter\s+)(AuxWork|Aux|Ready|AutoIn|ManualIn|NotReady|Default|LoggedOut|PendingAux)") {
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].Aux += "ESTADO: $($matches[1].ToUpper())|"
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- Call created: registra dirección y limpia dict ---
                            elseif ($linea -match "(?i)(?:PhoneService_CallCreated:call=|Call created : )Id=(\d+).*?Outgoing=(True|False)") {
                                $IDLlamada = $matches[1]; $Sentido = $matches[2]
                                if ($Sentido -eq "True") { $DirLlamada[$IDLlamada] = "SALIENTE" } else { $DirLlamada[$IDLlamada] = "ENTRANTE" }
                                $ConexionesYaIniciadas.Remove($IDLlamada) | Out-Null
                            }
                            # --- WI.ADD: WorkItemImpl.VoiceInteractionListImpl.Add — tercera fuente de detección ---
                            # Más completo que XML y que InnerState=CONNECTED. Tiene todo en una sola línea.
                            # Actúa con o sin XML (guards previenen duplicación si XML ya lo hizo).
                            elseif ($linea -match "VoiceInteractionListImpl\.Add" -and $linea -match "\bid=(VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}))") {
                                $WiViIdFull = $matches[1]; $WiViUuid = $matches[2]
                                $WiOutbound = $linea -match "(?i)outbound=true"
                                $WiMissed   = $linea -match "(?i)missed=true"
                                # BUG FIX: el orden de campos en VoiceInteractionListImpl.Add varía entre versiones.
                                # Algunos logs ponen "state=" como primer campo: "[state=[type=Alerting,..." (sin coma previa).
                                # Otros lo ponen al final: "...ttyState=[...],state=[type=Alerting,...".
                                # Usando \bstate= (word boundary) el regex funciona en ambos casos sin capturar "ttyState=".
                                $WiEsAlerting = $linea -match "\bstate=\[.*?type=Alerting"
                                $WiEsActive   = $linea -match "\bstate=\[.*?type=Active"
                                # Solo procesamos: Alerting+Inbound (señal de llamada) o Active+Outbound (inicio saliente)
                                # Los Active+Inbound (llamada contestada) los maneja PRIMARY CONNECTED — los ignoramos aquí
                                if (-not (($WiEsAlerting -and -not $WiOutbound) -or ($WiEsActive -and $WiOutbound))) { continue }
                                $WiTopic = ""; if ($linea -match "topic=([^,\]]+)") { $WiTopic = $matches[1].Trim() }
                                $WiAddr  = ""; if ($linea -match "remoteAddress=([^,\]]+)") { $WiAddr = $matches[1].Trim() }
                                $WiName  = ""; if ($linea -match "remoteUserName=([^,\]]+)") { $WiName = $matches[1].Trim() }
                                $WiNum   = ($WiAddr -split ',')[0] -replace '\D',''
                                $WiTel   = if ($WiNum) { $WiNum } elseif ($WiName) { $WiName } else { "Desconocido" }
                                $WiInterna = ($WiNum.Length -gt 0 -and $WiNum.Length -le 6)
                                if ($WiEsAlerting -and -not $WiOutbound) {
                                    # INBOUND ALERTING
                                    # Normalizar teléfono para dedup cruzada (quita '+' y prefijo '9' de línea externa)
                                    $WiTelNorm = ($WiTel -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                    if (-not $Script:AlertingHoras.ContainsKey($WiViUuid) -and -not $AlertingPorTel.ContainsKey($WiTelNorm)) {
                                        # Primera fuente que ve este UUID y número → registrar y crear "señal de llamada"
                                        $Script:AlertingHoras[$WiViUuid] = $HoraLimpia
                                        $AlertingPorTel[$WiTelNorm] = $HoraLimpia
                                        $Script:UltimaHoraAlerting = $HoraLimpia; $Script:UltimoTopic = $WiTopic
                                        Init-Hora $HoraLimpia
                                        if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "señal de llamada") {
                                            $WiPerdida = if ($WiMissed) { " [LLAMADA PERDIDA]" } else { "" }
                                            if ($WiInterna) {
                                                $EventosTiempo[$HoraLimpia].Interpretacion = "Agente con señal de llamada (Entrante - INTERNA - $WiTel)$WiPerdida"
                                            } elseif ($WiTopic -ne "") {
                                                $EventosTiempo[$HoraLimpia].Interpretacion = "Agente con señal de llamada (Entrante - ACD - $WiTopic)$WiPerdida"
                                            } else {
                                                $EventosTiempo[$HoraLimpia].Interpretacion = "Agente con señal de llamada (Entrante - EXTERNA - $WiTel)$WiPerdida"
                                            }
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        }
                                        if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $WiTel }
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion += "[WI.ADD] $WiViIdFull | Tel:$WiTel | Entrante | Estado:Alerting | Topic:$WiTopic | Missed:$WiMissed`n"
                                    } else {
                                        # UUID ya registrado (XML) O teléfono ya tiene señal (fallback) → NO crear fila duplicada.
                                        # Si WI.ADD tiene Topic y el slot previo muestra "EXTERNA", mejorar la etiqueta.
                                        $HoraPrevia = if ($Script:AlertingHoras.ContainsKey($WiViUuid)) { $Script:AlertingHoras[$WiViUuid] } else { $AlertingPorTel[$WiTelNorm] }
                                        if (-not $Script:AlertingHoras.ContainsKey($WiViUuid)) { $Script:AlertingHoras[$WiViUuid] = $HoraPrevia }
                                        if ($EventosTiempo.ContainsKey($HoraPrevia)) {
                                            if ($WiTopic -ne "" -and $EventosTiempo[$HoraPrevia].Interpretacion -match "EXTERNA") {
                                                $WiPerdidaE = if ($WiMissed) { " [LLAMADA PERDIDA]" } else { "" }
                                                $EventosTiempo[$HoraPrevia].Interpretacion = "Agente con señal de llamada (Entrante - ACD - $WiTopic)$WiPerdidaE"
                                            }
                                            $EventosTiempo[$HoraPrevia].RawInterpretacion += "[WI.ADD] $WiViIdFull | Tel:$WiTel | Entrante | Estado:Alerting | Topic:$WiTopic | Missed:$WiMissed`n"
                                        }
                                    }
                                } else {
                                    # OUTBOUND ACTIVE: crear INICIO si XML no lo hizo
                                    Init-Hora $HoraLimpia
                                    if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "INICIO DE LLAMADA|LÍNEA ABIERTA") {
                                        if ($WiTel -eq "Desconocido" -or $WiTel -eq "") {
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR (Posible evasión)"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        } else {
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE LLAMADA (Saliente)"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        }
                                        if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $WiTel }
                                    }
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[WI.ADD] $WiViIdFull | Tel:$WiTel | Saliente | Estado:Active | Topic:$WiTopic | Missed:$WiMissed`n"
                                }
                            }
                            # --- ALERTING: solo si XML NO está cargado (fallback V21) ---
                            elseif (-not $XMLCargado -and $linea -match "(?i)(?:CallReceived|CallStateChanged).*?Id=\d+.*?State=Alerting" -and $linea -match "Outgoing=False" -and $linea -notmatch "InnerState=CONNECTED") {
                                Init-Hora $HoraLimpia
                                $NumRaw = "Desconocido"; if ($linea -match "RemoteParty=\[.*?,([^\]]+)\]") { $NumRaw = $matches[1] }
                                $FallTelNorm = ($NumRaw -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                $SegM1 = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-1)).ToString("HH:mm:ss")
                                $SegM2 = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-2)).ToString("HH:mm:ss")
                                $YaHayAlerting = ($EventosTiempo[$HoraLimpia].Interpretacion -match "señal de llamada") -or ($EventosTiempo.ContainsKey($SegM1) -and $EventosTiempo[$SegM1].Interpretacion -match "señal de llamada") -or ($EventosTiempo.ContainsKey($SegM2) -and $EventosTiempo[$SegM2].Interpretacion -match "señal de llamada") -or $AlertingPorTel.ContainsKey($FallTelNorm)
                                if (-not $YaHayAlerting) {
                                    $MsgA = if ($NumRaw.Length -gt 6) { "Agente con señal de llamada (Entrante - EXTERNA - $NumRaw)" } else { "Agente con señal de llamada (Entrante - INTERNA - $NumRaw)" }
                                    $EventosTiempo[$HoraLimpia].Interpretacion = $MsgA; $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                    if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $NumRaw }
                                    $Script:UltimaHoraAlerting = $HoraLimpia
                                    $AlertingPorTel[$FallTelNorm] = $HoraLimpia
                                }
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            elseif (-not $XMLCargado -and $linea -match "(?i)type=Alerting" -and $linea -match "(?i)outbound=False") {
                                Init-Hora $HoraLimpia
                                $TopicRaw = "Sin_Topic"; if ($linea -match "(?i)topic=([^,\]]*)") { $TopicRaw = $matches[1] }
                                $NumRaw = "Desconocido"; if ($linea -match "(?i)remoteAddress=([^,\]]+)") { $NumRaw = $matches[1] }
                                $Script:UltimoTopic = $TopicRaw
                                $EsInterna = ($linea -notmatch "(?i)inboundAcd=true" -or $NumRaw.Length -le 5)
                                $MsgAlerting = if ($EsInterna) { "Agente con señal de llamada (Entrante - INTERNA - $NumRaw)" } else { "Agente con señal de llamada (Entrante - ACD - $TopicRaw)" }
                                $SegM1 = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-1)).ToString("HH:mm:ss")
                                $SegM2 = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-2)).ToString("HH:mm:ss")
                                if ($EventosTiempo[$HoraLimpia].Interpretacion -match "señal de llamada") { $EventosTiempo[$HoraLimpia].Interpretacion = $EventosTiempo[$HoraLimpia].Interpretacion -replace "Agente con señal de llamada.*?(?=(?:  --->|$))", $MsgAlerting; $Script:UltimaHoraAlerting = $HoraLimpia }
                                elseif ($EventosTiempo.ContainsKey($SegM1) -and $EventosTiempo[$SegM1].Interpretacion -match "señal de llamada") { $EventosTiempo[$SegM1].Interpretacion = $EventosTiempo[$SegM1].Interpretacion -replace "Agente con señal de llamada.*?(?=(?:  --->|$))", $MsgAlerting; $Script:UltimaHoraAlerting = $SegM1 }
                                elseif ($EventosTiempo.ContainsKey($SegM2) -and $EventosTiempo[$SegM2].Interpretacion -match "señal de llamada") { $EventosTiempo[$SegM2].Interpretacion = $EventosTiempo[$SegM2].Interpretacion -replace "Agente con señal de llamada.*?(?=(?:  --->|$))", $MsgAlerting; $Script:UltimaHoraAlerting = $SegM2 }
                                else { $EventosTiempo[$HoraLimpia].Interpretacion = $MsgAlerting; $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold; $Script:UltimaHoraAlerting = $HoraLimpia }
                                if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $NumRaw }
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- CONFIRMACIÓN ACD (fallback, funciona con o sin XML) ---
                            elseif (-not $XMLCargado -and $linea -match "(?i)Setting Vi\.InboundAcd=true") {
                                Init-Hora $HoraLimpia
                                if ($Script:UltimaHoraAlerting -ne $null -and $EventosTiempo[$Script:UltimaHoraAlerting].Interpretacion -match "INTERNA") {
                                    $TopicVal = if ($Script:UltimoTopic -ne "Desconocido" -and $Script:UltimoTopic -ne $null -and $Script:UltimoTopic -ne "Sin_Topic") { $Script:UltimoTopic } else { "ACD" }
                                    $EventosTiempo[$Script:UltimaHoraAlerting].Interpretacion = $EventosTiempo[$Script:UltimaHoraAlerting].Interpretacion -replace "INTERNA - [^)]+", "ACD - $TopicVal"
                                }
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- PRIMARY CONNECTED: VoiceInteractionImpl type=Active (VI UUID → Ring Time exacto) ---
                            # Este evento es el correcto para cruzar UUID con ContactLog.xml.
                            # Formato real: VoiceInteractionImpl[VI#:UUID,cxt=UUID].StateImpl=[operations=[...],type=Active,...]
                            elseif ($linea -match "VoiceInteractionImpl\[VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}),cxt=([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\]\.StateImpl=\[.*?type=Active") {
                                $ViUuid = $matches[1]; $CxtUuid = $matches[2]
                                $ClaveVI = "VI_$ViUuid"
                                if (-not $ConexionesYaIniciadas.ContainsKey($ClaveVI)) {
                                    $ConexionesYaIniciadas[$ClaveVI] = $true
                                    $PrimaryConnectedSeconds[$HoraLimpia] = $true  # marca O(1) para SECONDARY guard
                                    # Slot con milisegundos: INICIO siempre en fila propia, separada del Alerting
                                    $SlotInicio = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotInicio
                                    Init-Hora $HoraLimpia   # asegura slot de segundos para rama SALIENTE
                                    # Ring Time (precisión de segundos; AlertingHoras almacena HH:mm:ss)
                                    $RingTimeStr = ""
                                    if ($Script:AlertingHoras.ContainsKey($ViUuid)) {
                                        try {
                                            $AlertingBase = ($Script:AlertingHoras[$ViUuid] -split ',')[0]
                                            $TAlerta = [datetime]::ParseExact($AlertingBase, "HH:mm:ss", $null)
                                            $TConect = [datetime]::ParseExact($HoraLimpia, "HH:mm:ss", $null)
                                            $RingSeg = [int]($TConect - $TAlerta).TotalSeconds
                                            if ($RingSeg -ge 0 -and $RingSeg -le 300) { $RingTimeStr = " [Ring: ${RingSeg}s]" }
                                        } catch {}
                                    }
                                    # Si el UUID está en AlertingHoras → ENTRANTE; si no → SALIENTE
                                    if ($Script:AlertingHoras.ContainsKey($ViUuid)) {
                                        # Entrante: INICIO siempre en su propio slot ms (nunca fusionado con señal de llamada)
                                        if ($EventosTiempo[$SlotInicio].Interpretacion -notmatch "INICIO DE LLAMADA") {
                                            $EventosTiempo[$SlotInicio].Interpretacion = "$symPlay INICIO DE LLAMADA (Entrante)$RingTimeStr"
                                        }
                                        $EventosTiempo[$SlotInicio].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        # Evidencia de dirección para el diagnóstico
                                        $AlertHoraRef = $Script:AlertingHoras[$ViUuid]
                                        $EventosTiempo[$SlotInicio].RawInterpretacion += "[DIR:ENTRANTE] PRIMARY_CONNECTED: AlertingHoras[VI:$ViUuid] = señal en $AlertHoraRef`n"
                                        $EventosTiempo[$SlotInicio].RawInterpretacion += "$linea`n"
                                        # Alerting terminó → liberar teléfono para que una rellamada genere nueva señal
                                        $AlertHoraX = $Script:AlertingHoras[$ViUuid]
                                        if ($EventosTiempo.ContainsKey($AlertHoraX)) {
                                            $TelClearX = ($EventosTiempo[$AlertHoraX].Tel -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                            $AlertingPorTel.Remove($TelClearX) | Out-Null
                                            # Bloquear SECONDARY CONNECTED para este teléfono (ya fue procesado por PRIMARY)
                                            if ($TelClearX -ne "" -and $TelClearX -ne "Desconocido") { $PhoneYaEnInicio[$TelClearX] = $true }
                                        }
                                    } else {
                                        # Saliente: slot de segundos (XML y WI.ADD ya mandan en ese slot)
                                        if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "INICIO DE LLAMADA|LÍNEA ABIERTA") {
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symPlay INICIO DE LLAMADA (Saliente)"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        }
                                        # Evidencia de dirección para el diagnóstico
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion += "[DIR:SALIENTE] PRIMARY_CONNECTED: VI UUID '$ViUuid' NO estaba en AlertingHoras`n"
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                                    }
                                }
                            }
                            # --- SECONDARY CONNECTED: InnerState=CONNECTED (usa ID numérico, sin UUID) ---
                            # Fallback por si type=Active no apareció. Guarda RemoteParty y previene duplicados.
                            elseif ($linea -match "(?i)(?:PhoneService_CallStateChanged|PhoneService_CallUpdated):call=Id=(\d+).*?State=Active.*?InnerState=CONNECTED" -and $linea -notmatch "(?i)State=Disconnected") {
                                $CallID = $matches[1]
                                Init-Hora $HoraLimpia
                                $NumRaw = "Desconocido"; if ($linea -match "RemoteParty=\[.*?,([^\]]+)\]") { $NumRaw = $matches[1] }
                                $NumRawNorm = ($NumRaw -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                if ($EventosTiempo[$HoraLimpia].Tel -eq "-" -or $EventosTiempo[$HoraLimpia].Tel -match "Desconocido") { $EventosTiempo[$HoraLimpia].Tel = $NumRaw }
                                # Guard: ventana de ±6 segundos para detectar si PRIMARY ya actuó.
                                # El log puede tener hasta 3-4 segundos de diferencia entre VoiceInteractionImpl type=Active
                                # y PhoneService_CallStateChanged InnerState=CONNECTED para la misma llamada.
                                $tBase2 = [datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null)
                                $YaHayInicio = $false
                                for ($i2 = 0; $i2 -le 6 -and -not $YaHayInicio; $i2++) {
                                    $tCheck2 = ($tBase2.AddSeconds(-$i2)).ToString("HH:mm:ss")
                                    if ($PrimaryConnectedSeconds.ContainsKey($tCheck2)) { $YaHayInicio = $true }
                                    elseif ($EventosTiempo.ContainsKey($tCheck2) -and $EventosTiempo[$tCheck2].Interpretacion -match "INICIO DE LLAMADA") { $YaHayInicio = $true }
                                }
                                # SECONDARY solo actúa si PRIMARY (type=Active) no se encargó ya.
                                # Guard adicional: si PRIMARY registró el teléfono en $PhoneYaEnInicio, esta llamada
                                # ya fue procesada aunque el CallID numérico no coincida con la clave UUID del primario.
                                # Esto bloquea re-disparos de PhoneService_CallUpdated minutos después (p.ej. por mute/unmute).
                                $SecundarioOK = (-not $YaHayInicio) -and
                                                (-not $ConexionesYaIniciadas.ContainsKey($CallID)) -and
                                                (-not ($NumRawNorm -ne "Desconocido" -and $PhoneYaEnInicio.ContainsKey($NumRawNorm))) -and
                                                ($linea -match "Outgoing=False") -and
                                                ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "señal de llamada")
                                if ($SecundarioOK) {
                                    $ConexionesYaIniciadas[$CallID] = $true
                                    # Marcar teléfono para bloquear disparos futuros del mismo CallUpdated
                                    if ($NumRawNorm -ne "Desconocido") { $PhoneYaEnInicio[$NumRawNorm] = $true }
                                    # Fallback ring time por proximidad temporal
                                    $RingTimeStr = ""
                                    if ($Script:UltimaHoraAlerting -ne $null) {
                                        try {
                                            $TAlerta = [datetime]::ParseExact($Script:UltimaHoraAlerting, "HH:mm:ss", $null)
                                            $TConect = [datetime]::ParseExact($HoraLimpia, "HH:mm:ss", $null)
                                            $RingSeg = [int]($TConect - $TAlerta).TotalSeconds
                                            if ($RingSeg -ge 0 -and $RingSeg -le 60) { $RingTimeStr = " [Ring: ${RingSeg}s]" }
                                        } catch {}
                                    }
                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symPlay INICIO DE LLAMADA (Entrante)$RingTimeStr"
                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                    # Evidencia de dirección para el diagnóstico
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[DIR:ENTRANTE] SECONDARY_CONNECTED: Outgoing=False, sin PRIMARY previo en ventana ±6s`n"
                                }
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                            }
                            # --- MakeCall (respaldo dirección saliente) ---
                            elseif ($linea -match "MakeCall\((.*?)\)") {
                                $NumMake = $matches[1] -replace '\D',''; if ($NumMake) { $ListaMakeCall += "$NumMake|$HoraLimpia" }
                                Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- FALLBACK FIN: Call StateChanged OldState=Active,NewState=Disconnected ---
                            # Evento UUID-based que confirma fin de llamada. No escribe FIN aquí porque PASO 4
                            # corre antes que PASO 5 (donde vive OnRequestEndSession que define manual vs normal).
                            # Solo guarda sesión→hora en dict; el sweep post-PASO-5 aplica el FIN si
                            # ProcessSessionEndedEvent no lo cubrió.
                            elseif ($linea -match "Call StateChanged Call\[Id=[0-9a-fA-F\-]+,ConnectionId=(\d+),OldState=Active,NewState=Disconnected\]") {
                                $SesD = $matches[1]
                                if (-not $CallStateDisconnected.ContainsKey($SesD)) {
                                    Init-Hora $HoraLimpia
                                    $CallStateDisconnected[$SesD] = $HoraLimpia
                                }
                            }
                        }

                        # Fuera del filtro de fecha — corren para TODAS las líneas del archivo
                        if ($HoraLimpia -ne "" -and $linea -match "Session_LoginAgent failed") { Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n" }
                        if ($linea -match "(?i)Call ended.*?Id=(\d+)") { $ConexionesYaIniciadas.Remove($matches[1]) | Out-Null }

                        # Radares de crash y errores
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)ThreadAbortException|Subproceso anulado") {
                            Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].AppLog = "¡AVAYA ALERTA!: Cierre Forzado (X o Taskmgr)"
                            $EventosTiempo[$HoraLimpia].ColorApp = [System.Drawing.Color]::Red; $EventosTiempo[$HoraLimpia].RawAppLog += "$linea`n"
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)VoiceInteraction_Transfer failed|Transfer failed --->|StartTransferCallAction") {
                            Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].AppLog = "¡AVAYA ALERTA!: Falla Crítica en Transferencia"
                            $EventosTiempo[$HoraLimpia].ColorApp = [System.Drawing.Color]::Red; $EventosTiempo[$HoraLimpia].RawAppLog += "$linea`n"
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)timed out|response is null") {
                            Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].AppLog = "¡RED/SISTEMA!: Posible Desincronización (Timeout/Null)"
                            $EventosTiempo[$HoraLimpia].ColorApp = [System.Drawing.Color]::Orange; $EventosTiempo[$HoraLimpia].RawAppLog += "$linea`n"
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "System\.Exception") {
                            Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].AppLog = "¡ERR CRÍTICO!: System.Exception"
                            $EventosTiempo[$HoraLimpia].ColorApp = [System.Drawing.Color]::DarkRed; $EventosTiempo[$HoraLimpia].RawAppLog += "$linea`n"; $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                        }
                        elseif ($HoraLimpia -ne "" -and $linea -match "(?i)\bException\b|\bFATAL\b|\bERROR\b" -and $linea -notmatch "VoiceInteraction_Transfer|ThreadAbortException|System\.Exception") {
                            Init-Hora $HoraLimpia
                            if ([string]::IsNullOrEmpty($EventosTiempo[$HoraLimpia].AppLog)) {
                                $EventosTiempo[$HoraLimpia].AppLog = "Evento sospechoso en código base (Revisar RAW)"
                                $EventosTiempo[$HoraLimpia].ColorApp = [System.Drawing.Color]::Yellow
                            }
                            $EventosTiempo[$HoraLimpia].RawAppLog += "$linea`n"
                        }
                    }
                }
            }

            # ================================================================
            # PASO 5: ENDPOINT LOGS
            # ================================================================
            $lblStatus.Text = "PASO 5/6: Leyendo EndpointLog (sesiones, hold, fin de llamada)..."; $Form.Refresh()
            $ArchivosLog = Get-ChildItem -Path $DirFinal -Filter "EndpointLog.txt*" | Sort-Object { if ($_.Name -match "\.(\d+)$") { [int]$matches[1] } else { -1 } } -Descending
            $MapeoTel = @{}; $RawMapeoTel = @{}; $DirLlamada = @{}; $CurrentExt = "Desconocida"
            $LlamadasVistas = @{}; $CuelguesVistos = @{}; $CuelguesManuales = @{}; $EstadoSesion = @{}
            $ConsultaTransf = @{}; $ConsultaConf = @{}   # sesId → destino: sesiones de consulta (Transferencia/Conferencia)

            if ($ArchivosLog) {
                foreach ($Archivo in $ArchivosLog) {
                    $LineasLog = Get-Content -Path $Archivo.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                    if (-not $LineasLog) { continue }

                    # Primera pasada: extensiones y logins
                    foreach ($linea in $LineasLog) {
                        if ($linea -match "StartUserRegistration: server: '[^']+', extension: '(\d+)'") {
                            $ExtEncontrada = $matches[1]
                            $HoraFirma = if ($linea -match "\d{2}:\d{2}:\d{2}") { $matches[0] } else { "---" }
                            $Firma = "$ExtEncontrada ($HoraFirma)"; $CurrentExt = $ExtEncontrada
                            if ($ListaExtensiones -notcontains $Firma) { $ListaExtensiones += $Firma }
                        }
                        if ($linea -match "\+?564(3\d{5})\d{6}\b") {
                            $LoginExtraido = $matches[1]
                            $HoraFirma = if ($linea -match "\d{2}:\d{2}:\d{2}") { $matches[0] } else { "---" }
                            $FirmaUnicaLog = "$LoginExtraido ($HoraFirma)"; $YaRegistrado = $false
                            foreach ($item in $ListaLogins) { if ($item -match "^$LoginExtraido") { $YaRegistrado = $true; break } }
                            if (-not $YaRegistrado) { $ListaLogins += $FirmaUnicaLog }
                        }
                    }

                    # ── Mini-pasada: sesiones de consulta (Transferencia / Conferencia) ──
                    # OnRequestTransferSession + "To phone#" → siguiente UpdateHistoryRecord vacío = consulta de transferencia
                    # Conference_Merged: nFirstCall=X nConsultCall=Y                       → Y = consulta de conferencia
                    $_TrPend = $false; $_TrDest = ""; $_CfPhones = @{}
                    foreach ($linea in $LineasLog) {
                        if ($linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*([^.]+)\.") {
                            $s = $matches[1]; $n = $matches[2].Trim()
                            if ($n -ne "") { $_CfPhones[$s] = $n }
                        }
                        if ($linea -match "OnRequestTransferSession\(\) entered from sessionId=\d+") {
                            $_TrPend = $true; $_TrDest = ""
                        }
                        if ($_TrPend -and $linea -match "\bTo phone#: '([^']+)'") { $_TrDest = $matches[1] }
                        if ($_TrPend -and $linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*\.") {
                            $s = $matches[1]
                            if (-not $ConsultaTransf.ContainsKey($s)) {
                                $ConsultaTransf[$s] = if ($_TrDest -ne "") { $_TrDest } else { "Desconocido" }
                            }
                            $_TrPend = $false
                        }
                        if ($linea -match "Conference_Merged: nFirstCall:\s*\d+\s+nConsultCall:\s*(\d+)") {
                            $s = $matches[1]
                            if (-not $ConsultaConf.ContainsKey($s)) {
                                $ConsultaConf[$s] = if ($_CfPhones[$s]) { $_CfPhones[$s] } else { "Desconocido" }
                            }
                        }
                    }

                    $LastXMLEvent = $null; $LastXMLTime = $null

                    # Segunda pasada: análisis principal
                    foreach ($linea in $LineasLog) {
                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2}).*?\].*<(HeldEvent|UnheldEvent)") { $LastXMLEvent = $matches[2]; $LastXMLTime = $matches[1] }
                        elseif ($LastXMLEvent -ne $null -and $linea -match "<connectionId>(\d+)</connectionId>") {
                            $Ses = $matches[1]; Init-Hora $LastXMLTime; $EventosTiempo[$LastXMLTime].Sesion = $Ses; $EventosTiempo[$LastXMLTime].Tel = if ($MapeoTel[$Ses]) { $MapeoTel[$Ses] } else { "Desconocido" }
                            if ($LastXMLEvent -eq "HeldEvent") { $EventosTiempo[$LastXMLTime].Agente = "|| HOLD (Confirmación del Sistema)"; $EventosTiempo[$LastXMLTime].ColorAgente = [System.Drawing.Color]::Yellow } else { $EventosTiempo[$LastXMLTime].Agente = "$symRes UNHOLD (Confirmación del Sistema)"; $EventosTiempo[$LastXMLTime].ColorAgente = [System.Drawing.Color]::LightGoldenrodYellow }
                            $EventosTiempo[$LastXMLTime].RawAgente += "<$LastXMLEvent>`n$linea`n"; $LastXMLEvent = $null
                        }

                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2})") {
                            $HoraLimpia = $matches[1]

                            # Mapeo de dirección
                            if ($linea -match "GenerateIncomingCall: callIndex:\s*(\d+)") {
                                $DirLlamada[$matches[1]] = "ENTRANTE"
                                # Evidencia para el diagnóstico: registrar en el slot actual
                                if ($HoraLimpia -ne "") { Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAux += "[DIR:ENTRANTE] PASO5: GenerateIncomingCall callIndex=$($matches[1])`n$linea`n" }
                            }

                            if ($linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*([^.]+)\.") {
                                $SesID = $matches[1]; $NumObj = $matches[2].Trim()
                                $MapeoTel[$SesID] = $NumObj; $RawMapeoTel[$SesID] = $linea
                                $NumLimpio = $NumObj -replace '\D',''
                                if ($DirLlamada[$SesID] -ne "ENTRANTE") {
                                    foreach ($m in $ListaMakeCall) {
                                        $DatosMake = $m -split '\|'; $NumM = $DatosMake[0]; $HoraM = $DatosMake[1]
                                        if ($NumM -match $NumLimpio -and $NumLimpio.Length -gt 4) {
                                            try { $t1=[datetime]::ParseExact($HoraM,"HH:mm:ss",$null); $t2=[datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null); if ([math]::Abs(($t2-$t1).TotalSeconds) -le 10) { $DirLlamada[$SesID]="SALIENTE"; break } } catch {}
                                        }
                                    }
                                }
                            }

                            # Protocolo de Login — Amnesia V21
                            if ($linea -match "DIAL: \[\d+\] NormalizeNumber return '(\+?564\d+)'") {
                                $LoginTel = $matches[1]
                                $ExtEncontrada = ""; $RawOneX = ""; $HoraMatch = $null
                                for ($i = 0; $i -le 10; $i++) {
                                    try {
                                        $HoraTest = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-$i)).ToString("HH:mm:ss")
                                        if ($EventosTiempo.ContainsKey($HoraTest) -and $EventosTiempo[$HoraTest].Aux -match "INTENTO_FIRMA_ONEX_EXT:(\d+)") {
                                            $ExtEncontrada = $matches[1]; $HoraMatch = $HoraTest
                                            if ($EventosTiempo[$HoraTest].RawAux -match "(?i)Attempting to Login.*") { $RawOneX = $matches[0] }
                                            break
                                        }
                                    } catch {}
                                }
                                # Cerrar sesiones zombie antes del nuevo login
                                $LlavesZombi = @($EstadoSesion.Keys)
                                foreach ($SesZombi in $LlavesZombi) {
                                    if ($EstadoSesion[$SesZombi] -eq "ACTIVA") {
                                        $HorasSesion = @(); $LlavesTiempo = @($EventosTiempo.Keys)
                                        foreach ($h in $LlavesTiempo) { if ($EventosTiempo[$h].Sesion -eq $SesZombi) { $HorasSesion += $h } }
                                        if ($HorasSesion.Count -gt 0) {
                                            $UltimaHoraConocida = ($HorasSesion | Sort-Object)[-1]; Init-Hora $UltimaHoraConocida
                                            if ($EventosTiempo[$UltimaHoraConocida].Interpretacion -notmatch "jamás se cerró|ABORTADA") {
                                                $EventosTiempo[$UltimaHoraConocida].Interpretacion += " `n---> [X] SESIÓN ABORTADA (Firma forzada a las $HoraLimpia)"
                                                $EventosTiempo[$UltimaHoraConocida].ColorInterpretacion = [System.Drawing.Color]::Crimson
                                            }
                                        }
                                        $EstadoSesion[$SesZombi] = "CERRADA"
                                    }
                                }
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Yellow
                                $EventosTiempo[$HoraLimpia].Tel = $LoginTel
                                if ($ExtEncontrada -ne "") {
                                    $EventosTiempo[$HoraLimpia].Interpretacion = "Usuario intentando firmarse en la Ext. $ExtEncontrada"
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[$HoraMatch] (OneX): $RawOneX`n[$HoraLimpia] (Endpoint): $linea`n"
                                    $EventosTiempo[$HoraMatch].Aux = $EventosTiempo[$HoraMatch].Aux -replace "INTENTO_FIRMA_ONEX_EXT:\d+\|", ""
                                } else {
                                    $EventosTiempo[$HoraLimpia].Interpretacion = "Usuario intentando firmarse (Recuperado por PBX)"
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[$HoraLimpia] (Endpoint): $linea`n"
                                }
                                continue
                            }

                            Init-Hora $HoraLimpia
                            $EvA = ""; $ColorA = [System.Drawing.Color]::White; $EvApp = ""; $ColorApp = [System.Drawing.Color]::White

                            if ($linea -match "UpdateHistoryRecord: SessionId=\s*(\d+)") {
                                $Ses = $matches[1]; $EventosTiempo[$HoraLimpia].Sesion = $Ses
                                $TelActual = if ($MapeoTel[$Ses]) { $MapeoTel[$Ses] } else { "Desconocido" }
                                # Normalizar número para deduplicación: quita '+' y el '9' de acceso a línea externa
                                # Ej: 95539991927 → 5539991927  |  +5539991927 → 5539991927  (mismo número, distintos formatos)
                                $TelNorm = ($TelActual -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                $FirmaUnica = "$Ses-$TelNorm"; $EstadoSesion[$Ses] = "ACTIVA"

                                if (-not $LlamadasVistas[$FirmaUnica]) {

                                    # ── CASO ESPECIAL: teléfono real llega para una sesión que ya tiene LÍNEA ABIERTA ──
                                    # El primer UpdateHistoryRecord llega sin teléfono (RemoteUserAddress vacío) y crea
                                    # "LÍNEA ABIERTA SIN MARCAR". El segundo, 1 segundo después, trae el número real.
                                    # En lugar de crear una fila nueva, actualizamos el slot existente.
                                    $UpgradeHecho = $false
                                    if ($TelActual -ne "Desconocido" -and $TelActual -ne "" -and
                                        $SesionHoraInicio.ContainsKey($Ses) -and
                                        $EventosTiempo.ContainsKey($SesionHoraInicio[$Ses]) -and
                                        $EventosTiempo[$SesionHoraInicio[$Ses]].Interpretacion -match "LÍNEA ABIERTA SIN MARCAR") {

                                        $HoraPrev = $SesionHoraInicio[$Ses]
                                        $EventosTiempo[$HoraPrev].Tel = $TelActual
                                        $EventosTiempo[$HoraPrev].Interpretacion = "$symUp INICIO DE LLAMADA (Saliente)"
                                        $EventosTiempo[$HoraPrev].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        $EventosTiempo[$HoraPrev].RawInterpretacion += "$linea`n"
                                        $LlamadasVistas[$FirmaUnica] = $true
                                        $UpgradeHecho = $true
                                        # Limpiar el slot actual (ya no necesita crear eventos propios)
                                        $EventosTiempo.Remove($HoraLimpia) | Out-Null
                                    }

                                    if (-not $UpgradeHecho) {
                                    $LlamadasVistas[$FirmaUnica] = $true; $EventosTiempo[$HoraLimpia].Tel = $TelActual

                                    if ($TelActual -match "^\+?564(3\d{5})\d{6}$") { $EvA = "" }
                                    else {
                                        if ($DirLlamada[$Ses] -eq "SALIENTE") {
                                            # Fallback: si PASO 4 ya puso el label pero sin registrar $SesionHoraInicio
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -match "LÍNEA ABIERTA SIN MARCAR" -and
                                                -not $SesionHoraInicio.ContainsKey($Ses)) {
                                                $SesionHoraInicio[$Ses] = $HoraLimpia
                                            }
                                            # Solo crear INICIO si XML no lo hizo ya
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "señal de llamada|intentando firmarse|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                                                if ($TelActual -match "Desconocido") {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR (Posible evasión)"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                                    # Registrar esta hora para el posible upgrade posterior
                                                    if (-not $SesionHoraInicio.ContainsKey($Ses)) { $SesionHoraInicio[$Ses] = $HoraLimpia }
                                                } else {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE LLAMADA (Saliente)"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                                    if (-not $SesionHoraInicio.ContainsKey($Ses)) { $SesionHoraInicio[$Ses] = $HoraLimpia }
                                                }
                                            }
                                        } elseif ($DirLlamada[$Ses] -ne "ENTRANTE") {
                                            # Fallback: si PASO 4 ya puso LÍNEA ABIERTA en este slot (llamada saliente via
                                            # botón físico / lampara de línea, sin MakeCall en log), registrar la hora ahora
                                            # para que el upgrade pueda ejecutarse cuando llegue el teléfono real en el
                                            # siguiente UpdateHistoryRecord — aunque el guard -notmatch lo bloquee abajo.
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -match "LÍNEA ABIERTA SIN MARCAR" -and
                                                -not $SesionHoraInicio.ContainsKey($Ses)) {
                                                $SesionHoraInicio[$Ses] = $HoraLimpia
                                            }
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "intentando firmarse|INICIO DE LLAMADA|LÍNEA ABIERTA|señal de llamada") {
                                                if ($TelActual -match "Desconocido") {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR (Posible evasión)"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                                } elseif ($TelActual -match "^\d{1,6}$") {
                                                    # Número corto (1-6 dígitos) = extensión interna: no pasa por ACD, no genera señal de llamada
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE SESIÓN (Llamada Interna — Extensión $TelActual)"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::DarkKhaki
                                                } else {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE SESIÓN (Interna/Sistema)"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::DarkKhaki
                                                }
                                            }
                                        }
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                                        $EvA = "UpdateHistoryRecord: SessionID=[$Ses]"; $ColorA = [System.Drawing.Color]::White
                                    }
                                    }  # fin if (-not $UpgradeHecho)
                                }
                            }
                            elseif ($linea -match "OnRequestHoldSession\(\)\. [Ss]ession id=\s*(\d+)") { $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses; $EventosTiempo[$HoraLimpia].Tel=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}; $EvA="|| HOLD MANUAL (Clic del Agente)"; $ColorA=[System.Drawing.Color]::Yellow }
                            elseif ($linea -match "OnRequestUnholdSession\(\)\. [Ss]ession id=\s*(\d+)") { $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses; $EventosTiempo[$HoraLimpia].Tel=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}; $EvA="$symRes UNHOLD MANUAL (Clic del Agente)"; $ColorA=[System.Drawing.Color]::LightGoldenrodYellow }
                            elseif ($linea -match "OnRequestEndSession\(\)\. [Ss]ession=\s*(\d+)") {
                                $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses
                                $TelActual=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}; $EventosTiempo[$HoraLimpia].Tel=$TelActual
                                $CuelguesManuales["$Ses-$TelActual"]=$true; $EvA="$symStop CUELGUE MANUAL (Clic)"; $ColorA=[System.Drawing.Color]::LightCoral
                            }
                            elseif ($linea -match "ProcessSessionEndedEvent: Entry\. connectinoId = (\d+)") {
                                $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses
                                $TelActual=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}
                                $FirmaUnica="$Ses-$TelActual"; $EstadoSesion[$Ses]="CERRADA"
                                # Liberar teléfono del guard SECONDARY para que una rellamada futura genere nuevo INICIO
                                $TelFinNorm = ($TelActual -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                if ($TelFinNorm -ne "Desconocido" -and $TelFinNorm -ne "") { $PhoneYaEnInicio.Remove($TelFinNorm) | Out-Null }
                                $SesionHoraInicio.Remove($Ses) | Out-Null   # limpiar tracking de UpdateHistoryRecord

                                if (-not $CuelguesVistos[$FirmaUnica]) {
                                    $CuelguesVistos[$FirmaUnica]=$true
                                    if ($TelActual -match "^\+?564(3\d{5})\d{6}$") {
                                        $EvA="Entry.ConnectinoID=$Ses"; $ColorA=[System.Drawing.Color]::DarkCyan; $EventosTiempo[$HoraLimpia].Tel="-"
                                    } else {
                                        $EventosTiempo[$HoraLimpia].Tel=$TelActual
                                        if ($CuelguesManuales[$FirmaUnica]) {
                                            if ($TelActual -match "Desconocido" -and $DirLlamada[$Ses] -ne "ENTRANTE") {
                                                $EventosTiempo[$HoraLimpia].Interpretacion="$symStop CUELGUE MANUAL (Línea abierta sin marcar)"
                                                $EventosTiempo[$HoraLimpia].ColorInterpretacion=[System.Drawing.Color]::Orange
                                            } else {
                                                $EventosTiempo[$HoraLimpia].Interpretacion="$symStop FIN DE LLAMADA MANUAL (Colgada por el Asesor)"
                                                $EventosTiempo[$HoraLimpia].ColorInterpretacion=[System.Drawing.Color]::LightCoral
                                            }
                                        } else {
                                            if ($TelActual -match "Desconocido" -and $DirLlamada[$Ses] -ne "ENTRANTE") {
                                                $EventosTiempo[$HoraLimpia].Interpretacion="¡EVASIÓN! Línea abandonada sin marcar (Timeout/Tapón)"
                                                $EventosTiempo[$HoraLimpia].ColorInterpretacion=[System.Drawing.Color]::Red
                                            } else {
                                                $EventosTiempo[$HoraLimpia].Interpretacion="$symStop FIN DE LLAMADA NORMAL"
                                                $EventosTiempo[$HoraLimpia].ColorInterpretacion=[System.Drawing.Color]::DarkGray
                                            }
                                        }
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion+="$linea`n"; $EvA="ProcessSessionEndedEvent: connectinoId=[$Ses]"; $ColorA=[System.Drawing.Color]::White
                                    }
                                }
                            }
                            elseif ($linea -match "Message type= MuteMediaRequest")   { $EvA="x MUTE MANUAL (Silenció)"; $ColorA=[System.Drawing.Color]::Orange }
                            elseif ($linea -match "Message type= UnMuteMediaRequest") { $EvA="o UNMUTE MANUAL (Abrió)"; $ColorA=[System.Drawing.Color]::OrangeRed }
                            elseif ($linea -match "Message type= TransferSessionRequest" -or $linea -match "OnRequestTransferSession") { $EvA="$symArr TRANSFERENCIA INICIADA"; $ColorA=[System.Drawing.Color]::Plum }
                            elseif ($linea -match "Message type= LogoutRequest") { $EvA="¦ ASESOR SOLICITÓ DESFIRMARSE (Clic en Salir)"; $ColorA=[System.Drawing.Color]::LightCoral }
                            elseif ($linea -match "SetPhoneDisplay = <Transferencia realizada") { $EvA="$symOK TRANSFERENCIA COMPLETADA (PBX Confirmó)"; $ColorA=[System.Drawing.Color]::LimeGreen }
                            elseif ($linea -match "Conference_Merged: nFirstCall:\s*(\d+)\s+nConsultCall:\s*(\d+)") { $EvA="$symOK CONFERENCIA ESTABLECIDA (Sesión $($matches[1]) + Sesión $($matches[2]))"; $ColorA=[System.Drawing.Color]::MediumOrchid }

                            if ($EvA -ne "") {
                                # Proteger eventos de alta prioridad (HOLD/UNHOLD/CUELGUE/MUTE/TRANSFERENCIA/DESFIRMARSE)
                                # de ser sobreescritos por eventos de menor prioridad (p.ej. UpdateHistoryRecord o
                                # ProcessSessionEndedEvent) que ocurren en el mismo segundo.
                                if ($EventosTiempo[$HoraLimpia].Agente -notmatch "HOLD|UNHOLD|CUELGUE MANUAL|MUTE|TRANSFERENCIA INICIADA|DESFIRMARSE") {
                                    $EventosTiempo[$HoraLimpia].Agente=$EvA; $EventosTiempo[$HoraLimpia].ColorAgente=$ColorA
                                }
                                $EventosTiempo[$HoraLimpia].RawAgente+="$linea`n"
                            }
                            if ($linea -match "closeSignalingChannel") { $EvApp="¡AVAYA ALERTA!: Caída de Túnel Principal"; $ColorApp=[System.Drawing.Color]::Red }
                            if ($EvApp -ne "") { $EventosTiempo[$HoraLimpia].AppLog=$EvApp; $EventosTiempo[$HoraLimpia].ColorApp=$ColorApp; if($EventosTiempo[$HoraLimpia].Tel -eq "-"){$EventosTiempo[$HoraLimpia].Tel="RED/AVAYA"}; $EventosTiempo[$HoraLimpia].RawAppLog+="$linea`n" }
                        }
                    }
                }
            }

            # ================================================================
            # LIMPIEZA POST-PASO5: Consultas de Transferencia/Conferencia + LÍNEA ABIERTA huérfana
            # ① Sesiones marcadas como consulta → relabeling correcto en la grilla.
            # ② LÍNEA ABIERTA cuya misma sesión ya tiene un INICIO en los próximos 10 s
            #   (upgrade tardío de PASO 5) → eliminar la fila duplicada creada por PASO 4.
            # Nota: si PASO 4 creó la fila con Sesion="-" (distinto segundo de PASO 5),
            # esa fila no puede identificarse por sesión y queda fuera del barrido.
            # ================================================================
            $LimHoras = @($EventosTiempo.Keys | Sort-Object)
            foreach ($HoraLA in $LimHoras) {
                if (-not $EventosTiempo.ContainsKey($HoraLA)) { continue }
                $ObjLA = $EventosTiempo[$HoraLA]
                $SesLA = $ObjLA.Sesion
                if (-not $SesLA -or $SesLA -eq "-" -or $SesLA -eq "") { continue }
                # ① Consulta de Transferencia
                if ($ConsultaTransf.ContainsKey($SesLA) -and
                    $ObjLA.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR|INICIO DE LLAMADA|INICIO DE SESIÓN") {
                    $EventosTiempo[$HoraLA].Interpretacion = "$symArr CONSULTA DE TRANSFERENCIA $symArr $($ConsultaTransf[$SesLA])"
                    $EventosTiempo[$HoraLA].ColorInterpretacion = [System.Drawing.Color]::Plum
                    continue
                }
                # ② Consulta de Conferencia
                if ($ConsultaConf.ContainsKey($SesLA) -and
                    $ObjLA.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR|INICIO DE LLAMADA|INICIO DE SESIÓN") {
                    $EventosTiempo[$HoraLA].Interpretacion = "$symArr CONSULTA DE CONFERENCIA $symArr $($ConsultaConf[$SesLA])"
                    $EventosTiempo[$HoraLA].ColorInterpretacion = [System.Drawing.Color]::Orchid
                    continue
                }
                # ③ LÍNEA ABIERTA huérfana (de PASO 4) cuya sesión ya tiene INICIO en PASO 5
                if ($ObjLA.Interpretacion -notmatch "LÍNEA ABIERTA SIN MARCAR") { continue }
                $Found = $false
                for ($fwd = 1; $fwd -le 10; $fwd++) {
                    try {
                        $FwdH = ([datetime]::ParseExact($HoraLA,"HH:mm:ss",$null).AddSeconds($fwd)).ToString("HH:mm:ss")
                        if ($EventosTiempo.ContainsKey($FwdH) -and
                            $EventosTiempo[$FwdH].Sesion -eq $SesLA -and
                            $EventosTiempo[$FwdH].Interpretacion -match "INICIO DE LLAMADA") {
                            $Found = $true; break
                        }
                    } catch {}
                }
                if ($Found) { $EventosTiempo.Remove($HoraLA) | Out-Null }
            }

            # ================================================================
            # FALLBACK FIN: Call StateChanged OldState=Active,NewState=Disconnected
            # Aplica FIN para sesiones que PASO 5 (ProcessSessionEndedEvent) no cubrió.
            # En este punto $CuelguesVistos y $CuelguesManuales ya están poblados por PASO 5.
            # ================================================================
            foreach ($SesD in $CallStateDisconnected.Keys) {
                $TelD   = if ($MapeoTel[$SesD]) { $MapeoTel[$SesD] } else { "Desconocido" }
                $FirmaD = "$SesD-$TelD"
                if (-not $CuelguesVistos[$FirmaD]) {
                    $HoraD = $CallStateDisconnected[$SesD]
                    Init-Hora $HoraD
                    $CuelguesVistos[$FirmaD]          = $true
                    $EventosTiempo[$HoraD].Sesion      = $SesD
                    $EventosTiempo[$HoraD].Tel         = $TelD
                    if ($CuelguesManuales[$FirmaD]) {
                        $EventosTiempo[$HoraD].Interpretacion      = "$symStop FIN DE LLAMADA MANUAL (Colgada por el Asesor)"
                        $EventosTiempo[$HoraD].ColorInterpretacion = [System.Drawing.Color]::LightCoral
                    } else {
                        $EventosTiempo[$HoraD].Interpretacion      = "$symStop FIN DE LLAMADA NORMAL"
                        $EventosTiempo[$HoraD].ColorInterpretacion = [System.Drawing.Color]::DarkGray
                    }
                    $EventosTiempo[$HoraD].RawInterpretacion += "[Fallback] Call StateChanged OldState=Active,NewState=Disconnected`n"
                }
            }

            # ================================================================
            # DETECTOR DE LLAMADAS ZOMBIE (sin cierre detectado)
            # ================================================================
            foreach ($SesZombi in $LlamadasVistas.Keys) {
                $HorasSesion = @()
                foreach ($h in $EventosTiempo.Keys) { if ($EventosTiempo[$h].Sesion -eq $SesZombi) { $HorasSesion += $h } }
                if ($HorasSesion.Count -gt 0) {
                    $UltimaHoraConocida = ($HorasSesion | Sort-Object)[-1]; Init-Hora $UltimaHoraConocida
                    if ($EventosTiempo[$UltimaHoraConocida].Interpretacion -notmatch "jamás se cerró") {
                        $EventosTiempo[$UltimaHoraConocida].Interpretacion += " `n---> [!] ¡PELIGRO! Sesión $SesZombi jamás se cerró (Posible Crash/Timeout)"
                        $EventosTiempo[$UltimaHoraConocida].ColorInterpretacion = [System.Drawing.Color]::Crimson
                    }
                }
            }

            # Etiqueta de extensiones y logins
            $StrExt = if ($ListaExtensiones.Count -gt 0) { $ListaExtensiones -join " | " } else { "N/A" }
            $StrLog = if ($ListaLogins.Count -gt 0) { $ListaLogins -join " | " } else { "N/A" }
            $lblExtension.Text = "Exts: $StrExt   ---   Login: $StrLog"

            # ================================================================
            # PASO 6: CEREBRO FORENSE V22 — RENDERIZADO
            # ================================================================
            $lblStatus.Text = "PASO 6/6: Renderizando Cerebro Forense..."; $Form.Refresh()
            $HorasOrdenadas = $EventosTiempo.Keys | Sort-Object
            $CurrentAux = "$symUser Estado: USUARIO DESFIRMADO"; $CurrentColor = [System.Drawing.Color]::Gray; $CurrentReasonCode = ""
            $RecienFirmado = $false; $ValidandoLogin = $false; $LoginFallido = $false
            $LlamadaActiva = $false; $SesionActual = "-"

            $DictRC = @{ "0"="DEFAULT"; "1"="COMIDA"; "2"="BAÑO"; "3"="LLAMADA SALIDA"; "4"="CAPACITACION"; "5"="SERVICIOS ESPECIALES"; "6"="COBRANZA"; "7"="SEGUIMIENTO"; "8"="RETRO"; "9"="SISTEMAS" }

            $GridResultados.SuspendLayout()
            if ($HorasOrdenadas -ne $null -and $HorasOrdenadas.Count -gt 0) {
                foreach ($H in $HorasOrdenadas) {
                    $Obj = $EventosTiempo[$H]
                    if ($null -eq $Obj -or $Obj -isnot [hashtable]) { continue }   # guard: slots siempre deben ser hashtable
                    if ($Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "" -and $Obj.Aux -eq "" -and $Obj.Interpretacion -eq "" -and $Obj.Ispeac -eq "" -and $Obj.RawAux -notmatch "Session_LoginAgent failed") { continue }

                    if ($Obj.Interpretacion -match "INICIO DE LLAMADA|LÍNEA ABIERTA") { $LlamadaActiva = $true; $SesionActual = $Obj.Sesion }
                    elseif ($Obj.Interpretacion -match "FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") { $LlamadaActiva = $false; $SesionActual = "-" }

                    $Interp = $Obj.Interpretacion; $ColorInterp = $Obj.ColorInterpretacion

                    # Activar ValidandoLogin desde XML (LoggedIn) o desde Amnesia V21
                    # Guard $RecienFirmado: evita re-trigger si Avaya emite LOGGEDIN desde 2 fuentes (XML + OneXAgent)
                    if ($Obj.Aux -match "ESTADO: LOGGEDIN" -and -not $RecienFirmado) { $ValidandoLogin = $true }
                    if ($Interp -match "Usuario intentando firmarse") { $ValidandoLogin = $true; $RecienFirmado = $false }

                    # Prioridad de interpretación (de mayor a menor)
                    if ($Obj.SysLog -match "PANTALLAZO AZUL|Apagado Sucio|Corte energía|Memoria Virtual Agotada|Tarjeta de Red desconectada" -or $Obj.AppLog -match "Caída de Túnel Principal") {
                        $Interp = "FALLA TÉCNICA (Justificado / Caída de Sistema o Red)"; $ColorInterp = [System.Drawing.Color]::LimeGreen
                    }
                    elseif ($Obj.AppLog -match "Cierre Forzado") { $Interp = "Cierre forzado de AvayaOne-X Agent"; $ColorInterp = [System.Drawing.Color]::Red }
                    elseif ($Obj.AppLog -match "System.Exception") { $Interp = "Falla grave en llamada"; $ColorInterp = [System.Drawing.Color]::Red }
                    elseif ($Obj.Audio -match "Dispositivo de Audio Desconectado") { $Interp = "¡ALERTA CRÍTICA! Diadema desconectada físicamente"; $ColorInterp = [System.Drawing.Color]::Red }
                    elseif ($Obj.Agente -match "CUELGUE MANUAL|LLAMADA COLGADA POR EL ASESOR") {
                        if ($Interp -match "Línea abierta sin marcar") { $Interp = "Precaución: El asesor abrió y cerró línea sin marcar"; $ColorInterp = [System.Drawing.Color]::Orange }
                        else { $Interp = "¡ALERTA! El asesor finalizó la llamada deliberadamente"; $ColorInterp = [System.Drawing.Color]::OrangeRed }
                    }
                    elseif ($Obj.Agente -match "UNHOLD MANUAL") {
                        if ($Interp -match "INICIO DE LLAMADA|LÍNEA ABIERTA") {
                            # UNHOLD coincide con nueva llamada — INICIO tiene prioridad visual
                            $HoldSesU = ""; if ($Obj.RawAgente -match "OnRequestUnholdSession.*?[Ss]ession\s+id=\s*(\d+)") { $HoldSesU = $matches[1] }
                            $Interp += if ($HoldSesU) { "  +  Retoma Ses:$HoldSesU" } else { "  +  Retoma llamada anterior" }
                        } else { $Interp = "Se retoma la llamada"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    }
                    elseif ($Obj.Agente -match "HOLD MANUAL") {
                        if ($Interp -match "INICIO DE LLAMADA|LÍNEA ABIERTA") {
                            # AutoHold al iniciar nueva llamada — INICIO tiene prioridad visual; anotar sesión en espera
                            $HoldSesH = ""; if ($Obj.RawAgente -match "OnRequestHoldSession.*?[Ss]ession\s+id=\s*(\d+)") { $HoldSesH = $matches[1] }
                            $Interp += if ($HoldSesH) { "  +  Hold Ses:$HoldSesH en espera" } else { "  +  Llamada anterior en Hold" }
                        } else { $Interp = "Llamada en Hold"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    }
                    elseif ($Obj.Agente -match "x MUTE MANUAL") { $Interp = "Mute activado"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    elseif ($Obj.Agente -match "o UNMUTE MANUAL") { $Interp = "Mute desactivado"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    elseif ($Obj.AppLog -match "APP CRASH") { $Interp = "Aplicación congelada o cerrada inesperadamente"; $ColorInterp = [System.Drawing.Color]::Orange }
                    elseif ($Obj.Aux -match "SISTEMA_LOGOUT") { $Interp = "El asesor forzó el cierre de su sesión (Logout)"; $ColorInterp = [System.Drawing.Color]::LightSkyBlue; $RecienFirmado = $false }
                    elseif ($Interp -match "INICIO DE LLAMADA" -and $Obj.Tel -match "^565$") { $Interp = "Asesor solicita desfirmarse"; $ColorInterp = [System.Drawing.Color]::IndianRed }
                    elseif ($Interp -match "FIN DE LLAMADA NORMAL" -and $Obj.Tel -match "^565$") { $Interp = "Asesor desfirmado"; $ColorInterp = [System.Drawing.Color]::SkyBlue }
                    elseif ($Obj.RawAux -match "Session_LoginAgent failed") { $Interp = "Fallo en el intento de firmarse"; $ColorInterp = [System.Drawing.Color]::LightCoral; $LoginFallido = $true; $ValidandoLogin = $false }

                    # Lógica de login dinámico
                    elseif ($ValidandoLogin -eq $true) {
                        if ($Obj.RawAux -match "(?i)Enter\s+Aux" -or $Obj.Aux -match "ESTADO: AUX|ESTADO: PENDINGAUX|NUEVO_RC") {
                            $Interp = "Usuario firmado exitosamente"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $ValidandoLogin = $false; $RecienFirmado = $true
                            $CodigoRC = ""; if ($Obj.RawAux -match "(?i)ReasonCode[=\[>:\s]*(\d+)") { $CodigoRC = $matches[1] }
                            # Intentar también desde XML
                            if ($CodigoRC -eq "" -and $Obj.Aux -match "NUEVO_RC_NOMBRE:([^|]+)") { $CodigoRC = $matches[1].Trim() }
                            $NombreRC = if ($CodigoRC -ne "" -and $DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } elseif ($CodigoRC -ne "" -and $CodigoRC -ne "default" -and $CodigoRC -ne "Default") { $CodigoRC } else { "DEFAULT" }
                            $CurrentReasonCode = $CodigoRC
                            if ($CodigoRC -eq "0" -or $CodigoRC -eq "" -or $CodigoRC -match "(?i)^default$") {
                                $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue
                            } else {
                                $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange
                            }
                            $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        }
                        elseif ($Obj.RawAux -match "(?i)AgentStateChanged.*newState=Ready" -or $Obj.Aux -match "ESTADO: READY|ESTADO: AUTOIN|ESTADO: MANUALIN") {
                            $Interp = "Usuario firmado exitosamente"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $ValidandoLogin = $false; $RecienFirmado = $true
                        }
                    }

                    # Estados de agente
                    # Guard: los eventos de llamada tienen prioridad sobre los cambios de estado del agente
                    # cuando coinciden en el mismo segundo (frecuente en XML donde Ready y llamada son simultáneos)
                    elseif ($Obj.Aux -match "GUI_READY_CONFIRMADO" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        $Interp = "Asesor se cambia a Disponible (Confirmado por clic)"; $ColorInterp = [System.Drawing.Color]::Yellow
                    }
                    elseif ($Obj.RawAux -match "(?i)AgentStateChanged.*newState=Ready" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        $Interp = "Asesor se cambia a Disponible"; $ColorInterp = [System.Drawing.Color]::Yellow
                    }
                    elseif ($Obj.Aux -match "ESTADO: READY" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        # ESTADO: READY del XML = notificación de Avaya, no es clic del asesor.
                        # Solo actualizamos la columna Aux en silencio. Sin interpretación visible.
                        # El único evento confiable de "cambió a disponible" es GUI_READY_CONFIRMADO.
                        $CurrentAux = "$symUser Estado: DISPONIBLE"; $CurrentColor = [System.Drawing.Color]::LimeGreen
                        $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        # $Interp queda vacío → fila se filtra si no tiene otros eventos en ese segundo
                    }
                    elseif ($Obj.Aux -match "NUEVO_RC_NOMBRE:([^|]+)" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        # AUX desde XML (texto, no numérico)
                        $NombreRC = $matches[1].Trim()
                        if ($NombreRC -match "(?i)^default$" -or $NombreRC -eq "") { $NombreRC = "DEFAULT" }
                        # Siempre actualizar estado del asesor (independiente de si hay FIN simultáneo)
                        $CurrentAux = if ($NombreRC -eq "DEFAULT") { "$symUser Estado: DEFAULT" } else { "$symUser Estado: AUXILIAR ($NombreRC)" }
                        $CurrentColor = if ($NombreRC -eq "DEFAULT") { [System.Drawing.Color]::CadetBlue } else { [System.Drawing.Color]::Orange }
                        $CurrentReasonCode = $NombreRC; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        if ($Interp -match "FIN DE LLAMADA|CUELGUE MANUAL") {
                            # FIN tiene prioridad visual — agregar nota compacta sin reemplazar
                            if ($NombreRC -ne "DEFAULT") { $Interp += "  +  Auxiliar [$NombreRC]" }
                        } else {
                            $Interp = "Asesor se cambia a Auxiliar [$NombreRC]"; $ColorInterp = [System.Drawing.Color]::Orange
                            if ($NombreRC -eq "DEFAULT" -and $RecienFirmado -eq $true) { $Interp = "" }
                        }
                    }
                    elseif ($Obj.Aux -match "NUEVO_RC:(\d+)" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        # AUX desde OneXAgent (numérico)
                        $CodigoRC = $matches[1]
                        $NombreRC = if ($DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } else { $CodigoRC }
                        if ($NombreRC -eq "0" -or $NombreRC -eq "") { $NombreRC = "DEFAULT" }
                        $CurrentReasonCode = $CodigoRC
                        # Siempre actualizar estado del asesor (independiente de si hay FIN simultáneo)
                        $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        if ($Interp -match "FIN DE LLAMADA|CUELGUE MANUAL") {
                            # FIN tiene prioridad visual — agregar nota compacta sin reemplazar
                            if ($CodigoRC -ne "0") { $Interp += "  +  Auxiliar [$NombreRC]" }
                            else { $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        } else {
                            $Interp = "Asesor se cambia a Auxiliar [$NombreRC] (Confirmado por clic)"; $ColorInterp = [System.Drawing.Color]::Orange
                            if ($CodigoRC -eq "0" -and $RecienFirmado -eq $true) { $Interp = ""; $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        }
                    }
                    elseif ($Obj.RawAux -match "(?i)Enter\s+Aux" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        $CodigoRC = ""; if ($Obj.RawAux -match "(?i)ReasonCode[=\[>:\s]*(\d+)") { $CodigoRC = $matches[1] }
                        $NombreRC = if ($CodigoRC -ne "" -and $DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } elseif ($CodigoRC -ne "") { $CodigoRC } else { "" }
                        # Siempre actualizar estado del asesor (independiente de si hay FIN simultáneo)
                        if ($NombreRC -ne "") { $CurrentReasonCode = $CodigoRC; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        else                  { $CurrentAux = "$symUser Estado: AUXILIAR"; $CurrentColor = [System.Drawing.Color]::LightCoral; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        if ($CodigoRC -eq "0") { $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        if ($Interp -match "FIN DE LLAMADA|CUELGUE MANUAL") {
                            # FIN tiene prioridad visual — agregar nota compacta sin reemplazar
                            if ($CodigoRC -ne "0") { $Interp += if ($NombreRC -ne "") { "  +  Auxiliar [$NombreRC]" } else { "  +  Auxiliar" } }
                        } else {
                            if ($NombreRC -ne "") { $Interp = "Asesor se cambia a Auxiliar [$NombreRC]"; $ColorInterp = [System.Drawing.Color]::Orange }
                            else                  { $Interp = "Asesor se cambia a Auxiliar";             $ColorInterp = [System.Drawing.Color]::Orange }
                            if ($CodigoRC -eq "0" -and $RecienFirmado) { $Interp = ""; $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        }
                    }

                    # Acumulador de estado Aux (tokens separados por |)
                    if ($Obj.Aux -ne "" -and $Obj.RawAux -notmatch "(?i)Enter\s+Aux") {
                        $ArrEstados = $Obj.Aux -split "\|" | Where-Object { $_.Trim() -ne "" }
                        foreach ($Item in $ArrEstados) {
                            if ($Item -match "NUEVO_RC_NOMBRE:(.+)") {
                                $NRC = $matches[1].Trim()
                                if ($NRC -match "(?i)^default$" -or $NRC -eq "") { $CurrentReasonCode = "0"; $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue }
                                else { $CurrentReasonCode = $NRC; $CurrentAux = "$symUser Estado: AUXILIAR ($NRC)"; $CurrentColor = [System.Drawing.Color]::Orange }
                            }
                            elseif ($Item -match "RC: (\d+)") { $CurrentReasonCode = $matches[1]; $NombreRC = if ($DictRC.ContainsKey($CurrentReasonCode)) { $DictRC[$CurrentReasonCode] } else { "RC: $CurrentReasonCode" }; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange }
                            elseif ($Item -match "DEFAULT") { $CurrentReasonCode = "0"; $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue }
                            elseif ($Item -match "SISTEMA_LOGOUT") { $CurrentReasonCode = "10"; $CurrentAux = "$symUser Estado: EN PROCESO DE DESFIRMARSE"; $CurrentColor = [System.Drawing.Color]::LightCoral }
                            elseif ($Item -match "ESTADO: PENDINGAUX") { $CurrentAux = "$symUser Estado: PENDIENTE AUXILIAR"; $CurrentColor = [System.Drawing.Color]::MediumOrchid }
                            elseif ($Item -match "ESTADO: READY|AUTOIN|MANUALIN") { $CurrentReasonCode = ""; $CurrentAux = "$symUser Estado: DISPONIBLE"; $CurrentColor = [System.Drawing.Color]::LimeGreen }
                            elseif ($Item -match "ESTADO: LOGGEDOUT") { $CurrentReasonCode = ""; $CurrentAux = "$symUser Estado: USUARIO DESFIRMADO"; $CurrentColor = [System.Drawing.Color]::Gray }
                            elseif ($Item -match "ESTADO: AUX|ESTADO: NOTREADY") { if ($CurrentReasonCode -ne "" -and $CurrentReasonCode -ne "0") { $NombreRC = if ($DictRC.ContainsKey($CurrentReasonCode)) { $DictRC[$CurrentReasonCode] } else { "RC: $CurrentReasonCode" }; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange } else { $CurrentAux = "$symUser Estado: AUXILIAR"; $CurrentColor = [System.Drawing.Color]::LightCoral } }
                        }
                    }

                    if (-not ($ValidandoLogin -eq $false -and $Interp -match "Usuario firmado exitosamente")) {
                        $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                    }

                    if ($Obj.RawIspeac -ne "" -and $Interp -eq "" -and $LlamadaActiva) {
                        $Interp = "Llamada en curso (Analizando Calidad)"; $ColorInterp = [System.Drawing.Color]::MediumSpringGreen; $Obj.Sesion = $SesionActual
                    }

                    $Obj.Interpretacion = $Interp; $Obj.ColorInterpretacion = $ColorInterp

                    if ($RecienFirmado -eq $true -and $Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "" -and $Obj.Ispeac -eq "" -and $Obj.Interpretacion -eq "") { if ($Obj.Aux -match "Estado: AUXILIAR") { $RecienFirmado = $false; continue } }
                    if ($Obj.Interpretacion -eq "" -and $Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "" -and $Obj.Ispeac -eq "") { if ($Obj.Aux -match "Estado:") { continue } }
                    if ($Interp -match "Usuario firmado exitosamente") { $RecienFirmado = $true } elseif ($Obj.Interpretacion -ne "" -or $Obj.Agente -ne "" -or $Obj.Audio -ne "" -or $Obj.Ispeac -ne "") { $RecienFirmado = $false }

                    $Row = $GridResultados.Rows.Add()
                    $GridResultados.Rows[$Row].Cells["Hora"].Value          = $H
                    $GridResultados.Rows[$Row].Cells["Sesion"].Value        = $Obj.Sesion
                    if ($Obj.Sesion -ne "-" -and $Obj.Sesion -ne "") { $GridResultados.Rows[$Row].Cells["Sesion"].Style.ForeColor = [System.Drawing.Color]::Cyan; $GridResultados.Rows[$Row].Cells["Sesion"].ToolTipText = "Clic para aislar y ver el historial completo de esta llamada" }
                    $GridResultados.Rows[$Row].Cells["Telefono"].Value      = $Obj.Tel
                    if ($Obj.Sesion -ne "-" -and $Obj.Sesion -ne "" -and $RawMapeoTel.ContainsKey($Obj.Sesion)) { $GridResultados.Rows[$Row].Cells["Telefono"].ToolTipText = "Clic para ver Log Original de extracción del número"; $GridResultados.Rows[$Row].Cells["Telefono"].Tag = $RawMapeoTel[$Obj.Sesion].Trim() }
                    $GridResultados.Rows[$Row].Cells["Interpretacion"].Value = $Obj.Interpretacion; $GridResultados.Rows[$Row].Cells["Interpretacion"].Style.ForeColor = $Obj.ColorInterpretacion
                    $GridResultados.Rows[$Row].Cells["EvAgente"].Value      = $Obj.Agente;  $GridResultados.Rows[$Row].Cells["EvAgente"].Style.ForeColor  = $Obj.ColorAgente
                    $GridResultados.Rows[$Row].Cells["EvAudio"].Value       = $Obj.Audio;   $GridResultados.Rows[$Row].Cells["EvAudio"].Style.ForeColor   = $Obj.ColorAudio
                    $GridResultados.Rows[$Row].Cells["EvAux"].Value         = $Obj.Aux;     $GridResultados.Rows[$Row].Cells["EvAux"].Style.ForeColor     = $Obj.ColorAux
                    $GridResultados.Rows[$Row].Cells["EvIspeac"].Value      = $Obj.Ispeac;  $GridResultados.Rows[$Row].Cells["EvIspeac"].Style.ForeColor  = $Obj.ColorIspeac
                    $GridResultados.Rows[$Row].Cells["EvSysLog"].Value      = $Obj.SysLog;  $GridResultados.Rows[$Row].Cells["EvSysLog"].Style.ForeColor  = $Obj.ColorSys
                    $GridResultados.Rows[$Row].Cells["EvAppLog"].Value      = $Obj.AppLog;  $GridResultados.Rows[$Row].Cells["EvAppLog"].Style.ForeColor  = $Obj.ColorApp

                    if ($Obj.RawInterpretacion -ne "") { $GridResultados.Rows[$Row].Cells["Interpretacion"].ToolTipText = "Clic para ver Log Original"; $GridResultados.Rows[$Row].Cells["Interpretacion"].Tag = $Obj.RawInterpretacion.Trim() }
                    if ($Obj.RawAgente -ne "")         { $GridResultados.Rows[$Row].Cells["EvAgente"].ToolTipText = "Clic para ver Log Original"; $GridResultados.Rows[$Row].Cells["EvAgente"].Tag = $Obj.RawAgente.Trim() }
                    if ($Obj.RawAudio -ne "")          { $GridResultados.Rows[$Row].Cells["EvAudio"].ToolTipText = "Clic para ver Log Original"; $GridResultados.Rows[$Row].Cells["EvAudio"].Tag = $Obj.RawAudio.Trim() }
                    if ($Obj.RawAux -ne "")            { $GridResultados.Rows[$Row].Cells["EvAux"].ToolTipText = "Clic para ver Log Original"; $GridResultados.Rows[$Row].Cells["EvAux"].Tag = $Obj.RawAux.Trim() }
                    if ($Obj.RawIspeac -ne "")         { $GridResultados.Rows[$Row].Cells["EvIspeac"].ToolTipText = "Clic para ver Log Original Ispeac"; $GridResultados.Rows[$Row].Cells["EvIspeac"].Tag = $Obj.RawIspeac.Trim() }
                    if ($Obj.RawSysLog -ne "")         { $GridResultados.Rows[$Row].Cells["EvSysLog"].ToolTipText = "Clic para ver Log Original"; $GridResultados.Rows[$Row].Cells["EvSysLog"].Tag = $Obj.RawSysLog.Trim() }
                    if ($Obj.RawAppLog -ne "")         { $GridResultados.Rows[$Row].Cells["EvAppLog"].ToolTipText = "Clic para ver Log Original"; $GridResultados.Rows[$Row].Cells["EvAppLog"].Tag = $Obj.RawAppLog.Trim() }

                    if ($Obj.SysLog -ne "" -or ($Obj.AppLog -match "CRASH|¡AVAYA ALERTA!")) { $GridResultados.Rows[$Row].DefaultCellStyle.BackColor = [System.Drawing.Color]::DarkRed }
                }
            } else { $Row = $GridResultados.Rows.Add(); $GridResultados.Rows[$Row].Cells["EvAgente"].Value = "No se encontró actividad para esta fecha." }
            $GridResultados.ResumeLayout($false)

            if ($GridResultados.Rows.Count -gt 0) {
                $GridResultados.FirstDisplayedScrollingRowIndex = $GridResultados.Rows.Count - 1
                $GridResultados.CurrentCell = $null
                foreach ($row in $GridResultados.Rows) { if ($row.Cells["Interpretacion"].Value -eq "Llamada en curso (Analizando Calidad)") { $row.Visible = $chkVerCalidad.Checked } }
            }

        } else { $Row = $GridResultados.Rows.Add(); $GridResultados.Rows[$Row].Cells["EvAgente"].Value = "No se encontró el directorio de logs de Avaya."; $GridResultados.Rows[$Row].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Red }
    } catch { $Row = $GridResultados.Rows.Add(); $GridResultados.Rows[$Row].Cells["EvAgente"].Value = "Error al leer logs: $($_.Exception.Message) [Linea: $($_.InvocationInfo.ScriptLineNumber)]"; $GridResultados.Rows[$Row].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Red }
    finally { if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null } }

    $Script:SnapEventosTiempo = $EventosTiempo
    $Script:SnapAlertingHoras = $Script:AlertingHoras
    $lblStatus.Text = "Análisis completado para $TargetUser ($FechaVisualStr). XML cargado: $XMLCargado. Puedes usar EXTRACCIÓN RAW para ver logs crudos."
    $btnExtraccion.Enabled = $true; $btnBusqueda.Enabled = $true; $btnExportarCSV.Enabled = $true
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# ====================================================================
# DIAGNÓSTICO DE LLAMADA: LÓGICA DEL MENÚ CONTEXTUAL
# ====================================================================
$mnuDiag.Add_Click({
    $SelRow = if ($GridResultados.SelectedRows.Count -gt 0) { $GridResultados.SelectedRows[0] } else { $null }
    if ($null -eq $SelRow -or $null -eq $Script:SnapEventosTiempo) { return }

    $AlertingHora = $SelRow.Cells["Hora"].Value
    $AlertingDesc = $SelRow.Cells["Interpretacion"].Value
    $Snap   = $Script:SnapEventosTiempo
    $SnapAH = if ($null -ne $Script:SnapAlertingHoras) { $Script:SnapAlertingHoras } else { @{} }

    # ── Ventana modal ────────────────────────────────────────────────────
    $fDiag = New-Object System.Windows.Forms.Form
    $fDiag.Text            = "Diagnostico de llamada — señal $AlertingHora"
    $fDiag.Size            = New-Object System.Drawing.Size(900, 660)
    $fDiag.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $fDiag.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $fDiag.MaximizeBox     = $false
    $fDiag.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)

    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $rtb.ReadOnly   = $true
    $rtb.BackColor  = [System.Drawing.Color]::FromArgb(20,20,20)
    $rtb.ForeColor  = [System.Drawing.Color]::White
    $rtb.Font       = New-Object System.Drawing.Font("Consolas", 9)
    $rtb.WordWrap   = $false
    $rtb.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
    $fDiag.Controls.Add($rtb)

    # Colores
    $cOk   = [System.Drawing.Color]::LimeGreen
    $cFail = [System.Drawing.Color]::OrangeRed
    $cWarn = [System.Drawing.Color]::Gold
    $cInfo = [System.Drawing.Color]::Silver
    $cRaw  = [System.Drawing.Color]::FromArgb(150,150,255)
    $cHdr  = [System.Drawing.Color]::CornflowerBlue
    $cW    = [System.Drawing.Color]::White

    # Helper: agrega línea con color y negrita opcional
    # $args: [0]=texto, [1]=color (opcional), [2]=negrita bool (opcional)
    $AL = {
        $txt  = $args[0]
        $col  = if ($args.Count -gt 1 -and $null -ne $args[1]) { $args[1] } else { $cW }
        $bold = if ($args.Count -gt 2 -and $args[2]) { $true } else { $false }
        $rtb.SelectionStart  = $rtb.TextLength
        $rtb.SelectionLength = 0
        $rtb.SelectionColor  = $col
        $rtb.SelectionFont   = if ($bold) { New-Object System.Drawing.Font($rtb.Font.FontFamily, $rtb.Font.Size, [System.Drawing.FontStyle]::Bold) } else { $rtb.Font }
        $rtb.AppendText("$txt`n")
    }
    # Helper: escribe encabezado de sección
    $AH = {
        $rtb.SelectionStart = $rtb.TextLength; $rtb.SelectionLength = 0
        $rtb.SelectionColor = $cW; $rtb.SelectionFont = $rtb.Font; $rtb.AppendText("`n")
        $rtb.SelectionStart = $rtb.TextLength; $rtb.SelectionLength = 0
        $rtb.SelectionColor = $cHdr
        $rtb.SelectionFont  = New-Object System.Drawing.Font($rtb.Font.FontFamily, $rtb.Font.Size, [System.Drawing.FontStyle]::Bold)
        $rtb.AppendText("=== $($args[0]) ===`n")
    }

    # ── Encabezado ───────────────────────────────────────────────────────
    & $AL "DIAGNOSTICO DE INICIO DE LLAMADA" $cW $true
    & $AL "  Señal detectada : $AlertingDesc" $cWarn
    & $AL "  Hora de señal   : $AlertingHora" $cInfo

    # ════════════════════════════════════════════════════════════════
    # PASO 1: ¿Existe el slot en EventosTiempo?
    # ════════════════════════════════════════════════════════════════
    & $AH "PASO 1: Slot de 'señal de llamada' en datos analizados"
    if ($Snap.ContainsKey($AlertingHora)) {
        & $AL "  [OK] Slot [$AlertingHora] encontrado en EventosTiempo." $cOk
        $RawA = $Snap[$AlertingHora].RawInterpretacion
        if ($RawA -ne "") {
            & $AL "  RAW log del slot (fuentes que generaron este evento):" $cInfo
            foreach ($rl in ($RawA -split "`n")) { if ($rl.Trim() -ne "") { & $AL "    $rl" $cRaw } }
        } else { & $AL "  [Aviso] RawInterpretacion vacio para este slot." $cWarn }
    } else {
        & $AL "  [FALLO] Slot [$AlertingHora] NO existe en el snapshot." $cFail
        & $AL "  El evento puede haberse originado por un fallback sin datos guardados." $cWarn
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 2: Extracción del UUID de la llamada
    # ════════════════════════════════════════════════════════════════
    & $AH "PASO 2: Extraccion del UUID de la llamada (VI UUID)"
    $ViUuid = $null; $UuidFte = ""
    if ($Snap.ContainsKey($AlertingHora)) {
        $RawA = $Snap[$AlertingHora].RawInterpretacion
        if ($RawA -match "id=VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
            $ViUuid = $matches[1]; $UuidFte = "WI.ADD en RawInterpretacion"
        } elseif ($RawA -match "VoiceInteractionImpl\[VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
            $ViUuid = $matches[1]; $UuidFte = "VoiceInteractionImpl en RawInterpretacion"
        } elseif ($RawA -match "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
            $ViUuid = $matches[1]; $UuidFte = "UUID generico en log"
        }
    }
    # Reverse lookup en AlertingHoras si aún no se tiene UUID
    if ($null -eq $ViUuid -and $SnapAH.Count -gt 0) {
        foreach ($kv in $SnapAH.GetEnumerator()) {
            $hBase = ($kv.Value -split ',')[0]
            if ($hBase -eq $AlertingHora) { $ViUuid = $kv.Key; $UuidFte = "AlertingHoras (reverse lookup)"; break }
        }
    }
    if ($null -ne $ViUuid) {
        & $AL "  [OK] UUID extraido: $ViUuid" $cOk
        & $AL "       Fuente: $UuidFte" $cInfo
    } else {
        & $AL "  [FALLO] No se pudo extraer el UUID de esta señal de llamada." $cFail
        & $AL "  Causas posibles:" $cWarn
        & $AL "    * El evento se registro por fallback (CallReceived/CallStateChanged)" $cWarn
        & $AL "      que no almacena VI UUID en RawInterpretacion." $cWarn
        & $AL "    * El log no contenia WI.ADD ni VoiceInteractionImpl con UUID." $cWarn
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 3: ¿UUID registrado en AlertingHoras?
    # ════════════════════════════════════════════════════════════════
    & $AH "PASO 3: Registro en AlertingHoras (vinculo señal -> INICIO)"
    $AhRegistrado = $false
    if ($null -ne $ViUuid) {
        if ($SnapAH.ContainsKey($ViUuid)) {
            $AhRegistrado = $true
            & $AL "  [OK] UUID registrado en AlertingHoras." $cOk
            & $AL "       UUID  : $ViUuid" $cInfo
            & $AL "       Hora  : $($SnapAH[$ViUuid])" $cInfo
            & $AL "  El analizador vinculara PRIMARY CONNECTED Active con esta" $cInfo
            & $AL "  señal de llamada y generara el evento INICIO DE LLAMADA." $cInfo
        } else {
            & $AL "  [FALLO] UUID [$ViUuid] NO esta en AlertingHoras." $cFail
            & $AL "  Sin este registro, PRIMARY CONNECTED Active no puede" $cFail
            & $AL "  identificar la llamada como ENTRANTE y no generara INICIO." $cFail
        }
    } else { & $AL "  No aplicable (UUID no extraido en PASO 2)." $cWarn }

    # ════════════════════════════════════════════════════════════════
    # PASO 4: Búsqueda de INICIO DE LLAMADA en todos los slots
    # ════════════════════════════════════════════════════════════════
    & $AH "PASO 4: Busqueda de INICIO DE LLAMADA (Entrante) en todos los slots"
    $SlotsInicio  = @()
    $SlotsPorUuid = @()
    foreach ($kv in $Snap.GetEnumerator()) {
        if ($kv.Value.Interpretacion -match "INICIO DE LLAMADA \(Entrante\)") {
            $SlotsInicio += $kv.Key
            if ($null -ne $ViUuid -and $kv.Value.RawInterpretacion -match [regex]::Escape($ViUuid)) {
                $SlotsPorUuid += $kv.Key
            }
        }
    }
    $SlotsPorUuid = $SlotsPorUuid | Sort-Object
    $SlotsInicio  = $SlotsInicio  | Sort-Object

    if ($null -ne $ViUuid -and $SlotsPorUuid.Count -gt 0) {
        & $AL "  [OK] INICIO DE LLAMADA encontrado con UUID coincidente:" $cOk
        foreach ($s in $SlotsPorUuid) {
            & $AL "    Slot: $s  ->  $($Snap[$s].Interpretacion)" $cOk
            $rawI = $Snap[$s].RawInterpretacion
            if ($rawI -ne "") {
                foreach ($rl in ($rawI -split "`n")) { if ($rl.Trim() -ne "") { & $AL "      $rl" $cRaw } }
            }
        }
    } elseif ($null -ne $ViUuid) {
        & $AL "  [FALLO] No hay INICIO DE LLAMADA (Entrante) con UUID [$ViUuid]." $cFail
        if ($SlotsInicio.Count -gt 0) {
            & $AL "  (Existen otros INICIO en el dia sin ese UUID:)" $cWarn
            foreach ($s in $SlotsInicio) { & $AL "    $s  ->  $($Snap[$s].Interpretacion)" $cWarn }
        } else { & $AL "  (No hay ningun INICIO DE LLAMADA Entrante en el dia.)" $cWarn }
    } else {
        & $AL "  Sin UUID no es posible buscar con precision." $cWarn
        if ($SlotsInicio.Count -gt 0) {
            & $AL "  Todos los INICIO DE LLAMADA (Entrante) encontrados en el dia:" $cInfo
            foreach ($s in $SlotsInicio) { & $AL "    $s  ->  $($Snap[$s].Interpretacion)" $cInfo }
        } else { & $AL "  (No hay ningun INICIO DE LLAMADA Entrante en el dia.)" $cWarn }
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 5: CONCLUSIÓN Y DIAGNÓSTICO FINAL
    # ════════════════════════════════════════════════════════════════
    & $AH "PASO 5: CONCLUSION Y DIAGNOSTICO FINAL"
    if ($SlotsPorUuid.Count -gt 0) {
        & $AL "  [RESULTADO OK] INICIO DE LLAMADA fue generado correctamente." $cOk $true
        & $AL "  El agente contesto la llamada y el analizador lo registro." $cOk
        & $AL "" $cW
        & $AL "  Señal de llamada : $AlertingHora" $cInfo
        & $AL "  INICIO generado  : $($SlotsPorUuid[0])" $cInfo
        if ($AhRegistrado) {
            try {
                $hB = ($SnapAH[$ViUuid] -split ',')[0]
                $iB = ($SlotsPorUuid[0] -split ',')[0]
                $tA = [datetime]::ParseExact($hB,"HH:mm:ss",$null)
                $tI = [datetime]::ParseExact($iB,"HH:mm:ss",$null)
                $ring = [int]($tI - $tA).TotalSeconds
                & $AL "  Ring time        : ${ring}s (entre señal y contestacion)" $cInfo
            } catch {}
        }
    } elseif ($null -eq $ViUuid) {
        & $AL "  [PROBLEMA] UUID de la llamada no extraido." $cFail $true
        & $AL "  Sin UUID es imposible vincular la señal con el evento de contestacion." $cFail
        & $AL "" $cW
        & $AL "  ACCION: Revisar OneXAgent.log buscando WI.ADD o VoiceInteractionImpl" $cWarn
        & $AL "  cerca de las $AlertingHora para identificar el UUID manualmente." $cWarn
    } elseif (-not $AhRegistrado) {
        & $AL "  [PROBLEMA] UUID encontrado pero NO registrado en AlertingHoras." $cFail $true
        & $AL "  UUID: $ViUuid" $cFail
        & $AL "  El analizador detecto la señal pero no guardo el UUID de vinculo." $cFail
        & $AL "" $cW
        & $AL "  ACCION: Buscar en OneXAgent.log lineas 'VoiceInteractionListImpl.Add'" $cWarn
        & $AL "  que contengan UUID [$ViUuid] y state=Alerting." $cWarn
    } else {
        & $AL "  [PROBLEMA] PRIMARY CONNECTED Active no fue detectado." $cFail $true
        & $AL "  UUID         : $ViUuid" $cInfo
        & $AL "  AlertingHoras: registrado ($($SnapAH[$ViUuid]))" $cOk
        & $AL "  INICIO       : NO generado" $cFail
        & $AL "" $cW
        & $AL "  Causas posibles:" $cWarn
        & $AL "    * La llamada NO fue contestada (perdida o rechazada)." $cWarn
        & $AL "    * El evento VoiceInteractionImpl[...].StateImpl=[...,type=Active,...]" $cWarn
        & $AL "      no aparece en OneXAgent.log para este UUID." $cWarn
        & $AL "    * El archivo OneXAgent.log fue truncado (era muy pesado)." $cWarn
        & $AL "" $cW
        & $AL "  ACCION: Buscar en OneXAgent.log:" $cWarn
        & $AL "    'VoiceInteractionImpl[$ViUuid' con 'type=Active'" $cWarn
        & $AL "    Si existe, reportar el caso para revision del analizador." $cWarn
    }
    & $AL "" $cW
    & $AL "======================================================================" $cHdr
    $rtb.SelectionStart = 0; $rtb.ScrollToCaret()
    [void]$fDiag.ShowDialog($Form)
    $fDiag.Dispose()
})

# ====================================================================
# DIAGNÓSTICO: ¿POR QUÉ NO HAY FIN DE LLAMADA?
# ====================================================================
$mnuDiagFin.Add_Click({
    $SelRow = if ($GridResultados.SelectedRows.Count -gt 0) { $GridResultados.SelectedRows[0] } else { $null }
    if ($null -eq $SelRow -or $null -eq $Script:SnapEventosTiempo) { return }

    $InicioHora  = $SelRow.Cells["Hora"].Value
    $InicioDesc  = $SelRow.Cells["Interpretacion"].Value
    $Snap        = $Script:SnapEventosTiempo

    # Extraer Sesion y Tel del slot INICIO
    $InicioSlot  = if ($Snap.ContainsKey($InicioHora)) { $Snap[$InicioHora] } else { $null }
    $SesId       = if ($null -ne $InicioSlot -and $InicioSlot.Sesion -ne "-" -and $InicioSlot.Sesion -ne "") { $InicioSlot.Sesion } else { "" }
    $TelRaw      = if ($null -ne $InicioSlot -and $InicioSlot.Tel -ne "-" -and $InicioSlot.Tel -ne "") { $InicioSlot.Tel } else { "" }
    $TelNormI    = ($TelRaw -replace '^\+','') -replace '^9(\d{10,})','$1'

    # ── Ventana modal ────────────────────────────────────────────────────
    $fDiagFin = New-Object System.Windows.Forms.Form
    $fDiagFin.Text            = "Diagnostico FIN DE LLAMADA — inicio $InicioHora"
    $fDiagFin.Size            = New-Object System.Drawing.Size(920, 680)
    $fDiagFin.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $fDiagFin.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $fDiagFin.MaximizeBox     = $false
    $fDiagFin.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)

    $rtbF = New-Object System.Windows.Forms.RichTextBox
    $rtbF.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $rtbF.ReadOnly   = $true
    $rtbF.BackColor  = [System.Drawing.Color]::FromArgb(20,20,20)
    $rtbF.ForeColor  = [System.Drawing.Color]::White
    $rtbF.Font       = New-Object System.Drawing.Font("Consolas", 9)
    $rtbF.WordWrap   = $false
    $rtbF.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
    $fDiagFin.Controls.Add($rtbF)

    # Colores
    $cOk   = [System.Drawing.Color]::LimeGreen
    $cFail = [System.Drawing.Color]::OrangeRed
    $cWarn = [System.Drawing.Color]::Gold
    $cInfo = [System.Drawing.Color]::Silver
    $cRaw  = [System.Drawing.Color]::FromArgb(150,150,255)
    $cHdr  = [System.Drawing.Color]::CornflowerBlue
    $cW    = [System.Drawing.Color]::White

    $BL = {
        $txt  = $args[0]
        $col  = if ($args.Count -gt 1 -and $null -ne $args[1]) { $args[1] } else { $cW }
        $bold = if ($args.Count -gt 2 -and $args[2]) { $true } else { $false }
        $rtbF.SelectionStart  = $rtbF.TextLength
        $rtbF.SelectionLength = 0
        $rtbF.SelectionColor  = $col
        $rtbF.SelectionFont   = if ($bold) { New-Object System.Drawing.Font($rtbF.Font.FontFamily, $rtbF.Font.Size, [System.Drawing.FontStyle]::Bold) } else { $rtbF.Font }
        $rtbF.AppendText("$txt`n")
    }
    $BH = {
        $rtbF.SelectionStart = $rtbF.TextLength; $rtbF.SelectionLength = 0
        $rtbF.SelectionColor = $cW; $rtbF.SelectionFont = $rtbF.Font; $rtbF.AppendText("`n")
        $rtbF.SelectionStart = $rtbF.TextLength; $rtbF.SelectionLength = 0
        $rtbF.SelectionColor = $cHdr
        $rtbF.SelectionFont  = New-Object System.Drawing.Font($rtbF.Font.FontFamily, $rtbF.Font.Size, [System.Drawing.FontStyle]::Bold)
        $rtbF.AppendText("=== $($args[0]) ===`n")
    }

    # ── Encabezado ───────────────────────────────────────────────────────
    & $BL "DIAGNOSTICO: ¿POR QUE NO HAY FIN DE LLAMADA?" $cW $true
    & $BL "  Evento de inicio : $InicioDesc" $cWarn
    & $BL "  Hora de inicio   : $InicioHora" $cInfo
    & $BL "  Sesion ID        : $(if ($SesId -ne '') { $SesId } else { '(no detectada)' })" $cInfo
    & $BL "  Telefono         : $(if ($TelRaw -ne '') { $TelRaw } else { '(no detectado)' })" $cInfo

    # ════════════════════════════════════════════════════════════════
    # PASO 1: Verificar el slot INICIO
    # ════════════════════════════════════════════════════════════════
    & $BH "PASO 1: Datos del slot INICIO en EventosTiempo"
    if ($null -ne $InicioSlot) {
        & $BL "  [OK] Slot [$InicioHora] encontrado en EventosTiempo." $cOk
        if ($InicioSlot.RawInterpretacion -ne "") {
            & $BL "  Fuentes que generaron el INICIO:" $cInfo
            foreach ($rl in ($InicioSlot.RawInterpretacion -split "`n")) {
                if ($rl.Trim() -ne "") { & $BL "    $rl" $cRaw }
            }
        } else {
            & $BL "  [Aviso] RawInterpretacion vacio para este slot." $cWarn
        }
    } else {
        & $BL "  [AVISO] Slot [$InicioHora] no encontrado en el snapshot." $cWarn
        & $BL "  El evento fue generado por la capa de presentacion (no por el analizador base)." $cWarn
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 2: Buscar FIN DE LLAMADA posterior con misma sesión o teléfono
    # ════════════════════════════════════════════════════════════════
    & $BH "PASO 2: Busqueda de FIN DE LLAMADA / CUELGUE posterior"

    # Ordenar todas las horas del snapshot cronológicamente y filtrar las posteriores al INICIO
    $TodasHoras = $Snap.Keys | Sort-Object
    $HorasPost  = $TodasHoras | Where-Object { $_ -gt $InicioHora }

    $FinEncontrado   = $false
    $FinHora         = ""
    $FinDesc         = ""
    $FinSesion       = ""
    $FinTel          = ""
    $FinRaw          = ""
    $MatchPorSesion    = $false
    $MatchPorTel       = $false
    $FinSesionMismatch = $false   # true cuando el match es solo por tel y los IDs de sesion difieren

    foreach ($h in $HorasPost) {
        $slot = $Snap[$h]
        $interp = $slot.Interpretacion
        # Buscar solo eventos de fin real: FIN DE LLAMADA o CUELGUE MANUAL.
        # "LÍNEA ABIERTA SIN MARCAR (Posible evasión)" es un INICIO saliente, NO un fin — excluir.
        if ($interp -notmatch "FIN DE LLAMADA|CUELGUE MANUAL") { continue }

        # Verificar si corresponde a la misma llamada
        $slotSes    = $slot.Sesion
        $slotTelRaw = $slot.Tel
        $slotTelN   = ($slotTelRaw -replace '^\+','') -replace '^9(\d{10,})','$1'

        $porSes = ($SesId -ne "" -and $slotSes -eq $SesId)
        $porTel = ($TelNormI -ne "" -and $TelNormI -ne "Desconocido" -and $slotTelN -eq $TelNormI -and $slotTelN -ne "")

        if ($porSes -or $porTel) {
            $FinEncontrado  = $true
            $FinHora        = $h
            $FinDesc        = $interp
            $FinSesion      = $slotSes
            $FinTel         = $slotTelRaw
            $FinRaw         = $slot.RawInterpretacion
            $MatchPorSesion = $porSes
            $MatchPorTel    = $porTel
            # Detectar match dudoso: solo teléfono coincide pero sesiones son distintas y conocidas
            # → podría ser otra llamada al mismo número, no el fin de ESTA llamada
            $FinSesionMismatch = ($porTel -and -not $porSes -and
                                  $SesId -ne "" -and
                                  $slotSes -ne "-" -and $slotSes -ne "")
            break
        }
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 3: Resultado de la búsqueda
    # ════════════════════════════════════════════════════════════════
    & $BH "PASO 3: Resultado de la busqueda"

    if ($FinEncontrado -and -not $FinSesionMismatch) {
        # ── Caso A: match confiable (por Sesion ID, o tel cuando no hay sesion en el INICIO) ──
        & $BL "  [ENCONTRADO] Existe un evento de fin de llamada posterior." $cOk $true
        & $BL "" $cW
        & $BL "  Hora del FIN    : $FinHora" $cOk
        & $BL "  Tipo de fin     : $FinDesc" $cOk
        & $BL "  Sesion del FIN  : $FinSesion" $cInfo
        & $BL "  Telefono del FIN: $FinTel" $cInfo
        & $BL "  Coincidencia    : $(if ($MatchPorSesion -and $MatchPorTel) { 'Sesion ID + Telefono' } elseif ($MatchPorSesion) { 'Sesion ID' } else { 'Telefono normalizado' })" $cInfo

        # Calcular duración si es posible
        try {
            $tInicio = [datetime]::ParseExact($InicioHora, "HH:mm:ss", $null)
            $tFin    = [datetime]::ParseExact($FinHora,    "HH:mm:ss", $null)
            $dur     = ($tFin - $tInicio).TotalSeconds
            & $BL "  Duracion aprox. : $([int]$dur) segundos ($([math]::Floor($dur/60))m $([int]($dur%60))s)" $cOk
        } catch { & $BL "  Duracion aprox. : No calculable (formato de hora inesperado)" $cWarn }

        & $BL "" $cW
        if ($FinRaw -ne "") {
            & $BL "  RAW del evento FIN (fuentes de datos):" $cInfo
            foreach ($rl in ($FinRaw -split "`n")) {
                if ($rl.Trim() -ne "") { & $BL "    $rl" $cRaw }
            }
        } else {
            & $BL "  [Aviso] No hay RawInterpretacion para el slot FIN." $cWarn
        }
        & $BL "" $cW
        & $BL "  CONCLUSION: El FIN DE LLAMADA SI existe en los datos analizados." $cOk $true
        & $BL "  Si no lo ves en la grilla, puede ser que:" $cWarn
        & $BL "    * Hay un filtro activo en la grilla que oculta ese segundo." $cWarn
        & $BL "    * El FIN esta en un segundo sin otros eventos y fue omitido por el filtro de fila vacia." $cWarn
        & $BL "    * La grilla no se ha actualizado; intenta volver a analizar." $cWarn

    } elseif ($FinEncontrado -and $FinSesionMismatch) {
        # ── Caso B: match dudoso — mismo teléfono pero distinto ID de sesión ──
        & $BL "  [ADVERTENCIA] Se encontro un evento con el mismo telefono pero DISTINTO ID de sesion." $cWarn $true
        & $BL "  Es probable que sea una llamada DIFERENTE al mismo numero, no el fin de ESTA llamada." $cWarn
        & $BL "" $cW
        & $BL "  Evento encontrado (posiblemente otra llamada):" $cInfo
        & $BL "    Hora           : $FinHora" $cWarn
        & $BL "    Tipo           : $FinDesc" $cWarn
        & $BL "    Sesion del FIN : $FinSesion  ← DISTINTA a la del INICIO ($SesId)" $cFail
        & $BL "    Telefono       : $FinTel  (normalizado: $( ($FinTel -replace '^\+','') -replace '^9(\d{10,})','$1' ))" $cWarn
        & $BL "" $cW
        if ($FinRaw -ne "") {
            & $BL "  RAW del evento dudoso:" $cInfo
            foreach ($rl in ($FinRaw -split "`n")) {
                if ($rl.Trim() -ne "") { & $BL "    $rl" $cRaw }
            }
        }
        & $BL "" $cW
        & $BL "  CONCLUSION: NO se puede confirmar el FIN DE LLAMADA para esta sesion." $cFail $true
        & $BL "  El evento hallado pertenece a la sesion $FinSesion, no a la sesion $SesId." $cFail
        & $BL "  Continua el analisis de señales de desconexion en los pasos siguientes." $cWarn
        & $BL "" $cW

        # Continuar igual que el caso de NO ENCONTRADO: pasos 4 y 5
        & $BH "PASO 4: Analisis de señales de desconexion en logs posteriores"
        $SeñalesEncontradas4 = @()
        $LogoutPost4     = $false; $CrashPost4      = $false
        $DisconnectPost4 = $false; $ProcessEndPost4 = $false
        $UltimaHoraLog4  = ""
        foreach ($h4 in $HorasPost) {
            $slot4 = $Snap[$h4]; $UltimaHoraLog4 = $h4
            $rawTotal4 = "$($slot4.RawInterpretacion)$($slot4.RawAgente)$($slot4.RawAux)$($slot4.RawSysLog)$($slot4.RawAppLog)"
            if ($rawTotal4 -match "State=Disconnected|Disconnected.*?$($SesId)|(?i)disconnect") {
                $DisconnectPost4 = $true; $SeñalesEncontradas4 += "[${h4}] State=Disconnected detectado en logs"
            }
            if ($rawTotal4 -match "ProcessSessionEndedEvent") {
                $ProcessEndPost4 = $true; $SeñalesEncontradas4 += "[${h4}] ProcessSessionEndedEvent detectado"
            }
            if ($slot4.Interpretacion -match "LOGOUT|Firmado|firmado|cierre.*sesion" -or $rawTotal4 -match "(?i)SISTEMA_LOGOUT|logoff|logout") {
                $LogoutPost4 = $true; $SeñalesEncontradas4 += "[${h4}] Evento de cierre de sesion del agente"
            }
            if ($rawTotal4 -match "System\.Exception|Application Error|crash|(?i)fatalexception") {
                $CrashPost4 = $true; $SeñalesEncontradas4 += "[${h4}] Posible crash o error de aplicacion"
            }
        }
        if ($SeñalesEncontradas4.Count -gt 0) {
            & $BL "  Se encontraron señales en slots posteriores:" $cWarn
            foreach ($s4 in $SeñalesEncontradas4) { & $BL "    * $s4" $cWarn }
        } else {
            & $BL "  No se encontraron señales de desconexion para la sesion $SesId en slots posteriores." $cInfo
        }
        if ($UltimaHoraLog4 -ne "") { & $BL "  Ultimo segundo con datos en el snapshot: $UltimaHoraLog4" $cInfo }

        & $BH "PASO 5: Causas posibles del FIN DE LLAMADA faltante"
        & $BL "  [1] Log truncado — el FIN ocurrio despues del ultimo segundo registrado ($UltimaHoraLog4)" $cFail $true
        & $BL "      ACCION: Verificar si hay archivo de log rotado (.log.1) con mas eventos." $cInfo
        & $BL "" $cW
        & $BL "  [2] Llamada aun activa al momento de la captura del log" $cWarn $true
        & $BL "      ACCION: Confirmar con el asesor si la llamada habia terminado al tomar el log." $cInfo
        & $BL "" $cW
        & $BL "  [3] El ID de sesion cambio (transferencia o consulta)" $cWarn $true
        & $BL "      ACCION: Buscar en OneXAgent.log 'ProcessSessionEndedEvent' con ID distinto a $SesId." $cInfo
        & $BL "" $cW
        if ($LogoutPost4) {
            & $BL "  [4] * El agente cerro sesion con la llamada en curso" $cFail $true
            & $BL "      Se detecto un evento de cierre de sesion posterior al INICIO." $cFail; & $BL "" $cW
        }
        if ($CrashPost4) {
            & $BL "  [5] * Crash de la aplicacion OneX" $cFail $true
            & $BL "      Se detecto un error de aplicacion posterior al INICIO." $cFail; & $BL "" $cW
        }
        & $BL "  RESUMEN:" $cHdr $true
        & $BL "    State=Disconnected detectado    : $(if ($DisconnectPost4) { 'SI' } else { 'NO' })" $(if ($DisconnectPost4) { $cWarn } else { $cInfo })
        & $BL "    ProcessSessionEndedEvent detectado: $(if ($ProcessEndPost4) { 'SI' } else { 'NO' })" $(if ($ProcessEndPost4) { $cWarn } else { $cInfo })
        & $BL "    Cierre de sesion de agente post.: $(if ($LogoutPost4) { 'SI' } else { 'NO' })" $(if ($LogoutPost4) { $cFail } else { $cInfo })
        & $BL "    Crash/Error de aplicacion post. : $(if ($CrashPost4) { 'SI' } else { 'NO' })" $(if ($CrashPost4) { $cFail } else { $cInfo })

    } else {
        & $BL "  [NO ENCONTRADO] No se encontro evento de FIN DE LLAMADA / CUELGUE posterior" $cFail $true
        & $BL "  que corresponda a esta llamada (misma sesion o mismo telefono)." $cFail
        & $BL "" $cW
        if ($SesId -eq "") {
            & $BL "  [Aviso] No se pudo extraer el ID de sesion del slot INICIO." $cWarn
            & $BL "  La busqueda se baso solo en telefono normalizado: '$TelNormI'" $cWarn
        } else {
            & $BL "  Criterios de busqueda utilizados:" $cInfo
            & $BL "    Sesion ID        : $SesId" $cInfo
            & $BL "    Telefono norm.   : $TelNormI" $cInfo
        }

        # ════════════════════════════════════════════════════════════════
        # PASO 4: Análisis de señales de desconexión en logs crudos
        # ════════════════════════════════════════════════════════════════
        & $BH "PASO 4: Analisis de señales de desconexion en logs posteriores"

        $SeñalesEncontradas = @()
        $LogoutPost     = $false
        $CrashPost      = $false
        $DisconnectPost = $false
        $ProcessEndPost = $false
        $UltimaHoraLog  = ""

        foreach ($h in $HorasPost) {
            $slot = $Snap[$h]
            $UltimaHoraLog = $h

            # Comprobar si el slot tiene datos relevantes
            $rawTotal = "$($slot.RawInterpretacion)$($slot.RawAgente)$($slot.RawAux)$($slot.RawSysLog)$($slot.RawAppLog)"

            if ($rawTotal -match "State=Disconnected|Disconnected.*?$($SesId)|(?i)disconnect") {
                $DisconnectPost = $true
                $SeñalesEncontradas += "[${h}] State=Disconnected detectado en logs"
            }
            if ($rawTotal -match "ProcessSessionEndedEvent") {
                $ProcessEndPost = $true
                $SeñalesEncontradas += "[${h}] ProcessSessionEndedEvent detectado"
            }
            if ($slot.Interpretacion -match "LOGOUT|Firmado|firmado|cierre.*sesion" -or $rawTotal -match "(?i)SISTEMA_LOGOUT|logoff|logout") {
                $LogoutPost = $true
                $SeñalesEncontradas += "[${h}] Evento de cierre de sesion del agente"
            }
            if ($rawTotal -match "System\.Exception|Application Error|crash|(?i)fatalexception") {
                $CrashPost = $true
                $SeñalesEncontradas += "[${h}] Posible crash o error de aplicacion"
            }
        }

        if ($SeñalesEncontradas.Count -gt 0) {
            & $BL "  Se encontraron señales de desconexion en slots posteriores:" $cWarn
            foreach ($s in $SeñalesEncontradas) { & $BL "    * $s" $cWarn }
        } else {
            & $BL "  No se encontraron señales de State=Disconnected, ProcessSessionEndedEvent," $cInfo
            & $BL "  ni logout/crash en ninguno de los $($HorasPost.Count) slots posteriores al INICIO." $cInfo
        }

        if ($UltimaHoraLog -ne "") {
            & $BL "" $cW
            & $BL "  Ultimo segundo con datos en el snapshot: $UltimaHoraLog" $cInfo
        }

        # ════════════════════════════════════════════════════════════════
        # PASO 5: Conclusión y causas posibles
        # ════════════════════════════════════════════════════════════════
        & $BH "PASO 5: Causas posibles del FIN DE LLAMADA faltante"

        & $BL "  Las causas mas comunes son:" $cWarn $true
        & $BL "" $cW

        & $BL "  [1] Log truncado (causa mas frecuente)" $cFail $true
        & $BL "      El archivo OneXAgent.log tiene un tamaño maximo." $cInfo
        & $BL "      Si la llamada termino despues del ultimo segundo registrado ($UltimaHoraLog)," $cInfo
        & $BL "      el evento de desconexion quedo fuera del archivo capturado." $cInfo
        & $BL "      ACCION: Verificar si el log continua en un archivo rotado (p.ej. .log.1)." $cInfo
        & $BL "" $cW

        & $BL "  [2] Llamada aun activa al momento de la captura" $cWarn $true
        & $BL "      Si el log fue tomado con la llamada en curso, no existira FIN." $cInfo
        & $BL "      ACCION: Confirmar con el asesor si la llamada termino antes del horario del log." $cInfo
        & $BL "" $cW

        & $BL "  [3] El analizador no detecto el evento de desconexion" $cWarn $true
        & $BL "      El patron ProcessSessionEndedEvent o State=Disconnected puede venir" $cInfo
        & $BL "      con un ID de sesion diferente si hubo transferencia o consulta." $cInfo
        if ($SesId -ne "") {
            & $BL "      ACCION: Buscar en OneXAgent.log el texto 'ProcessSessionEndedEvent'" $cInfo
            & $BL "              y verificar si 'connectinoId' es distinto a $SesId." $cInfo
        }
        & $BL "" $cW

        if ($LogoutPost) {
            & $BL "  [4] * El agente cerro sesion con la llamada en curso" $cFail $true
            & $BL "      Se detecto un evento de cierre de sesion del agente posterior al INICIO." $cFail
            & $BL "      El agente puede haber forzado su desfirmado mientras estaba en llamada." $cInfo
            & $BL "" $cW
        }
        if ($CrashPost) {
            & $BL "  [5] * Crash o error grave de la aplicacion OneX" $cFail $true
            & $BL "      Se detecto un error de aplicacion posterior al INICIO." $cFail
            & $BL "      Un cierre inesperado puede interrumpir el registro del FIN DE LLAMADA." $cInfo
            & $BL "" $cW
        }
        if ($DisconnectPost -and -not $ProcessEndPost) {
            & $BL "  [6] * State=Disconnected sin ProcessSessionEndedEvent" $cWarn $true
            & $BL "      Se vio señal de desconexion pero sin el evento de fin de sesion." $cWarn
            & $BL "      Puede ser un cuelgue de la linea sin cierre formal del lado del agente." $cInfo
            & $BL "" $cW
        }

        & $BL "  RESUMEN:" $cHdr $true
        & $BL "    Señales de desconexion encontradas : $($SeñalesEncontradas.Count)" $cInfo
        & $BL "    State=Disconnected en logs         : $(if ($DisconnectPost) { 'SI' } else { 'NO' })" $(if ($DisconnectPost) { $cWarn } else { $cInfo })
        & $BL "    ProcessSessionEndedEvent encontrado: $(if ($ProcessEndPost) { 'SI' } else { 'NO' })" $(if ($ProcessEndPost) { $cWarn } else { $cInfo })
        & $BL "    Cierre de sesion de agente post.   : $(if ($LogoutPost) { 'SI' } else { 'NO' })" $(if ($LogoutPost) { $cFail } else { $cInfo })
        & $BL "    Crash/Error de aplicacion post.    : $(if ($CrashPost) { 'SI' } else { 'NO' })" $(if ($CrashPost) { $cFail } else { $cInfo })
    }

    & $BL "" $cW
    & $BL "======================================================================" $cHdr
    $rtbF.SelectionStart = 0; $rtbF.ScrollToCaret()
    [void]$fDiagFin.ShowDialog($Form)
    $fDiagFin.Dispose()
})

# ====================================================================
# DIAGNÓSTICO: ¿POR QUÉ FINALIZÓ ESTA LLAMADA? (MANUAL o NORMAL)
# ====================================================================
$mnuDiagExplicaFin.Add_Click({
    $SelRow = if ($GridResultados.SelectedRows.Count -gt 0) { $GridResultados.SelectedRows[0] } else { $null }
    if ($null -eq $SelRow -or $null -eq $Script:SnapEventosTiempo) { return }

    $FinHora  = $SelRow.Cells["Hora"].Value
    $FinDesc  = $SelRow.Cells["Interpretacion"].Value
    $Snap     = $Script:SnapEventosTiempo
    $FinSlot  = if ($Snap.ContainsKey($FinHora)) { $Snap[$FinHora] } else { $null }

    $SesId    = if ($null -ne $FinSlot -and $FinSlot.Sesion -ne "-" -and $FinSlot.Sesion -ne "") { $FinSlot.Sesion } else { "" }
    $TelRaw   = if ($null -ne $FinSlot -and $FinSlot.Tel -ne "-" -and $FinSlot.Tel -ne "") { $FinSlot.Tel } else { "" }
    $TelNorm  = ($TelRaw -replace '^\+','') -replace '^9(\d{10,})$','$1'
    $RawInterp = if ($null -ne $FinSlot) { $FinSlot.RawInterpretacion } else { "" }
    $RawAgen   = if ($null -ne $FinSlot) { $FinSlot.RawAgente }         else { "" }

    # ── Ventana modal ────────────────────────────────────────────────────
    $fEF = New-Object System.Windows.Forms.Form
    $fEF.Text            = "Diagnostico FIN — $FinHora"
    $fEF.Size            = New-Object System.Drawing.Size(940, 700)
    $fEF.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $fEF.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $fEF.MaximizeBox     = $false
    $fEF.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)

    $rtbEF = New-Object System.Windows.Forms.RichTextBox
    $rtbEF.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $rtbEF.ReadOnly   = $true
    $rtbEF.BackColor  = [System.Drawing.Color]::FromArgb(20,20,20)
    $rtbEF.ForeColor  = [System.Drawing.Color]::White
    $rtbEF.Font       = New-Object System.Drawing.Font("Consolas", 9)
    $rtbEF.WordWrap   = $false
    $rtbEF.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
    $fEF.Controls.Add($rtbEF)

    # Colores
    $cOk   = [System.Drawing.Color]::LimeGreen
    $cFail = [System.Drawing.Color]::OrangeRed
    $cWarn = [System.Drawing.Color]::Gold
    $cInfo = [System.Drawing.Color]::Silver
    $cRaw  = [System.Drawing.Color]::FromArgb(150,150,255)
    $cHdr  = [System.Drawing.Color]::CornflowerBlue
    $cW    = [System.Drawing.Color]::White
    $cMan  = [System.Drawing.Color]::LightCoral
    $cNrm  = [System.Drawing.Color]::DarkGray

    $EL = {
        $txt  = $args[0]
        $col  = if ($args.Count -gt 1 -and $null -ne $args[1]) { $args[1] } else { $cW }
        $bold = if ($args.Count -gt 2 -and $args[2]) { $true } else { $false }
        $rtbEF.SelectionStart  = $rtbEF.TextLength
        $rtbEF.SelectionLength = 0
        $rtbEF.SelectionColor  = $col
        $rtbEF.SelectionFont   = if ($bold) { New-Object System.Drawing.Font($rtbEF.Font.FontFamily, $rtbEF.Font.Size, [System.Drawing.FontStyle]::Bold) } else { $rtbEF.Font }
        $rtbEF.AppendText("$txt`n")
    }
    $EH = {
        $rtbEF.SelectionStart = $rtbEF.TextLength; $rtbEF.SelectionLength = 0
        $rtbEF.SelectionColor = $cW; $rtbEF.SelectionFont = $rtbEF.Font; $rtbEF.AppendText("`n")
        $rtbEF.SelectionStart = $rtbEF.TextLength; $rtbEF.SelectionLength = 0
        $rtbEF.SelectionColor = $cHdr
        $rtbEF.SelectionFont  = New-Object System.Drawing.Font($rtbEF.Font.FontFamily, $rtbEF.Font.Size, [System.Drawing.FontStyle]::Bold)
        $rtbEF.AppendText("=== $($args[0]) ===`n")
    }

    # Determinar tipo de fin
    $EsFallback = ($RawInterp -match "\[Fallback\]")
    $EsManual   = ($FinDesc -match "MANUAL|CUELGUE")
    $EsEvasion  = ($FinDesc -match "EVASIÓN|abandonada")

    # ── Encabezado ───────────────────────────────────────────────────────
    & $EL "DIAGNOSTICO: ¿POR QUE FINALIZO ESTA LLAMADA?" $cW $true
    & $EL "" $cW
    $colorTipo = if ($EsManual) { $cMan } elseif ($EsEvasion) { $cFail } else { $cNrm }
    & $EL "  Clasificacion   : $FinDesc" $colorTipo $true
    & $EL "  Hora de FIN     : $FinHora" $cInfo
    & $EL "  Sesion ID       : $(if ($SesId -ne '') { $SesId } else { '(no detectada)' })" $cInfo
    & $EL "  Telefono        : $(if ($TelRaw -ne '') { $TelRaw } else { '(no detectado)' })" $cInfo
    if ($EsFallback) {
        & $EL "  Origen del FIN  : FALLBACK (Call StateChanged Disconnected — ProcessSessionEndedEvent no aparecio)" $cWarn $true
    } else {
        & $EL "  Origen del FIN  : PASO 5 EndpointLog (ProcessSessionEndedEvent — cierre formal de sesion)" $cOk
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 1: Señales encontradas en el slot FIN
    # ════════════════════════════════════════════════════════════════
    & $EH "PASO 1: Señales de desconexion encontradas en este segundo"

    $rawTodo = "$RawInterp$RawAgen$($FinSlot.RawAux)$($FinSlot.RawSysLog)$($FinSlot.RawAppLog)"

    $HayOnRequest   = $rawTodo -match "OnRequestEndSession"
    $HayProcEnd     = $rawTodo -match "ProcessSessionEndedEvent"
    $HayCallState   = $rawTodo -match "Call StateChanged.*?NewState=Disconnected|\[Fallback\]"
    $HayPhoneSvc    = $rawTodo -match "PhoneService_CallStateChanged.*?State=Disconnected|PhoneService_CallUpdated.*?State=Disconnected"

    if ($HayOnRequest) {
        & $EL "  [SEÑAL] OnRequestEndSession() detectado" $cOk $true
        & $EL "          -> El asesor presiono el boton Colgar en la aplicacion OneX." $cInfo
        & $EL "          -> Este evento es exclusivo del agente: no puede dispararse si el cliente cuelga." $cInfo
    }
    if ($HayProcEnd) {
        & $EL "  [SEÑAL] ProcessSessionEndedEvent detectado" $cOk $true
        & $EL "          -> Confirmacion formal del sistema Avaya: la sesion H.323 fue cerrada." $cInfo
        & $EL "          -> Este es el evento primario para registrar el FIN DE LLAMADA." $cInfo
    }
    if ($HayCallState) {
        & $EL "  [SEÑAL] Call StateChanged OldState=Active,NewState=Disconnected detectado" $(if ($EsFallback) { $cWarn } else { $cOk }) $true
        if ($EsFallback) {
            & $EL "          -> Usada como FALLBACK porque ProcessSessionEndedEvent no fue encontrado." $cWarn
        } else {
            & $EL "          -> Señal UUID-based de desconexion (complementaria al evento principal)." $cInfo
        }
        & $EL "          -> NewState=Disconnected es un estado terminal: el HOLD produce NewState=Held," $cInfo
        & $EL "             por lo que esta señal NO puede ser un falso positivo por hold o mute." $cInfo
    }
    if ($HayPhoneSvc) {
        & $EL "  [SEÑAL] PhoneService_CallStateChanged/CallUpdated State=Disconnected detectado" $cOk $true
        & $EL "          -> Señal numerica (ConnectionId) de desconexion en la capa de telefonia." $cInfo
    }
    if (-not $HayOnRequest -and -not $HayProcEnd -and -not $HayCallState -and -not $HayPhoneSvc) {
        & $EL "  [AVISO] No se encontraron señales de desconexion explicitas en el RAW de este slot." $cWarn
        & $EL "          El FIN pudo haberse generado por inferencia del sistema de analisis." $cWarn
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 2: ¿Por qué MANUAL vs NORMAL?
    # ════════════════════════════════════════════════════════════════
    & $EH "PASO 2: Criterio de clasificacion MANUAL vs NORMAL"

    if ($EsEvasion) {
        & $EL "  Clasificacion: EVASION / LINEA ABANDONADA" $cFail $true
        & $EL "  Se cumplio ProcessSessionEndedEvent SIN que se detectara el telefono del cliente." $cInfo
        & $EL "  Ademas, la llamada NO era entrante confirmada." $cInfo
        & $EL "  Interpretacion: la linea se mantuvo abierta sin asociarse a un numero, lo que" $cWarn
        & $EL "  sugiere que el agente nunca registro correctamente la llamada saliente." $cWarn
    } elseif ($EsManual) {
        & $EL "  Clasificacion: FIN MANUAL — el asesor colgó la llamada" $cMan $true
        & $EL "" $cW
        if ($HayOnRequest) {
            & $EL "  RAZON PRINCIPAL: Se encontro OnRequestEndSession() en los logs." $cOk
            & $EL "  Este evento solo aparece cuando el agente hace clic en el boton Colgar de OneX." $cInfo
            & $EL "  El sistema lo registra ANTES de que Avaya procese el cierre de sesion," $cInfo
            & $EL "  y cuando luego llega ProcessSessionEndedEvent, el analizador ve la bandera" $cInfo
            & $EL "  CuelguesManuales activa y clasifica el FIN como MANUAL." $cInfo
        } else {
            & $EL "  [Aviso] OnRequestEndSession no se encontro en el RAW de este slot." $cWarn
            & $EL "  La clasificacion MANUAL proviene del campo Interpretacion del slot." $cWarn
            & $EL "  Es posible que OnRequestEndSession haya caido en otro segundo adyacente." $cWarn
        }
        & $EL "" $cW
        & $EL "  Flujo que llevo a MANUAL:" $cInfo
        & $EL "    1. Asesor presiona Colgar en OneX Agent" $cInfo
        & $EL "    2. OneX emite OnRequestEndSession() → EndpointLog" $cInfo
        & $EL "    3. Avaya CM procesa el Release y emite ProcessSessionEndedEvent → EndpointLog" $cInfo
        & $EL "    4. Analizador verifica bandera de CuelguesManuales → MANUAL confirmado" $cInfo
    } else {
        & $EL "  Clasificacion: FIN NORMAL — el cliente (u otro extremo) colgó" $cNrm $true
        & $EL "" $cW
        if (-not $HayOnRequest) {
            & $EL "  RAZON PRINCIPAL: ProcessSessionEndedEvent llego SIN OnRequestEndSession previo." $cOk
            & $EL "  Cuando el cliente cuelga, Avaya CM cierra la sesion H.323 sin que el agente" $cInfo
            & $EL "  haya emitido una solicitud de cuelgue. El analizador no encuentra la bandera" $cInfo
            & $EL "  CuelguesManuales activa, por lo tanto clasifica el FIN como NORMAL." $cInfo
        } else {
            & $EL "  [Aviso] Se encontro OnRequestEndSession en el mismo segundo, pero la" $cWarn
            & $EL "  clasificacion final fue NORMAL. Puede haber un desfase de milisegundos" $cWarn
            & $EL "  entre el OnRequest y el ProcessSessionEnded que los pone en slots distintos." $cWarn
        }
        & $EL "" $cW
        if ($EsFallback) {
            & $EL "  Nota adicional: Este FIN fue detectado via FALLBACK (Call StateChanged" $cWarn
            & $EL "  OldState=Active,NewState=Disconnected). ProcessSessionEndedEvent no aparecio." $cWarn
            & $EL "  Sin OnRequestEndSession disponible, se asume FIN NORMAL." $cWarn
        } else {
            & $EL "  Flujo que llevo a NORMAL:" $cInfo
            & $EL "    1. Cliente (u otro extremo) cuelga la llamada" $cInfo
            & $EL "    2. Avaya CM emite Release → ProcessSessionEndedEvent en EndpointLog" $cInfo
            & $EL "    3. Analizador verifica CuelguesManuales → vacío → FIN NORMAL" $cInfo
        }
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 3: INICIO correspondiente y duración de la llamada
    # ════════════════════════════════════════════════════════════════
    & $EH "PASO 3: Llamada de INICIO correspondiente y duracion"

    $TodasHoras  = $Snap.Keys | Sort-Object
    $HorasAnte   = $TodasHoras | Where-Object { $_ -lt $FinHora }
    $InicioHoraF = ""; $InicioDescF = ""; $InicioSesF = ""; $InicioTelF = ""
    $InicioEncontrado = $false

    # Buscar hacia atrás el INICIO con misma sesión o teléfono
    foreach ($h in ($HorasAnte | Sort-Object -Descending)) {
        $slt  = $Snap[$h]
        $itp  = $slt.Interpretacion
        if ($itp -notmatch "INICIO DE LLAMADA|LÍNEA ABIERTA") { continue }
        $sSes = $slt.Sesion
        $sTel = ($slt.Tel -replace '^\+','') -replace '^9(\d{10,})$','$1'
        $porS = ($SesId -ne "" -and $sSes -eq $SesId)
        $porT = ($TelNorm -ne "" -and $TelNorm -ne "Desconocido" -and $sTel -eq $TelNorm)
        if ($porS -or $porT) {
            $InicioHoraF = $h; $InicioDescF = $itp; $InicioSesF = $sSes; $InicioTelF = $slt.Tel
            $InicioEncontrado = $true; break
        }
    }

    if ($InicioEncontrado) {
        & $EL "  [OK] INICIO correspondiente encontrado:" $cOk
        & $EL "    Hora INICIO : $InicioHoraF" $cOk
        & $EL "    Tipo INICIO : $InicioDescF" $cInfo
        & $EL "    Sesion      : $InicioSesF" $cInfo
        & $EL "    Telefono    : $InicioTelF" $cInfo
        try {
            $tI  = [datetime]::ParseExact($InicioHoraF, "HH:mm:ss", $null)
            $tF  = [datetime]::ParseExact($FinHora,      "HH:mm:ss", $null)
            $dur = ($tF - $tI).TotalSeconds
            & $EL "" $cW
            & $EL "    Duracion de la llamada: $([int]$dur) segundos ($([math]::Floor($dur/60))m $([int]($dur % 60))s)" $cOk $true
        } catch { & $EL "    Duracion: No calculable (formato de hora inesperado)" $cWarn }
    } else {
        & $EL "  [Aviso] No se encontro un INICIO DE LLAMADA previo con la misma sesion o telefono." $cWarn
        & $EL "  Puede que el INICIO haya sido registrado antes del rango de tiempo analizado," $cWarn
        & $EL "  o que los identificadores de sesion y telefono no coincidan con ningun slot anterior." $cWarn
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 4: RAW completo del slot FIN
    # ════════════════════════════════════════════════════════════════
    & $EH "PASO 4: Lineas RAW del slot FIN (todas las fuentes)"

    $seccionesRaw = [ordered]@{
        "RawInterpretacion (EndpointLog / Fallback)" = $RawInterp
        "RawAgente (OneXAgent — OnRequest, Hold, etc.)" = $RawAgen
        "RawAux (Estado Auxiliar)"                   = $FinSlot.RawAux
        "RawSysLog (SysLog)"                         = $FinSlot.RawSysLog
        "RawAppLog (AppLog / Errores)"               = $FinSlot.RawAppLog
    }
    $HayRaw = $false
    foreach ($sec in $seccionesRaw.Keys) {
        $contenido = $seccionesRaw[$sec]
        if ($contenido -and $contenido.Trim() -ne "") {
            & $EL "  [$sec]" $cHdr
            foreach ($rl in ($contenido -split "`n")) {
                if ($rl.Trim() -ne "") { & $EL "    $rl" $cRaw }
            }
            $HayRaw = $true
        }
    }
    if (-not $HayRaw) {
        & $EL "  [Aviso] No hay lineas RAW registradas para este slot." $cWarn
        & $EL "  El FIN fue generado por inferencia a nivel de renderizado." $cWarn
    }

    & $EL "" $cW
    & $EL "======================================================================" $cHdr
    $rtbEF.SelectionStart = 0; $rtbEF.ScrollToCaret()
    [void]$fEF.ShowDialog($Form)
    $fEF.Dispose()
})

# ====================================================================
# DIAGNÓSTICO: ¿POR QUÉ ES ENTRANTE O SALIENTE?
# ====================================================================
$mnuDiagDir.Add_Click({
    $SelRow = if ($GridResultados.SelectedRows.Count -gt 0) { $GridResultados.SelectedRows[0] } else { $null }
    if ($null -eq $SelRow -or $null -eq $Script:SnapEventosTiempo) { return }

    $InicioHora = $SelRow.Cells["Hora"].Value
    $InicioDesc = $SelRow.Cells["Interpretacion"].Value
    $Snap       = $Script:SnapEventosTiempo
    $InicioSlot = if ($Snap.ContainsKey($InicioHora)) { $Snap[$InicioHora] } else { $null }

    $SesId    = if ($null -ne $InicioSlot -and $InicioSlot.Sesion -ne "-" -and $InicioSlot.Sesion -ne "") { $InicioSlot.Sesion } else { "" }
    $TelRaw   = if ($null -ne $InicioSlot) { $InicioSlot.Tel } else { "" }
    $RawInterp = if ($null -ne $InicioSlot) { $InicioSlot.RawInterpretacion } else { "" }
    $RawAux    = if ($null -ne $InicioSlot) { $InicioSlot.RawAux } else { "" }

    # ── Ventana modal ────────────────────────────────────────────────────
    $fDD = New-Object System.Windows.Forms.Form
    $fDD.Text            = "Diagnostico Direccion — $InicioHora"
    $fDD.Size            = New-Object System.Drawing.Size(980, 750)
    $fDD.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $fDD.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $fDD.MaximizeBox     = $false
    $fDD.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)

    $rtbDD = New-Object System.Windows.Forms.RichTextBox
    $rtbDD.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $rtbDD.ReadOnly   = $true
    $rtbDD.BackColor  = [System.Drawing.Color]::FromArgb(20,20,20)
    $rtbDD.ForeColor  = [System.Drawing.Color]::White
    $rtbDD.Font       = New-Object System.Drawing.Font("Consolas", 9)
    $rtbDD.WordWrap   = $false
    $rtbDD.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
    $fDD.Controls.Add($rtbDD)

    # Colores
    $cOk   = [System.Drawing.Color]::LimeGreen
    $cFail = [System.Drawing.Color]::OrangeRed
    $cWarn = [System.Drawing.Color]::Gold
    $cInfo = [System.Drawing.Color]::Silver
    $cRaw  = [System.Drawing.Color]::FromArgb(150,150,255)
    $cHdr  = [System.Drawing.Color]::CornflowerBlue
    $cW    = [System.Drawing.Color]::White
    $cEnt  = [System.Drawing.Color]::LimeGreen
    $cSal  = [System.Drawing.Color]::DeepSkyBlue

    $DL = {
        $txt  = $args[0]
        $col  = if ($args.Count -gt 1 -and $null -ne $args[1]) { $args[1] } else { $cW }
        $bold = if ($args.Count -gt 2 -and $args[2]) { $true } else { $false }
        $rtbDD.SelectionStart  = $rtbDD.TextLength
        $rtbDD.SelectionLength = 0
        $rtbDD.SelectionColor  = $col
        $rtbDD.SelectionFont   = if ($bold) { New-Object System.Drawing.Font($rtbDD.Font.FontFamily, $rtbDD.Font.Size, [System.Drawing.FontStyle]::Bold) } else { $rtbDD.Font }
        $rtbDD.AppendText("$txt`n")
    }
    $DH = {
        $rtbDD.SelectionStart = $rtbDD.TextLength; $rtbDD.SelectionLength = 0
        $rtbDD.SelectionColor = $cW; $rtbDD.SelectionFont = $rtbDD.Font; $rtbDD.AppendText("`n")
        $rtbDD.SelectionStart = $rtbDD.TextLength; $rtbDD.SelectionLength = 0
        $rtbDD.SelectionColor = $cHdr
        $rtbDD.SelectionFont  = New-Object System.Drawing.Font($rtbDD.Font.FontFamily, $rtbDD.Font.Size, [System.Drawing.FontStyle]::Bold)
        $rtbDD.AppendText("=== $($args[0]) ===`n")
    }

    # Determinar dirección
    $EsEntrante = ($InicioDesc -match "Entrante")
    $EsSaliente = ($InicioDesc -match "Saliente")
    $colorDir   = if ($EsEntrante) { $cEnt } elseif ($EsSaliente) { $cSal } else { $cWarn }

    # Extraer VI UUID del raw (evidencia PRIMARY_CONNECTED)
    $ViUuidDir = ""
    if ($RawInterp -match "\[DIR:(?:ENTRANTE|SALIENTE)\] PRIMARY_CONNECTED: (?:AlertingHoras\[VI:|VI UUID ')([0-9a-fA-F\-]+)") { $ViUuidDir = $matches[1] }

    # Extraer hora alerting de la evidencia
    $AlertingHoraDir = ""
    if ($RawInterp -match "\[DIR:ENTRANTE\] PRIMARY_CONNECTED: AlertingHoras\[VI:[0-9a-fA-F\-]+\] = señal en ([0-9:]+(?:,[0-9]+)?)") { $AlertingHoraDir = $matches[1] }

    # ── Encabezado ───────────────────────────────────────────────────────
    & $DL "DIAGNOSTICO: ¿POR QUE ES ENTRANTE O SALIENTE?" $cW $true
    & $DL "" $cW
    & $DL "  Clasificacion   : $InicioDesc" $colorDir $true
    & $DL "  Hora del INICIO : $InicioHora" $cInfo
    & $DL "  Sesion ID       : $(if ($SesId -ne '') { $SesId } else { '(no detectada)' })" $cInfo
    & $DL "  Telefono        : $(if ($TelRaw -ne '' -and $TelRaw -ne '-') { $TelRaw } else { '(no detectado)' })" $cInfo
    & $DL "" $cW

    # ════════════════════════════════════════════════════════════════
    # PASO 1: ¿Qué mecanismo determinó la dirección?
    # ════════════════════════════════════════════════════════════════
    & $DH "PASO 1: Mecanismo que determino la direccion"

    $TienePrimary   = ($RawInterp -match "\[DIR:(?:ENTRANTE|SALIENTE)\] PRIMARY_CONNECTED")
    $TieneSecondary = ($RawInterp -match "\[DIR:ENTRANTE\] SECONDARY_CONNECTED")
    $TieneGenInc    = ($RawAux    -match "\[DIR:ENTRANTE\] PASO5: GenerateIncomingCall") -or
                      ($RawInterp -match "\[DIR:ENTRANTE\] PASO5: GenerateIncomingCall")
    $TieneWiAdd     = ($RawInterp -match "\[WI\.ADD\].*Entrante.*Alerting")

    if ($TienePrimary) {
        if ($EsEntrante) {
            & $DL "  [OK] PRIMARY CONNECTED (VoiceInteractionImpl type=Active)" $cOk $true
            & $DL "       Este es el mecanismo principal y mas preciso." $cInfo
            & $DL "       Funciona cruzando el VI UUID del estado Active con la tabla AlertingHoras." $cInfo
            & $DL "       Si el UUID estaba en AlertingHoras (registrado al detectar Alerting+Inbound)" $cInfo
            & $DL "       significa que el agente contesto una llamada ENTRANTE." $cInfo
        } else {
            & $DL "  [!] PRIMARY CONNECTED (VoiceInteractionImpl type=Active)" $cWarn $true
            & $DL "       El VI UUID NO fue encontrado en AlertingHoras." $cWarn
            & $DL "       Esto puede ocurrir si el WI.ADD Alerting no fue procesado correctamente." $cWarn
            & $DL "       O si el agente inició la llamada de forma saliente." $cInfo
        }
    } elseif ($TieneSecondary) {
        & $DL "  [OK] SECONDARY CONNECTED (PhoneService_CallStateChanged InnerState=CONNECTED)" $cOk $true
        & $DL "       Mecanismo de respaldo cuando PRIMARY no actua." $cInfo
        & $DL "       Se confirmo con Outgoing=False en el evento de telefonía." $cInfo
    } else {
        & $DL "  [?] Mecanismo no identificado en el RAW de este slot." $cWarn $true
        & $DL "      El INICIO pudo haber sido creado por WI.ADD, XML, o PASO 5 (EndpointLog)." $cInfo
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 2: Evidencia de Alerting (señal previa de llamada)
    # ════════════════════════════════════════════════════════════════
    & $DH "PASO 2: Señal previa de llamada (Alerting)"

    # Buscar slot de señal de llamada en los 15 segundos anteriores
    $AlertingSlotEnc = $null; $AlertingHoraEnc = ""
    $TodasHD = @($Snap.Keys | Sort-Object)
    $HoraBase = ($InicioHora -split ',')[0]
    try {
        $tBase = [datetime]::ParseExact($HoraBase, "HH:mm:ss", $null)
        foreach ($h in ($TodasHD | Sort-Object -Descending)) {
            $hBase = ($h -split ',')[0]
            try {
                $tH = [datetime]::ParseExact($hBase, "HH:mm:ss", $null)
                $diff = ($tBase - $tH).TotalSeconds
                if ($diff -lt 0) { continue }
                if ($diff -gt 30) { break }
                if ($Snap[$h].Interpretacion -match "señal de llamada") {
                    $AlertingSlotEnc = $Snap[$h]; $AlertingHoraEnc = $h; break
                }
            } catch {}
        }
    } catch {}

    if ($AlertingSlotEnc -ne $null) {
        $DiffSeg = 0
        try { $DiffSeg = [int]([datetime]::ParseExact($HoraBase,"HH:mm:ss",$null) - [datetime]::ParseExact(($AlertingHoraEnc -split ',')[0],"HH:mm:ss",$null)).TotalSeconds } catch {}
        & $DL "  [OK] Señal detectada $DiffSeg segundo(s) antes del INICIO:" $cOk
        & $DL "       Hora     : $AlertingHoraEnc" $cInfo
        & $DL "       Tipo     : $($AlertingSlotEnc.Interpretacion)" $cEnt
        if ($AlertingSlotEnc.RawInterpretacion -match "\[WI\.ADD\] (VI\d+:[0-9a-fA-F\-]+).*?Tel:([^|]+).*?Topic:([^|]+)") {
            & $DL "       Fuente   : WI.ADD | VI=$($matches[1]) | Tel=$($matches[2].Trim()) | Topic=$($matches[3].Trim())" $cInfo
        }
        if ($EsEntrante -and $ViUuidDir -ne "") {
            & $DL "       UUID registrado en AlertingHoras: $ViUuidDir" $cOk
            & $DL "       PRIMARY CONNECTED pudo verificar ese UUID → ENTRANTE confirmado" $cOk
        }
    } else {
        if ($EsEntrante) {
            & $DL "  [!] No se encontro señal de llamada en los 30 seg anteriores." $cWarn
            & $DL "      Si la llamada es Entrante, la señal pudo caer fuera del rango de busqueda" $cWarn
            & $DL "      o fue detectada por SECONDARY en lugar de PRIMARY." $cInfo
        } else {
            & $DL "  [OK] No hay señal de Alerting previa (esperado para llamadas salientes)." $cOk
            & $DL "       El agente inicio la llamada — no hay periodo de timbrado entrante." $cInfo
        }
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 3: Evidencia de EndpointLog (GenerateIncomingCall / MakeCall)
    # ════════════════════════════════════════════════════════════════
    & $DH "PASO 3: Evidencia de EndpointLog (PASO 5)"

    if ($TieneGenInc) {
        & $DL "  [OK] GenerateIncomingCall detectado en el slot" $cOk $true
        & $DL "       PASO 5 (EndpointLog) confirmó que Avaya CM generó la sesión como ENTRANTE." $cInfo
        & $DL "       Este evento solo aparece para llamadas que el switch enruta al agente." $cInfo
    } else {
        # Buscar en slots adyacentes (±3 seg)
        $GenIncEnc = $false
        try {
            $tB2 = [datetime]::ParseExact($HoraBase,"HH:mm:ss",$null)
            for ($di = -3; $di -le 3 -and -not $GenIncEnc; $di++) {
                $hTest = $tB2.AddSeconds($di).ToString("HH:mm:ss")
                if ($Snap.ContainsKey($hTest) -and ($Snap[$hTest].RawAux -match "\[DIR:ENTRANTE\] PASO5: GenerateIncomingCall" -or $Snap[$hTest].RawInterpretacion -match "\[DIR:ENTRANTE\] PASO5: GenerateIncomingCall")) {
                    & $DL "  [OK] GenerateIncomingCall detectado en slot adyacente ($hTest, desfase ${di}s)" $cOk
                    & $DL "       PASO 5 confirmo sesion ENTRANTE para este rango de tiempo." $cInfo
                    $GenIncEnc = $true
                }
            }
        } catch {}
        if (-not $GenIncEnc) {
            if ($EsEntrante) {
                & $DL "  [?] GenerateIncomingCall no encontrado en slots adyacentes (±3s)." $cWarn
                & $DL "      Puede estar en otro segundo o no haber quedado registrado en este slot." $cWarn
                & $DL "      La direccion ENTRANTE fue determinada por PASO 4 (OneXAgent.log)." $cInfo
            } else {
                & $DL "  [OK] GenerateIncomingCall no detectado (esperado para salientes)." $cOk
                & $DL "       El agente marcó el número — Avaya no genera este evento." $cInfo
                # Buscar MakeCall en raw
                if ($RawInterp -match "MakeCall\(([^)]+)\)") {
                    & $DL "  [OK] MakeCall($($matches[1])) detectado → llamada saliente iniciada por el agente." $cOk
                }
            }
        }
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 4: Resumen y flujo lógico
    # ════════════════════════════════════════════════════════════════
    & $DH "PASO 4: Resumen y flujo de clasificacion"

    if ($EsEntrante) {
        & $DL "  RESULTADO: LLAMADA ENTRANTE" $cEnt $true
        & $DL "" $cW
        & $DL "  Flujo que llevo a ENTRANTE:" $cInfo
        & $DL "    1. Avaya CM envía la llamada al agente" $cInfo
        & $DL "    2. OneXAgent.log: WI.ADD con state=Alerting + outbound=False" $cInfo
        & $DL "       → El UUID del VI queda registrado en AlertingHoras" $cInfo
        & $DL "    3. El agente contesta (AutoAnswer o manual)" $cInfo
        & $DL "    4. OneXAgent.log: VoiceInteractionImpl[VI:UUID].StateImpl=[type=Active]" $cInfo
        & $DL "       → El analizador busca el UUID en AlertingHoras → encontrado → ENTRANTE" $cInfo
        if ($TieneGenInc) {
            & $DL "    5. EndpointLog: GenerateIncomingCall → DirLlamada[SesionID]=ENTRANTE (confirmacion)" $cInfo
        }
        if ($ViUuidDir -ne "") {
            & $DL "" $cW
            & $DL "  VI UUID involucrado : $ViUuidDir" $cOk
        }
        if ($AlertingHoraDir -ne "") {
            & $DL "  Señal detectada en  : $AlertingHoraDir" $cOk
        }
    } else {
        & $DL "  RESULTADO: LLAMADA SALIENTE" $cSal $true
        & $DL "" $cW
        & $DL "  Flujo que llevo a SALIENTE:" $cInfo
        & $DL "    1. El agente marcó el número (MakeCall)" $cInfo
        & $DL "    2. OneXAgent.log: VoiceInteractionImpl[VI:UUID].StateImpl=[type=Active]" $cInfo
        if ($TienePrimary) {
            & $DL "       → El UUID del VI NO fue encontrado en AlertingHoras" $cWarn
            & $DL "       → No hubo Alerting+Inbound previo → se asume SALIENTE" $cInfo
        }
        & $DL "    3. EndpointLog: MakeCall vinculado por numero y proximidad temporal" $cInfo
        if ($ViUuidDir -ne "") {
            & $DL "" $cW
            & $DL "  VI UUID involucrado : $ViUuidDir" $cSal
        }
    }

    # ════════════════════════════════════════════════════════════════
    # PASO 5: Datos RAW del slot INICIO
    # ════════════════════════════════════════════════════════════════
    & $DH "PASO 5: Lineas RAW del slot INICIO (todas las fuentes)"

    $secRaw = [ordered]@{
        "RawInterpretacion (OneXAgent / EndpointLog)" = $RawInterp
        "RawAgente (OnRequest, Hold, etc.)"           = $InicioSlot.RawAgente
        "RawAux (PASO 5 / GenerateIncomingCall)"      = $RawAux
    }
    $HayRawD = $false
    foreach ($sec in $secRaw.Keys) {
        $cont = $secRaw[$sec]
        if ($cont -and $cont.Trim() -ne "") {
            & $DL "  [$sec]" $cHdr
            foreach ($rl in ($cont -split "`n")) { if ($rl.Trim() -ne "") { & $DL "    $rl" $cRaw } }
            $HayRawD = $true
        }
    }
    if (-not $HayRawD) {
        & $DL "  [Aviso] No hay lineas RAW en este slot. El INICIO fue generado por inferencia." $cWarn
    }

    & $DL "" $cW
    & $DL "======================================================================" $cHdr
    $rtbDD.SelectionStart = 0; $rtbDD.ScrollToCaret()
    [void]$fDD.ShowDialog($Form)
    $fDD.Dispose()
})

# ====================================================================
# EXPORTAR GRID A CSV
# ====================================================================
$btnExportarCSV.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Archivo CSV (*.csv)|*.csv"
    $sfd.Title  = "Exportar resultados a CSV"
    $NombreBase = if ($dtpFecha.Value) { "Auditoria_$($dtpFecha.Value.ToString('yyyy-MM-dd'))" } else { "Auditoria" }
    $sfd.FileName = $NombreBase
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $sb = New-Object System.Text.StringBuilder

        # Encabezados (solo columnas visibles)
        $colsVisibles = @($GridResultados.Columns | Where-Object { $_.Visible })
        $headers = $colsVisibles | ForEach-Object { '"' + $_.HeaderText.Replace('"','""') + '"' }
        [void]$sb.AppendLine($headers -join ",")

        # Filas
        foreach ($row in $GridResultados.Rows) {
            if ($row.IsNewRow) { continue }
            $values = $colsVisibles | ForEach-Object {
                $v = if ($null -ne $row.Cells[$_.Name].Value) { $row.Cells[$_.Name].Value.ToString() } else { "" }
                '"' + $v.Replace('"','""') + '"'
            }
            [void]$sb.AppendLine($values -join ",")
        }

        $enc = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($sfd.FileName, $sb.ToString(), $enc)
        [System.Windows.Forms.MessageBox]::Show("CSV exportado correctamente:`n$($sfd.FileName)", "Exportacion exitosa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al exportar: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ====================================================================
# MÓDULO 6: CLICK EN CELDA
# ====================================================================
$GridResultados.Add_CellClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $ColName = $GridResultados.Columns[$e.ColumnIndex].Name
        $ColumnasRegulares = @("Interpretacion","EvAgente","EvAudio","EvAux","EvSysLog","EvAppLog","EvIspeac")

        if ($ColName -eq "Telefono" -and $GridResultados.Rows[$e.RowIndex].Cells["Interpretacion"].Value -match "INICIO DE LLAMADA \(Entrante\)") {
            $TelClick  = $GridResultados.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
            $HoraClick = $GridResultados.Rows[$e.RowIndex].Cells["Hora"].Value
            $LogCrudo  = $GridResultados.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Tag

            $Mensaje  = if ($LogCrudo) { "LOG DE EXTRACCIÓN ORIGINAL:`n$LogCrudo`n`n" } else { "" }
            $Mensaje += "--- Validación AutoAnswer ---`n"
            $Encontrado = $false

            if ($Script:DirFinalGlobal -and $TelClick -ne "-" -and $TelClick -ne "Desconocido") {
                $XMLPath = Get-ChildItem -Path $Script:DirFinalGlobal -Filter "ContactLog.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
                if (-not $XMLPath) {
                    $ParentDir = Split-Path $Script:DirFinalGlobal -Parent
                    $XMLPath = Get-ChildItem -Path $ParentDir -Filter "ContactLog.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
                }
                if ($XMLPath -and (Test-Path $XMLPath)) {
                    try {
                        $xml = [xml](Get-Content $XMLPath -Encoding UTF8)
                        foreach ($log in $xml.ContactLogs.ContactLog) {
                            $Items = $log.ContactLogItem
                            if ($Items -eq $null) { continue }
                            if ($Items -isnot [System.Array]) { $Items = @($Items) }
                            foreach ($item in $Items) {
                                if ($item -eq $null -or $item.Type -ne "Voice") { continue }
                                $TelLimpio = $TelClick -replace '\D',''
                                $UriLimpia = $item.Uri -replace '\D',''
                                if ($UriLimpia -and $TelLimpio -and ($UriLimpia -match $TelLimpio -or $TelLimpio -match $UriLimpia)) {
                                    $HoraLogObj = [datetimeoffset]::Parse($log.CreateTime).LocalDateTime
                                    $T1 = [datetime]::ParseExact($HoraLogObj.ToString("HH:mm:ss"), "HH:mm:ss", $null)
                                    $T2 = [datetime]::ParseExact($HoraClick, "HH:mm:ss", $null)
                                    if ([math]::Abs(($T1 - $T2).TotalSeconds) -le 120) {
                                        $VI_ID = $item.Id
                                        $Modo = if ($Script:ModoContestacion.ContainsKey($VI_ID)) { $Script:ModoContestacion[$VI_ID] } else { "No se detectó Auto Accepting ni AnswerVoiceInteraction (Posible llamada interna o de sistema)" }
                                        $Mensaje += "ID Interno (ContactLog): $VI_ID`n"
                                        $Mensaje += "Método de Contestación: $Modo`n"
                                        $Encontrado = $true; break
                                    }
                                }
                            }
                            if ($Encontrado) { break }
                        }
                    } catch {}
                }
            }
            if (-not $Encontrado) { $Mensaje += "No se pudo cruzar el ID en el ContactLog.xml para esta hora exacta." }
            [System.Windows.Forms.MessageBox]::Show($Mensaje, "Auditoría de Llamada", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        elseif ($ColumnasRegulares -contains $ColName) {
            $LogCrudo = $GridResultados.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Tag
            if ($LogCrudo) { [System.Windows.Forms.MessageBox]::Show("EVIDENCIA DEL LOG ORIGINAL:`n`n$LogCrudo", "Análisis de Logs Avaya OneX Agent - Detalle", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
        }
        elseif ($ColName -eq "Sesion") {
            $SesionID = $GridResultados.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
            if ($SesionID -ne "-" -and $SesionID -ne "") {
                $SubForm = New-Object System.Windows.Forms.Form
                $SubForm.Text = "Analizador Avaya - Aislamiento de Llamada (Sesión ID: $SesionID)"
                $SubForm.Size = New-Object System.Drawing.Size(1300, 450); $SubForm.StartPosition = "CenterParent"; $SubForm.BackColor = $ColorFondo; $SubForm.ForeColor = $ColorTexto
                $SubGrid = New-Object System.Windows.Forms.DataGridView
                $SubGrid.Dock = "Fill"; $SubGrid.BackgroundColor = $ColorPanel; $SubGrid.AllowUserToAddRows = $false; $SubGrid.RowHeadersVisible = $false; $SubGrid.ReadOnly = $true; $SubGrid.AutoSizeColumnsMode = "Fill"; $SubGrid.DefaultCellStyle.BackColor = $ColorPanel; $SubGrid.DefaultCellStyle.ForeColor = $ColorTexto; $SubGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(28,28,28); $SubGrid.ColumnHeadersDefaultCellStyle.ForeColor = $ColorTexto; $SubGrid.EnableHeadersVisualStyles = $false
                foreach ($col in $GridResultados.Columns) { $SubGrid.Columns.Add($col.Name, $col.HeaderText) | Out-Null; $SubGrid.Columns[$col.Name].FillWeight = $col.FillWeight; $SubGrid.Columns[$col.Name].Visible = $col.Visible }
                foreach ($row in $GridResultados.Rows) { if ($row.Cells["Sesion"].Value -eq $SesionID) { $newIdx = $SubGrid.Rows.Add(); foreach ($col in $GridResultados.Columns) { $SubGrid.Rows[$newIdx].Cells[$col.Name].Value = $row.Cells[$col.Name].Value; $SubGrid.Rows[$newIdx].Cells[$col.Name].Style.ForeColor = $row.Cells[$col.Name].Style.ForeColor } } }
                $SubForm.Controls.Add($SubGrid); $SubForm.ShowDialog() | Out-Null
            }
        }
    }
})

# ====================================================================
# MÓDULO 7: EXTRACCIÓN RAW (al milisegundo)
# ====================================================================
$btnExtraccion.Add_Click({
    if ($Script:DirFinalGlobal -eq "") { [System.Windows.Forms.MessageBox]::Show("Primero debes realizar una auditoría (Botón 2).", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }

    $FormTiempo = New-Object System.Windows.Forms.Form; $FormTiempo.Text = "Configurar Aislamiento RAW"; $FormTiempo.Size = New-Object System.Drawing.Size(350, 200); $FormTiempo.StartPosition = "CenterParent"; $FormTiempo.BackColor = $ColorPanel; $FormTiempo.ForeColor = $ColorTexto
    $lblT1 = New-Object System.Windows.Forms.Label; $lblT1.Text = "Hora Inicio (HH:mm:ss):"; $lblT1.Location = New-Object System.Drawing.Point(20, 20); $lblT1.AutoSize = $true
    $txtT1 = New-Object System.Windows.Forms.TextBox; $txtT1.Location = New-Object System.Drawing.Point(160, 18); $txtT1.Text = "00:00:00"; $txtT1.BackColor = $ColorFondo; $txtT1.ForeColor = [System.Drawing.Color]::LimeGreen
    $lblT2 = New-Object System.Windows.Forms.Label; $lblT2.Text = "Hora Fin (HH:mm:ss):"; $lblT2.Location = New-Object System.Drawing.Point(20, 60); $lblT2.AutoSize = $true
    $txtT2 = New-Object System.Windows.Forms.TextBox; $txtT2.Location = New-Object System.Drawing.Point(160, 58); $txtT2.Text = "23:59:59"; $txtT2.BackColor = $ColorFondo; $txtT2.ForeColor = [System.Drawing.Color]::LimeGreen
    $btnGenerar = New-Object System.Windows.Forms.Button; $btnGenerar.Text = "EXTRAER LOGS RAW"; $btnGenerar.Location = New-Object System.Drawing.Point(100, 100); $btnGenerar.BackColor = [System.Drawing.Color]::Maroon; $btnGenerar.FlatStyle = "Flat"; $btnGenerar.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $FormTiempo.Controls.AddRange(@($lblT1, $txtT1, $lblT2, $txtT2, $btnGenerar))
    if ($FormTiempo.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $HoraInicioStr = $txtT1.Text.Trim(); $HoraFinStr = $txtT2.Text.Trim()
    if ($HoraInicioStr -notmatch "^\d{2}:\d{2}:\d{2}$" -or $HoraFinStr -notmatch "^\d{2}:\d{2}:\d{2}$") { [System.Windows.Forms.MessageBox]::Show("Formato de hora inválido. Usa HH:mm:ss", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); return }

    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $lblStatus.Text = "Módulo RAW: Entrelazando líneas al milisegundo..."; $Form.Refresh()

    if ($Script:RutaManual -eq "") {
        $TargetIP = $txtIP.Text.Trim()
        if (-not $TargetIP -or -not $Script:Creds) { [System.Windows.Forms.MessageBox]::Show("Se perdió la sesión. Vuelve a ejecutar la auditoría.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); $Form.Cursor = [System.Windows.Forms.Cursors]::Default; return }
        if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null }
        try { New-PSDrive -Name $Script:DriveName -PSProvider FileSystem -Root "\\$TargetIP\c$" -Credential $Script:Creds -EA Stop | Out-Null }
        catch { [System.Windows.Forms.MessageBox]::Show("No se pudo reconectar al equipo.", "Error de Red", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); $Form.Cursor = [System.Windows.Forms.Cursors]::Default; return }
    }

    try {
        $RawLogsExtraidos = [System.Collections.ArrayList]::new()
        $F1 = $dtpFecha.Value.ToString("MM/dd/yyyy"); $F2 = $dtpFecha.Value.ToString("dd/MM/yyyy"); $F3 = $dtpFecha.Value.ToString("yyyy-MM-dd"); $F4 = $dtpFecha.Value.ToString("M/d/yyyy"); $F5 = $dtpFecha.Value.ToString("d/M/yyyy")
        $FechaOmni = "(?:$([regex]::Escape($F1))|$([regex]::Escape($F2))|$([regex]::Escape($F3))|$([regex]::Escape($F4))|$([regex]::Escape($F5)))"
        $Script:LineaID = 0

        function Extraer-Lineas {
            param([string]$Filtro, [string]$NombreColumna, [switch]$EsOneX)
            $Archivos = $null
            if ($EsOneX) { $Archivos = Get-ChildItem -Path $Script:DirFinalGlobal -EA SilentlyContinue | Where-Object { (-not $_.PSIsContainer) -and ($_.Name -match "(?i)one-?x.*\.log" -or $_.Name -match "(?i)one-?x.*\.txt") } | Sort-Object LastWriteTime }
            else { $Archivos = Get-ChildItem -Path $Script:DirFinalGlobal -Filter $Filtro -EA SilentlyContinue | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime }
            if (-not $Archivos) { return }
            foreach ($Archivo in $Archivos) {
                $Lineas = Get-Content -Path $Archivo.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                if (-not $Lineas) { continue }
                $UltimaHoraVista = ""
                foreach ($linea in $Lineas) {
                    if ($linea -match "^\[?(?:$FechaOmni).*?(\d{2}:\d{2}:\d{2})(?:[.,:](\d{1,3}))?") { $ms = if ($matches[2]) { $matches[2].PadRight(3,'0') } else { "000" }; $UltimaHoraVista = "$($matches[1]).$ms" }
                    elseif ($linea -match "^\[\d{1,2}/\d{1,2}/\d{4}\s+(\d{2}:\d{2}:\d{2})(?:[.,:](\d{1,3}))?\]") { $ms = if ($matches[2]) { $matches[2].PadRight(3,'0') } else { "000" }; $UltimaHoraVista = "$($matches[1]).$ms" }
                    if ($UltimaHoraVista -ne "") {
                        $HoraBase = $UltimaHoraVista.Substring(0,8)
                        if ($HoraBase -ge $HoraInicioStr -and $HoraBase -le $HoraFinStr) { [void]$RawLogsExtraidos.Add([PSCustomObject]@{ HoraStr=$UltimaHoraVista; Tipo=$NombreColumna; Texto=$linea; ID=$Script:LineaID++ }) }
                    }
                }
            }
        }

        Extraer-Lineas -Filtro "EndpointLog.txt*" -NombreColumna "Endpoint"
        Extraer-Lineas -NombreColumna "OneX" -EsOneX
        Extraer-Lineas -Filtro "AudioLog.txt*" -NombreColumna "Audio"
        Extraer-Lineas -Filtro "IspeacLog.txt*" -NombreColumna "Ispeac"
        $LogsOrdenados = $RawLogsExtraidos | Sort-Object HoraStr, ID

        $FormRaw = New-Object System.Windows.Forms.Form; $FormRaw.Text = "Correlación RAW ($HoraInicioStr — $HoraFinStr)"; $FormRaw.Size = New-Object System.Drawing.Size(1600, 800); $FormRaw.StartPosition = "CenterParent"; $FormRaw.BackColor = $ColorFondo
        $GridRaw = New-Object System.Windows.Forms.DataGridView; $GridRaw.Size = New-Object System.Drawing.Size(1550, 680); $GridRaw.Location = New-Object System.Drawing.Point(15, 15)
        $GridRaw.BackgroundColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridRaw.AllowUserToAddRows = $false; $GridRaw.RowHeadersVisible = $false; $GridRaw.ReadOnly = $true; $GridRaw.AutoSizeColumnsMode = "Fill"
        $GridRaw.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridRaw.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(10,10,10); $GridRaw.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White; $GridRaw.EnableHeadersVisualStyles = $false
        $GridRaw.Columns.Add("Hora","Hora Rastreada") | Out-Null; $GridRaw.Columns["Hora"].FillWeight = 8
        $GridRaw.Columns.Add("Endpoint","Endpoint.log (Cyan)") | Out-Null; $GridRaw.Columns["Endpoint"].FillWeight = 23
        $GridRaw.Columns.Add("OneX","AvayaOneX.log (Orquídea)") | Out-Null; $GridRaw.Columns["OneX"].FillWeight = 23
        $GridRaw.Columns.Add("Audio","Audio.log (Verde)") | Out-Null; $GridRaw.Columns["Audio"].FillWeight = 23
        $GridRaw.Columns.Add("Ispeac","IspeacLog.txt (Naranja)") | Out-Null; $GridRaw.Columns["Ispeac"].FillWeight = 23
        $GridRaw.SuspendLayout()
        foreach ($item in $LogsOrdenados) {
            $r = $GridRaw.Rows.Add(); $GridRaw.Rows[$r].Cells["Hora"].Value = $item.HoraStr; $GridRaw.Rows[$r].Cells["Hora"].Style.ForeColor = [System.Drawing.Color]::White
            $GridRaw.Rows[$r].Cells[$item.Tipo].Value = $item.Texto
            switch ($item.Tipo) { "Endpoint"{$GridRaw.Rows[$r].Cells["Endpoint"].Style.ForeColor=[System.Drawing.Color]::Cyan} "OneX"{$GridRaw.Rows[$r].Cells["OneX"].Style.ForeColor=[System.Drawing.Color]::MediumOrchid} "Audio"{$GridRaw.Rows[$r].Cells["Audio"].Style.ForeColor=[System.Drawing.Color]::LimeGreen} "Ispeac"{$GridRaw.Rows[$r].Cells["Ispeac"].Style.ForeColor=[System.Drawing.Color]::Orange} }
        }
        $GridRaw.ResumeLayout()

        $btnExportarCSV = New-Object System.Windows.Forms.Button; $btnExportarCSV.Text = "EXPORTAR RESULTADOS A CSV"; $btnExportarCSV.Location = New-Object System.Drawing.Point(15, 710); $btnExportarCSV.Size = New-Object System.Drawing.Size(250, 35); $btnExportarCSV.BackColor = [System.Drawing.Color]::DarkGreen; $btnExportarCSV.ForeColor = [System.Drawing.Color]::White; $btnExportarCSV.FlatStyle = "Flat"; $btnExportarCSV.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $btnExportarCSV.Add_Click({ $sd = New-Object System.Windows.Forms.SaveFileDialog; $sd.Filter = "Archivo CSV (*.csv)|*.csv"; $sd.FileName = "Extraccion_RAW_Avaya_$($HoraInicioStr.Replace(':',''))_a_$($HoraFinStr.Replace(':','')).csv"; if ($sd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $CsvData=@(); foreach($row in $GridRaw.Rows){$CsvData+=[PSCustomObject]@{Hora=$row.Cells["Hora"].Value;EndpointLog=$row.Cells["Endpoint"].Value;AvayaOneXLog=$row.Cells["OneX"].Value;AudioLog=$row.Cells["Audio"].Value;IspeacLog=$row.Cells["Ispeac"].Value}}; $CsvData|Export-Csv -Path $sd.FileName -NoTypeInformation -Encoding UTF8; [System.Windows.Forms.MessageBox]::Show("¡Exportación exitosa!","Éxito",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) } })

        $lblBuscarRAW = New-Object System.Windows.Forms.Label; $lblBuscarRAW.Text = "Buscar en Logs:"; $lblBuscarRAW.Location = New-Object System.Drawing.Point(300, 718); $lblBuscarRAW.AutoSize = $true; $lblBuscarRAW.ForeColor = [System.Drawing.Color]::White
        $txtBuscarRAW = New-Object System.Windows.Forms.TextBox; $txtBuscarRAW.Location = New-Object System.Drawing.Point(400, 715); $txtBuscarRAW.Size = New-Object System.Drawing.Size(200, 25); $txtBuscarRAW.BackColor = [System.Drawing.Color]::FromArgb(45,45,48); $txtBuscarRAW.ForeColor = [System.Drawing.Color]::White
        $btnBuscarRAW = New-Object System.Windows.Forms.Button; $btnBuscarRAW.Text = "FILTRAR"; $btnBuscarRAW.Location = New-Object System.Drawing.Point(610, 713); $btnBuscarRAW.Size = New-Object System.Drawing.Size(100, 30); $btnBuscarRAW.BackColor = [System.Drawing.Color]::Teal; $btnBuscarRAW.ForeColor = [System.Drawing.Color]::White; $btnBuscarRAW.FlatStyle = "Flat"
        $btnLimpiarRAW = New-Object System.Windows.Forms.Button; $btnLimpiarRAW.Text = "LIMPIAR"; $btnLimpiarRAW.Location = New-Object System.Drawing.Point(720, 713); $btnLimpiarRAW.Size = New-Object System.Drawing.Size(100, 30); $btnLimpiarRAW.BackColor = [System.Drawing.Color]::Gray; $btnLimpiarRAW.ForeColor = [System.Drawing.Color]::White; $btnLimpiarRAW.FlatStyle = "Flat"
        $btnBuscarRAW.Add_Click({ $termino=$txtBuscarRAW.Text.Trim(); if($termino -eq ""){return}; $FormRaw.Cursor=[System.Windows.Forms.Cursors]::WaitCursor; $GridRaw.SuspendLayout(); $GridRaw.CurrentCell=$null; foreach($row in $GridRaw.Rows){$match=$false;foreach($cell in $row.Cells){if($cell.Value -and $cell.Value.ToString() -match [regex]::Escape($termino)){$match=$true;break}};$row.Visible=$match}; $GridRaw.ResumeLayout(); $FormRaw.Cursor=[System.Windows.Forms.Cursors]::Default })
        $btnLimpiarRAW.Add_Click({ $txtBuscarRAW.Text=""; $FormRaw.Cursor=[System.Windows.Forms.Cursors]::WaitCursor; $GridRaw.SuspendLayout(); $GridRaw.CurrentCell=$null; foreach($row in $GridRaw.Rows){$row.Visible=$true}; $GridRaw.ResumeLayout(); $FormRaw.Cursor=[System.Windows.Forms.Cursors]::Default })

        $FormRaw.Controls.AddRange(@($GridRaw,$btnExportarCSV,$lblBuscarRAW,$txtBuscarRAW,$btnBuscarRAW,$btnLimpiarRAW))
        $lblStatus.Text = "Aislamiento completado. Visualizando $($LogsOrdenados.Count) líneas crudas entrelazadas."
        if ($LogsOrdenados.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No se encontró actividad en ese rango.", "Grid Vacío", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) }
        else { $FormRaw.ShowDialog() | Out-Null }
    } catch { $Row=$GridResultados.Rows.Add(); $GridResultados.Rows[$Row].Cells["EvAgente"].Value="Error al leer logs: $($_.Exception.Message)"; $GridResultados.Rows[$Row].DefaultCellStyle.ForeColor=[System.Drawing.Color]::Red }
    finally { if($Script:RutaManual -eq ""){if(Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue){Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue|Out-Null}}; $Form.Cursor=[System.Windows.Forms.Cursors]::Default }
})

# ====================================================================
# MÓDULO 8: BÚSQUEDA RÁPIDA
# ====================================================================
$btnBusqueda.Add_Click({
    if ($Script:DirFinalGlobal -eq "") { [System.Windows.Forms.MessageBox]::Show("Primero debes realizar una auditoría (Botón 2).", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }

    $FormB = New-Object System.Windows.Forms.Form; $FormB.Text = "Búsqueda Rápida de Eventos"; $FormB.Size = New-Object System.Drawing.Size(1200, 600); $FormB.StartPosition = "CenterParent"; $FormB.BackColor = $ColorPanel; $FormB.ForeColor = $ColorTexto
    $lblP = New-Object System.Windows.Forms.Label; $lblP.Text = "Palabras a buscar (separadas por coma):"; $lblP.Location = New-Object System.Drawing.Point(20,20); $lblP.AutoSize = $true
    $txtP = New-Object System.Windows.Forms.TextBox; $txtP.Location = New-Object System.Drawing.Point(260,18); $txtP.Size = New-Object System.Drawing.Size(300,25); $txtP.BackColor = $ColorFondo; $txtP.ForeColor = [System.Drawing.Color]::LimeGreen
    $lblF = New-Object System.Windows.Forms.Label; $lblF.Text = "Fecha a escanear:"; $lblF.Location = New-Object System.Drawing.Point(580,20); $lblF.AutoSize = $true
    $dtpF = New-Object System.Windows.Forms.DateTimePicker; $dtpF.Location = New-Object System.Drawing.Point(690,18); $dtpF.Size = New-Object System.Drawing.Size(110,25); $dtpF.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom; $dtpF.CustomFormat = "dd/MM/yyyy"; $dtpF.Value = $dtpFecha.Value
    $btnEjecutar = New-Object System.Windows.Forms.Button; $btnEjecutar.Text = "BUSCAR"; $btnEjecutar.Location = New-Object System.Drawing.Point(820,16); $btnEjecutar.Size = New-Object System.Drawing.Size(100,28); $btnEjecutar.BackColor = [System.Drawing.Color]::Teal; $btnEjecutar.FlatStyle = "Flat"; $btnEjecutar.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $lblEstatusB = New-Object System.Windows.Forms.Label; $lblEstatusB.Location = New-Object System.Drawing.Point(20,50); $lblEstatusB.Size = New-Object System.Drawing.Size(1140,20); $lblEstatusB.ForeColor = [System.Drawing.Color]::Yellow; $lblEstatusB.Text = "Ejemplo: LogoutRequest, AgentState, closeSignalingChannel"
    $GridB = New-Object System.Windows.Forms.DataGridView; $GridB.Location = New-Object System.Drawing.Point(20,75); $GridB.Size = New-Object System.Drawing.Size(1140,470)
    $GridB.BackgroundColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridB.AllowUserToAddRows = $false; $GridB.RowHeadersVisible = $false; $GridB.ReadOnly = $true; $GridB.AutoSizeColumnsMode = "Fill"
    $GridB.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridB.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(10,10,10); $GridB.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White; $GridB.EnableHeadersVisualStyles = $false
    $GridB.Columns.Add("Hora","Hora") | Out-Null; $GridB.Columns["Hora"].FillWeight = 10
    $GridB.Columns.Add("Archivo","Archivo Origen") | Out-Null; $GridB.Columns["Archivo"].FillWeight = 15
    $GridB.Columns.Add("Log","Línea de Log Encontrada") | Out-Null; $GridB.Columns["Log"].FillWeight = 75

    $btnEjecutar.Add_Click({
        $terminosInput = $txtP.Text.Trim()
        if ($terminosInput -eq "") { [System.Windows.Forms.MessageBox]::Show("Ingresa al menos una palabra.","Aviso",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning); return }
        $FormB.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $lblEstatusB.Text = "Buscando..."; $FormB.Refresh(); $GridB.Rows.Clear()
        if ($Script:RutaManual -eq "") {
            $TargetIP = $txtIP.Text.Trim()
            if (-not $TargetIP -or -not $Script:Creds) { [System.Windows.Forms.MessageBox]::Show("Se perdió la sesión.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error); $FormB.Cursor=[System.Windows.Forms.Cursors]::Default; return }
            if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null }
            try { New-PSDrive -Name $Script:DriveName -PSProvider FileSystem -Root "\\$TargetIP\c$" -Credential $Script:Creds -EA Stop | Out-Null }
            catch { [System.Windows.Forms.MessageBox]::Show("No se pudo reconectar.","Error de Red",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error); $FormB.Cursor=[System.Windows.Forms.Cursors]::Default; return }
        }
        $TerminosLimpios = $terminosInput -split "," | ForEach-Object { [regex]::Escape($_.Trim()) } | Where-Object { $_ -ne "" }
        $RegexBusqueda = "(?i)($($TerminosLimpios -join '|'))"

        function Buscar-En-Archivos { param([string]$Filtro,[string]$NombreFuente,[switch]$EsOneX)
            $Archivos = $null
            if ($EsOneX) { $Archivos = Get-ChildItem -Path $Script:DirFinalGlobal -EA SilentlyContinue | Where-Object{(-not $_.PSIsContainer)-and($_.Name -match "(?i)one-?x.*\.log"-or $_.Name -match "(?i)one-?x.*\.txt")} }
            else { $Archivos = Get-ChildItem -Path $Script:DirFinalGlobal -Filter $Filtro -EA SilentlyContinue | Where-Object{-not $_.PSIsContainer} }
            if (-not $Archivos) { return @() }
            $ResultadosLocales = [System.Collections.ArrayList]::new()
            foreach ($Archivo in $Archivos) {
                $Lineas = Get-Content -Path $Archivo.FullName -Encoding UTF8 -ReadCount 1000 -EA SilentlyContinue
                if (-not $Lineas) { continue }
                $UltimaHoraVista = "---"
                foreach ($bloque in $Lineas) { foreach ($linea in $bloque) {
                    if ($linea -match "(?:^|\[|\s)(\d{2}:\d{2}:\d{2})(?:[.,:]\d{1,3})?") { $UltimaHoraVista = $matches[1] }
                    if ($linea -match $RegexBusqueda) { [void]$ResultadosLocales.Add([PSCustomObject]@{Hora=$UltimaHoraVista;Archivo=$NombreFuente;Texto=$linea}) }
                }}
            }
            return $ResultadosLocales
        }
        try {
            $ResultadosRAW = [System.Collections.ArrayList]::new()
            $R1=Buscar-En-Archivos -Filtro "EndpointLog.txt*" -NombreFuente "Endpoint.log"; if($R1){$ResultadosRAW.AddRange($R1)}
            $R2=Buscar-En-Archivos -NombreFuente "AvayaOneX.log" -EsOneX; if($R2){$ResultadosRAW.AddRange($R2)}
            $R3=Buscar-En-Archivos -Filtro "AudioLog.txt*" -NombreFuente "Audio.log"; if($R3){$ResultadosRAW.AddRange($R3)}
            $R4=Buscar-En-Archivos -Filtro "IspeacLog.txt*" -NombreFuente "Ispeac.log"; if($R4){$ResultadosRAW.AddRange($R4)}
            $ResultadosOrdenados = $ResultadosRAW | Sort-Object Hora
            $GridB.SuspendLayout()
            foreach ($item in $ResultadosOrdenados) {
                $r=$GridB.Rows.Add(); $GridB.Rows[$r].Cells["Hora"].Value=$item.Hora; $GridB.Rows[$r].Cells["Archivo"].Value=$item.Archivo; $GridB.Rows[$r].Cells["Log"].Value=$item.Texto
                switch ($item.Archivo) {"Endpoint.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::Cyan} "AvayaOneX.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::MediumOrchid} "Audio.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::LimeGreen} "Ispeac.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::Orange}}
            }
            $GridB.ResumeLayout()
            $lblEstatusB.Text = "Búsqueda completada. $($ResultadosOrdenados.Count) coincidencias."
        } catch { $lblEstatusB.Text = "Error: $($_.Exception.Message)" }
        finally { if($Script:RutaManual -eq ""){if(Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue){Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue|Out-Null}}; $FormB.Cursor=[System.Windows.Forms.Cursors]::Default }
    })

    $FormB.Controls.AddRange(@($lblP,$txtP,$lblF,$dtpF,$btnEjecutar,$lblEstatusB,$GridB))
    $FormB.ShowDialog() | Out-Null
})

# ====================================================================
# MÓDULO 9: HISTORIAL DE REINICIOS (BOTÓN OCULTO)
# ====================================================================
$btnReinicios.Add_Click({
    $TargetIP = $txtIP.Text.Trim()
    if (-not $TargetIP) { return }
    if (-not $Script:Creds) { try { $Script:Creds = Get-Credential -UserName "local\soporte" -Message "Credenciales WMI para $TargetIP" -EA Stop } catch { return } }
    $FormDias = New-Object System.Windows.Forms.Form; $FormDias.Text = "Días a consultar"; $FormDias.Size = New-Object System.Drawing.Size(300,150); $FormDias.StartPosition = "CenterParent"; $FormDias.BackColor = $ColorPanel; $FormDias.ForeColor = $ColorTexto
    $lblDias = New-Object System.Windows.Forms.Label; $lblDias.Text = "¿Cuántos días hacia atrás deseas revisar?"; $lblDias.Location = New-Object System.Drawing.Point(20,20); $lblDias.AutoSize = $true
    $txtDias = New-Object System.Windows.Forms.TextBox; $txtDias.Location = New-Object System.Drawing.Point(100,45); $txtDias.Size = New-Object System.Drawing.Size(50,25); $txtDias.Text = "15"; $txtDias.BackColor = $ColorFondo; $txtDias.ForeColor = [System.Drawing.Color]::LimeGreen
    $btnAceptarDias = New-Object System.Windows.Forms.Button; $btnAceptarDias.Text = "CONSULTAR"; $btnAceptarDias.Location = New-Object System.Drawing.Point(75,80); $btnAceptarDias.DialogResult = [System.Windows.Forms.DialogResult]::OK; $btnAceptarDias.BackColor = [System.Drawing.Color]::Teal; $btnAceptarDias.FlatStyle = "Flat"
    $FormDias.Controls.AddRange(@($lblDias,$txtDias,$btnAceptarDias))
    if ($FormDias.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    [int]$DiasConsulta = if ($txtDias.Text -match "^\d+$") { $txtDias.Text } else { 15 }
    $lblStatus.Text = "Consultando historial de reinicios en $TargetIP..."; $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $Form.Refresh()

    try {
        $NetCred = $Script:Creds.GetNetworkCredential()
        $Usr = if ($NetCred.Domain) { "$($NetCred.Domain)\$($NetCred.UserName)" } else { $NetCred.UserName }
        $Pwd = $NetCred.Password
        net use \\$TargetIP\C$ $Pwd /user:$Usr 2>&1 | Out-Null
        $RutaPublicaSMB = "\\$TargetIP\C$\Users\Public"

$ScriptPayload = @"
& {
    try {
        `$Dias = $DiasConsulta
        `$Inicio = (Get-Date).AddDays(-`$Dias)
        Write-Output "BUSCANDO REINICIOS EN `$env:COMPUTERNAME (Ultimos `$Dias dias)..."
        Write-Output "------------------------------------------------------------"
        `$eventos = Get-WinEvent -FilterHashtable @{LogName='System'; ID=6005,6006,12,13,6008,41; StartTime=`$Inicio} -ErrorAction Stop | Sort-Object TimeCreated
        `$eventosBSOD = @(Get-WinEvent -FilterHashtable @{LogName='System'; ID=1001; StartTime=`$Inicio} -ErrorAction SilentlyContinue)
        if (`$eventos) {
            `$arranque = `$null
            foreach (`$e in `$eventos) {
                if (`$e.Id -eq 6005 -or `$e.Id -eq 12) { if (`$arranque -eq `$null) { `$arranque = `$e.TimeCreated } }
                elseif ((`$e.Id -eq 6006 -or `$e.Id -eq 13) -and `$arranque -ne `$null) {
                    `$apagado = `$e.TimeCreated; `$duracion = `$apagado - `$arranque
                    if (`$duracion.TotalMinutes -lt 1) { continue }
                    `$diasOn=[math]::Round(`$duracion.TotalHours,1); `$Tipo=if(`$e.Id -eq 13){"HIBERNACION/INICIO RAPIDO"}else{"REINICIO LIMPIO"}
                    `$strArr=`$arranque.ToString('dd/MM/yyyy HH:mm:ss'); `$strApg=`$apagado.ToString('dd/MM/yyyy HH:mm:ss')
                    if (`$duracion.Days -ge 7) { Write-Output "[ALERTA] Encendido `$(`$duracion.Days) dias (`$diasOn hrs). De `$strArr a `$strApg -> [`$Tipo]" }
                    else { Write-Output "[OK] Encendido `$(`$duracion.Days) dias (`$diasOn hrs). De `$strArr a `$strApg -> [`$Tipo]" }
                    `$arranque = `$null
                }
                elseif ((`$e.Id -eq 6008 -or `$e.Id -eq 41) -and `$arranque -ne `$null) {
                    `$horaError=`$e.TimeCreated; `$duracionParcial=`$horaError - `$arranque
                    if (`$duracionParcial.TotalMinutes -lt 1) { continue }
                    `$esBSOD=`$false; foreach(`$bsod in `$eventosBSOD){if([math]::Abs((`$bsod.TimeCreated-`$horaError).TotalMinutes) -le 5){`$esBSOD=`$true;break}}
                    `$EtiquetaFalla=if(`$esBSOD){"PANTALLA AZUL (Fallo de Sistema)"}else{"BOTONAZO / CORTE DE ENERGIA"}
                    Write-Output "[PELIGRO] Encendido `$(`$duracionParcial.Days) dias (`$([math]::Round(`$duracionParcial.TotalHours,1)) hrs). De `$(`$arranque.ToString('dd/MM/yyyy HH:mm:ss')) a `$(`$horaError.ToString('dd/MM/yyyy HH:mm:ss')) -> [`$EtiquetaFalla]"
                    `$arranque = `$null
                }
            }
            if (`$arranque -ne `$null) {
                `$duracionActual=(Get-Date)-`$arranque
                Write-Output "------------------------------------------------------------"
                Write-Output "[ACTUAL] El equipo lleva encendido `$(`$duracionActual.Days) dias (`$([math]::Round(`$duracionActual.TotalHours,1)) hrs) desde `$(`$arranque.ToString('dd/MM/yyyy HH:mm:ss'))."
            }
        } else { Write-Output "[INFO] No se encontraron eventos en el rango de fechas." }
        Write-Output "------------------------------------------------------------"
    } catch { Write-Output "[ERROR INTERNO] No se pudo procesar el Visor de Eventos." }
} 2>&1 | Out-File -FilePath "C:\Users\Public\ResReinicios.txt" -Encoding UTF8
'LISTO' | Out-File -FilePath "C:\Users\Public\ResReinicios.done" -Encoding UTF8
"@
        $ScriptPayload | Out-File -FilePath "$RutaPublicaSMB\CazaReinicios.ps1" -Encoding UTF8 -Force
        Remove-Item -Path "$RutaPublicaSMB\ResReinicios.done" -EA SilentlyContinue -Force
        $CmdCrearTarea = "schtasks.exe /Create /S $TargetIP /U `"$Usr`" /P `"$Pwd`" /RU `"SYSTEM`" /TN `"ValidarReinicios`" /TR `"powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\Users\Public\CazaReinicios.ps1`" /SC ONCE /ST 00:00 /F"
        Invoke-Expression $CmdCrearTarea 2>&1 | Out-Null
        $CmdEjecutarTarea = "schtasks.exe /Run /S $TargetIP /U `"$Usr`" /P `"$Pwd`" /TN `"ValidarReinicios`""
        Invoke-Expression $CmdEjecutarTarea 2>&1 | Out-Null
        $Cronometro = [System.Diagnostics.Stopwatch]::StartNew()
        while (-not (Test-Path "$RutaPublicaSMB\ResReinicios.done")) {
            if ($Cronometro.Elapsed.TotalSeconds -ge 90) { break }
            [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100
        }
        $Cronometro.Stop()
        $TextoFinal = "No se pudieron obtener los datos."
        if (Test-Path "$RutaPublicaSMB\ResReinicios.txt") { $TextoFinal = Get-Content "$RutaPublicaSMB\ResReinicios.txt" | Out-String }
        $CmdBorrarTarea = "schtasks.exe /Delete /S $TargetIP /U `"$Usr`" /P `"$Pwd`" /TN `"ValidarReinicios`" /F"
        Invoke-Expression $CmdBorrarTarea 2>&1 | Out-Null
        Remove-Item -Path "$RutaPublicaSMB\CazaReinicios.ps1" -EA SilentlyContinue -Force
        Remove-Item -Path "$RutaPublicaSMB\ResReinicios.txt" -EA SilentlyContinue -Force
        Remove-Item -Path "$RutaPublicaSMB\ResReinicios.done" -EA SilentlyContinue -Force
        net use \\$TargetIP\C$ /delete 2>&1 | Out-Null
        $FormRes = New-Object System.Windows.Forms.Form; $FormRes.Text = "Historial de Reinicios - $TargetIP"; $FormRes.Size = New-Object System.Drawing.Size(750,500); $FormRes.StartPosition = "CenterParent"; $FormRes.BackColor = [System.Drawing.Color]::FromArgb(20,20,20)
        $rtb = New-Object System.Windows.Forms.RichTextBox; $rtb.Dock = "Fill"; $rtb.BackColor = [System.Drawing.Color]::FromArgb(15,15,15); $rtb.Font = New-Object System.Drawing.Font("Consolas",11); $rtb.ReadOnly = $true
        $FormRes.Controls.Add($rtb)
        $FormRes.Add_Shown({
            $rtb.Text = $TextoFinal
            foreach ($linea in ($TextoFinal -split "`n")) {
                $lineaLimpia = $linea -replace "`r",""
                if ([string]::IsNullOrWhiteSpace($lineaLimpia)) { continue }
                $startPos = $rtb.Find($lineaLimpia,[System.Windows.Forms.RichTextBoxFinds]::None)
                if ($startPos -ge 0) {
                    $rtb.Select($startPos,$lineaLimpia.Length)
                    if ($lineaLimpia -match "\[ALERTA\]|\[PELIGRO\]|\[ERROR") { $rtb.SelectionColor = [System.Drawing.Color]::LightCoral }
                    elseif ($lineaLimpia -match "\[OK\]|\[ACTUAL\]") { $rtb.SelectionColor = [System.Drawing.Color]::LimeGreen }
                    else { $rtb.SelectionColor = [System.Drawing.Color]::White }
                }
            }
            $rtb.Select(0,0)
        })
        $FormRes.ShowDialog() | Out-Null
        $lblStatus.Text = "Consulta de reinicios finalizada."
    } catch { [System.Windows.Forms.MessageBox]::Show("Error al conectar a $($TargetIP): $($_.Exception.Message)","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error); $lblStatus.Text = "Error consultando reinicios." }
    finally { $Form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# ====================================================================
$Form.ShowDialog() | Out-Null
