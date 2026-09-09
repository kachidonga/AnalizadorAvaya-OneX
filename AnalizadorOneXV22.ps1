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
$symPhone = [char]9742    # ☎  contestaron (saliente conectada)
$symPause = [char]10074   # ❚  el asesor pone la llamada en espera (hold del agente)

# --- VARIABLES GLOBALES DE SESIÓN ---
$Script:Creds          = $null
$Script:DriveName      = "UnidadAuditoria"
$Script:CurrentIP      = ""
$Script:RutaManual     = ""
$Script:DirFinalGlobal = ""
$Script:ModoContestacion    = @{}   # VI_ID -> modo de contestación
$Script:AlertingHoras         = @{}   # UUID(36) -> HH:mm:ss del alerting por llamada
$Script:AlertingHorasConsumed  = @{}   # HH:mm:ss -> $true: alerting ya usada por un PRIMARY_CONNECTED; bloquea duplicados mid-call
$Script:AlertingPorCxtUuid    = @{}   # callUUID(cxt=) -> HH:mm:ss — puente CM Auto-Answer para PRIMARY_CONNECTED
$Script:CxtToConnId           = @{}   # callUUID(cxt=) -> ConnectionId numérico — puente para bloquear SECONDARY tardíos
$Script:AlertingHorasRedirect = @{}   # HH:mm:ss -> HH:mm:ss,ms: cuando señal se mueve al slot ms (auto-in mismo segundo)

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

$btnBuscarUsr = New-Object System.Windows.Forms.Button; $btnBuscarUsr.Text = "1. BUSCAR USUARIOS"; $btnBuscarUsr.Location = New-Object System.Drawing.Point(215, 16); $btnBuscarUsr.Size = New-Object System.Drawing.Size(125, 28); $btnBuscarUsr.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnBuscarUsr.FlatStyle = "Flat"; $btnBuscarUsr.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$lblUsuario = New-Object System.Windows.Forms.Label; $lblUsuario.Text = "Usuario Windows:"; $lblUsuario.Location = New-Object System.Drawing.Point(365, 20); $lblUsuario.AutoSize = $true
$cmbUsuarios = New-Object System.Windows.Forms.ComboBox; $cmbUsuarios.Location = New-Object System.Drawing.Point(475, 18); $cmbUsuarios.Size = New-Object System.Drawing.Size(130, 25); $cmbUsuarios.BackColor = $ColorPanel; $cmbUsuarios.ForeColor = $ColorTexto; $cmbUsuarios.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$lblFecha = New-Object System.Windows.Forms.Label; $lblFecha.Text = "Fecha:"; $lblFecha.Location = New-Object System.Drawing.Point(615, 20); $lblFecha.AutoSize = $true
$dtpFecha = New-Object System.Windows.Forms.DateTimePicker; $dtpFecha.Location = New-Object System.Drawing.Point(660, 18); $dtpFecha.Size = New-Object System.Drawing.Size(105, 25)
$dtpFecha.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom; $dtpFecha.CustomFormat = "dd/MM/yyyy"
$dtpFecha.MinDate = [datetime]::new(2020, 1, 1); $dtpFecha.MaxDate = [datetime]::new(2035, 12, 31)

$btnAnalizar = New-Object System.Windows.Forms.Button; $btnAnalizar.Text = "2. ANALIZAR"; $btnAnalizar.Location = New-Object System.Drawing.Point(775, 16); $btnAnalizar.Size = New-Object System.Drawing.Size(125, 28); $btnAnalizar.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnAnalizar.FlatStyle = "Flat"; $btnAnalizar.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnAnalizar.Enabled = $false

$btnRutaManual = New-Object System.Windows.Forms.Button; $btnRutaManual.Text = "3. RUTA MANUAL"; $btnRutaManual.Location = New-Object System.Drawing.Point(905, 16); $btnRutaManual.Size = New-Object System.Drawing.Size(125, 28); $btnRutaManual.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnRutaManual.FlatStyle = "Flat"; $btnRutaManual.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$btnInfoPC = New-Object System.Windows.Forms.Button; $btnInfoPC.Text = "4. INFO PC"; $btnInfoPC.Location = New-Object System.Drawing.Point(1035, 16); $btnInfoPC.Size = New-Object System.Drawing.Size(125, 28); $btnInfoPC.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnInfoPC.FlatStyle = "Flat"; $btnInfoPC.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnInfoPC.Visible = $false

$btnExtraccion = New-Object System.Windows.Forms.Button; $btnExtraccion.Text = "5. EXTRAER LOG"; $btnExtraccion.Location = New-Object System.Drawing.Point(1035, 16); $btnExtraccion.Size = New-Object System.Drawing.Size(125, 28); $btnExtraccion.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnExtraccion.FlatStyle = "Flat"; $btnExtraccion.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnExtraccion.Enabled = $false
$btnExtraccion.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$btnBusqueda = New-Object System.Windows.Forms.Button; $btnBusqueda.Text = "6. BÚSQUEDA EN LOGS"; $btnBusqueda.Location = New-Object System.Drawing.Point(1165, 16); $btnBusqueda.Size = New-Object System.Drawing.Size(125, 28); $btnBusqueda.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnBusqueda.FlatStyle = "Flat"; $btnBusqueda.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnBusqueda.Enabled = $false
$btnBusqueda.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$btnReinicios = New-Object System.Windows.Forms.Button; $btnReinicios.Text = "7. REINICIOS"; $btnReinicios.Location = New-Object System.Drawing.Point(1295, 16); $btnReinicios.Size = New-Object System.Drawing.Size(125, 28); $btnReinicios.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnReinicios.FlatStyle = "Flat"; $btnReinicios.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnReinicios.Visible = $false
$btnReinicios.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$btnExportarCSV = New-Object System.Windows.Forms.Button; $btnExportarCSV.Text = "8. EXPORTAR"; $btnExportarCSV.Location = New-Object System.Drawing.Point(1295, 16); $btnExportarCSV.Size = New-Object System.Drawing.Size(125, 28); $btnExportarCSV.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnExportarCSV.FlatStyle = "Flat"; $btnExportarCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnExportarCSV.Enabled = $false
$btnExportarCSV.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

$chkVerCalidad = New-Object System.Windows.Forms.CheckBox; $chkVerCalidad.Text = "Ver Calidad de Red"; $chkVerCalidad.Location = New-Object System.Drawing.Point(1435, 20); $chkVerCalidad.ForeColor = [System.Drawing.Color]::Cyan; $chkVerCalidad.AutoSize = $true; $chkVerCalidad.Checked = $false
$chkVerCalidad.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# Botón para revisar posibles errores del código base (fuera del timeline). Aparece tras analizar, si hay.
$btnRevisarErrores = New-Object System.Windows.Forms.Button; $btnRevisarErrores.Text = "⚠ Revisar posibles errores"; $btnRevisarErrores.Location = New-Object System.Drawing.Point(1310, 78); $btnRevisarErrores.Size = New-Object System.Drawing.Size(210, 26); $btnRevisarErrores.BackColor = [System.Drawing.Color]::FromArgb(80,70,20); $btnRevisarErrores.ForeColor = [System.Drawing.Color]::Khaki; $btnRevisarErrores.FlatStyle = "Flat"; $btnRevisarErrores.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold); $btnRevisarErrores.Visible = $false
$btnRevisarErrores.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# Botón para revisar VACÍOS de log (silencio total en Endpoint+AvayaOneX = posible congelamiento del
# proceso OneXAgent.exe) detectados en el DÍA COMPLETO. Aparece tras analizar, si hay. (Pablo, 08/2026.)
$btnVacios = New-Object System.Windows.Forms.Button; $btnVacios.Text = "🔍 Vacíos en logs"; $btnVacios.Location = New-Object System.Drawing.Point(1090, 78); $btnVacios.Size = New-Object System.Drawing.Size(210, 26); $btnVacios.BackColor = [System.Drawing.Color]::FromArgb(20,50,80); $btnVacios.ForeColor = [System.Drawing.Color]::LightSkyBlue; $btnVacios.FlatStyle = "Flat"; $btnVacios.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold); $btnVacios.Visible = $false
$btnVacios.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

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
$GridResultados.Columns.Add("EvDtmf", "DTMF") | Out-Null; $GridResultados.Columns["EvDtmf"].FillWeight = 8
$GridResultados.Columns.Add("Interpretacion", "Actividad del Agente") | Out-Null; $GridResultados.Columns["Interpretacion"].FillWeight = 20
$GridResultados.Columns.Add("EvAgente", "Endpoint.log") | Out-Null; $GridResultados.Columns["EvAgente"].FillWeight = 15
$GridResultados.Columns.Add("EvAudio", "Audio.log") | Out-Null; $GridResultados.Columns["EvAudio"].FillWeight = 15   # Se reconsideró ocultarla: al esconderla, los slots de "Línea abierta/cerrada" (que solo pueblan Audio) quedaban como filas vacías. Se mantiene visible.
$GridResultados.Columns.Add("EvAux", "AvayaOneX.log") | Out-Null; $GridResultados.Columns["EvAux"].FillWeight = 10
$GridResultados.Columns.Add("EvIspeac", "Calidad Red/Voz") | Out-Null; $GridResultados.Columns["EvIspeac"].FillWeight = 12; $GridResultados.Columns["EvIspeac"].Visible = $false
$GridResultados.Columns.Add("EvSysLog", "Log de Sistema") | Out-Null; $GridResultados.Columns["EvSysLog"].FillWeight = 10
$GridResultados.Columns.Add("EvAppLog", "Log de Aplicación") | Out-Null; $GridResultados.Columns["EvAppLog"].FillWeight = 10
foreach ($col in $GridResultados.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }

$Form.Controls.AddRange(@($lblIP, $txtIP, $btnBuscarUsr, $lblUsuario, $cmbUsuarios, $btnAnalizar, $lblFecha, $dtpFecha, $btnRutaManual, $btnInfoPC, $btnReinicios, $btnExtraccion, $btnBusqueda, $btnExportarCSV, $chkVerCalidad, $btnRevisarErrores, $btnVacios, $lblExtension, $lblStatus, $GridResultados))

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

[void]$ctxMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$mnuDetalleConf = New-Object System.Windows.Forms.ToolStripMenuItem
$mnuDetalleConf.Text = "Ver detalle: Como se establecio la Conferencia?"
$mnuDetalleConf.Enabled = $false
[void]$ctxMenu.Items.Add($mnuDetalleConf)
$mnuDetalleTransf = New-Object System.Windows.Forms.ToolStripMenuItem
$mnuDetalleTransf.Text = "Ver detalle: Como se completo la Transferencia?"
$mnuDetalleTransf.Enabled = $false
[void]$ctxMenu.Items.Add($mnuDetalleTransf)

$GridResultados.ContextMenuStrip = $ctxMenu

# Cancelar apertura del menú si ningún ítem está habilitado
$ctxMenu.Add_Opening({
    param($sender, $e)
    if (-not $mnuDiag.Enabled -and -not $mnuDiagFin.Enabled -and -not $mnuDiagExplicaFin.Enabled -and -not $mnuDiagDir.Enabled -and -not $mnuDetalleConf.Enabled -and -not $mnuDetalleTransf.Enabled) { $e.Cancel = $true }
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
            $CeldaAgente = $GridResultados.Rows[$hit.RowIndex].Cells["EvAgente"].Value
            $mnuDiag.Enabled           = ($CeldaInterp -match "señal de llamada") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDiagFin.Enabled        = ($CeldaInterp -match "INICIO DE LLAMADA|LÍNEA ABIERTA") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDiagExplicaFin.Enabled = ($CeldaInterp -match "FIN DE LLAMADA|CUELGUE MANUAL|Llamada finalizada por el agente") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDiagDir.Enabled        = ($CeldaInterp -match "INICIO DE LLAMADA") -and ($null -ne $Script:SnapEventosTiempo)
            $mnuDetalleConf.Enabled    = ($CeldaAgente -match "CONFERENCIA ESTABLECIDA")
            $mnuDetalleTransf.Enabled  = ($CeldaAgente -match "TRANSFERENCIA COMPLETADA")
        } else { $mnuDiag.Enabled = $false; $mnuDiagFin.Enabled = $false; $mnuDiagExplicaFin.Enabled = $false; $mnuDiagDir.Enabled = $false; $mnuDetalleConf.Enabled = $false; $mnuDetalleTransf.Enabled = $false }
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
    } catch {
        $Script:Creds = $null
        [System.Windows.Forms.MessageBox]::Show("Error al conectar: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
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

    $Script:RutaManual = ""   # limpiar ruta manual al cambiar a modo red

    # --- Loop de credenciales: reintenta automaticamente si falla autenticacion ---
    while ($true) {
        if ($null -eq $Script:Creds) {
            try { $Script:Creds = Get-Credential -UserName "local\soporte" -Message "Credenciales para \\$TargetIP\c$" -ErrorAction Stop } catch { $Script:Creds = $null }
            if ($null -eq $Script:Creds) { $lblStatus.Text = "Conexion cancelada."; break }
        }

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
            $lblStatus.Text = "Conexion exitosa. Selecciona un usuario y fecha, luego clic en ANALIZAR."
            break   # conexion exitosa: salir del loop
        } catch {
            $Script:Creds = $null
            [System.Windows.Forms.MessageBox]::Show("No se pudo conectar a $TargetIP.`n$($_.Exception.Message)", "Error de Conexion", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $lblStatus.Text = "Error al conectar. Vuelve a ingresar credenciales..."
            # el loop vuelve a pedir credenciales automaticamente
        } finally {
            if (Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue) { Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue | Out-Null }
        }
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
        # AUTO-DESCENSO: si la carpeta elegida no contiene los logs directamente (se eligió un
        # nivel arriba, ej. "...\Avaya" en vez de "...\Avaya\one-X Agent\2.5\Log Files"),
        # buscar EndpointLog/OneXAgent recursivamente y usar SU carpeta. Evita el
        # "Archivos de log leídos: 0" por apuntar al nivel equivocado.
        $TieneLogs = Get-ChildItem -LiteralPath $Script:RutaManual -Filter "EndpointLog.txt*" -EA SilentlyContinue
        if (-not $TieneLogs) {
            $LogHallado = Get-ChildItem -LiteralPath $Script:RutaManual -Recurse -EA SilentlyContinue |
                Where-Object { -not $_.PSIsContainer -and ($_.Name -like "EndpointLog.txt*" -or $_.Name -match "(?i)^one-?x.*\.(log|txt)") } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($LogHallado) { $Script:RutaManual = $LogHallado.DirectoryName }
        }
        $lblStatus.Text = "Ruta manual cargada: $($Script:RutaManual)"
        Reset-Entorno
        # Auto-ajuste del DateTimePicker según fechas disponibles en los logs
        $LogsDisponibles = Get-ChildItem -Path $Script:RutaManual -Filter "EndpointLog.txt*" -EA SilentlyContinue
        if ($LogsDisponibles) {
            $FechasLog = @()
            foreach ($lf in $LogsDisponibles) {
                $contenido = Get-Content $lf.FullName -TotalCount 50 -EA SilentlyContinue
                foreach ($linea in $contenido) {
                    if ($linea -match "(\d{2}/\d{2}/\d{4})") { try { $FechasLog += [datetime]::ParseExact($matches[1],"MM/dd/yyyy",$null) } catch {} }   # Avaya usa MM/dd/yyyy (gringo), NO dd/MM — antes auto-seteaba el día equivocado y la extracción no hallaba nada
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
    # Logs Avaya usan formato MM/dd/yyyy (formato gringo). Solo se incluyen variantes MM/dd:
    # F1 = MM/dd/yyyy (con ceros), F4 = M/d/yyyy (sin ceros), F3 = yyyy-MM-dd (ISO, no ambiguo).
    # Se excluyen dd/MM y d/M para evitar colisiones: "05/06/2026" (Jun 5 en dd/MM)
    # matchearía entradas del 6 de mayo escritas como "05/06/2026" en MM/dd.
    $F1 = $dtpFecha.Value.ToString("MM/dd/yyyy"); $F3 = $dtpFecha.Value.ToString("yyyy-MM-dd"); $F4 = $dtpFecha.Value.ToString("M/d/yyyy")
    $FechaOmni = "(?:$([regex]::Escape($F1))|$([regex]::Escape($F3))|$([regex]::Escape($F4)))"

    $GridResultados.Rows.Clear(); $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $EventosTiempo       = @{}
    $ListaExtensiones    = @()
    $ListaLogins         = @()
    # Buffer de posibles errores del código base (ERROR/Exception/FATAL genéricos). NO se pintan en el
    # timeline (inflaban con ruido benigno y "Evento sospechoso" alarmaba en pantalla compartida); se
    # revisan aparte con el botón "Revisar posibles errores".
    $Script:ErroresSospechosos = @()
    $btnRevisarErrores.Visible = $false
    # Buscador de VACÍOS de log (día completo): ms-del-día de TODA línea con timestamp en Endpoint+
    # AvayaOneX (proceso principal) e IspeacLog (proceso independiente), recolectados sin releer
    # archivos durante PASO 3/4/5. List[Int64] en vez de array +=: evita el costo O(n²) en logs de
    # un día completo (pueden ser decenas de miles de líneas).
    $Script:TsProceso = [System.Collections.Generic.List[Int64]]::new()
    $Script:TsIspeac  = [System.Collections.Generic.List[Int64]]::new()
    $Script:VaciosDetectados = @()
    $btnVacios.Visible = $false
    $Script:ModoContestacion    = @{}
    $Script:AlertingHoras         = @{}
    $Script:AlertingHorasConsumed  = @{}
    $Script:AlertingPorCxtUuid    = @{}
    $Script:CxtToConnId           = @{}
    $Script:AlertingHorasRedirect = @{}
    $DirLlamada          = @{}
    $ListaMakeCall       = @()
    $ConexionesYaIniciadas = @{}
    $PrimaryConnectedSeconds = @{}   # O(1) guard para SECONDARY CONNECTED
    $AlertingPorTel = @{}            # telNorm → horaLimpia: unifica dedup entre XML, WI.ADD y fallback
    $PhoneYaEnInicio = @{}           # telNorm → $true: bloquea SECONDARY para teléfono ya conectado por PRIMARY
    $SesionHoraInicio = @{}          # sesionId → horaSlot: para actualizar LÍNEA ABIERTA cuando llega tel real
    $CallStateDisconnected = @{}     # sesionId → horaSlot: fallback FIN cuando ProcessSessionEndedEvent no aparece
    $MapeoTelP4 = @{}                # sesionId → tel: capturado en PASO 4 desde "Call ended" cuando EndpointLog no tiene la sesión
    $HoldImplicito = @{}             # sesionId → LISTA de horaSlot: cada hold automático del sistema (sin clic de Hold). Navegando entre líneas una llamada se retiene varias veces.
    $RetomaImplicita = @{}           # sesionId → LISTA de horaSlot: cada vez que la llamada sale del hold (Inactive→Active) al volver a su línea
    $CallCreatedMs = @()             # ms absolutos de cada "Call created": distingue "abre línea" (nace llamada) de "cambia a la línea" (se retoma una existente)
    $DtmfPresses = @()               # lista de pulsaciones DTMF del asesor durante una llamada: @{Sim;Ses;Hora}. constant 19-28=0-9, 29='*', 30='#'
    $DtmfConstPend = $null           # constant del último "input digit constant" en espera de su nCallIndex (sesión)
    $DtmfHoraPend  = $null           # slot ms de esa pulsación pendiente
    $NewCallSlots  = @()             # slots "HH:mm:ss,fff" de cada NewCallHandler (asesor captura número en caja + Enter)
    $DialedApply   = @()             # números pasados por ApplyDialingRulesToNumber: @{Ms;Num}. Se emparejan con NewCallSlots para mostrar el número en la etiqueta
    $ViDisconnected = @{}            # viUUID → horaSlot: FIN por VI StateImpl type=Disconnected (cubre cuelgue en hold sin EndpointLog/ProcessSessionEndedEvent)
    $ViFinEnHold = @{}               # viUUID → $true: el VI estaba en hold (Inactive) al desconectarse → CLIENTE COLGÓ EN HOLD
    $ViLastState = @{}               # viUUID → "Active"/"Inactive": último estado vivo del VI antes del Disconnected
    $SesionesBridge = @{}            # sesionId → horaSlot: sesiones bridge creadas por auto-transferencia del sistema
    $BridgeSlot = @{}                # sesionId → slot ms exacto de la fila "Sesión bridge": permite retirarla si luego resulta ser una llamada real (ver retract en PASO 5)
    $ViSinBotones = @{}              # viUUID → @{Slot;Raw}: VI entrante Alerting con answer=False que NUNCA fue aceptada (ni Auto Accepting ni AnswerVoiceInteraction) → UI sin botones para contestar (caso dleal 02/07/2026)
    $Script:FallaRecepSlot = $null   # slot activo de "FALLA EN LA RECEPCIÓN DE LLAMADA": mientras esté seteado, las líneas sin fecha (stack trace del Exception) se anexan a su Raw
    $Script:FallaRecepLineas = 0     # contador de líneas de stack capturadas (tope 12)
    $UltimaTransfGUIHora = ""        # hora del último GUI Method TransferHandler/ConsultationHandler (PASO 4)
    # Diccionario de Reason Codes (motivos de Auxiliar). Movido aquí (antes vivía solo en PASO 6/render) para
    # que PASO 4 también pueda armar la etiqueta con nombre al capturar el clic de "Auxiliar con motivo"
    # (ver $SlotAuxConMotivo más abajo — Opción B, Pablo 08/2026).
    $DictRC = @{ "0"="DEFAULT"; "1"="COMIDA"; "2"="BAÑO"; "3"="LLAMADA SALIDA"; "4"="CAPACITACION"; "5"="SERVICIOS ESPECIALES"; "6"="COBRANZA"; "7"="SEGUIMIENTO"; "8"="RETRO"; "9"="SISTEMAS" }
    # Slot ms protegido del botón "Auxiliar con motivo" (esquina superior izquierda): igual que ya existe para
    # "Disponible" (EnterReadyHandler crea su propia fila con ms, inmune a que Avaya, en el MISMO segundo,
    # dispare un parpadeo interno Ready→LoggedOut→Aux que contaminaba la fila de segundo compartida). Sin
    # esto, un reason code real podía "perderse" si ese parpadeo caía en el mismo segundo (caso Pablo, AUX
    # SISTEMAS/9, prueba controlada 28/08/2026).
    $SlotAuxConMotivo = ""
    $CodigoRCPendiente = ""
    # Milisegundos absolutos de un slot "HH:mm:ss[,fff]" — copia adelantada de la que ya existe más abajo
    # (línea ~2860), necesaria aquí porque PASO 4 corre ANTES de esa definición. Redefinirla allá abajo no
    # rompe nada (misma lógica, reasignación idempotente).
    $MsDeSlot = {
        param($s)
        if     ($s -match "^(\d{2}:\d{2}:\d{2}),(\d{1,3})$") { $hh = $matches[1]; $mm = [int]$matches[2] }
        elseif ($s -match "^(\d{2}:\d{2}:\d{2})$")           { $hh = $matches[1]; $mm = 0 }
        else { return -1 }
        try { return [int]([datetime]::ParseExact($hh,"HH:mm:ss",$null).TimeOfDay.TotalSeconds) * 1000 + $mm } catch { return -1 }
    }
    # --- Auxiliar solicitado CON LLAMADA ACTIVA (estado "PendingAux" de Avaya) ---
    # Si el asesor elige Auxiliar mientras sigue en llamada, Avaya no lo mete de inmediato: lo deja
    # "PendingAux" y aplica el cambio real (LoggedOut→Aux) hasta que cuelga — normalmente varios segundos
    # después, en un parpadeo oldState=PendingAux;newState=LoggedOut seguido de oldState=LoggedOut;newState=Aux
    # en el mismo segundo. Antes esto NO se distinguía de un Auxiliar normal (Punto 2, Pablo 28/08/2026):
    # $UltimoMotivoElegidoP4  = espejo de $UltimoMotivoElegido (que vive en PASO 6/render) pero disponible
    #                           aquí en PASO 4, para poder etiquetar la finalización diferida con el motivo.
    # $UltimoPendingAuxLogoutSlot = slot ms del "oldState=PendingAux;newState=LoggedOut" más reciente — la
    #                           ventana (≤3s) para reconocer que el "LoggedOut→Aux" que sigue es la
    #                           finalización diferida de ESE PendingAux, no un Aux nuevo sin relación.
    $UltimoMotivoElegidoP4 = ""
    $UltimoPendingAuxLogoutSlot = ""
    $UltimoSlotAuxClicMs = ""   # slot ms del último "Asesor se cambia a Auxiliar... (Confirmado por clic)" — por si hay que reetiquetarlo a "pendiente"
    # --- Intento de desfirme (LogoutAgentHandler) que puede FALLAR ---
    # Si al pedir el logout hay una llamada activa/entrante justo en ese instante, Avaya puede rechazarlo
    # ("Session_LogoutAgent failed...Call not disconnected") y el asesor queda pegado en un PendingAux
    # residual hasta que esa llamada termina — sin verse desfirmado en realidad (caso Pablo, 01/09/2026).
    $UltimoSlotDesfirmeSolicitado = ""   # slot base (bare-hour) del último "LogoutAgent...code=ReasonCode[10]"
    $DesfirmeEnProceso = $false          # true entre la solicitud y el "GUI Method ENDED: LogoutAgentHandler" (éxito o fallo)
    $PendingAuxEsResiduoDesfirme = $false # true si el PendingAux en curso es residuo de un intento de desfirme fallido, no un Auxiliar real
    $DesfirmeYaConfirmado = $false        # true una vez que YA se creó la fila "Asesor ya se encuentra desfirmado" para el intento en curso (evita duplicarla si Avaya parpadea LoggedOut→Aux→LoggedOut al cerrar)
    # OJO: separado de $DesfirmeEnProceso a propósito — "GUI Method ENDED: LogoutAgentHandler" SIEMPRE
    # llega ANTES (o en el mismo ms) que la confirmación real "newState=LoggedOut" que le sigue, así que
    # si se usara la misma bandera para ambas cosas, limpiarla en ENDED apagaría la confirmación antes de
    # que pudiera dispararse. $DesfirmeEnProceso sigue protegiendo solo el PendingAux (Ronda 1); esta
    # bandera vive un poco más: desde la solicitud hasta la confirmación real o el fallo.
    $DesfirmeEsperandoConfirmacion = $false
    # --- Mecanismo 2 de desfirme: marcado manual "565"+dígito (ej. "565"+"3"), Pablo 09/09/2026 ---
    # No deja ninguna línea de "intento" limpia (ver hallazgo abajo) — se rastrea viendo cuándo una
    # llamada saliente marca EXACTO "565" y luego se extiende con un dígito más en esa MISMA llamada.
    $CandidatoIdLlamada565 = ""   # Id de la llamada saliente que llegó a marcar "565" exacto
    $CandidatoSlot565      = ""  # slot ms del momento "565"+N (el "intento" real, si se confirma)
    $CandidatoMs565        = -1  # ms absolutos de ese slot, para la ventana de 10s
    $CandidatoLinea565     = ""  # línea cruda del "565"+N, para el RawInterpretacion del intento
    $GuiHoldSeg = @{}; $GuiUnholdSeg = @{}; $GuiEndSeg = @{}   # segundo del clic GUI (Hold/UnHold/EndCall) → sella "✓ clic confirmado" en el evento manual del EndpointLog
    $TransfClicSeg   = @{}           # segundo (HH:mm:ss) del clic GUI "TransferCallHandler STARTED" → etiqueta "Asesor presiona botón Transferir" (distingue manual vs automática) y evita duplicar con OnRequestTransferSession
    $EndTransferSeg  = @{}           # segundo (HH:mm:ss) de "End Executing method Transfer" (PASO 4/OneXAgent) → detecta retomado automático de la llamada tras fallo de transferencia
    $LineaAppPorSeg  = @{}           # segundo (HH:mm:ss) → letra de call-appearance (a/b/c…) de MapCallAppToBtnIndex (PASO 5/EndpointLog) → número de línea en "Asesor abre línea N"
    $HoldFalloSes    = @{}           # sesión → slot del fallo: hold que NO se pudo completar (HoldSessionCommand response null / VoiceInteraction_Hold failed) → relabela su "HOLD MANUAL" como fallido
    $HoldMetodoSes   = @{}           # ConnId → LISTA de slots de "Begin Executing method Hold(...ConnId=N...)": hold que el asesor ejecuta por una vía (hotkey/CTI) que NO deja HoldCallHandler ni OnRequestHoldSession → distingue hold del AGENTE del auto-hold del SISTEMA (cambio de línea)
    $AddCallSlots    = @()           # slots "HH:mm:ss,fff" del botón "Agregar llamada" (AddCallHandler): ese botón pone la llamada actual en AutoHold; el barrido reetiqueta ese hold cercano como "por Agregar llamada" (no es un clic de hold suelto)
    $LineaAppBtn     = @{}           # buttonIndex (7,8,9…) → número de línea (1,2,3…), derivado de MapCallAppToBtnIndex(letra)
    $ActiveLineSeq   = @()           # cronología de la línea activa: @{Ms;Btn} de "found activeLine: N" y <lineAppearanceId>N</lineAppearanceId>
    # --- Caso "logeo autónomo": distinguir CIERRE NORMAL de CAÍDA DE RED + RECONEXIÓN AUTOMÁTICA ---
    # (Cierre normal: requestor='manual' + closeSignalingChannel cat:0/cod:0 "closed by the application" + ExitHandler/Shutdown.
    #  Caída+reconexión: closeSignalingChannel cat:2/cod:10060 + RASKeepaliveFailed + LinkRecoveryProgressEvent +
    #  CompleteLoginRequest bLInkRecovery=1 + LoginAgent. El bLInkRecovery=1 prueba que la app estaba ABIERTA.)
    # LISTAS (no un solo slot): un MISMO log puede contener VARIAS sesiones/cierres/caídas el mismo día
    # (ej. Pablo probó un cierre normal y minutos después cerró con la X, sin borrar logs). Se colecta
    # TODA ocurrencia y el barrido COLAPSA cada lista por episodios (por hueco de tiempo) → una fila por
    # episodio. Antes se guardaba solo el PRIMER slot → el segundo cierre se perdía. Cada elemento: @{Slot;Raw}.
    $CierreAppList      = @()        # cierres de app (Shutdown()/ExitHandler/PhoneService shutdown) — PASO 4
    $DesfirmeFallidoList = @()       # intentos de desfirme fallidos ("Session_LogoutAgent failed") — PASO 4; la fila
                                      # real se arma DESPUÉS de PASO 5 (ver barrido pre-PASO 6) porque el ms heredado
                                      # puede coincidir con un FIN DE LLAMADA de otra sesión que PASO 5 aún no ha
                                      # escrito — escribir aquí de una vez se perdía cuando ese FIN llegaba después
                                      # y sobreescribía la fila sin avisar (Pablo, caso 01/09/2026).
    $DesfirmeManualList = @()        # Logoff requestor='manual' — PASO 5
    $RedCaidaList       = @()        # caída de red (closeSignalingChannel cat≠0 / RASKeepaliveFailed) — PASO 5
    $ReconIntentoList   = @()        # LinkRecoveryProgressEvent (recuperación en curso) — PASO 5
    $ReconOkList        = @()        # CompleteLoginRequest ... bLInkRecovery=1 (reconexión automática) — PASO 5
    $ReconRefirmaList   = @()        # LoginAgent()/AgentStateChanged LoggedOut→Aux tras recovery — PASO 4
    # ThreadAbortException/"Subproceso anulado": ambiguo (ocurre en recuperación de red Y en teardown de cierre).
    # Se interpreta en el barrido según haya o no un cierre de app cercano.
    $HiloAbortadoList   = @()
    # LÍNEA ABIERTA SIN MARCAR — confirmación DEFINITIVA del endpoint: al cerrar el historial de la sesión
    # avisa "this record has no far-end address" = nunca hubo dirección de destino = jamás se marcó.
    # Validado con 0 falsos positivos en los logs donde SÍ se marcó. Lista (no dict) porque los IDs de
    # sesión se reciclan y hay que aparear cada apertura con SU cierre. — PASO 5
    $SinDestinoFin      = @()        # @{Ses;Slot}
    # Contrario a $SinDestinoFin: sesiones que SÍ resolvieron un número real (far-end address). Se usa para
    # SUPRIMIR una fila "LÍNEA ABIERTA SIN MARCAR" preliminar cuando su sesión sí marcó (ej. destino de una
    # transferencia cuya pata phantom se llevó el registro de consulta). @{Ses;Slot}, acotado por tiempo (IDs reciclan).
    $FarEndResueltoSlots = @()
    # Primer número REAL por sesión (PASO 5). Distingue una transferencia inter-agente REAL (la MISMA sesión
    # cambia de número: del asesor originador al cliente) de un mismo cliente que vuelve a llamar (llamada
    # nueva que nace directo con su número, sin cambio). El label "Número de cliente actualizado" solo debe
    # salir cuando ESTA sesión cambió. Se libera en ProcessSessionEndedEvent (los IDs se reciclan).
    $_PrimerFonoSes  = @{}
    $_endHistSes = $null; $_endHistSlot = $null   # pendiente: sesión del EndHistoryRecord en curso
    # CONTESTARON (saliente): instante en que la VoiceInteraction recibe la dirección del otro lado
    # (= "me contestaron"), coincidente al ms con la transición interna Alerting→Active. Se captura en
    # PASO 4 y el barrido post-PASO5 separa "Marcando/Timbrando" del "Contestaron". @{Vi;Cxt;Slot;Num}.
    $ContestoSaliente   = @()
    $RAvistoVI          = @{}         # "VIx|cxt" → $true: 1ª aparición de RemoteAddress por VI. La llave lleva el cxt (GUID único) porque los números de VI se reciclan en cada login
    # Corroboración por EVENTO del "contestaron": el log debe mostrar la transición Alerting→Active en el
    # mismo instante. Aparece SOLO cuando una llamada pasó de timbrando a activa (las aperturas de línea
    # sin marcar nunca la generan). Se usa en vez de un umbral de segundos. Guarda slots "HH:mm:ss,fff".
    $AlertingActivaSlots = @()
    # Respaldo para SALIENTES INTERNAS (extensión a extensión): estas llamadas conectan directo
    # (State=New→Active, SIN pasar por Alerting) porque no hay timbrado real que loguear — por eso
    # nunca disparan ni el RemoteAddress= de $ContestoSaliente ni la corroboración Alerting→Active de
    # $AlertingActivaSlots, y se quedan sin "(habló mm:ss)" en el FIN. sesión(ConnectionId) → primer
    # slot "HH:mm:ss,fff" en que esa sesión llega a State=Active,InnerState=CONNECTED,Outgoing=True.
    $ConexionActivaSaliente = @{}
    # DISPONIBLE por botón favorito "AUTO IN": ese botón dispara un FAC (código de función) que la central
    # marca como una saliente cortita (conecta al instante del Ready y se cuelga sola en ~4s). El botón
    # "correcto" de Disponible usa la vía CTI/API y NO genera esa llamada. $FacAutoIn = el número del FAC
    # (562 en esta central; cámbialo si otra sede usa otro código). $AutoInFacSlots = slots "HH:mm:ss,fff"
    # donde se vio esa saliente al FAC; el render marca el "Disponible" cercano como "botón favorito AUTO-IN".
    $FacAutoIn          = "562"
    $AutoInFacSlots     = @()
    # Ventanas del método Login() (registro de estación): durante ese registro Avaya crea una pata saliente
    # VACÍA (RemoteParty=[,]) que nace y muere en ~1s, ANTES de que el agente esté firmado, y se cuela como
    # "LÍNEA ABIERTA SIN MARCAR" + "¡EVASIÓN!". No es un abandono real (no se puede abandonar una línea sin
    # estar firmado). $LoginWins = @{B;E} (slots "HH:mm:ss,fff" de Begin/End del Login()); el render suprime
    # las filas "sin marcar"/EVASIÓN que caigan dentro. $_loginBeginSlot = Begin pendiente de cerrar.
    $LoginWins          = @()
    $_loginBeginSlot    = ""
    # Lecturas de calidad de red (IspeacLog RTCP). NO crean filas propias (eso saturaba el timeline):
    # se recolectan aquí y un barrido pre-render las ADJUNTA a las filas de evento YA EXISTENTES
    # (la más cercana en el tiempo). $LossReadings = pérdida de paquetes (%), $RttReadings = lag ida/vuelta (ms).
    $LossReadings       = @()   # @{Slot="HH:mm:ss,fff"; Pct}
    $RttReadings        = @()   # @{Slot="HH:mm:ss,fff"; Rtt}
    # "Cierre con asesor firmado": el cierre desfirmó a un agente que seguía FIRMADO. Señal = transición a
    # LoggedOut (desde Ready/Aux/…) o EnterAuxHandler que ocurre AL/DESPUÉS del ExitHandler (el cierre la
    # disparó). Si el desfirme fue ANTES del ExitHandler (menú Cerrar sesión) = cierre NORMAL. OJO: el botón X
    # y el menú "Salir" son IDÉNTICOS en el log (ambos = ExitHandler); lo detectable es si el asesor estaba
    # firmado al cerrar, NO el mecanismo. Y forceLogoff=true aparece en TODO cierre → no discrimina. — PASO 4
    $CierreFirmadoList  = @()
    $ConnIdToSeñalSlot = @{}        # sessionId → ms-slot: señal creada por OldState=New,NewState=Alerting (PhoneService Alerting lo actualiza con Tel)
    $TransferEntranteOrigen = @{}    # CaId → @{Tel;Nombre;SlotReal}: posible ORIGEN de transferencia entrante (quien transfirió)
    $TransferEntranteConfirm = @{}   # CaId → $true: InBoundConsultTransferCompleted confirmó la transferencia entrante
    $UltimoCandidatoTransfer = $null # CaId del último candidato de origen de transferencia (para ligar el marcador sin ConnId)
    $Script:UltimaHoraAlerting = $null
    $Script:UltimoTopic        = $null
    $XMLCargado = $false

    function Init-Hora ($hora) {
        if (-not $EventosTiempo.ContainsKey($hora)) {
            $EventosTiempo[$hora] = @{
                Interpretacion=""; Agente=""; Audio=""; Aux=""; Ispeac=""; SysLog=""; AppLog=""; Dtmf=""
                RawInterpretacion=""; RawAgente=""; RawAudio=""; RawAux=""; RawIspeac=""; RawSysLog=""; RawAppLog=""; RawDtmf=""
                Sesion="-"; Tel="-"; ViId=""; Topic=""
                ColorInterpretacion=[System.Drawing.Color]::White; ColorAgente=[System.Drawing.Color]::White
                ColorAudio=[System.Drawing.Color]::White; ColorAux=[System.Drawing.Color]::White
                ColorIspeac=[System.Drawing.Color]::White; ColorSys=[System.Drawing.Color]::White; ColorApp=[System.Drawing.Color]::White; ColorDtmf=[System.Drawing.Color]::White
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
                                    # Tema 1: la señal ENTRANTE va a un ms-slot con el timestamp real del XML.
                                    # El XML trae precisión de milisegundos (ej: 09:34:11.7147515 → "09:34:11,714"),
                                    # que antes se descartaba mandando la señal al slot base. Con el ms-slot la señal
                                    # ya no cae al slot base (no se ordena antes de eventos ms del mismo segundo) y
                                    # queda protegida de que otro evento del segundo base la sobreescriba.
                                    # La SALIENTE se mantiene en el slot base: su INICIO/LÍNEA ABIERTA real lo refinan PASO 4/5.
                                    $MsXML     = $ItemTimeLocal.ToString("fff")
                                    $SlotVoice = if ($IsOutbound) { $HoraLimpia } else { "$HoraLimpia,$MsXML" }
                                    Init-Hora $SlotVoice
                                    $EventosTiempo[$SlotVoice].ViId  = $ViId
                                    $EventosTiempo[$SlotVoice].Topic = $Topic
                                    if ($EventosTiempo[$SlotVoice].Tel -eq "-") { $EventosTiempo[$SlotVoice].Tel = $TelDisplay }

                                    if ($IsOutbound) {
                                        # --- LLAMADA SALIENTE ---
                                        if ($TelDisplay -match "Desconocido|^$") {
                                            $EventosTiempo[$SlotVoice].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR"
                                            $EventosTiempo[$SlotVoice].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        } else {
                                            $EventosTiempo[$SlotVoice].Interpretacion = "$symUp INICIO DE LLAMADA (Saliente)"
                                            $EventosTiempo[$SlotVoice].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
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
                                        $EventosTiempo[$SlotVoice].Interpretacion      = $MsgAlert
                                        $EventosTiempo[$SlotVoice].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        $Script:UltimaHoraAlerting = $SlotVoice
                                        $Script:UltimoTopic        = $Topic
                                        # Guardar UUID para cruzar con CONNECTED en OneXAgent (ahora apunta al ms-slot)
                                        if ($ViId -match "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
                                            $Script:AlertingHoras[$matches[1]] = $SlotVoice
                                        }
                                        # Registrar teléfono normalizado para dedup cruzada con WI.ADD/fallback
                                        $TelNormS = ($TelDisplay -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                        $AlertingPorTel[$TelNormS] = $SlotVoice
                                    }
                                    $DirXML = if ($IsOutbound) { "Saliente" } else { "Entrante" }
                                    $EventosTiempo[$SlotVoice].RawInterpretacion += "[XML] $ViId | Tel:$TelDisplay | $DirXML | Dur:${DurSeg}s | Topic:$Topic`n"
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
                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2})(?::(\d{3}))?") {
                            $HoraLimpia = $matches[1]; $MsAudio = if ($matches[2]) { $matches[2] } else { "000" }
                            Init-Hora $HoraLimpia
                            if ($linea -match "(?i)sActiveWave(In|Out)Device\s+'([^']+)'") {
                                $Dispositivo = $matches[2].Trim()
                                $EventosTiempo[$HoraLimpia].Interpretacion = "Extensión en línea y conectada a dispositivos de audio"
                                $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Cyan
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                                if ($EventosTiempo[$HoraLimpia].Audio -eq "") { $EventosTiempo[$HoraLimpia].Audio = "$symMusic HW: $Dispositivo" } elseif ($EventosTiempo[$HoraLimpia].Audio -notmatch [regex]::Escape($Dispositivo)) { $EventosTiempo[$HoraLimpia].Audio += " / $Dispositivo" }
                                $EventosTiempo[$HoraLimpia].ColorAudio = [System.Drawing.Color]::DeepSkyBlue; $EventosTiempo[$HoraLimpia].RawAudio += "$linea`n"
                            }
                            # Línea abierta/cerrada a slot ms (el AudioLog trae HH:mm:ss:fff) para que ordenen bien dentro del segundo
                            elseif ($linea -match "StartSession: Start ISPEAC audio session") { $SlotAu = "$HoraLimpia,$MsAudio"; Init-Hora $SlotAu; $EventosTiempo[$SlotAu].Audio = "$symMusic Línea abierta"; $EventosTiempo[$SlotAu].ColorAudio = [System.Drawing.Color]::LimeGreen; $EventosTiempo[$SlotAu].RawAudio += "$linea`n" }
                            elseif ($linea -match "EndCall: Session ended") { $SlotAu = "$HoraLimpia,$MsAudio"; Init-Hora $SlotAu; $EventosTiempo[$SlotAu].Audio = "$symStop Línea cerrada"; $EventosTiempo[$SlotAu].ColorAudio = [System.Drawing.Color]::Gray; $EventosTiempo[$SlotAu].RawAudio += "$linea`n" }
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
                        # Se CAPTURA la hora CON milisegundos (…:fff) y se recolecta la lectura, SIN crear slot
                        # (antes cada lectura RTCP hacía su propia fila y saturaba). Un barrido pre-render las
                        # adjunta a las filas de evento ya existentes. Parse InvariantCulture: el SO es es-MX
                        # (separador decimal ','), un [double] normal malinterpretaría "1.007080".
                        if ($linea -match "^\[$FechaOmni\s+(\d{2}:\d{2}:\d{2}):(\d{3})\]") {
                            $hI = $matches[1]; $msI = $matches[2]
                            # Buscador de vacíos: registra CUALQUIER línea de Ispeac (proceso independiente
                            # de OneXAgent.exe), para saber si seguía vivo durante un vacío del proceso principal.
                            if ($hI -match '^(\d{2}):(\d{2}):(\d{2})$') { $Script:TsIspeac.Add(([int64]$matches[1])*3600000L + ([int64]$matches[2])*60000L + ([int64]$matches[3])*1000L + [int64]$msI) }
                            if ($linea -match "FRACTION DROPPED = ([\d.]+)") {
                                $vFrac = [double]::Parse($matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
                                $LossReadings += [pscustomobject]@{ Slot = "$hI,$msI"; Pct = [math]::Round($vFrac * 100, 1); Raw = $linea }
                            }
                            elseif ($linea -match "RECEIVED RTCP ROUND TRIP DELAY = ([\d.]+)") {
                                $vRtt = [double]::Parse($matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
                                $RttReadings += [pscustomobject]@{ Slot = "$hI,$msI"; Rtt = [math]::Round($vRtt, 1); Raw = $linea }
                            }
                        }
                    }
                }
            }

            # ================================================================
            # PRE-PASADA: Detectar actualizaciones de teléfono mid-call en EndpointLog
            # ================================================================
            # Cuando una llamada llega vía transferencia inter-agentes, CM envía primero el número
            # del agente originador y luego actualiza (Active→Active) con el número real del cliente.
            # Eso genera dos UpdateHistoryRecord para la misma sesión con teléfonos distintos.
            # Si PASO 4 (OneXAgent.log) recibe un nuevo VoiceInteraction para el número actualizado,
            # lo trataría como nueva llamada → segundo INICIO DE LLAMADA falso.
            # Esta pre-pasada detecta los teléfonos "mid-call" para que PASO 4 los suprima.
            $Script:MidCallPhones = @{}
            $ArchivosLogPre = Get-ChildItem -Path $DirFinal -Filter "EndpointLog.txt*" -EA SilentlyContinue | Sort-Object { if ($_.Name -match "\.(\d+)$") { [int]$matches[1] } else { -1 } } -Descending
            if ($ArchivosLogPre) {
                foreach ($ArchivoP in $ArchivosLogPre) {
                    $LineasPre = Get-Content -Path $ArchivoP.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                    if (-not $LineasPre) { continue }
                    # CRÍTICO: reiniciar por archivo. Los IDs de sesión (ej: 38) se reutilizan entre
                    # archivos rotados y también dentro del mismo archivo cuando una sesión termina
                    # y el ID se recicla para otra llamada. Sin reinicio, un teléfono de una sesión
                    # anterior quedaría en $_SesionPrimerFono y haría que el teléfono legítimo de
                    # la siguiente sesión con mismo ID se marcara falsamente como mid-call update.
                    $_SesionPrimerFono = @{}
                    foreach ($lineaP in $LineasPre) {
                        # Cuando una sesión termina, liberar su entrada para que si el ID se reutiliza
                        # en la misma sesión del archivo, empiece limpio.
                        if ($lineaP -match "ProcessSessionEndedEvent: Entry\. connectinoId = (\d+)") {
                            $_SesionPrimerFono.Remove($matches[1]) | Out-Null
                            continue
                        }
                        # Solo procesar líneas del día analizado con UpdateHistoryRecord que tenga número
                        if ($lineaP -match "^\[?$FechaOmni" -and
                            $lineaP -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*([^.]+)\.") {
                            $pSes = $matches[1]; $pFon = $matches[2].Trim()
                            if ($pFon -eq "") { continue }
                            $pFonN = ($pFon -replace '^\+','') -replace '^9(\d{10,})$','$1'
                            if (-not $_SesionPrimerFono.ContainsKey($pSes)) {
                                $_SesionPrimerFono[$pSes] = $pFonN          # primer teléfono para esta sesión activa
                            } elseif ($_SesionPrimerFono[$pSes] -ne $pFonN) {
                                # Misma sesión activa recibe un número diferente: es actualización mid-call
                                $Script:MidCallPhones[$pFonN] = $true
                            }
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
                    $HoraLimpia = ""; $MsLimpio = "000"; $_loginBeginSlot = ""
                    foreach ($linea in $LineasX) {
                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2})(?:,(\d{3}))?") {
                            $HoraLimpia = $matches[1]; $MsLimpio = if ($matches[2]) { $matches[2] } else { "000" }
                            # Buscador de vacíos: registra CUALQUIER línea con timestamp (AvayaOneX.log), sea o
                            # no relevante para el timeline — el punto es medir el "pulso" real del proceso.
                            if ($HoraLimpia -match '^(\d{2}):(\d{2}):(\d{2})$') { $Script:TsProceso.Add(([int64]$matches[1])*3600000L + ([int64]$matches[2])*60000L + ([int64]$matches[3])*1000L + [int64]$MsLimpio) }

                            # --- Intento de desfirme (LogoutAgentHandler) que puede FALLAR ---
                            # "LogoutAgent:session=...;code=ReasonCode[10]" es la solicitud (misma línea que ya
                            # capturaba el bucket de reason codes más abajo como "SISTEMA_LOGOUT|" en el slot
                            # BASE $HoraLimpia). Si unos segundos después llega "Session_LogoutAgent failed"
                            # (Avaya rechazó el logout — típicamente "Call not disconnected" por una llamada
                            # activa/entrante justo en ese instante), se marca ese MISMO slot base como fallido,
                            # para que el render muestre el fallo real en vez de "Asesor se desfirma". OJO: la
                            # línea del fallo es continuación de un stack trace y NO trae timestamp propio — su
                            # detección vive MÁS ABAJO, fuera del filtro de fecha (mismo patrón que ya usa
                            # "Session_LoginAgent failed"), no aquí.
                            if ($linea -match "(?i)LogoutAgent:session=.*?;code=ReasonCode\[10\]") {
                                $UltimoSlotDesfirmeSolicitado = $HoraLimpia
                                $DesfirmeEnProceso = $true
                                $DesfirmeYaConfirmado = $false
                                $DesfirmeEsperandoConfirmacion = $true
                                $CandidatoSlot565 = ""   # descarta cualquier candidato pendiente del mecanismo 2 (Ronda siguiente)
                                # Fila propia con ms (Pablo, caso 08/09/2026): el mecanismo interno de logout
                                # también dispara una llamada "fantasma" a la extensión 565 (ver Ronda 3) — si
                                # el EndpointLog crea su propio "INICIO DE LLAMADA" en el MISMO slot base (bare-
                                # hour) que esta solicitud, esa clasificación puede ganarle en el render y
                                # "Asesor tratando de desfirmarse" nunca se ve. Se le da su propia fila ms,
                                # inmune a esa colisión — mismo patrón ya usado para las otras 3 filas de
                                # desfirme (confirmación, fallo, residuo de PendingAux).
                                $SlotDesfirmeSolicitud = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotDesfirmeSolicitud
                                if ($EventosTiempo[$SlotDesfirmeSolicitud].Interpretacion -eq "") {
                                    $EventosTiempo[$SlotDesfirmeSolicitud].Interpretacion      = "Asesor tratando de desfirmarse"
                                    $EventosTiempo[$SlotDesfirmeSolicitud].ColorInterpretacion = [System.Drawing.Color]::LightSkyBlue
                                    $EventosTiempo[$SlotDesfirmeSolicitud].RawInterpretacion  += "$linea`n"
                                }
                            }
                            # --- Intento de desfirme (Mecanismo 3: botón favorito "Desfirmarse", Pablo 09/09/2026) ---
                            # Este botón (tipo "abrv-dial", marca el FAC 565 por su cuenta) nunca pasa por
                            # LogoutAgentHandler ni deja la línea "LogoutAgent:...code=ReasonCode[10]" de arriba —
                            # Avaya lo resuelve solo. La única señal inequívoca es "Feature : DESFIRMARSE" (el
                            # nombre interno del botón), pero aparece 2 veces por cada uso: al presionarlo Y al
                            # soltarlo — el guard "-not $DesfirmeEsperandoConfirmacion" evita disparar 2 veces.
                            # Comparte la MISMA confirmación real que ya captura el bloque de abajo (con
                            # $DesfirmeEsperandoConfirmacion activa) — no hace falta nada adicional para eso.
                            if ($linea -match "(?i)Feature\s*:\s*DESFIRMARSE" -and -not $DesfirmeEsperandoConfirmacion) {
                                $DesfirmeEnProceso = $true
                                $DesfirmeYaConfirmado = $false
                                $DesfirmeEsperandoConfirmacion = $true
                                $CandidatoSlot565 = ""   # descarta cualquier candidato pendiente del mecanismo 2 (Ronda siguiente)
                                $SlotDesfirmeSolicitud = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotDesfirmeSolicitud
                                if ($EventosTiempo[$SlotDesfirmeSolicitud].Interpretacion -eq "") {
                                    $EventosTiempo[$SlotDesfirmeSolicitud].Interpretacion      = "Asesor tratando de desfirmarse"
                                    $EventosTiempo[$SlotDesfirmeSolicitud].ColorInterpretacion = [System.Drawing.Color]::LightSkyBlue
                                    $EventosTiempo[$SlotDesfirmeSolicitud].RawInterpretacion  += "$linea`n"
                                }
                            }
                            if ($linea -match "GUI Method ENDED: LogoutAgentHandler") {
                                # El intento concluyó (haya fallado o no) — deja de "proteger" el próximo
                                # PendingAux como residuo de desfirme; uno nuevo solo empieza con otra solicitud.
                                # OJO: esto SIEMPRE llega antes (o junto con) la confirmación real que sigue —
                                # NO limpia $DesfirmeEsperandoConfirmacion (ver declaración de esa bandera).
                                $DesfirmeEnProceso = $false
                            }
                            # --- Confirmación REAL del desfirme (Pablo, 09/09/2026) ---
                            # A diferencia de la solicitud (que puede fallar o quedar "PendingAux" si hay una
                            # llamada activa — ver arriba, ese caso NUNCA llega a "newState=LoggedOut" mientras
                            # $DesfirmeEsperandoConfirmacion sigue activa: se queda atorado en PendingAux, y esa
                            # bandera se apaga en el fallo — ver fuera del filtro de fecha, más abajo), este es
                            # el primer "...;newState=LoggedOut" que SÍ llega durante un intento en curso: la
                            # confirmación real de que el asesor quedó desfirmado. Se le da su propia fila ms
                            # (no la del segundo compartido de la solicitud/ENDED) y solo se crea UNA vez por
                            # intento — Avaya a veces parpadea LoggedOut→Aux→LoggedOut como limpieza interna al
                            # cerrar sesión; solo la PRIMERA cuenta como confirmación.
                            if ($linea -match "(?i)oldState\s*=\s*\w+\s*;\s*newState\s*=\s*LoggedOut" -and $DesfirmeEsperandoConfirmacion -and -not $DesfirmeYaConfirmado) {
                                $SlotDesfirmeConfirmado = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotDesfirmeConfirmado
                                if ($EventosTiempo[$SlotDesfirmeConfirmado].Interpretacion -eq "") {
                                    $EventosTiempo[$SlotDesfirmeConfirmado].Interpretacion      = "Asesor ya se encuentra desfirmado"
                                    $EventosTiempo[$SlotDesfirmeConfirmado].ColorInterpretacion = [System.Drawing.Color]::LightSkyBlue
                                    $EventosTiempo[$SlotDesfirmeConfirmado].RawInterpretacion  += "[DESFIRME CONFIRMADO] $linea`n"
                                }
                                $DesfirmeYaConfirmado = $true
                                $DesfirmeEsperandoConfirmacion = $false
                            }

                            # --- Intento de desfirme (Mecanismo 2: marcado manual "565"[+dígito], Pablo 09/09/2026) ---
                            # Este mecanismo tampoco pasa por LogoutAgentHandler ni por el botón favorito
                            # "DESFIRMARSE": el asesor disca a mano el FAC de logout (565) desde el teclado del
                            # teléfono, normalmente seguido de un dígito más (prueba controlada de Pablo: "565"+
                            # "3"). No hay ninguna línea de "intento" limpia como en los otros 2 mecanismos — el
                            # número se arma dígito por dígito en el propio texto de la llamada saliente ("Call
                            # updated : Id=N,...RemoteParty=[565,]..." → "...RemoteParty=[5653,]..." un rato
                            # después). Se rastrea así: en cuanto el número marcado llega EXACTO a "565" en una
                            # llamada saliente, se guarda su Id Y ya se marca como candidato (cubre el caso de
                            # que el agente marque solo "565" sin dígito extra — esto reemplaza/unifica el
                            # mecanismo viejo que vivía en el render vía EndpointLog, ver Ronda 8); si esa MISMA
                            # llamada se extiende con un dígito más (565+N), se refina el candidato a ESE instante
                            # (más preciso: ahí es donde de verdad se completó el FAC+código). Si dentro de los
                            # 10s siguientes llega una confirmación real (oldState=...;newState=LoggedOut) que el
                            # mecanismo 1/3 no esté ya cubriendo, se imprimen AMBOS eventos retroactivamente: el
                            # intento (en el instante del candidato) y la confirmación (en el instante del
                            # LoggedOut). Si no llega en esa ventana, no se imprime nada — pudo ser una llamada
                            # real a una extensión que solo empieza con "565".
                            if ($linea -match "(?i)Call (?:updated|created)\s*:\s*Id=(\d+),.*?RemoteParty=\[(565\d*),\]Outgoing=True" -and -not $DesfirmeEsperandoConfirmacion) {
                                $IdLlamada565 = $matches[1]; $DigitosMarcados565 = $matches[2]
                                if ($DigitosMarcados565 -eq "565") { $CandidatoIdLlamada565 = $IdLlamada565 }
                                if ($IdLlamada565 -eq $CandidatoIdLlamada565) {
                                    $CandidatoSlot565  = "$HoraLimpia,$MsLimpio"
                                    $CandidatoMs565    = & $MsDeSlot $CandidatoSlot565
                                    $CandidatoLinea565 = $linea
                                }
                            }
                            if ($linea -match "(?i)oldState\s*=\s*\w+\s*;\s*newState\s*=\s*LoggedOut" -and -not $DesfirmeEsperandoConfirmacion -and $CandidatoSlot565 -ne "") {
                                $MsAhora565 = & $MsDeSlot "$HoraLimpia,$MsLimpio"
                                if ($CandidatoMs565 -ge 0 -and $MsAhora565 -ge 0 -and ($MsAhora565 - $CandidatoMs565) -ge 0 -and ($MsAhora565 - $CandidatoMs565) -le 10000) {
                                    Init-Hora $CandidatoSlot565
                                    if ($EventosTiempo[$CandidatoSlot565].Interpretacion -eq "") {
                                        $EventosTiempo[$CandidatoSlot565].Interpretacion      = "Asesor tratando de desfirmarse"
                                        $EventosTiempo[$CandidatoSlot565].ColorInterpretacion = [System.Drawing.Color]::LightSkyBlue
                                        $EventosTiempo[$CandidatoSlot565].RawInterpretacion  += "$CandidatoLinea565`n"
                                    }
                                    $SlotConfirm565 = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotConfirm565
                                    if ($EventosTiempo[$SlotConfirm565].Interpretacion -eq "") {
                                        $EventosTiempo[$SlotConfirm565].Interpretacion      = "Asesor ya se encuentra desfirmado"
                                        $EventosTiempo[$SlotConfirm565].ColorInterpretacion = [System.Drawing.Color]::LightSkyBlue
                                        $EventosTiempo[$SlotConfirm565].RawInterpretacion  += "[DESFIRME CONFIRMADO] $linea`n"
                                    }
                                }
                                $CandidatoSlot565 = ""; $CandidatoMs565 = -1; $CandidatoIdLlamada565 = ""; $CandidatoLinea565 = ""
                            }

                            # --- Auxiliar solicitado CON LLAMADA ACTIVA: estado "PendingAux" (Punto 2, Pablo 28/08/2026) ---
                            # Independiente de la cadena elseif de abajo (igual que el contador de vacíos arriba):
                            # necesita ver TODAS las líneas relevantes, no solo la que "gane" la clasificación de esa línea.
                            if ($linea -match "(?i)newState\s*=\s*PendingAux" -and $DesfirmeEnProceso) {
                                # Este PendingAux es residuo de un intento de desfirme que Avaya aún no resuelve
                                # (con llamada activa) — NO es un Auxiliar elegido por el asesor. No crear la fila
                                # de "se detectó un auxiliar sin código"; se etiqueta correctamente al resolverse
                                # más abajo (oldState=LoggedOut;newState=Aux).
                                $PendingAuxEsResiduoDesfirme = $true
                            }
                            elseif ($linea -match "(?i)newState\s*=\s*PendingAux" -and $UltimoSlotAuxClicMs -ne "") {
                                # Confirma que el clic de hace un momento (EnterAuxWithReasonCodeHandler ENDED) NO se
                                # aplicó de inmediato — reetiquetar esa fila de "confirmado" a "pendiente".
                                if ($EventosTiempo.ContainsKey($UltimoSlotAuxClicMs) -and $EventosTiempo[$UltimoSlotAuxClicMs].Interpretacion -match "\(Confirmado por clic\)$") {
                                    $EventosTiempo[$UltimoSlotAuxClicMs].Interpretacion = $EventosTiempo[$UltimoSlotAuxClicMs].Interpretacion -replace "\(Confirmado por clic\)$", "(clic registrado — pendiente, la llamada seguía activa)"
                                    $EventosTiempo[$UltimoSlotAuxClicMs].RawInterpretacion += "[PENDINGAUX] $linea`n"
                                }
                                $UltimoSlotAuxClicMs = ""   # ya reetiquetado, no repetir en próximas líneas PendingAux del mismo tramo
                            }
                            elseif ($linea -match "(?i)newState\s*=\s*PendingAux") {
                                # TrabAux (botón favorito) elegido CON llamada activa: a diferencia del botón de la
                                # esquina superior, este NUNCA pasa por EnterAuxWithReasonCodeHandler ni deja un
                                # "Enter Aux;code=..." — solo "aux-work Button transitioned...to wink" seguido
                                # directo de newState=PendingAux. Sin este bloque, el clic no dejaba NINGÚN rastro
                                # visible (ni siquiera al colgar se sabía que fue un clic real en ese instante) —
                                # reportado por Pablo, log LogsAvaya-172.18.224.150_183848_a_184126_28082026.csv.
                                $SlotAuxFavPend = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotAuxFavPend
                                if ($EventosTiempo[$SlotAuxFavPend].Interpretacion -eq "") {
                                    # Ronda siguiente (Pablo): se quita el "Detectado [X]" — un caso real mostró que
                                    # la inferencia (último motivo elegido) puede no coincidir con lo que Avaya CMS
                                    # aplicó de verdad. Solo se reporta el hecho, sin adivinar el motivo.
                                    $EventosTiempo[$SlotAuxFavPend].Interpretacion      = "Asesor se cambia a Auxiliar — se detectó un auxiliar sin código (clic registrado — pendiente, la llamada seguía activa)"
                                    $EventosTiempo[$SlotAuxFavPend].ColorInterpretacion = [System.Drawing.Color]::Orange
                                    $EventosTiempo[$SlotAuxFavPend].RawInterpretacion  += "[PENDINGAUX-FAVORITO] $linea`n"
                                }
                            }
                            if ($linea -match "(?i)oldState\s*=\s*PendingAux\s*;\s*newState\s*=\s*LoggedOut") {
                                $UltimoPendingAuxLogoutSlot = "$HoraLimpia,$MsLimpio"
                            }
                            if ($linea -match "(?i)oldState\s*=\s*LoggedOut\s*;\s*newState\s*=\s*Aux" -and $UltimoPendingAuxLogoutSlot -ne "") {
                                $msPend = & $MsDeSlot $UltimoPendingAuxLogoutSlot
                                $msNow  = & $MsDeSlot "$HoraLimpia,$MsLimpio"
                                if ($msPend -ge 0 -and $msNow -ge 0 -and ($msNow - $msPend) -ge 0 -and ($msNow - $msPend) -le 3000) {
                                    # Esta es la finalización REAL, diferida, del Auxiliar que se pidió durante la
                                    # llamada — dale su propia fila ms (en vez de la fila de segundo compartida) para
                                    # que ordene bien contra el FIN DE LLAMADA que la disparó (Punto 3, mismo caso).
                                    $SlotAuxDiferido = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotAuxDiferido
                                    if ($EventosTiempo[$SlotAuxDiferido].Interpretacion -eq "") {
                                        if ($PendingAuxEsResiduoDesfirme) {
                                            # No es un Auxiliar real: es la limpieza del PendingAux que dejó un
                                            # intento de desfirme fallido (Pablo, caso 01/09/2026) — la llamada
                                            # activa impidió el logout, y al colgar Avaya simplemente libera al
                                            # asesor de vuelta a Aux (no a un motivo elegido).
                                            $EventosTiempo[$SlotAuxDiferido].Interpretacion      = "Se libera el estado — el intento de desfirme anterior no se completó (la llamada lo impidió)"
                                            $EventosTiempo[$SlotAuxDiferido].ColorInterpretacion = [System.Drawing.Color]::LightSkyBlue
                                            $EventosTiempo[$SlotAuxDiferido].RawInterpretacion  += "[DESFIRME-RESIDUO LIBERADO] $linea`n"
                                        } else {
                                            $NombreDif = if ($UltimoMotivoElegidoP4 -ne "" -and $DictRC.ContainsKey($UltimoMotivoElegidoP4)) { $DictRC[$UltimoMotivoElegidoP4] } else { "" }
                                            # Pablo (ronda siguiente): "(solicitud pendiente aplicada al colgar)" sonaba
                                            # poco firme. Cambiado a "Auxiliar pendiente aplicado [X]" — más directo.
                                            # OJO: el dedup del render busca este texto ("Auxiliar pendiente aplicado",
                                            # igual que "Se libera el estado" arriba) para reconocer esta fila como
                                            # hermana ms y no duplicar el evento en la fila de segundo compartido — si
                                            # se vuelve a cambiar la redacción, actualizar también esa búsqueda
                                            # (~línea con "se cambia a Auxiliar|se cambia a Default").
                                            $EventosTiempo[$SlotAuxDiferido].Interpretacion      = if ($NombreDif -ne "") { "Auxiliar pendiente aplicado [$NombreDif]" } else { "Auxiliar pendiente aplicado" }
                                            $EventosTiempo[$SlotAuxDiferido].ColorInterpretacion = [System.Drawing.Color]::Orange
                                            $EventosTiempo[$SlotAuxDiferido].RawInterpretacion  += "[PENDINGAUX->AUX] $linea`n"
                                        }
                                    }
                                }
                                $UltimoPendingAuxLogoutSlot = ""   # consumida, no reusar en próximas líneas
                                $PendingAuxEsResiduoDesfirme = $false   # consumido, resetear para el próximo PendingAux normal
                            }

                            # --- Caso "logeo autónomo": señales de PASO 4 (captura, se emiten en barrido post-PASO5) ---
                            # Independiente de la cadena elseif de abajo: solo guarda slots, no pinta filas aquí.
                            # Cierre de la app (clic en la X / apagado del OneX):
                            if ($linea -match "GUI Method (?:STARTED|ENDED): ExitHandler" -or $linea -match "Begin Executing method Shutdown\(\)" -or $linea -match "PhoneService shutdown") {
                                # Tipo: "inicio" = clic de cerrar (puede repetirse si la app está colgada);
                                #       "fin"    = apagado definitivo (aquí termina el episodio de cierre).
                                $tipoC = if ($linea -match "PhoneService shutdown" -or $linea -match "Begin Executing method Shutdown\(\)") { "fin" }
                                         elseif ($linea -match "GUI Method STARTED: ExitHandler") { "inicio" } else { "otro" }
                                $CierreAppList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio"; Raw = $linea; Tipo = $tipoC }
                            }
                            # Re-firma automática tras recuperación (solo se USA si hay bLInkRecovery=1 cercano; ver barrido):
                            if ($linea -match "Begin Executing method LoginAgent\(" -or $linea -match "AgentStateChanged:\s*oldState=LoggedOut;\s*newState=Aux") {
                                $ReconRefirmaList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio"; Raw = $linea }
                            }
                            # Señales de "asesor firmado al cerrar": EnterAuxHandler (la app lo parquea) o una transición
                            # a LoggedOut desde un estado activo. En el barrido se exige que ocurran AL/DESPUÉS del cierre
                            # (si el desfirme fue antes = menú Cerrar sesión = cierre normal).
                            if ($linea -match "GUI Method STARTED: EnterAuxHandler" -or ($linea -match "AgentStateChanged:\s*oldState=(\w+);\s*newState=LoggedOut" -and $matches[1] -ne "LoggedOut")) {
                                $CierreFirmadoList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio"; Raw = $linea }
                            }
                            # CONTESTARON (saliente): la VoiceInteraction recibe la dirección del otro lado =
                            # el instante en que "me contestaron". Coincide al ms con Alerting→Active. Solo la
                            # PRIMERA vez por VI (luego re-loguea). El barrido post-PASO5 lo separa del marcado.
                            if ($linea -match "VoiceInteractionImpl\[(VI\d+):[0-9a-fA-F\-]+,cxt=([0-9a-fA-F\-]{36})\]\.RemoteAddress=(\d{3,})") {
                                $viRA = $matches[1]; $cxtRA = $matches[2]; $numRA = $matches[3]
                                # La llave DEBE incluir el cxt (GUID de la llamada, único e irrepetible): los
                                # números de VoiceInteraction (VI3, VI6…) REINICIAN en cada login, así que
                                # keyear solo por VI hacía que un "VI3" de la mañana (p.ej. el marcado FAC de
                                # firma) descartara el "VI3" real de la tarde y esa llamada se quedara sin
                                # su fila "Contestaron".
                                $claveRA = "$viRA|$cxtRA"
                                if (-not $RAvistoVI.ContainsKey($claveRA)) {
                                    $RAvistoVI[$claveRA] = $true
                                    $ContestoSaliente += [pscustomobject]@{ Vi = $viRA; Cxt = $cxtRA; Slot = "$HoraLimpia,$MsLimpio"; Num = $numRA }
                                }
                            }
                            # Señal de corroboración: la llamada pasó de timbrando a activa (= contestaron de verdad).
                            if ($linea -match "CheckScreenPop_impl\(\):\s*previousState = Alerting;\s*currentState = Active") {
                                $AlertingActivaSlots += "$HoraLimpia,$MsLimpio"
                            }
                            # Respaldo saliente interna: "Call updated : Id=X,...,State=Active,InnerState=CONNECTED,...Outgoing=True"
                            # (sin fase de Alerting previa en el log). Solo se guarda la 1ª vez por sesión.
                            if ($linea -match "Call updated\s*:\s*Id=(\d+),.*State=Active,InnerState=CONNECTED.*Outgoing=True") {
                                $sesCA = $matches[1]
                                if (-not $ConexionActivaSaliente.ContainsKey($sesCA)) { $ConexionActivaSaliente[$sesCA] = "$HoraLimpia,$MsLimpio" }
                            }
                            # FAC del botón favorito AUTO-IN: saliente al $FacAutoIn (RemoteParty=[562,]Outgoing=True).
                            # Registra su slot para que el render marque el "Disponible" cercano como "botón favorito".
                            if ($linea -match ("RemoteParty=\[" + [regex]::Escape($FacAutoIn) + ",\]Outgoing=True")) {
                                $AutoInFacSlots += "$HoraLimpia,$MsLimpio"
                            }
                            # Ventana del método Login() (registro de estación). OJO: "Login(?!Agent)" excluye el
                            # LoginAgent (firma real, que va después). El End cierra la ventana; si faltara, el
                            # Begin del LoginAgent la cierra como red de seguridad (siempre ocurre tras el Login()).
                            if ($linea -match "Begin Executing method Login(?!Agent)") {
                                $_loginBeginSlot = "$HoraLimpia,$MsLimpio"
                            }
                            elseif ($linea -match "End Executing method Login(?!Agent)|Begin Executing method LoginAgent") {
                                if ($_loginBeginSlot -ne "") { $LoginWins += [pscustomobject]@{ B = $_loginBeginSlot; E = "$HoraLimpia,$MsLimpio" }; $_loginBeginSlot = "" }
                            }

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
                            # --- Modo de contestación (auto/manual), indexado por UUID del VI ---
                            # AUTO  = "Auto Accepting"          (el CM auto-contestó).
                            # MANUAL= "AnswerVoiceInteraction"  (efecto del clic en AnswerCallHandler, que
                            #         llega 1 ms después y SÍ trae el UUID — el GUI handler no lo trae).
                            # Son mutuamente excluyentes: una llamada trae uno u otro, nunca ambos.
                            elseif ($linea -match "(?i)(AnswerVoiceInteraction|Auto Accepting).*?VI\d+:([0-9a-fA-F\-]{36})") {
                                $AccionDetectada = $matches[1]; $IdVoice = $matches[2]
                                $Script:ModoContestacion[$IdVoice] = if ($AccionDetectada -like "*Answer*") { "MANUAL" } else { "AUTO" }
                                # La llamada SÍ fue aceptada (auto o manual) → descartar candidato "sin botones"
                                $ViSinBotones.Remove($IdVoice) | Out-Null
                            }
                            # --- Disponible via Auto-In (sin clic GUI, no hay EnterReadyHandler) ---
                            # Patrón: Invoke(fnu=auto-in...state=off) = el agente activa auto-in (sale de Aux → Ready)
                            # state=off: la bandera pasa de OFF a ON (= activar auto-in = ponerse disponible)
                            elseif ($linea -match "(?i)Begin Executing method Invoke\(fnu=auto-in.*?state=off") {
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                                $EventosTiempo[$HoraLimpia].Aux += "AUTOIN_DISP_SLOT|"
                                $SlotAutoIn = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotAutoIn
                                if ($EventosTiempo[$SlotAutoIn].Interpretacion -eq "") {
                                    $EventosTiempo[$SlotAutoIn].Interpretacion      = "Asesor se cambia a Disponible usando botón favorito"
                                    $EventosTiempo[$SlotAutoIn].ColorInterpretacion = [System.Drawing.Color]::Yellow
                                    $EventosTiempo[$SlotAutoIn].RawInterpretacion  += "$linea`n"
                                }
                            }
                            # --- GUI Ready/Aux confirmado ---
                            elseif ($linea -match "GUI Method (STARTED|ENDED): EnterReadyHandler") {
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                                if ($matches[1] -eq "ENDED") {
                                    # Marcar slot de segundos para compatibilidad con el render loop
                                    $EventosTiempo[$HoraLimpia].Aux += "GUI_READY_CONFIRMADO|"
                                    # Crear slot ms-precisión propio: así siempre aparece en su propia fila
                                    # aunque coincida con "señal de llamada" u otro evento en el mismo segundo.
                                    # Los GUI Method son el evento más confiable del clic del agente.
                                    $SlotDisp = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotDisp
                                    if ($EventosTiempo[$SlotDisp].Interpretacion -eq "") {
                                        $EventosTiempo[$SlotDisp].Interpretacion      = "Asesor se cambia a Disponible (Confirmado por clic)"
                                        $EventosTiempo[$SlotDisp].ColorInterpretacion = [System.Drawing.Color]::Yellow
                                        $EventosTiempo[$SlotDisp].RawInterpretacion  += "$linea`n"
                                    }
                                }
                            }
                            elseif ($linea -match "GUI Method (STARTED|ENDED): EnterAuxWithReasonCodeHandler") {
                                Init-Hora $HoraLimpia
                                if ($matches[1] -eq "STARTED") {
                                    # Opción B: abrir la ventana + crear su propia fila ms, igual que EnterReadyHandler
                                    # ya hace con "Disponible" — inmune a que un parpadeo de estado en el mismo
                                    # segundo (ver Opción A arriba) contamine la fila y se pierda el reason code.
                                    $SlotAuxConMotivo = "$HoraLimpia,$MsLimpio"; $CodigoRCPendiente = ""
                                    Init-Hora $SlotAuxConMotivo
                                } else {
                                    $EventosTiempo[$HoraLimpia].Aux += "GUI_AUX_CONFIRMADO|"
                                    if ($SlotAuxConMotivo -ne "" -and $CodigoRCPendiente -ne "") {
                                        $NombreRCSlot = if ($DictRC.ContainsKey($CodigoRCPendiente)) { $DictRC[$CodigoRCPendiente] } else { $CodigoRCPendiente }
                                        if ($EventosTiempo[$SlotAuxConMotivo].Interpretacion -eq "") {
                                            $EventosTiempo[$SlotAuxConMotivo].Interpretacion      = if ($CodigoRCPendiente -eq "0") { "Asesor se cambia a Default (Confirmado por clic)" } else { "Asesor se cambia a Auxiliar [$NombreRCSlot] (Confirmado por clic)" }
                                            $EventosTiempo[$SlotAuxConMotivo].ColorInterpretacion = [System.Drawing.Color]::Orange
                                            $EventosTiempo[$SlotAuxConMotivo].RawInterpretacion  += "[CONFIRMADO POR CLIC] EnterAuxWithReasonCodeHandler + code=$CodigoRCPendiente ($NombreRCSlot)`n"
                                        }
                                        if ($CodigoRCPendiente -ne "0") { $UltimoMotivoElegidoP4 = $CodigoRCPendiente }
                                        # Punto 2 (Pablo, 28/08/2026): si la llamada seguía activa, Avaya NO aplica el
                                        # cambio de inmediato — lo deja "PendingAux" (se detecta unas líneas después,
                                        # ver el "if" independiente más abajo). Se recuerda esta fila para poder
                                        # reetiquetarla de "(Confirmado por clic)" a "(clic registrado — pendiente,
                                        # la llamada seguía activa)" si de verdad resulta ser el caso.
                                        $UltimoSlotAuxClicMs = $SlotAuxConMotivo
                                    }
                                    # Cerrar la ventana siempre (con o sin código encontrado) para no arrastrarla a
                                    # un próximo EnterAux suelto que no venga de este clic.
                                    $SlotAuxConMotivo = ""; $CodigoRCPendiente = ""
                                }
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            elseif ($linea -match "(?i)GUI Method (STARTED|ENDED): (?:TransferCallHandler|TransferHandler|ConsultationHandler|CompleteTransferHandler)") {
                                # TransferCallHandler = el clic real del asesor en "Transferir" (log moderno).
                                # Capturarlo es clave: alimenta $UltimaTransfGUIHora, que distingue transferencia
                                # MANUAL (con GUI) de AUTOMÁTICA (sin GUI) en OnRequestTransferSession (~1848).
                                $UltimaTransfGUIHora = $HoraLimpia
                                # Segundo del clic (solo STARTED): la fila visible "Asesor presiona botón Transferir"
                                # se pinta en PASO 5 cuando OnRequestTransferSession cae en este mismo segundo.
                                if ($matches[1] -eq "STARTED") { $TransfClicSeg[$HoraLimpia] = $true }
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- AddCallHandler: el asesor agregó una 2ª llamada (CONSULTA para transferir/conferenciar) ---
                            elseif ($linea -match "GUI Method (STARTED|ENDED): AddCallHandler") {
                                if ($matches[1] -eq "STARTED") {
                                    $SlotAdd = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotAdd
                                    if ($EventosTiempo[$SlotAdd].Agente -eq "") {
                                        $EventosTiempo[$SlotAdd].Agente      = "$symArr El asesor agregó una llamada (consulta)"
                                        $EventosTiempo[$SlotAdd].ColorAgente = [System.Drawing.Color]::MediumOrchid
                                    }
                                    $EventosTiempo[$SlotAdd].RawAgente += "$linea`n"
                                    $AddCallSlots += $SlotAdd   # para reetiquetar el AutoHold que dispara este botón
                                }
                                Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- NewCallHandler: el asesor inició una llamada NUEVA (saliente) desde la GUI ---
                            elseif ($linea -match "GUI Method (STARTED|ENDED): NewCallHandler") {
                                if ($matches[1] -eq "STARTED") {
                                    $SlotNew = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotNew
                                    if ($EventosTiempo[$SlotNew].Agente -eq "") {
                                        $EventosTiempo[$SlotNew].Agente      = "$symUp Asesor captura número y da Enter"
                                        $EventosTiempo[$SlotNew].ColorAgente = [System.Drawing.Color]::LightSkyBlue
                                    }
                                    # Guardar el slot para, tras PASO 5, inyectar el número tecleado
                                    # (ApplyDialingRulesToNumber llega ~ms después en el EndpointLog).
                                    $NewCallSlots += $SlotNew
                                    $EventosTiempo[$SlotNew].RawAgente += "$linea`n"
                                }
                                Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- Hold/UnHold/EndCallHandler: clics GUI que CONFIRMAN la acción manual. No crean fila
                            #     (el evento real lo pinta el EndpointLog: HOLD/UNHOLD/CUELGUE MANUAL); solo registran
                            #     el segundo del clic para sellar "✓ clic confirmado" y dejan la evidencia cruda. ---
                            elseif ($linea -match "GUI Method (STARTED|ENDED): (HoldCallHandler|UnHoldCallHandler|EndCallHandler)") {
                                if ($matches[1] -eq "STARTED") {
                                    switch ($matches[2]) {
                                        "HoldCallHandler"   { $GuiHoldSeg[$HoraLimpia]   = $true }
                                        "UnHoldCallHandler" { $GuiUnholdSeg[$HoraLimpia] = $true }
                                        "EndCallHandler"    { $GuiEndSeg[$HoraLimpia]    = $true }
                                    }
                                }
                                Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAgente += "$linea`n"
                            }
                            # --- Hold ejecutado por método directo (hotkey/CTI): "Begin Executing method Hold(...ConnId=N...)".
                            #     No pasa por HoldCallHandler ni OnRequestHoldSession, así que el detector normal no lo ve.
                            #     Se registra por ConnId+slot para distinguir, en el sweep de hold implícito, un hold del
                            #     ASESOR de un auto-hold del sistema (abrir 2da línea, que NO genera este método). ---
                            elseif ($linea -match "Begin Executing method Hold\(Call\[Id=[0-9a-fA-F\-]+,ConnId=(\d+)") {
                                $cHM = $matches[1]
                                if (-not $HoldMetodoSes.ContainsKey($cHM)) { $HoldMetodoSes[$cHM] = @() }
                                $HoldMetodoSes[$cHM] += "$HoraLimpia,$MsLimpio"
                            }
                            elseif ($linea -match "GUI Method (STARTED|ENDED): ConferenceCallDragDropHandler") {
                                $SlotDD = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotDD
                                $LabelDD = if ($matches[1] -eq "STARTED") { "Inicio de Conferencia Drag/Drop" } else { "Fin de Conferencia Drag/Drop" }
                                if ($EventosTiempo[$SlotDD].Agente -eq "") {
                                    $EventosTiempo[$SlotDD].Agente = $LabelDD
                                    $EventosTiempo[$SlotDD].ColorAgente = [System.Drawing.Color]::MediumOrchid
                                }
                                $EventosTiempo[$SlotDD].RawAgente += "$linea`n"
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            elseif ($linea -match "GUI Method (STARTED|ENDED): ConferenceCallHandler$") {
                                if ($matches[1] -eq "STARTED") {
                                    $SlotConferenciaBoton = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotConferenciaBoton
                                    if ($EventosTiempo[$SlotConferenciaBoton].Agente -eq "") {
                                        $EventosTiempo[$SlotConferenciaBoton].Agente      = "Inicio de Conferencia (Botón)"
                                        $EventosTiempo[$SlotConferenciaBoton].ColorAgente = [System.Drawing.Color]::MediumOrchid
                                    }
                                    $EventosTiempo[$SlotConferenciaBoton].RawAgente += "$linea`n"
                                    Init-Hora $HoraLimpia
                                    $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                                } else {
                                    if ($SlotConferenciaBoton) { $EventosTiempo[$SlotConferenciaBoton].RawAgente += "$linea`n" }
                                    $SlotConferenciaBoton = ""
                                }
                            }
                            elseif ($SlotConferenciaBoton -and $linea -match "Begin Executing method InitiateConference\(") {
                                $EventosTiempo[$SlotConferenciaBoton].RawAgente += "$linea`n"
                            }
                            # --- ReasonCode y estados del agente (respaldo) ---
                            elseif ($linea -match "(?i)WorkServiceImpl EnterAux:session=.*?;code=(\d+)") {
                                Init-Hora $HoraLimpia; $Rc = $matches[1]
                                $EventosTiempo[$HoraLimpia].Aux += "NUEVO_RC:$Rc|"; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                                # Opción B: si este código llegó DENTRO de la ventana de un clic de "Auxiliar con
                                # motivo" (EnterAuxWithReasonCodeHandler STARTED..ENDED), recordarlo para armar la
                                # fila ms protegida al cerrar esa ventana (ver más abajo).
                                if ($SlotAuxConMotivo -ne "") { $CodigoRCPendiente = $Rc }
                            }
                            elseif ($linea -match "(?i)ReasonCode[=\[>:\s]*(\d+)") {
                                Init-Hora $HoraLimpia; $Rc = $matches[1]
                                if ($Rc -eq "0") { $EventosTiempo[$HoraLimpia].Aux += "DEFAULT|" } elseif ($Rc -eq "10") { $EventosTiempo[$HoraLimpia].Aux += "SISTEMA_LOGOUT|" } else { $EventosTiempo[$HoraLimpia].Aux += "RC: $Rc|" }
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                                if ($SlotAuxConMotivo -ne "") { $CodigoRCPendiente = $Rc }
                            }
                            # Estado genérico del agente (respaldo). OJO: usa (?<!old) para que "oldState=X;newState=Y"
                            # SIEMPRE capture Y (el estado NUEVO), nunca X — antes, como -match toma la PRIMERA
                            # coincidencia y "oldState=" aparece textualmente ANTES que "newState=" en la misma línea,
                            # esta regla capturaba el estado VIEJO. Avaya a veces dispara un parpadeo interno real
                            # (ej. Ready→LoggedOut→Aux en <20ms, todo en el mismo segundo) al cambiar de Auxiliar con
                            # motivo desde la esquina superior izquierda; con la captura vieja, ese parpadeo generaba
                            # un "ESTADO: READY|" falso que en el render (rama "ESTADO: READY") ganaba por orden de
                            # prioridad y borraba el reason code real ya capturado en la MISMA fila (caso Pablo, AUX
                            # SISTEMAS/9 desaparecido, prueba controlada 28/08/2026). Opción A. Pablo, 08/2026.
                            elseif ($linea -match "(?i)(?:UpdateSessionState:\s*Agent State\s*|AgentState:\s*|(?<!old)State\s*[:=]\s*|Enter\s+)(AuxWork|Aux|Ready|AutoIn|ManualIn|NotReady|Default|LoggedOut|PendingAux)") {
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].Aux += "ESTADO: $($matches[1].ToUpper())|"
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- Call created: registra dirección; detecta sesiones bridge del sistema (Signal A) ---
                            elseif ($linea -match "(?i)(?:PhoneService_CallCreated:call=|Call created : )Id=(\d+).*?Outgoing=(True|False)") {
                                $IDLlamada = $matches[1]; $Sentido = $matches[2]
                                # Instante en que NACE una llamada: el sweep de apertura de línea lo usa para saber
                                # si el clic abrió una línea nueva o solo cambió a una que ya tenía llamada.
                                try { $CallCreatedMs += ([int]([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).TimeOfDay.TotalSeconds) * 1000 + [int]$MsLimpio) } catch {}
                                if ($Sentido -eq "True" -and $linea -match "RemoteParty=\[,\]") {
                                    # Signal A: sesión bridge/phantom — Avaya CM crea una llamada saliente sin destino real
                                    # para ejecutar una transferencia automática; produce tono de ringback audible.
                                    # UNA sola fila por sesión: Avaya loguea la creación DOS veces ("Call created :" y
                                    # "PhoneService_CallCreated:call="). Si caen en ms distintos se creaban dos filas y el
                                    # retract (que guarda un solo $BridgeSlot) limpiaba solo la última → falso positivo.
                                    $SesionesBridge[$IDLlamada] = $HoraLimpia
                                    $DirLlamada[$IDLlamada] = "BRIDGE"
                                    if (-not $BridgeSlot.ContainsKey($IDLlamada)) {
                                        $SlotB = "$HoraLimpia,$MsLimpio"
                                        $BridgeSlot[$IDLlamada] = $SlotB
                                        Init-Hora $SlotB
                                        $EventosTiempo[$SlotB].Sesion = $IDLlamada
                                        $EventosTiempo[$SlotB].Interpretacion = "[!] Sesión bridge del sistema (auto-transferencia)"
                                        $EventosTiempo[$SlotB].ColorInterpretacion = [System.Drawing.Color]::Orange
                                    }
                                    $EventosTiempo[$BridgeSlot[$IDLlamada]].RawInterpretacion += "$linea`n"
                                } else {
                                    if ($Sentido -eq "True") { $DirLlamada[$IDLlamada] = "SALIENTE" } else { $DirLlamada[$IDLlamada] = "ENTRANTE" }
                                }
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
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR"
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
                            # --- CONFIRMACIÓN ACD: "Setting Vi.InboundAcd=true" es el veredicto REAL de que la
                            # llamada entró por la cola ACD (más confiable que el largo del número o el topic;
                            # la línea VoiceInteractionListImpl.Add que crea la señal trae inboundAcd=False, y el
                            # TRUE llega después en esta línea aparte). Reetiqueta la señal de ESE VI de
                            # INTERNA/EXTERNA → "ACD - <topic>". Funciona con o sin XML; cruza por el UUID del VI.
                            elseif ($linea -match "(?i)Setting Vi\.InboundAcd=true") {
                                Init-Hora $HoraLimpia
                                $AcdSlot = $null
                                if ($linea -match "VI=VoiceInteractionImpl\[VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})") {
                                    $AcdViUuid = $matches[1]
                                    if ($Script:AlertingHoras.ContainsKey($AcdViUuid)) {
                                        $AcdSlot = $Script:AlertingHoras[$AcdViUuid]
                                        if ($Script:AlertingHorasRedirect.ContainsKey($AcdSlot)) { $AcdSlot = $Script:AlertingHorasRedirect[$AcdSlot] }
                                    }
                                }
                                if (-not $AcdSlot -and $Script:UltimaHoraAlerting -ne $null) { $AcdSlot = $Script:UltimaHoraAlerting }
                                if ($AcdSlot -and $EventosTiempo.ContainsKey($AcdSlot) -and
                                    $EventosTiempo[$AcdSlot].Interpretacion -match "señal de llamada" -and
                                    $EventosTiempo[$AcdSlot].Interpretacion -notmatch "Entrante - ACD") {
                                    $TopicReal = if ($EventosTiempo[$AcdSlot].Topic -and $EventosTiempo[$AcdSlot].Topic.Trim() -ne "") { $EventosTiempo[$AcdSlot].Topic.Trim() }
                                                 elseif ($Script:UltimoTopic -and ($Script:UltimoTopic -notin @("Desconocido","Sin_Topic")) -and $Script:UltimoTopic.Trim() -ne "") { $Script:UltimoTopic.Trim() }
                                                 else { "" }
                                    $NuevoTagAcd = if ($TopicReal -ne "") { "ACD - $TopicReal" } else { "ACD" }
                                    $EventosTiempo[$AcdSlot].Interpretacion = $EventosTiempo[$AcdSlot].Interpretacion -replace "(INTERNA|EXTERNA) - [^)]+", $NuevoTagAcd
                                }
                                $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"; $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- TRANSFERENCIA ENTRANTE CONFIRMADA: InBoundConsultTransferCompleted ---
                            # Confirma que la llamada llegó al asesor vía transferencia consultiva. La línea no trae
                            # ConnId, pero el agente atiende una llamada a la vez → se liga al último candidato de
                            # origen de transferencia registrado por el handler CM AUTO-ANSWER SEÑAL.
                            elseif ($linea -match "(?i)InBoundConsultTransferCompleted") {
                                if ($null -ne $UltimoCandidatoTransfer) { $TransferEntranteConfirm[$UltimoCandidatoTransfer] = $true }
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- SALIENTE: Call StateChanged OldState=Ringing,NewState=Active ---
                            # Permite obtener el ConnectionId numérico de llamadas salientes para
                            # mostrarlo en el campo Sesion del INICIO DE LLAMADA (Saliente).
                            elseif ($linea -match "Call StateChanged Call\[Id=([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}),ConnectionId=(\d+),OldState=Ringing,NewState=Active\]") {
                                $CxUuidSal = $matches[1]; $CxSesIdSal = $matches[2]
                                if (-not $Script:CxtToConnId.ContainsKey($CxUuidSal)) {
                                    $Script:CxtToConnId[$CxUuidSal] = $CxSesIdSal
                                }
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- CM AUTO-ANSWER BRIDGE: Call StateChanged OldState=New,NewState=Alerting ---
                            # Con "Se requiere uso de respuesta automática CM" activo, el CM ya responde la llamada
                            # a nivel de conmutador antes de que el agente toque algo. El WI.ADD nunca llega a
                            # state=Alerting en el modelo OneXAgent (llega directo a Active), así que AlertingHoras
                            # nunca se puebla y PRIMARY_CONNECTED clasificaría como Saliente erróneamente.
                            # Este evento sí aparece siempre que el teléfono timbra. Capturamos el callUUID (Id=)
                            # que luego figura como cxt= en PRIMARY_CONNECTED para reconstruir el puente.
                            elseif ($linea -match "Call StateChanged Call\[Id=([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}),ConnectionId=(\d+),OldState=New,NewState=Alerting\]") {
                                $CxUuid = $matches[1]; $CxSesId = $matches[2]
                                # Crear señal de llamada en ms-slot (solo si WI.ADD no la creó ya en el slot base).
                                # OldState=New aparece para TODAS las llamadas ACD; WI.ADD solo cuando NO hay
                                # CM Auto-Answer activo (sin auto-answer, WI llega directo a type=Active y nunca
                                # genera type=Alerting). Si WI.ADD ya puso señal en slot base, no crear ms-slot.
                                $SlotSenal = "$HoraLimpia,$MsLimpio"
                                # Revisar todo el segundo (slot base + ms-slots): la señal del XML ahora vive en un
                                # ms-slot (tema 1), no en el slot base, así que un check solo-base crearía un duplicado.
                                $BaseYaTieneSeñal = $false
                                foreach ($kSenal in @($EventosTiempo.Keys)) {
                                    if (($kSenal -eq $HoraLimpia -or $kSenal -match "^$([regex]::Escape($HoraLimpia)),\d+$") -and $EventosTiempo[$kSenal].Interpretacion -match "señal de llamada") { $BaseYaTieneSeñal = $true; break }
                                }
                                if (-not $BaseYaTieneSeñal) {
                                    Init-Hora $SlotSenal
                                    if ($EventosTiempo[$SlotSenal].Interpretacion -notmatch "señal de llamada|INICIO DE LLAMADA") {
                                        $EventosTiempo[$SlotSenal].Interpretacion      = "Agente con señal de llamada (Entrante - ACD)"
                                        $EventosTiempo[$SlotSenal].ColorInterpretacion = [System.Drawing.Color]::Gold
                                        $EventosTiempo[$SlotSenal].RawInterpretacion  += "$linea`n"
                                        $ConnIdToSeñalSlot[$CxSesId] = $SlotSenal
                                        $DirLlamada[$CxSesId] = "ENTRANTE"
                                    }
                                }
                                # Solo registrar si WI.ADD no lo hizo ya (evita sobreescribir hora más precisa).
                                # Si creamos ms-slot señal, apuntar a él para que PRIMARY propague Tel al INICIO.
                                if (-not $Script:AlertingPorCxtUuid.ContainsKey($CxUuid)) {
                                    $Script:AlertingPorCxtUuid[$CxUuid] = if ($ConnIdToSeñalSlot.ContainsKey($CxSesId)) { $SlotSenal } else { $HoraLimpia }
                                }
                                # Mapear cxtUUID → ConnectionId numérico para que PRIMARY_CONNECTED pueda
                                # bloquear SECONDARY tardíos del mismo CallID (ej: inter-agente Active→Active
                                # que llega 20-30s después) no creen un INICIO duplicado.
                                $Script:CxtToConnId[$CxUuid] = $CxSesId
                                Init-Hora $HoraLimpia
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                                # Orden cronológico auto-in + señal mismo segundo:
                                # Si el agente se puso Disponible vía auto-in (AUTOIN_DISP_SLOT está en este segundo)
                                # Y hay señal de llamada en el slot base, la señal llegó DESPUÉS del auto-in
                                # (ej: auto-in a 579ms, llamada a 965ms). Pero el slot base "HH:mm:ss" ordena antes
                                # que cualquier slot ms "HH:mm:ss,xxx", mostrando señal primero — orden incorrecto.
                                # Fix: mover la señal a un slot ms con el timestamp exacto del alerting,
                                # así "T,965" ordena DESPUÉS de "T,579" (Disponible) → orden correcto en la grilla.
                                if ($EventosTiempo[$HoraLimpia].Aux -match "AUTOIN_DISP_SLOT" -and
                                    $EventosTiempo[$HoraLimpia].Interpretacion -match "señal de llamada") {
                                    $SlotAlerta = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotAlerta
                                    # Copiar datos de señal al slot ms (no se copia Aux/RawAux: son del estado del agente, no de la llamada)
                                    $EventosTiempo[$SlotAlerta].Interpretacion      = $EventosTiempo[$HoraLimpia].Interpretacion
                                    $EventosTiempo[$SlotAlerta].ColorInterpretacion = $EventosTiempo[$HoraLimpia].ColorInterpretacion
                                    $EventosTiempo[$SlotAlerta].Tel                 = $EventosTiempo[$HoraLimpia].Tel
                                    $EventosTiempo[$SlotAlerta].RawInterpretacion   = $EventosTiempo[$HoraLimpia].RawInterpretacion
                                    # Limpiar señal del slot base (Aux/RawAux permanecen para que PASO 5 pueda escribir sesión)
                                    $EventosTiempo[$HoraLimpia].Interpretacion      = ""
                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::White
                                    $EventosTiempo[$HoraLimpia].Tel                 = "-"
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion   = "[SEÑAL MOVIDA → $SlotAlerta (auto-in mismo segundo)]`n"
                                    # Actualizar $AlertingPorTel al slot ms
                                    $TelMovNorm = ($EventosTiempo[$SlotAlerta].Tel -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                    if ($TelMovNorm -ne "-" -and $TelMovNorm -ne "") {
                                        if ($AlertingPorTel.ContainsKey($TelMovNorm)) { $AlertingPorTel[$TelMovNorm] = $SlotAlerta }
                                    }
                                    # Actualizar $Script:AlertingPorCxtUuid al slot ms (PRIMARY_CONNECTED lo usará si AlertingHoras no está)
                                    $Script:AlertingPorCxtUuid[$CxUuid] = $SlotAlerta
                                    # Guardar redirect: segundos → ms, para que PRIMARY_CONNECTED encuentre el slot correcto
                                    # aunque $Script:AlertingHoras apunte a la clave de segundos (seteada por XML/WI.ADD antes)
                                    $Script:AlertingHorasRedirect[$HoraLimpia] = $SlotAlerta
                                }
                            }
                            # --- CM AUTO-ANSWER SEÑAL: PhoneService_CallStateChanged State=Alerting,Outgoing=False ---
                            # Aparece SIEMPRE que el teléfono timbra, independientemente de la configuración de
                            # auto-answer. Con CM Auto-Answer activo, WI.ADD no genera type=Alerting, por lo que
                            # este es el único evento que puede crear la fila "señal de llamada".
                            #
                            # Con CM Auto-Answer el CM responde en el conmutador ANTES que OneX actualice estado,
                            # por lo que VoiceInteractionImpl type=Active (PRIMARY_CONNECTED) puede llegar al log
                            # ANTES que este PhoneService State=Alerting. Si ya existe un INICIO(Saliente) en el
                            # slot (porque PRIMARY no encontró AlertingHoras y clasificó saliente), este handler
                            # lo CORRIGE a Entrante en lugar de sobreescribirlo con una señal de llamada.
                            elseif ($linea -match "(?i)PhoneService_CallStateChanged:call=Id=(\d+).*?State=Alerting.*?InnerState=INIT.*?RemoteParty=\[([^\]]*)\]Outgoing=False") {
                                $CaId = $matches[1]; $CaParty = $matches[2]
                                $CaParts = $CaParty -split ','
                                $CaNum   = if ($CaParts.Count -ge 2) { ($CaParts[1] -replace '\D','') } else { ($CaParts[0] -replace '\D','') }
                                $CaTel   = if ($CaNum) { $CaNum } else { "Desconocido" }
                                $CaTelN  = ($CaTel -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                Init-Hora $HoraLimpia
                                # Guard doble: AlertingPorTel y PhoneYaEnInicio evitan duplicados cuando
                                # este evento aparece tardíamente (tras call end) o hay eventos repetidos.
                                if (-not $AlertingPorTel.ContainsKey($CaTelN) -and -not $PhoneYaEnInicio.ContainsKey($CaTelN)) {
                                    if ($ConnIdToSeñalSlot.ContainsKey($CaId)) {
                                        # OldState=New ya creó señal en ms-slot → solo actualizar Tel (no crear duplicado en slot base).
                                        # La señal ms-slot sobrevive a ProcessSessionEndedEvent de PASO 5, que solo escribe al slot base.
                                        $SlotEx = $ConnIdToSeñalSlot[$CaId]
                                        if ($EventosTiempo.ContainsKey($SlotEx)) {
                                            $CaLabel = if ($CaTel.Length -le 6) { "INTERNA" } else { "ACD/EXTERNA" }
                                            $EventosTiempo[$SlotEx].Interpretacion      = "Agente con señal de llamada (Entrante - $CaLabel - $CaTel) [CM Auto-Answer]"
                                            $EventosTiempo[$SlotEx].ColorInterpretacion = [System.Drawing.Color]::Gold
                                            if ($EventosTiempo[$SlotEx].Tel -eq "-") { $EventosTiempo[$SlotEx].Tel = $CaTel }
                                            $EventosTiempo[$SlotEx].RawInterpretacion  += "[CM-AUTO] PhoneService_CallStateChanged State=Alerting | Tel actualizado | SessionId=$CaId | Tel=$CaTel`n"
                                        }
                                        $AlertingPorTel[$CaTelN]   = $SlotEx
                                        $Script:UltimaHoraAlerting = $HoraLimpia
                                        $DirLlamada[$CaId]         = "ENTRANTE"
                                    } elseif ($EventosTiempo[$HoraLimpia].Interpretacion -match "INICIO DE LLAMADA \(Saliente\)") {
                                        # PRIMARY_CONNECTED llegó primero y clasificó como Saliente porque AlertingHoras
                                        # aún no tenía el UUID (Call StateChanged tampoco había llegado todavía).
                                        # PhoneService confirma ahora que la llamada es Entrante → corregir dirección.
                                        $EventosTiempo[$HoraLimpia].Interpretacion = $EventosTiempo[$HoraLimpia].Interpretacion -replace "INICIO DE LLAMADA \(Saliente\)", "INICIO DE LLAMADA (Entrante) [CM Auto-Answer]"
                                        $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $CaTel }
                                        $AlertingPorTel[$CaTelN]    = $HoraLimpia
                                        $Script:UltimaHoraAlerting  = $HoraLimpia
                                        $DirLlamada[$CaId]          = "ENTRANTE"
                                        $PhoneYaEnInicio[$CaTelN]   = $true   # bloquear SECONDARY y señales futuras de este tel
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion += "[CM-AUTO] PhoneService_CallStateChanged State=Alerting | Corrección Saliente→Entrante | SessionId=$CaId | Tel=$CaTel`n"
                                    } elseif ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "señal de llamada|INICIO DE LLAMADA") {
                                        # Antes de crear señal en el slot base, revisar TODO el segundo (slot base + ms-slots).
                                        # Con CM Auto-Answer la señal REAL (ACD, con el número de cliente y el topic) suele vivir en
                                        # un ms-slot creado por el XML/WI.ADD, mientras este evento sólo trae el número de la cola/VDN
                                        # (ej: 300079 — el número del cliente llega ~1 s después en un CallUpdated), que se etiquetaría
                                        # erróneamente como INTERNA. Si ya hay una señal/INICIO en el segundo → NO duplicar: adjuntar
                                        # la evidencia al slot existente y apuntar los guards a él.
                                        # Preferir la fila de "señal de llamada"; si no hay, usar el INICIO del mismo segundo.
                                        $SlotSeñalSeg = $null; $SlotInicioSeg = $null
                                        foreach ($kSeg in @($EventosTiempo.Keys)) {
                                            if ($kSeg -eq $HoraLimpia -or $kSeg -match "^$([regex]::Escape($HoraLimpia)),\d+$") {
                                                if ($EventosTiempo[$kSeg].Interpretacion -match "señal de llamada") { $SlotSeñalSeg = $kSeg; break }
                                                elseif ($EventosTiempo[$kSeg].Interpretacion -match "INICIO DE LLAMADA") { $SlotInicioSeg = $kSeg }
                                            }
                                        }
                                        if (-not $SlotSeñalSeg) { $SlotSeñalSeg = $SlotInicioSeg }
                                        if ($SlotSeñalSeg) {
                                            # Conservar el dato [CM Auto-Answer] en la fila REAL (antes vivía sólo en la fila duplicada que se eliminó).
                                            if ($EventosTiempo[$SlotSeñalSeg].Interpretacion -match "señal de llamada|INICIO DE LLAMADA" -and $EventosTiempo[$SlotSeñalSeg].Interpretacion -notmatch "\[CM Auto-Answer\]") {
                                                $EventosTiempo[$SlotSeñalSeg].Interpretacion += " [CM Auto-Answer]"
                                            }
                                            $EventosTiempo[$SlotSeñalSeg].RawInterpretacion += "[CM-AUTO] PhoneService_CallStateChanged State=Alerting | SessionId=$CaId | Tel cola=$CaTel (señal real ya existe en el segundo, no se duplica)`n"
                                            $AlertingPorTel[$CaTelN]    = $SlotSeñalSeg
                                            $Script:UltimaHoraAlerting  = $SlotSeñalSeg
                                            $DirLlamada[$CaId]          = "ENTRANTE"
                                            # Posible ORIGEN de transferencia entrante: este evento trae un número interno (cola/ext)
                                            # distinto al de la señal real (cliente). Se confirma como transferencia sólo si más
                                            # adelante aparece InBoundConsultTransferCompleted (evita falsos positivos en ACD normal).
                                            $TelRealSeg = ($EventosTiempo[$SlotSeñalSeg].Tel -replace '\D','')
                                            if ($CaTel -ne "Desconocido" -and $CaTel.Length -le 6 -and ($CaTel -replace '\D','') -ne $TelRealSeg) {
                                                $CaNombreOrig = if ($CaParts.Count -ge 1) { (($CaParts[0] -replace '^a=','').Trim() -replace '\s+[\d.]+$','').Trim() } else { "" }
                                                $TransferEntranteOrigen[$CaId] = @{ Tel = $CaTel; Nombre = $CaNombreOrig; SlotReal = $SlotSeñalSeg }
                                                $UltimoCandidatoTransfer = $CaId
                                            }
                                        } else {
                                            # Caso fallback real: ni XML, ni OldState=New, ni WI.ADD crearon señal → crear en slot base.
                                            $CaLabel = if ($CaTel.Length -le 6) { "INTERNA" } else { "ACD/EXTERNA" }
                                            $EventosTiempo[$HoraLimpia].Interpretacion      = "Agente con señal de llamada (Entrante - $CaLabel - $CaTel) [CM Auto-Answer]"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                            if ($EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $CaTel }
                                            $AlertingPorTel[$CaTelN]    = $HoraLimpia
                                            $Script:UltimaHoraAlerting  = $HoraLimpia
                                            $DirLlamada[$CaId]          = "ENTRANTE"
                                            $EventosTiempo[$HoraLimpia].RawInterpretacion += "[CM-AUTO] PhoneService_CallStateChanged State=Alerting | SessionId=$CaId | Tel=$CaTel`n"
                                        }
                                    }
                                } elseif ($ConnIdToSeñalSlot.ContainsKey($CaId) -and $AlertingPorTel.ContainsKey($CaTelN)) {
                                    # WI.ADD ya creó señal en slot base (AlertingPorTel seteado) y OldState=New también
                                    # creó un ms-slot señal → eliminar el ms-slot duplicado para no mostrar dos filas.
                                    $SlotDup = $ConnIdToSeñalSlot[$CaId]
                                    if ($EventosTiempo.ContainsKey($SlotDup)) {
                                        $SlotDupObj = $EventosTiempo[$SlotDup]
                                        if ($SlotDupObj.Agente -eq "" -and $SlotDupObj.Audio -eq "" -and $SlotDupObj.Ispeac -eq "") {
                                            $EventosTiempo.Remove($SlotDup) | Out-Null
                                        } else {
                                            $SlotDupObj.Interpretacion = ""; $SlotDupObj.ColorInterpretacion = [System.Drawing.Color]::White
                                        }
                                    }
                                    $ConnIdToSeñalSlot.Remove($CaId) | Out-Null
                                }
                                $EventosTiempo[$HoraLimpia].RawAux += "$linea`n"
                            }
                            # --- PRIMARY CONNECTED: VoiceInteractionImpl type=Active (VI UUID → Ring Time exacto) ---
                            # Este evento es el correcto para cruzar UUID con ContactLog.xml.
                            # Formato real: VoiceInteractionImpl[VI#:UUID,cxt=UUID].StateImpl=[operations=[...],type=Active,...]
                            elseif ($linea -match "VoiceInteractionImpl\[VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}),cxt=([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\]\.StateImpl=\[.*?type=Active") {
                                $ViUuid = $matches[1]; $CxtUuid = $matches[2]
                                # CM Auto-Answer fallback: si WI.ADD nunca registró el viUUID en AlertingHoras
                                # (porque el WI nunca pasó por state=Alerting), pero Call StateChanged sí capturó
                                # el callUUID que coincide con este cxt=, usamos ese puente para clasificar ENTRANTE.
                                # Lo registramos en AlertingHoras ahora para que toda la lógica subsiguiente funcione igual.
                                if (-not $Script:AlertingHoras.ContainsKey($ViUuid) -and $Script:AlertingPorCxtUuid.ContainsKey($CxtUuid)) {
                                    $Script:AlertingHoras[$ViUuid] = $Script:AlertingPorCxtUuid[$CxtUuid]
                                }
                                $ClaveVI = "VI_$ViUuid"
                                if (-not $ConexionesYaIniciadas.ContainsKey($ClaveVI)) {
                                    $ConexionesYaIniciadas[$ClaveVI] = $true
                                    $PrimaryConnectedSeconds[$HoraLimpia] = $true  # marca O(1) para SECONDARY guard
                                    # Bloquear también el CallID numérico para que SECONDARY tardíos (ej: inter-agente
                                    # Active→Active que llega 20-30s después) no creen un INICIO duplicado.
                                    # OldState=New,NewState=Alerting ya mapeó: cxtUUID → ConnectionId numérico.
                                    if ($Script:CxtToConnId.ContainsKey($CxtUuid)) {
                                        $ConexionesYaIniciadas[$Script:CxtToConnId[$CxtUuid]] = $true
                                    }
                                    # Slot con milisegundos: INICIO siempre en fila propia, separada del Alerting
                                    $SlotInicio = "$HoraLimpia,$MsLimpio"
                                    Init-Hora $SlotInicio
                                    $EventosTiempo[$SlotInicio].ViId = $ViUuid   # liga el INICIO con su modo de contestación (ModoContestacion)
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
                                        # Guard: si la hora de alerting de este UUID ya fue "consumida" por otro
                                        # PRIMARY_CONNECTED anterior, este VoiceInteraction es un duplicado generado
                                        # por CM al actualizar el teléfono mid-call (transferencia inter-agentes).
                                        # Usar la hora de alerting como clave es correcto: UUID_A y UUID_B comparten
                                        # la misma hora de alerting (12:04:34) porque UUID_B hereda ese timestamp
                                        # del WI.ADD que reutiliza el slot original.
                                        $AlertHoraX = $Script:AlertingHoras[$ViUuid]
                                        # Redirect: si la señal fue movida a un slot ms (auto-in mismo segundo),
                                        # apuntar a la nueva clave para encontrar Tel y marcar consumed correctamente.
                                        if ($Script:AlertingHorasRedirect.ContainsKey($AlertHoraX)) {
                                            $AlertHoraX = $Script:AlertingHorasRedirect[$AlertHoraX]
                                        }
                                        $TelClearX  = ""
                                        if ($EventosTiempo.ContainsKey($AlertHoraX)) {
                                            $TelClearX = ($EventosTiempo[$AlertHoraX].Tel -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                        }
                                        $EsMidCallUpdate = $Script:AlertingHorasConsumed.ContainsKey($AlertHoraX)
                                        if (-not $EsMidCallUpdate) {
                                            # Entrante normal: INICIO siempre en su propio slot ms (nunca fusionado con señal de llamada)
                                            if ($EventosTiempo[$SlotInicio].Interpretacion -notmatch "INICIO DE LLAMADA") {
                                                $EventosTiempo[$SlotInicio].Interpretacion = "$symPlay INICIO DE LLAMADA (Entrante)$RingTimeStr"
                                            }
                                            $EventosTiempo[$SlotInicio].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                            # Propagar teléfono del slot señal al INICIO ms-slot
                                            if ($TelClearX -ne "" -and $EventosTiempo[$SlotInicio].Tel -eq "-") {
                                                $EventosTiempo[$SlotInicio].Tel = $TelClearX
                                            }
                                            # Evidencia de dirección para el diagnóstico
                                            $AlertHoraRef = $Script:AlertingHoras[$ViUuid]
                                            $EventosTiempo[$SlotInicio].RawInterpretacion += "[DIR:ENTRANTE] PRIMARY_CONNECTED: AlertingHoras[VI:$ViUuid] = señal en $AlertHoraRef`n"
                                            $EventosTiempo[$SlotInicio].RawInterpretacion += "$linea`n"
                                            # Marcar esta hora de alerting como consumida → UUID_B queda bloqueado
                                            $Script:AlertingHorasConsumed[$AlertHoraX] = $true
                                            # Absorber INICIO duplicado de SECONDARY: ocurre cuando PhoneService_CallStateChanged
                                            # State=Active,InnerState=CONNECTED dispara 1-2ms ANTES de VoiceInteractionImpl type=Active
                                            # en el mismo segundo. SECONDARY no ve el flag de PRIMARY aún y crea su propio INICIO.
                                            # Solución: copiar Tel del slot base al ms-slot y limpiar el INICIO del slot base.
                                            if ($EventosTiempo.ContainsKey($HoraLimpia) -and
                                                $EventosTiempo[$HoraLimpia].Interpretacion -match "INICIO DE LLAMADA" -and
                                                $EventosTiempo[$SlotInicio].Interpretacion -match "INICIO DE LLAMADA") {
                                                if ($EventosTiempo[$SlotInicio].Tel -eq "-" -and $EventosTiempo[$HoraLimpia].Tel -ne "-") {
                                                    $EventosTiempo[$SlotInicio].Tel = $EventosTiempo[$HoraLimpia].Tel
                                                }
                                                $EventosTiempo[$HoraLimpia].Interpretacion      = ""
                                                $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::White
                                                $EventosTiempo[$HoraLimpia].RawInterpretacion  += "[SECONDARY ABSORBIDO] INICIO duplicado eliminado — PRIMARY ms-slot: $SlotInicio`n"
                                            }
                                            # Sesión en ms-slot INICIO: el ConnectionId numérico seteado en OldState=New,NewState=Alerting
                                            # es el mismo SessionId que PASO 5 usará en UpdateHistoryRecord/ProcessSessionEndedEvent.
                                            if ($Script:CxtToConnId.ContainsKey($CxtUuid) -and $EventosTiempo[$SlotInicio].Sesion -eq "-") {
                                                $EventosTiempo[$SlotInicio].Sesion = $Script:CxtToConnId[$CxtUuid]
                                            }
                                        } else {
                                            # Mid-call update: CM creó un nuevo VoiceInteraction UUID para la misma
                                            # llamada (transferencia inter-agentes). La hora de alerting ya fue usada
                                            # por el UUID anterior → suprimir este INICIO duplicado.
                                            $EventosTiempo[$SlotInicio].RawInterpretacion += "[MID-CALL SUPRIMIDO] AlertingHora '$AlertHoraX' ya fue consumida por otro VoiceInteraction (transferencia inter-agentes). INICIO duplicado suprimido.`n"
                                            $EventosTiempo[$SlotInicio].RawInterpretacion += "$linea`n"
                                            # Si el slot creado para este INICIO está vacío, eliminarlo para no dejar fila en blanco
                                            $sObj = $EventosTiempo[$SlotInicio]
                                            if ($sObj -and $sObj.Interpretacion -eq "" -and $sObj.RawInterpretacion -match "MID-CALL SUPRIMIDO" -and
                                                $sObj.Agente -eq "" -and $sObj.Tel -eq "-" -and $sObj.Audio -eq "" -and $sObj.Ispeac -eq "") {
                                                $EventosTiempo.Remove($SlotInicio) | Out-Null
                                            }
                                        }
                                        # Alerting terminó → liberar teléfono y bloquear SECONDARY (siempre, independiente de mid-call)
                                        $AlertingPorTel.Remove($TelClearX) | Out-Null
                                        if ($TelClearX -ne "" -and $TelClearX -ne "Desconocido") { $PhoneYaEnInicio[$TelClearX] = $true }
                                    } else {
                                        # Saliente: usar slot con ms para mostrar milisegundos en el grid
                                        $EventosTiempo[$SlotInicio].Interpretacion      = "$symPlay INICIO DE LLAMADA (Saliente)"
                                        $EventosTiempo[$SlotInicio].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                        # Heredar Tel y Sesion del slot base (PASO 1 los escribió ahí)
                                        if ($EventosTiempo[$HoraLimpia].Tel -ne "-" -and $EventosTiempo[$SlotInicio].Tel -eq "-") {
                                            $EventosTiempo[$SlotInicio].Tel = $EventosTiempo[$HoraLimpia].Tel
                                        }
                                        if ($Script:CxtToConnId.ContainsKey($CxtUuid)) {
                                            $EventosTiempo[$SlotInicio].Sesion = $Script:CxtToConnId[$CxtUuid]
                                        } elseif ($EventosTiempo[$HoraLimpia].Sesion -ne "-") {
                                            $EventosTiempo[$SlotInicio].Sesion = $EventosTiempo[$HoraLimpia].Sesion
                                        }
                                        # Evidencia de dirección al slot con ms
                                        $EventosTiempo[$SlotInicio].RawInterpretacion += $EventosTiempo[$HoraLimpia].RawInterpretacion
                                        $EventosTiempo[$SlotInicio].RawInterpretacion += "[DIR:SALIENTE] PRIMARY_CONNECTED: VI UUID '$ViUuid' NO estaba en AlertingHoras`n"
                                        $EventosTiempo[$SlotInicio].RawInterpretacion += "$linea`n"
                                        # Absorber slot base: limpiar Interpretacion para que no aparezca en el grid
                                        $EventosTiempo[$HoraLimpia].Interpretacion      = ""
                                        $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::White
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion   = "[ABSORBIDO:$SlotInicio] INICIO movido al slot con ms`n"
                                    }
                                }
                            }
                            # --- BRIDGE CONNECTED: InnerState=BRIDGED_CONNECTED + RemoteParty=[,] (Signal A) ---
                            # BRIDGED_CONNECTED solo aparece en LLAMADAS BRIDGE cuando RemoteParty es vacío [,].
                            # En llamadas ACD normales, BRIDGED_CONNECTED sí aparece pero siempre con RemoteParty real.
                            # La combinación BRIDGED_CONNECTED + RemoteParty=[,] es exclusiva de sesiones bridge/phantom.
                            elseif ($linea -match "(?i)(?:PhoneService_CallStateChanged|PhoneService_CallUpdated):call=Id=(\d+).*?InnerState=BRIDGED_CONNECTED" -and $linea -match "RemoteParty=\[,\]") {
                                $BridgeId = $matches[1]
                                $SesionesBridge[$BridgeId] = $HoraLimpia
                                $DirLlamada[$BridgeId] = "BRIDGE"
                                # Una sola fila por sesión; se registra $BridgeSlot para que el retract la pueda limpiar.
                                if (-not $BridgeSlot.ContainsKey($BridgeId)) {
                                    $SlotBridge = "$HoraLimpia,$MsLimpio"
                                    $BridgeSlot[$BridgeId] = $SlotBridge
                                    Init-Hora $SlotBridge
                                    $EventosTiempo[$SlotBridge].Sesion = $BridgeId
                                    if ($EventosTiempo[$SlotBridge].Interpretacion -eq "") {
                                        $EventosTiempo[$SlotBridge].Interpretacion = "[!] Sesión bridge del sistema (auto-transferencia)"
                                        $EventosTiempo[$SlotBridge].ColorInterpretacion = [System.Drawing.Color]::Orange
                                    }
                                }
                                $EventosTiempo[$BridgeSlot[$BridgeId]].RawInterpretacion += "$linea`n"
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
                                $YaHayInicio = $false; $YaHayPrimary = $false
                                for ($i2 = 0; $i2 -le 6 -and -not $YaHayInicio; $i2++) {
                                    $tCheck2 = ($tBase2.AddSeconds(-$i2)).ToString("HH:mm:ss")
                                    if ($PrimaryConnectedSeconds.ContainsKey($tCheck2)) { $YaHayInicio = $true; $YaHayPrimary = $true }
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
                                            $TAlerta = [datetime]::ParseExact(($Script:UltimaHoraAlerting -split ',')[0], "HH:mm:ss", $null)
                                            $TConect = [datetime]::ParseExact($HoraLimpia, "HH:mm:ss", $null)
                                            $RingSeg = [int]($TConect - $TAlerta).TotalSeconds
                                            if ($RingSeg -ge 0 -and $RingSeg -le 60) { $RingTimeStr = " [Ring: ${RingSeg}s]" }
                                        } catch {}
                                    }
                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symPlay INICIO DE LLAMADA (Entrante)$RingTimeStr"
                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                                    # Evidencia de dirección para el diagnóstico
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[DIR:ENTRANTE] SECONDARY_CONNECTED: Outgoing=False, sin PRIMARY previo en ventana ±6s`n"
                                } elseif (-not $ConexionesYaIniciadas.ContainsKey($CallID) -and $YaHayPrimary) {
                                    # PRIMARY_CONNECTED actuó en la misma ventana ±6s: marcar el CallID como ya iniciado.
                                    # Propósito: cuando CM envía PhoneService_CallUpdated tardíos (ej: 20-30s después)
                                    # por una transferencia inter-agente que actualiza el RemoteParty mid-call,
                                    # esos eventos tienen el mismo CallID pero distinto teléfono y escaparían
                                    # los guards de YaHayInicio (fuera de ventana) y PhoneYaEnInicio (tel diferente).
                                    # Al marcar aquí, el primer CallUpdated bloqueado en ventana deja el CallID
                                    # protegido para todos los eventos tardíos que vengan después.
                                    $ConexionesYaIniciadas[$CallID] = $true
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[SECONDARY LOCKED] PRIMARY detectado en ventana ±6s → CallID=$CallID marcado para bloquear PhoneService_CallUpdated tardíos.`n"
                                } elseif (-not $ConexionesYaIniciadas.ContainsKey($CallID) -and
                                          $NumRawNorm -ne "Desconocido" -and $PhoneYaEnInicio.ContainsKey($NumRawNorm)) {
                                    # Fallback: el teléfono ya estaba en PhoneYaEnInicio (PRIMARY manejó esta sesión
                                    # pero PRIMARY_CONNECTED no coincidió en la misma ventana ±6s).
                                    $ConexionesYaIniciadas[$CallID] = $true
                                    $EventosTiempo[$HoraLimpia].RawInterpretacion += "[SECONDARY LOCKED por PhoneYaEnInicio] Tel '$NumRawNorm' → CallID=$CallID marcado.`n"
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
                                    # Guardar con milisegundos: el FIN del fallback hereda el ms real del
                                    # evento Disconnected (antes quedaba en slot base "sin ms", ambiguo al
                                    # combinarse con otros eventos del mismo segundo que sí traen ms).
                                    $CallStateDisconnected[$SesD] = "$HoraLimpia,$MsLimpio"
                                }
                            }
                            # --- APERTURA DE LÍNEA (FOQUITO): PressLineAppearance en OneXAgent.log ---
                            # El agente presionó el botón físico de línea del softphone (línea 2, 3...).
                            # No genera MakeCall — abre canal H.323 y pone en hold implícito la llamada activa.
                            elseif ($linea -match "Begin Executing method PressLineAppearance") {
                                $SlotPress = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotPress
                                $EventosTiempo[$SlotPress].Agente      = "↗ Apertura de línea desde botón"
                                $EventosTiempo[$SlotPress].ColorAgente = [System.Drawing.Color]::LightSkyBlue
                                $EventosTiempo[$SlotPress].RawAgente  += "$linea`n"
                            }
                            # --- FIN DEL HANDLER DE TRANSFERENCIA: marca el segundo en que el Transfer termina ---
                            # Si la transferencia falló, el sistema retoma la llamada original con un
                            # PressLineAppearance en este mismo segundo. El sweep post-PASO5 usa esto para
                            # NO confundir ese retomado automático con una apertura de línea manual del asesor.
                            elseif ($linea -match "End Executing method Transfer\b") {
                                $EndTransferSeg[$HoraLimpia] = $true
                            }
                            # --- HOLD IMPLÍCITO: OldState=Active,NewState=Inactive sin clic de Hold ---
                            # Avaya retiene automáticamente la llamada activa cuando el agente abre
                            # una segunda línea. Si EndpointLog no tiene OnRequestHoldSession para esta
                            # sesión, el sweep post-PASO5 añadirá "Hold automático del sistema".
                            elseif ($linea -match "Call StateChanged Call\[Id=[0-9a-fA-F\-]+,ConnectionId=(\d+),OldState=Active,NewState=Inactive\]") {
                                $SesH = $matches[1]
                                # LISTA de slots (antes solo el primero): navegando entre líneas la MISMA llamada se
                                # retiene varias veces y solo se pintaba el primer hold.
                                if (-not $HoldImplicito.ContainsKey($SesH)) { $HoldImplicito[$SesH] = @() }
                                $SlotHoldImp = "$HoraLimpia,$MsLimpio"
                                if ($HoldImplicito[$SesH] -notcontains $SlotHoldImp) {
                                    Init-Hora $SlotHoldImp
                                    $HoldImplicito[$SesH] += $SlotHoldImp
                                }
                            }
                            # --- RETOMAR: OldState=Inactive,NewState=Active = la llamada sale del hold ---
                            # Ocurre al volver a una línea que tenía una llamada en espera. Sin esto, el grid
                            # mostraba el cambio de audio pero nunca decía que la llamada se retomó.
                            elseif ($linea -match "Call StateChanged Call\[Id=[0-9a-fA-F\-]+,ConnectionId=(\d+),OldState=Inactive,NewState=Active\]") {
                                $SesR = $matches[1]
                                if (-not $RetomaImplicita.ContainsKey($SesR)) { $RetomaImplicita[$SesR] = @() }
                                $SlotRet = "$HoraLimpia,$MsLimpio"
                                if ($RetomaImplicita[$SesR] -notcontains $SlotRet) {
                                    Init-Hora $SlotRet
                                    $RetomaImplicita[$SesR] += $SlotRet
                                }
                            }
                            # --- FALLBACK TELÉFONO: capturar RemoteParty del "Call ended" del día analizado ---
                            # Pobla $MapeoTelP4 solo con líneas del día seleccionado (filtro de fecha activo).
                            # Permite que el sweep FALLBACK FIN muestre el teléfono cuando EndpointLog no tiene la sesión.
                            if ($linea -match "(?i)Call ended.*?Id=(\d+).*RemoteParty=\[.*?,([^\]]+)\]") {
                                $CeId = $matches[1]; $CeNum = ($matches[2] -replace '\D','')
                                if ($CeNum -and $CeNum.Length -ge 4 -and -not $MapeoTelP4.ContainsKey($CeId)) {
                                    $MapeoTelP4[$CeId] = $CeNum
                                }
                            }
                            # --- Rastreo de estado del VI para detectar CUELGUE EN HOLD ---
                            # El VI deja su estado en StateImpl=[...type=Active|Inactive|Disconnected...].
                            # Guardamos el último estado vivo; al llegar Disconnected, si el previo era
                            # Inactive (en hold) marcamos EnHold=$true. Es un 'if' independiente (no parte
                            # del elseif de PRIMARY, que consume type=Active) para no perder las
                            # transiciones de hold/unhold. El sweep post-PASO5 lo usa como fallback de FIN.
                            if ($linea -match "VoiceInteractionImpl\[VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}),cxt=[0-9a-fA-F\-]{36}\]\.StateImpl=\[.*?type=(Active|Inactive|Disconnected)") {
                                $viU = $matches[1]; $viSt = $matches[2]
                                if ($viSt -eq "Disconnected") {
                                    if (-not $ViDisconnected.ContainsKey($viU)) {
                                        $ViDisconnected[$viU] = "$HoraLimpia,$MsLimpio"
                                        $ViFinEnHold[$viU]    = ($ViLastState[$viU] -eq "Inactive")
                                    }
                                } else {
                                    $ViLastState[$viU] = $viSt
                                    # Si el VI llegó a Active, SÍ se contestó → no es candidato a "sin botones"
                                    if ($viSt -eq "Active") { $ViSinBotones.Remove($viU) | Out-Null }
                                }
                            }
                            # --- SIN BOTONES PARA TOMAR LA LLAMADA (caso dleal 02/07/2026) ---
                            # Toda llamada ACD auto-answer pasa por Alerting con answer=False (el botón manual
                            # se deshabilita porque el sistema la contestará). En sanas llega "Auto Accepting"
                            # ~50ms después y la contesta. Si el motor de auto-contestación tronó (FALLA EN LA
                            # RECEPCIÓN), ese Auto Accepting NUNCA llega y la llamada queda timbrando sin
                            # botones en la UI. Aquí solo se registra el candidato; el barrido post-PASO5
                            # descarta los que sí fueron contestados (ModoContestacion) o abandonos rápidos.
                            if ($linea -match "VoiceInteractionImpl\[VI\d+:([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}),cxt=([0-9a-fA-F\-]{36})\]\.StateImpl=\[operations=\[[^\]]*answer=False[^\]]*\],type=Alerting") {
                                $viUB = $matches[1]; $cxtUB = $matches[2]
                                if (-not $Script:ModoContestacion.ContainsKey($viUB) -and -not $ViSinBotones.ContainsKey($viUB)) {
                                    $ViSinBotones[$viUB] = @{ Slot = "$HoraLimpia,$MsLimpio"; Raw = $linea; Cxt = $cxtUB }
                                }
                            }
                        }

                        # Fuera del filtro de fecha — corren para TODAS las líneas del archivo
                        if ($HoraLimpia -ne "" -and $linea -match "Session_LoginAgent failed") {
                            Init-Hora $HoraLimpia
                            # La línea del fallo es continuación de un stack trace: NO trae timestamp propio.
                            # Su ms real vive en la línea anterior con hora (End Executing method LoginAgent),
                            # que $MsLimpio ya conserva (solo se actualiza en líneas con timestamp). Se antepone
                            # "[HH:mm:ss,fff]" para que el extractor de ms del render lo encuentre y el tooltip lo muestre.
                            $EventosTiempo[$HoraLimpia].RawAux += "[$HoraLimpia,$MsLimpio] $linea`n"
                        }
                        # Mismo patrón que "Session_LoginAgent failed" arriba, para el intento de DESFIRME:
                        # "Session_LogoutAgent failed...Logout Agent failed - Call not disconnected" es otra
                        # línea de continuación de stack trace sin timestamp propio.
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)Session_LogoutAgent failed|Logout Agent failed" -and $UltimoSlotDesfirmeSolicitado -ne "" -and $EventosTiempo.ContainsKey($UltimoSlotDesfirmeSolicitado)) {
                            # Se marca el slot de la solicitud (para el badge de Estado, que sí puede mostrar
                            # el desenlace en retrospectiva) pero YA NO se usa para reemplazar la Actividad de
                            # esa fila — Pablo (09/09/2026): el fallo es un HECHO DISTINTO, ocurrido varios
                            # segundos después, y "Asesor tratando de desfirmarse" debe seguir viéndose tal cual
                            # se vio en el momento del clic. El fallo se reporta en SU PROPIA fila ms (mismo
                            # patrón que "Se libera el estado" / "Asesor ya se encuentra desfirmado").
                            $EventosTiempo[$UltimoSlotDesfirmeSolicitado].Aux    += "DESFIRME_FALLIDO|"
                            $EventosTiempo[$UltimoSlotDesfirmeSolicitado].RawAux += "[$HoraLimpia,$MsLimpio] $linea`n"
                            # NO se crea la fila aquí todavía: el ms heredado ("End Executing method LogoutAgent")
                            # puede coincidir POR COINCIDENCIA con el ms de un FIN DE LLAMADA de OTRA sesión que
                            # vive en EndpointLog — y PASO 5 (que corre DESPUÉS de este punto) sobreescribe esa
                            # fila sin preguntar. Buscar "el próximo ms libre" aquí no servía de nada porque en
                            # este momento el FIN de PASO 5 TODAVÍA no existe — se ve libre, pero deja de estarlo
                            # más tarde (caso real de Pablo, 09/09/2026). Se guarda en $DesfirmeFallidoList y la
                            # fila se arma en el barrido justo antes de PASO 6, cuando PASO 4 Y PASO 5 ya
                            # terminaron de escribir todo — ahí "libre" sí significa libre de verdad.
                            $DesfirmeFallidoList += [pscustomobject]@{ Hora = $HoraLimpia; Ms = $MsLimpio; Raw = $linea }
                            # Ya se sabe que este intento falló — no esperar más una confirmación real que nunca
                            # va a llegar (el residuo de PendingAux, minutos después, NO cuenta como logout).
                            $DesfirmeEsperandoConfirmacion = $false
                        }
                        if ($linea -match "(?i)Call ended.*?Id=(\d+)") { $ConexionesYaIniciadas.Remove($matches[1]) | Out-Null }

                        # Radares de crash y errores
                        # OJO: ThreadAbortException / "Subproceso anulado" es AMBIGUO por sí solo. Es un abort de hilo .NET
                        # que Avaya lanza al reciclar sus PROPIOS hilos, y ocurre en DOS situaciones opuestas:
                        #   (a) recuperación de red / reinicio de servicio → la app SIGUE abierta.
                        #   (b) teardown de un cierre normal (Shutdown/ExitHandler) → la app SÍ se cerró.
                        # Un cierre real por Task Manager (TerminateProcess) NO deja este log (el proceso muere sin más).
                        # Antes se marcaba "Cierre Forzado" rojo (falso positivo en (a)); luego "la app NO se cerró"
                        # (falso en (b)). Ahora se CAPTURA y se decide en el barrido según haya o no cierre de app.
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)ThreadAbortException|Subproceso anulado") {
                            $HiloAbortadoList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio"; Raw = $linea }
                        }
                        # Falla Crítica GENERALIZADA de una operación telefónica: "VoiceInteraction_X failed" cubre
                        # Hold, Transfer y cualquier otra operación futura que use el mismo mecanismo interno de
                        # Avaya (Utils.CheckConditionAndHandleError lanzando PhoneServiceException "X Failed" cuando
                        # la llamada muere/cambia de estado mientras la operación seguía esperando una condición —
                        # confirmado con evidencia real en Transfer [16.4s de espera] y Hold [3.0s de espera]).
                        # Se captura la PALABRA que antecede a "Failed" ($OpFalloVI) y se incluye en la etiqueta:
                        # Pablo pidió explícitamente NO generalizar el texto a algo genérico como "Falla en
                        # Operación", para no perder el detalle de CUÁL operación fue (caso Hold Failed 08/08/2026).
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)VoiceInteraction_(\w+) failed") {
                            $OpFalloVI = $matches[1]
                            $SlotFalloVI = "$HoraLimpia,$MsLimpio"
                            Init-Hora $SlotFalloVI; $EventosTiempo[$SlotFalloVI].AppLog = "¡AVAYA ALERTA!: $OpFalloVI FAILED (Falla Crítica del sistema)"
                            $EventosTiempo[$SlotFalloVI].ColorApp = [System.Drawing.Color]::Red; $EventosTiempo[$SlotFalloVI].RawAppLog += "$linea`n"
                        }
                        # Variantes de fallo de transferencia SIN el prefijo "VoiceInteraction_" (mensajes/builds
                        # antiguos donde no hay palabra capturable) — se conserva la etiqueta específica de Transferencia.
                        elseif ($HoraLimpia -ne "" -and $linea -match "(?i)Transfer failed --->|StartTransferCallAction") {
                            $SlotFalloVI = "$HoraLimpia,$MsLimpio"
                            Init-Hora $SlotFalloVI; $EventosTiempo[$SlotFalloVI].AppLog = "¡AVAYA ALERTA!: TRANSFER FAILED (Falla Crítica del sistema)"
                            $EventosTiempo[$SlotFalloVI].ColorApp = [System.Drawing.Color]::Red; $EventosTiempo[$SlotFalloVI].RawAppLog += "$linea`n"
                        }
                        # Transfer() ejecutado con ORIGEN = DESTINO (mismo Call Id y ConnId en ambos argumentos):
                        # el método interno de Avaya se invocó SIN que se eligiera un destino real (no pasó por
                        # el submenú de "Ingresar valor"/contacto). Huella validada con 0 falsos positivos contra
                        # 2 transferencias exitosas (banco 205729: ConnSrc=47≠ConnDst=46; prueba plopezs 161304:
                        # ConnSrc=2≠ConnDst=3) vs el incidente jvazquezn 05/08/2026 (ConnSrc=45=ConnDst=45).
                        # No es un simple "olvidó marcar" del asesor: es Avaya iniciando la operación con un
                        # destino auto-referenciado, la causa raíz del congelamiento de ~16s reportado. Se marca
                        # como alerta de sistema (mismo prefijo "¡AVAYA ALERTA!" → fila resaltada en rojo) y se
                        # suma a $Script:ErroresSospechosos para que aparezca también en "Revisar posibles errores"
                        # al escanear logs de otros asesores/días.
                        if ($HoraLimpia -ne "" -and $linea -match "Begin Executing method Transfer\(Call\[Id=([0-9a-fA-F\-]+),ConnId=(\d+)[^\]]*\],Call\[Id=([0-9a-fA-F\-]+),ConnId=(\d+)[^\]]*\]\)") {
                            if ($matches[1] -eq $matches[3] -and $matches[2] -eq $matches[4]) {
                                $SlotSelfTransf = "$HoraLimpia,$MsLimpio"
                                Init-Hora $SlotSelfTransf
                                $EventosTiempo[$SlotSelfTransf].AppLog = "¡AVAYA ALERTA!: Transfer() con origen=destino (Sesión $($matches[2])) — no hubo selección real de destino"
                                $EventosTiempo[$SlotSelfTransf].ColorApp = [System.Drawing.Color]::Red
                                $EventosTiempo[$SlotSelfTransf].RawAppLog += "$linea`n"
                                $Script:ErroresSospechosos += [pscustomobject]@{ Hora = "$HoraLimpia,$MsLimpio"; Linea = "[TRANSFER ORIGEN=DESTINO] $($linea.Trim())" }
                            }
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "MoveSessionToConferenceCommand.*response is null|MoveSessionToConferenceRequest.*timed out") {
                            # MoveSessionToConferenceCommand response null = REAL ERROR: el CM no confirmó la unión.
                            # A diferencia de TransferSessionCommand (null = normal en ciega), aquí siempre es fallo.
                            $SlotErr = "$HoraLimpia,$MsLimpio"; Init-Hora $SlotErr
                            $EventosTiempo[$SlotErr].AppLog = "¡FALLO CONFERENCIA!: CM no confirmó la unión (Timeout 15s)"
                            $EventosTiempo[$SlotErr].ColorApp = [System.Drawing.Color]::Red
                            $EventosTiempo[$SlotErr].RawAppLog += "$linea`n"
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "ReconnectSessionCommand.*response is null|UnholdSessionRequest.*timed out") {
                            $SlotErr = "$HoraLimpia,$MsLimpio"; Init-Hora $SlotErr
                            if ([string]::IsNullOrEmpty($EventosTiempo[$SlotErr].AppLog)) {
                                $EventosTiempo[$SlotErr].AppLog = "Fallo al Retomar Llamada (CM no respondió)"
                                $EventosTiempo[$SlotErr].ColorApp = [System.Drawing.Color]::OrangeRed
                            }
                            $EventosTiempo[$SlotErr].RawAppLog += "$linea`n"
                        }
                        # --- FALLO AL RETENER (Hold): el PBX no confirmó el hold en 15 s (RC_RESPONSE_NULL) ---
                        # Distinto de Reconnect/Unhold (retomar). Típico al intentar retener una línea SIN
                        # llamada establecida (ej: 2ª línea con solo tono de marcar) → 15 s de timeout.
                        # Debe ir ANTES del catch-all genérico para no quedar como vago "Evento sospechoso".
                        # Se ancla SOLO a la excepción "Command failed [action=Hold, call=...]": es una única
                        # línea por fallo y trae el UUID de la llamada. (Antes también enganchaba
                        # "HoldSessionCommand response is null" → generaba la fila DUPLICADA.)
                        if ($HoraLimpia -ne "" -and $linea -match "Command failed \[action=Hold") {
                            $SlotHF = "$HoraLimpia,$MsLimpio"; Init-Hora $SlotHF
                            $EventosTiempo[$SlotHF].AppLog      = "¡FALLO AL RETENER!: la llamada no se pudo poner en espera (Timeout PBX 15 s)"
                            $EventosTiempo[$SlotHF].ColorApp    = [System.Drawing.Color]::Red
                            $EventosTiempo[$SlotHF].RawAppLog  += "$linea`n"
                            if ([string]::IsNullOrEmpty($EventosTiempo[$SlotHF].Interpretacion)) {
                                $EventosTiempo[$SlotHF].Interpretacion      = "$symStop FALLO AL RETENER (la llamada no se pudo poner en espera)"
                                $EventosTiempo[$SlotHF].ColorInterpretacion = [System.Drawing.Color]::Red
                            }
                            # Resolver la sesión por el UUID de la llamada para relabelar su "HOLD MANUAL".
                            if ($linea -match "call=([0-9a-fA-F\-]{36})" -and $Script:CxtToConnId.ContainsKey($matches[1])) {
                                $HoldFalloSes[$Script:CxtToConnId[$matches[1]]] = $SlotHF
                            }
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "TransferSessionCommand.*response is null|TransferSessionRequest.*timed out") {
                            Init-Hora $HoraLimpia
                            # response null = NORMAL para transferencia ciega (timeout 1s por diseño del CM)
                            # No se marca como error; solo se conserva en Raw para evidencia.
                            $EventosTiempo[$HoraLimpia].RawAppLog += "$linea`n"
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)timed out|response is null" `
                            -and $linea -match "\bERROR\b" `
                            -and $linea -notmatch "does not require a response" `
                            -and $linea -notmatch "TransferSessionCommand|TransferSessionRequest|MoveSessionToConferenceCommand|MoveSessionToConferenceRequest") {
                            $SlotErr = "$HoraLimpia,$MsLimpio"; Init-Hora $SlotErr
                            if ([string]::IsNullOrEmpty($EventosTiempo[$SlotErr].AppLog)) {
                                $EventosTiempo[$SlotErr].AppLog = "¡RED/SISTEMA!: Posible Desincronización (Timeout/Null)"
                                $EventosTiempo[$SlotErr].ColorApp = [System.Drawing.Color]::Orange
                            }
                            $EventosTiempo[$SlotErr].RawAppLog += "$linea`n"
                        }
                        # --- FALLA EN LA RECEPCIÓN DE LLAMADA: el motor de auto-contestación tronó ---
                        # Caso dleal 02/07/2026: "Exception ouccred while doing AutoAccpet" (typos de Avaya).
                        # El error deja rota la auto-contestación: la llamada que estaba timbrando queda
                        # huérfana (Alerting sin botones). Se detecta el ERROR de forma genérica y se
                        # capturan las líneas de stack que siguen (sin fecha) para el detalle de la celda.
                        if ($HoraLimpia -ne "" -and $linea -match "(?i)Exception\s+\w*\s*while doing AutoAcc") {
                            $Script:FallaRecepSlot = "$HoraLimpia,$MsLimpio"; Init-Hora $Script:FallaRecepSlot
                            $Script:FallaRecepLineas = 0
                            if ($EventosTiempo[$Script:FallaRecepSlot].Interpretacion -notmatch "FALLA EN LA RECEPCIÓN") {
                                $EventosTiempo[$Script:FallaRecepSlot].Interpretacion = "⚠ FALLA EN LA RECEPCIÓN DE LLAMADA."
                                $EventosTiempo[$Script:FallaRecepSlot].ColorInterpretacion = [System.Drawing.Color]::Red
                            }
                            if ($EventosTiempo[$Script:FallaRecepSlot].Tel -eq "-") { $EventosTiempo[$Script:FallaRecepSlot].Tel = "AVAYA/ONEX" }
                            $EventosTiempo[$Script:FallaRecepSlot].RawInterpretacion += "$linea`n"
                        }
                        elseif ($Script:FallaRecepSlot) {
                            if ($linea -match "^\[?$FechaOmni") { $Script:FallaRecepSlot = $null }
                            elseif ($Script:FallaRecepLineas -lt 12) {
                                # línea de stack (sin fecha) inmediatamente después del ERROR
                                $EventosTiempo[$Script:FallaRecepSlot].RawInterpretacion += "$linea`n"
                                $Script:FallaRecepLineas++
                            }
                        }
                        if ($HoraLimpia -ne "" -and $linea -match "System\.Exception") {
                            $SlotErr = "$HoraLimpia,$MsLimpio"; Init-Hora $SlotErr; $EventosTiempo[$SlotErr].AppLog = "¡ERR CRÍTICO!: System.Exception"
                            $EventosTiempo[$SlotErr].ColorApp = [System.Drawing.Color]::DarkRed; $EventosTiempo[$SlotErr].RawAppLog += "$linea`n"; $EventosTiempo[$SlotErr].RawInterpretacion += "$linea`n"
                        }
                        # LoginCommand/LogonResponse: el rawxml de la respuesta de login trae texto que dispara el
                        # catch-all (p.ej. campos de error con valor 0). No es un error → se excluye.
                        elseif ($HoraLimpia -ne "" -and $linea -match "(?i)\bException\b|\bFATAL\b|\bERROR\b" -and $linea -notmatch "VoiceInteraction_\w+ failed|ThreadAbortException|System\.Exception|TransferSessionCommand|TransferSessionRequest|LoginCommand|LogonResponse|LoginResponse") {
                            # NO se pinta en el timeline (era "Evento sospechoso", ruido benigno + alarmante en
                            # pantalla compartida). Se acumula en el buffer para revisarlo aparte bajo demanda.
                            $Script:ErroresSospechosos += [pscustomobject]@{ Hora = "$HoraLimpia,$MsLimpio"; Linea = $linea.Trim() }
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
            $TransferConsultaColgada = $false             # $true cuando Transfer_CompleteSetup trae nfirstCall==nSecondCall (la pata de consulta/destino ya colgó antes de completar)
            $TransferSinDestinoPorSes = @{}               # sesión cliente (from) → $true si "OnRequestTransferSession: To session id" == from (no se marcó un destino separado → transferencia SIN destino)
            $TransferSinDestino = $false                  # flag vigente en el CompleteSetup: distingue "no se marcó destino" de "el destino colgó"
            $DescForzadaSlot = @{}                        # segundo → slot: dedup de "otra sesión tomó la extensión" (URQ 2009/Force logoff by server dispara 3 líneas en el mismo segundo)

            if ($ArchivosLog) {
                foreach ($Archivo in $ArchivosLog) {
                    $LineasLog = Get-Content -Path $Archivo.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                    if (-not $LineasLog) { continue }

                    # Primera pasada: extensiones y logins
                    # ── Extensiones (StartUserRegistration): SIN filtro de fecha.
                    #    El evento de registración puede estar en logs rotados de días anteriores
                    #    si OneX no se ha reiniciado recientemente. Se deduplicca solo por número
                    #    de extensión para no mostrar spam de timestamps repetidos.
                    # ── Login number (564...): CON filtro de fecha exacta.
                    #    Evita que números de sesiones de prueba u otras fechas contaminen
                    #    la lista (ej. 399999 de un día de pruebas apareciendo en auditorías reales).
                    $ExtensionesVistas = @{}
                    foreach ($linea in $LineasLog) {
                        if ($linea -match "StartUserRegistration: server: '[^']+', extension: '(\d+)'") {
                            $ExtEncontrada = $matches[1]; $CurrentExt = $ExtEncontrada
                            if (-not $ExtensionesVistas.ContainsKey($ExtEncontrada)) {
                                $ExtensionesVistas[$ExtEncontrada] = $true
                                $ListaExtensiones += $ExtEncontrada
                            }
                        }
                        if ($linea -match "^\[?$FechaOmni" -and $linea -match "\+?564(3\d{5})\d{6}\b") {
                            $LoginExtraido = $matches[1]
                            $HoraFirma = if ($linea -match "\d{2}:\d{2}:\d{2}") { $matches[0] } else { "---" }
                            $FirmaUnicaLog = "$LoginExtraido ($HoraFirma)"; $YaRegistrado = $false
                            foreach ($item in $ListaLogins) { if ($item -match "^$LoginExtraido") { $YaRegistrado = $true; break } }
                            if (-not $YaRegistrado) { $ListaLogins += $FirmaUnicaLog }
                        }
                    }

                    # ── Mini-pasada: sesiones de consulta (Transferencia / Conferencia) ──
                    # OnRequestTransferSession + "To phone#":
                    #   → UpdateHistoryRecord vacío           = leg de consulta atendida  → $ConsultaTransf
                    #   → UpdateHistoryRecord con número      = transferencia ciega directa → $ConsultaTransf
                    # Conferencia: OnRequestMoveSessionToConference → UpdateHistoryRecord (SessionId HISTORY) → Conference_CompleteConf
                    #   Problema: nConsultCall en Conference_CompleteConf es el callIndex del módulo SESSION,
                    #   pero la grilla usa SessionId del módulo HISTORY (pueden diferir, ej: callIndex=40, SessionId=42).
                    #   Solución: capturar el SessionId de UpdateHistoryRecord mientras $_CfPend=true y usarlo en $ConsultaConf.
                    # $_CfPend  = true desde OnRequestMoveSessionToConference hasta Conference_CompleteConf
                    # $_CfDone  = true durante la ventana JUSTO DESPUÉS de Conference_CompleteConf:
                    #             captura el primer UpdateHistoryRecord vacío post-CompleteConf, que es la
                    #             sesión real que aparece en la grilla (puede ser distinta a la que llega antes).
                    # Problema raíz: la asesora trabaja turno nocturno (23:00–07:00). El mismo equipo
                    #   acumula sesiones del día anterior. Ejemplo: sesión 42 es una transferencia real del 28/05,
                    #   y queda en $ConsultaTransf["42"]. Al procesar el 29/05, la conferencia también usa
                    #   sesión 42 → $ConsultaTransf["42"] contamina la etiqueta. La solución: cuando se
                    #   confirma la conferencia, removemos de $ConsultaTransf TODAS las sesiones involucradas,
                    #   incluyendo la que llega después de Conference_CompleteConf.
                    # Filtro de fecha ($FechaOmni, igual que el resto de PASO 5): un archivo EndpointLog.txt
                    # rotado puede abarcar varios días, y Avaya reinicia/recicla el numero de SessionId en
                    # cada login. Sin este filtro, una transferencia/conferencia real de OTRO día puede
                    # contaminar $ConsultaTransf/$ConsultaConf con el mismo SessionId que hoy, mostrando un
                    # destino que nunca ocurrió hoy (caso: "MARCANDO A LA EXT. +5520819192", confirmado por
                    # Pablo como evento de un día previo con el mismo Id). Pablo, 18/08/2026.
                    $_TrPend = $false; $_TrDest = ""; $_CfPhones = @{}; $_CfPend = $false; $_CfDone = $false; $_TrFrom = ""
                    foreach ($linea in $LineasLog) {
                        if ($linea -notmatch "^\[?$FechaOmni") { continue }
                        if ($linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*([^.]+)\.") {
                            $s = $matches[1]; $n = $matches[2].Trim()
                            if ($n -ne "") {
                                $_CfPhones[$s] = $n
                                # Transferencia ciega: UpdateHistoryRecord llega con número inmediatamente
                                # (transferencia atendida llega vacío primero; ciega llega con destino directo)
                                if ($_TrPend -and -not $ConsultaTransf.ContainsKey($s)) {
                                    $ConsultaTransf[$s] = $n
                                    $_TrPend = $false
                                }
                            }
                        }
                        if ($linea -match "OnRequestTransferSession\(\) entered from sessionId=(\d+)") {
                            $_TrPend = $true; $_TrDest = ""; $_CfPend = $false; $_CfDone = $false; $_TrFrom = $matches[1]
                        }
                        # "To session id" == la sesión del cliente (from) ⇒ NO hubo una consulta/destino separado
                        # (transferencia SIN destino, ej: se dio Transferir sin marcar). Si es distinta ⇒ sí hubo consulta.
                        if ($linea -match "OnRequestTransferSession: To session id=(\d+)" -and $_TrFrom -ne "") {
                            $TransferSinDestinoPorSes[$_TrFrom] = ($_TrFrom -eq $matches[1])
                        }
                        if ($_TrPend -and $linea -match "\bTo phone#: '([^']+)'") { $_TrDest = $matches[1] }
                        if ($_TrPend -and $linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*\.") {
                            $s = $matches[1]
                            if (-not $ConsultaTransf.ContainsKey($s)) {
                                $ConsultaTransf[$s] = if ($_TrDest -ne "") { $_TrDest } else { "Desconocido" }
                            }
                            $_TrPend = $false
                        }
                        # Conferencia — INICIO: nueva conferencia cancela cualquier flag previo
                        if ($linea -match "OnRequestMoveSessionToConference entered") {
                            $_CfPend = $true; $_CfDone = $false; $_TrPend = $false
                        }
                        # Conferencia — CAPTURA de sesión (ANTES o JUSTO DESPUÉS de CompleteConf):
                        #   $_CfPend captura sesiones que Avaya registra mientras procesa el merge (ej: sesión 41).
                        #   $_CfDone captura la primera sesión que llega POST-CompleteConf (ej: sesión 42 = la de la grilla).
                        #   Ambas se agregan a $ConsultaConf y se remueven de $ConsultaTransf (elimina contaminación cross-día).
                        if (($_CfPend -or $_CfDone) -and $linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*\.") {
                            $sesId = $matches[1]
                            if (-not $ConsultaConf.ContainsKey($sesId)) {
                                $ConsultaConf[$sesId] = if ($_CfPhones[$sesId]) { $_CfPhones[$sesId] } else { "Desconocido" }
                            }
                            $ConsultaTransf.Remove($sesId)   # conferencia gana aunque el sesId venga de otro día
                            if ($_CfDone) { $_CfDone = $false }   # ventana post-CompleteConf: solo capturar una
                        }
                        # Conferencia — CompleteConf: el endpoint confirma el merge. Cambia modo a "post-merge"
                        # para capturar la sesión que Avaya registra en HISTORY justo después del merge.
                        if ($linea -match "Conference_CompleteConf: nFirstCall:\s*\d+,\s*nConsultCall:\s*(\d+)") {
                            $ConsultaTransf.Remove($matches[1])   # quitar callIndex también por si acaso
                            $_CfPend = $false; $_CfDone = $true
                        }
                    }

                    $LastXMLEvent = $null; $LastXMLTime = $null
                    $MsLimpio5 = "000"; $MuteSlotSeg = @{}; $UnmuteSlotSeg = @{}

                    # Segunda pasada: análisis principal
                    foreach ($linea in $LineasLog) {
                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2}).*?\].*<(HeldEvent|UnheldEvent)") { $LastXMLEvent = $matches[2]; $LastXMLTime = $matches[1] }
                        elseif ($LastXMLEvent -ne $null -and $linea -match "<connectionId>(\d+)</connectionId>") {
                            $Ses = $matches[1]; Init-Hora $LastXMLTime; $EventosTiempo[$LastXMLTime].Sesion = $Ses; $EventosTiempo[$LastXMLTime].Tel = if ($MapeoTel[$Ses]) { $MapeoTel[$Ses] } else { "Desconocido" }
                            if ($LastXMLEvent -eq "HeldEvent") { $EventosTiempo[$LastXMLTime].Agente = "|| HOLD (Confirmación del Sistema)"; $EventosTiempo[$LastXMLTime].ColorAgente = [System.Drawing.Color]::Yellow } else { $EventosTiempo[$LastXMLTime].Agente = "$symRes UNHOLD (Confirmación del Sistema)"; $EventosTiempo[$LastXMLTime].ColorAgente = [System.Drawing.Color]::LightGoldenrodYellow }
                            $EventosTiempo[$LastXMLTime].RawAgente += "<$LastXMLEvent>`n$linea`n"; $LastXMLEvent = $null
                        }

                        if ($linea -match "^\[?$FechaOmni.*?(\d{2}:\d{2}:\d{2})(?::(\d{3}))?") {
                            $HoraLimpia = $matches[1]; $MsLimpio5 = if ($matches[2]) { $matches[2] } else { "000" }
                            # Buscador de vacíos: registra CUALQUIER línea con timestamp (Endpoint.log), misma
                            # lista combinada que PASO 4 (ambos logs son del proceso OneXAgent.exe).
                            if ($HoraLimpia -match '^(\d{2}):(\d{2}):(\d{2})$') { $Script:TsProceso.Add(([int64]$matches[1])*3600000L + ([int64]$matches[2])*60000L + ([int64]$matches[3])*1000L + [int64]$MsLimpio5) }

                            # Letra de call-appearance (a/b/c… → línea 1/2/3…) para "Asesor abre línea N".
                            # Solo la que mapea a un buttonIndex válido (>=0); se descartan las de ' '/-1 (ruido de display).
                            if ($linea -match "MapCallAppToBtnIndex\(([a-zA-Z])\)\s*=>\s*buttonIndex:\s*(-?\d+)") {
                                $_lt = ([string]$matches[1]).ToLower(); $_btn = [int]$matches[2]
                                if ($_btn -ge 0) {
                                    $LineaAppPorSeg[$HoraLimpia] = $_lt
                                    # buttonIndex → número de línea (a=1, b=2, c=3…): el índice físico varía por plantilla, la letra no.
                                    if ($_lt -match "^[a-z]$" -and -not $LineaAppBtn.ContainsKey($_btn)) { $LineaAppBtn[$_btn] = [int][char]$_lt - [int][char]'a' + 1 }
                                }
                            }
                            # Cronología de la LÍNEA ACTIVA: permite saber a qué línea pertenece cada sesión de audio.
                            # ("found activeLine: N" al conmutar de línea; <lineAppearanceId>N</lineAppearanceId> al crear la llamada.)
                            if ($linea -match "found activeLine:\s*(\d+)" -or $linea -match "<lineAppearanceId>(\d+)</lineAppearanceId>") {
                                $_btnA = [int]$matches[1]
                                try { $_msA = [int]([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).TimeOfDay.TotalSeconds) * 1000 + [int]$MsLimpio5 } catch { $_msA = -1 }
                                if ($_msA -ge 0) { $ActiveLineSeq += [pscustomobject]@{ Ms = $_msA; Btn = $_btnA } }
                            }

                            # DTMF: dígitos que el asesor teclea DURANTE una llamada (marcar destino o navegar IVR).
                            # "input digit constant" da el símbolo (19-28=0-9, 29='*', 30='#'); la línea siguiente
                            # "nCallIndex N: call State" da la sesión. Los dígitos de firma/desfirme (FAC) NO traen
                            # nCallIndex, por eso quedan fuera. El sweep post-PASO5 los agrupa en secuencias.
                            if ($linea -match "OnRequestPressDigit: input digit constant:\s*(\d+)") {
                                $DtmfConstPend = [int]$matches[1]; $DtmfHoraPend = "$HoraLimpia,$MsLimpio5"
                            }
                            elseif ($linea -match "OnRequestPressDigit:.*nCallIndex\s+(\d+):\s*call State") {
                                if ($null -ne $DtmfConstPend) {
                                    $simDtmf = if ($DtmfConstPend -ge 19 -and $DtmfConstPend -le 28) { [string]($DtmfConstPend - 19) } elseif ($DtmfConstPend -eq 29) { "*" } elseif ($DtmfConstPend -eq 30) { "#" } else { "?" }
                                    $DtmfPresses += [pscustomobject]@{ Sim = $simDtmf; Ses = $matches[1]; Hora = $DtmfHoraPend }
                                    $DtmfConstPend = $null
                                }
                            }

                            # Número que el asesor teclea en la caja de texto y envía con Enter (método NewCallHandler).
                            # ApplyDialingRulesToNumber trae el número ya listo para marcar (con prefijo de salida).
                            # Se guarda con su ms para emparejarlo con el NewCallHandler más cercano tras PASO 5.
                            if ($linea -match "ApplyDialingRulesToNumber entered '([^']+)'") {
                                try { $_msAD = [int]([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).TimeOfDay.TotalSeconds) * 1000 + [int]$MsLimpio5 } catch { $_msAD = -1 }
                                if ($_msAD -ge 0) { $DialedApply += [pscustomobject]@{ Ms = $_msAD; Num = $matches[1] } }
                            }

                            # Mapeo de dirección
                            if ($linea -match "GenerateIncomingCall: callIndex:\s*(\d+)") {
                                $DirLlamada[$matches[1]] = "ENTRANTE"
                                # Evidencia para el diagnóstico: registrar en el slot actual
                                if ($HoraLimpia -ne "") { Init-Hora $HoraLimpia; $EventosTiempo[$HoraLimpia].RawAux += "[DIR:ENTRANTE] PASO5: GenerateIncomingCall callIndex=$($matches[1])`n$linea`n" }
                            }

                            # Signal A (EndpointLog): Call created con Outgoing=True y RemoteParty vacío = sesión bridge
                            if ($linea -match "(?i)Call created : Id=(\d+).*?Outgoing=(True|False)") {
                                $IDLlamadaEP = $matches[1]; $SentidoEP = $matches[2]
                                if ($SentidoEP -eq "True" -and $linea -match "RemoteParty=\[,\]") {
                                    $SesionesBridge[$IDLlamadaEP] = $HoraLimpia
                                    $DirLlamada[$IDLlamadaEP] = "BRIDGE"
                                    # Una sola fila por sesión (ver nota en el creador de PASO 4).
                                    if (-not $BridgeSlot.ContainsKey($IDLlamadaEP)) {
                                        $SlotBEP = "$HoraLimpia,$MsLimpio5"
                                        $BridgeSlot[$IDLlamadaEP] = $SlotBEP
                                        Init-Hora $SlotBEP
                                        $EventosTiempo[$SlotBEP].Sesion = $IDLlamadaEP
                                        if ($EventosTiempo[$SlotBEP].Interpretacion -eq "") {
                                            $EventosTiempo[$SlotBEP].Interpretacion = "[!] Sesión bridge del sistema (auto-transferencia)"
                                            $EventosTiempo[$SlotBEP].ColorInterpretacion = [System.Drawing.Color]::Orange
                                        }
                                    }
                                    $EventosTiempo[$BridgeSlot[$IDLlamadaEP]].RawInterpretacion += "$linea`n"
                                } else {
                                    if ($SentidoEP -eq "True") { $DirLlamada[$IDLlamadaEP] = "SALIENTE" } else { $DirLlamada[$IDLlamadaEP] = "ENTRANTE" }
                                }
                            }

                            if ($linea -match "UpdateHistoryRecord: SessionId=\s*(\d+),.*RemoteUserAddress=\s*([^.]+)\.") {
                                $SesID = $matches[1]; $NumObj = $matches[2].Trim()
                                $NumLimpio = $NumObj -replace '\D',''
                                # Dirección PLAUSIBLE = al menos 3 dígitos. Las patas que Avaya crea para AUTO-CONTESTAR
                                # una entrante (y los marcados a medias) reciben basura de 1 dígito: un '0', o el prefijo
                                # de salida '9'. Eso NO es un destino. Antes bastaba cualquier dígito → se retiraba el
                                # marcador bridge y la pata phantom terminaba pintada como "INICIO DE LLAMADA (Saliente)
                                # al 0" junto a la llamada entrante real. Extensiones (5 díg.) y códigos FAC (3) sí pasan.
                                $EsNumPlausible = ($NumLimpio.Length -ge 3)
                                if ($EsNumPlausible) { $MapeoTel[$SesID] = $NumObj; $RawMapeoTel[$SesID] = $linea; $FarEndResueltoSlots += [pscustomobject]@{ Ses = $SesID; Slot = "$HoraLimpia,$MsLimpio5" } }
                                # --- RETRACT bridge falso-positivo ---
                                # Si esta sesión fue marcada "Sesión bridge" en PASO 4/5 pero ahora tiene una
                                # dirección REAL en el historial (dígitos), NO es una pata phantom de transferencia:
                                # es una llamada normal (login FAC 564, marcado 955, número externo, favorito de línea)
                                # que se marcó bridge solo porque RemoteParty nace vacío y se puebla ~ms después.
                                # Las patas phantom reales NUNCA obtienen dirección (RemoteUserAddress vacío) → nunca se retiran.
                                if ($EsNumPlausible -and $SesionesBridge.ContainsKey($SesID)) {
                                    $SlotBr = $BridgeSlot[$SesID]
                                    if ($SlotBr -and $EventosTiempo.ContainsKey($SlotBr) -and $EventosTiempo[$SlotBr].Interpretacion -match "Sesión bridge") {
                                        $EventosTiempo[$SlotBr].Interpretacion = ""
                                        $EventosTiempo[$SlotBr].ColorInterpretacion = [System.Drawing.Color]::White
                                        $EventosTiempo[$SlotBr].Sesion = "-"
                                    }
                                    $SesionesBridge.Remove($SesID) | Out-Null
                                    $BridgeSlot.Remove($SesID) | Out-Null
                                    if ($DirLlamada[$SesID] -eq "BRIDGE") { $DirLlamada[$SesID] = "SALIENTE" }
                                }
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
                                # DEDUP del doble DIAL de login (Avaya dialea el FAC 564 dos veces): si ya hay un
                                # "Usuario intentando firmarse" en los últimos ~3s, no crear otra fila. Si este 2º
                                # dial encontró la Ext y el previo no, se mejora el previo. Va ANTES del cierre de
                                # zombies para no ejecutarlo dos veces (evita doble "SESIÓN ABORTADA").
                                $SlotPrevioIntento = $null
                                for ($iP = 0; $iP -le 3 -and -not $SlotPrevioIntento; $iP++) {
                                    try {
                                        $hP = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-$iP)).ToString("HH:mm:ss")
                                        foreach ($kP in @($EventosTiempo.Keys)) {
                                            if ((($kP -split ',')[0] -eq $hP) -and $EventosTiempo[$kP].Interpretacion -match "Usuario intentando firmarse") { $SlotPrevioIntento = $kP; break }
                                        }
                                    } catch {}
                                }
                                if ($SlotPrevioIntento) {
                                    if ($ExtEncontrada -ne "" -and $EventosTiempo[$SlotPrevioIntento].Interpretacion -notmatch "en la Ext\.") {
                                        $EventosTiempo[$SlotPrevioIntento].Interpretacion = "Usuario intentando firmarse en la Ext. $ExtEncontrada"
                                        if ($HoraMatch) { $EventosTiempo[$HoraMatch].Aux = $EventosTiempo[$HoraMatch].Aux -replace "INTENTO_FIRMA_ONEX_EXT:\d+\|", "" }
                                    }
                                    $EventosTiempo[$SlotPrevioIntento].RawInterpretacion += "[$HoraLimpia] (Endpoint, dial repetido del mismo login): $linea`n"
                                    continue
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
                                $Ses = $matches[1]
                                # Si el slot ya tiene un FIN/CUELGUE/EVASIÓN (ProcessSessionEndedEvent corrió primero al mismo segundo),
                                # no sobreescribir sus metadatos. Sesión/Tel/Interpretacion/Agente pertenecen a la sesión que terminó.
                                $FinEnSlot = $EventosTiempo[$HoraLimpia].Interpretacion -match "FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN"
                                if (-not $FinEnSlot) { $EventosTiempo[$HoraLimpia].Sesion = $Ses }
                                # Propagar Sesion al slot ms de señal si fue movida por auto-in en el mismo segundo
                                if ($Script:AlertingHorasRedirect.ContainsKey($HoraLimpia)) {
                                    $SlotMsSenial = $Script:AlertingHorasRedirect[$HoraLimpia]
                                    if ($EventosTiempo.ContainsKey($SlotMsSenial) -and $EventosTiempo[$SlotMsSenial].Sesion -eq "-") {
                                        $EventosTiempo[$SlotMsSenial].Sesion = $Ses
                                    }
                                }
                                $TelActual = if ($MapeoTel[$Ses]) { $MapeoTel[$Ses] } else { "Desconocido" }
                                # Normalizar número para deduplicación: quita '+' y el '9' de acceso a línea externa
                                # Ej: 95539991927 → 5539991927  |  +5539991927 → 5539991927  (mismo número, distintos formatos)
                                $TelNorm = ($TelActual -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                $FirmaUnica = "$Ses-$TelNorm"; $EstadoSesion[$Ses] = "ACTIVA"
                                # Fijar el PRIMER número real de la sesión (una sola vez; se libera en SessionEnded).
                                # Si un update posterior trae otro número → la sesión CAMBIÓ = transferencia inter-agente real.
                                if (-not $_PrimerFonoSes.ContainsKey($Ses) -and $TelActual -ne "Desconocido" -and $TelNorm -ne "") { $_PrimerFonoSes[$Ses] = $TelNorm }

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
                                        if (-not $FinEnSlot) { $EventosTiempo.Remove($HoraLimpia) | Out-Null }
                                    }

                                    if (-not $UpgradeHecho -and -not $FinEnSlot) {
                                    $LlamadasVistas[$FirmaUnica] = $true
                                    # Solo sobreescribir Tel si el slot aún no tiene número O si el nuevo valor es conocido
                                    if ($TelActual -ne "Desconocido" -or $EventosTiempo[$HoraLimpia].Tel -eq "-") { $EventosTiempo[$HoraLimpia].Tel = $TelActual }

                                    if ($TelActual -match "^\+?564(3\d{5})\d{6}$") { $EvA = "" }
                                    else {
                                        if ($DirLlamada[$Ses] -eq "SALIENTE") {
                                            # Fallback: si PASO 4 ya puso el label pero sin registrar $SesionHoraInicio
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -match "LÍNEA ABIERTA SIN MARCAR" -and
                                                -not $SesionHoraInicio.ContainsKey($Ses)) {
                                                $SesionHoraInicio[$Ses] = $HoraLimpia
                                            }
                                            # Solo crear INICIO si XML no lo hizo ya y no hay evento de TRANSFERENCIA/CONFERENCIA en el slot
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "señal de llamada|intentando firmarse|INICIO DE LLAMADA|LÍNEA ABIERTA|TRANSFERENCIA|CONFERENCIA" -and
                                                $EventosTiempo[$HoraLimpia].RawInterpretacion -notmatch "\[ABSORBIDO:") {
                                                if ($TelActual -match "Desconocido") {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR"
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
                                            if ($EventosTiempo[$HoraLimpia].Interpretacion -notmatch "intentando firmarse|INICIO DE LLAMADA|LÍNEA ABIERTA|señal de llamada|TRANSFERENCIA|CONFERENCIA") {
                                                if ($TelActual -match "Desconocido") {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp LÍNEA ABIERTA SIN MARCAR"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::Gold
                                                } elseif ($TelActual -match "^\d{1,6}$") {
                                                    # Número corto (1-6 dígitos) = extensión interna: no pasa por ACD, no genera señal de llamada.
                                                    # Tema 4: si PASO 4 ya creó un "INICIO DE LLAMADA (Saliente)" para esta misma sesión en otro
                                                    # slot (ms), NO duplicar con "INICIO DE SESIÓN": se enriquece ese slot con la extensión (y se
                                                    # corrige su Tel si quedó Desconocido) y este slot se deja sin label para que PASO 6 lo
                                                    # suprima como fila vacía de UpdateHistoryRecord.
                                                    $SlotInicioPrevio = $null
                                                    foreach ($kIni in @($EventosTiempo.Keys)) {
                                                        if ($kIni -ne $HoraLimpia -and $EventosTiempo[$kIni].Sesion -eq $Ses -and $EventosTiempo[$kIni].Interpretacion -match "INICIO DE LLAMADA") { $SlotInicioPrevio = $kIni; break }
                                                    }
                                                    if ($SlotInicioPrevio) {
                                                        if ($EventosTiempo[$SlotInicioPrevio].Interpretacion -notmatch "Extensión") {
                                                            $EventosTiempo[$SlotInicioPrevio].Interpretacion += " — Extensión $TelActual"
                                                        }
                                                        if ($EventosTiempo[$SlotInicioPrevio].Tel -eq "-" -or $EventosTiempo[$SlotInicioPrevio].Tel -eq "" -or $EventosTiempo[$SlotInicioPrevio].Tel -eq "Desconocido") {
                                                            $EventosTiempo[$SlotInicioPrevio].Tel = $TelActual
                                                        }
                                                        $EventosTiempo[$SlotInicioPrevio].RawInterpretacion += "[FUSIÓN tema4] INICIO DE SESIÓN interna (ext $TelActual) absorbido desde slot $HoraLimpia`n"
                                                    } else {
                                                        $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE SESIÓN (Llamada Interna — Extensión $TelActual)"
                                                        $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::DarkKhaki
                                                    }
                                                } else {
                                                    $EventosTiempo[$HoraLimpia].Interpretacion = "$symUp INICIO DE SESIÓN (Interna/Sistema)"
                                                    $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::DarkKhaki
                                                }
                                            }
                                        }
                                        # Actualización de teléfono mid-call: sesión ENTRANTE activa recibe número diferente.
                                        # Típico de transferencia inter-agente: CM reemplaza el número del agente originador
                                        # por el número real del cliente. No es una nueva llamada — la sesión continúa activa.
                                        # CLAVE: se exige que ESTA sesión haya CAMBIADO de número ($_PrimerFonoSes[$Ses] ≠ actual).
                                        # Antes bastaba con que el número estuviera en $MidCallPhones (lista GLOBAL por número), lo que
                                        # marcaba falsamente como inter-agente una llamada NUEVA de un cliente que ya había llamado
                                        # antes (mismo número, sesión distinta, sin cambio). Caso plopezs 28/07 (ses 45 real vs ses 46 falsa).
                                        if ($DirLlamada[$Ses] -eq "ENTRANTE" -and $TelActual -ne "Desconocido" -and
                                            $Script:MidCallPhones.ContainsKey($TelNorm) -and
                                            $_PrimerFonoSes.ContainsKey($Ses) -and $_PrimerFonoSes[$Ses] -ne $TelNorm -and
                                            $EventosTiempo[$HoraLimpia].Interpretacion -notmatch "intentando firmarse|INICIO DE LLAMADA|LÍNEA ABIERTA|señal de llamada") {
                                            $EventosTiempo[$HoraLimpia].Interpretacion = "$symArr Número de cliente actualizado (Transferencia inter-agente)"
                                            $EventosTiempo[$HoraLimpia].ColorInterpretacion = [System.Drawing.Color]::CornflowerBlue
                                        }
                                        $EventosTiempo[$HoraLimpia].RawInterpretacion += "$linea`n"
                                        $EvA = "UpdateHistoryRecord: SessionID=[$Ses]"; $ColorA = [System.Drawing.Color]::White
                                    }
                                    }  # fin if (-not $UpgradeHecho)
                                }
                            }
                            elseif ($linea -match "OnRequestHoldSession\(\)\. [Ss]ession id=\s*(\d+)") { $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses; $EventosTiempo[$HoraLimpia].Tel=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}; $EvA="|| HOLD MANUAL (Clic del Agente)"; $ColorA=[System.Drawing.Color]::Yellow }
                            elseif ($linea -match "OnRequestUnholdSession\(\)\. [Ss]ession id=\s*(\d+)") { $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses; $EventosTiempo[$HoraLimpia].Tel=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}; $EvA="$symRes UNHOLD MANUAL (Clic del Agente)"; $ColorA=[System.Drawing.Color]::LightGoldenrodYellow }
                            elseif ($linea -match "OnRequestTransferSession\(\) entered from sessionId=(\d+)") {
                                $SesT = $matches[1]
                                # Signal B: distinguir transferencia manual (GUI) de automática (sin GUI)
                                $EsAutoTransf = $true
                                if ($UltimaTransfGUIHora -ne "") {
                                    try {
                                        $TGUI    = [DateTime]::ParseExact($UltimaTransfGUIHora, "HH:mm:ss", $null)
                                        $TTransf = [DateTime]::ParseExact($HoraLimpia,           "HH:mm:ss", $null)
                                        if ([Math]::Abs(($TTransf - $TGUI).TotalSeconds) -le 5) { $EsAutoTransf = $false }
                                    } catch {}
                                }
                                if ($EsAutoTransf -and $SesionHoraInicio.ContainsKey($SesT)) {
                                    $SlotIni = $SesionHoraInicio[$SesT]
                                    Init-Hora $SlotIni
                                    $EventosTiempo[$SlotIni].Aux    += "TRANSF_AUTO|"
                                    $EventosTiempo[$SlotIni].RawAux += "[Auto-Transfer detectado en EndpointLog] $linea`n"
                                }
                                # Tema 6: la detección AUTO/manual se conserva (marca TRANSF_AUTO arriba y resalta el
                                # INICIO transferido), pero ya no se muestra "OnRequestTransferSession(...)" como evento
                                # en la grilla por ser ruido técnico. La línea cruda queda en el detalle Raw.
                                $EventosTiempo[$HoraLimpia].RawAgente += "$linea`n"
                            }
                            elseif ($linea -match "OnRequestEndSession\(\)\. [Ss]ession=\s*(\d+)") {
                                $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses
                                $TelActual=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}; $EventosTiempo[$HoraLimpia].Tel=$TelActual
                                $CuelguesManuales[$Ses]=$true; $EvA="$symStop CUELGUE MANUAL (Clic)"; $ColorA=[System.Drawing.Color]::LightCoral
                            }
                            elseif ($linea -match "ProcessSessionEndedEvent: Entry\. connectinoId = (\d+)") {
                                $Ses=$matches[1]; $EventosTiempo[$HoraLimpia].Sesion=$Ses
                                $TelActual=if($MapeoTel[$Ses]){$MapeoTel[$Ses]}else{"Desconocido"}
                                $FirmaUnica="$Ses-$TelActual"; $EstadoSesion[$Ses]="CERRADA"
                                # Liberar teléfono del guard SECONDARY para que una rellamada futura genere nuevo INICIO
                                $TelFinNorm = ($TelActual -replace '^\+','') -replace '^9(\d{10,})$','$1'
                                if ($TelFinNorm -ne "Desconocido" -and $TelFinNorm -ne "") { $PhoneYaEnInicio.Remove($TelFinNorm) | Out-Null }
                                $SesionHoraInicio.Remove($Ses) | Out-Null   # limpiar tracking de UpdateHistoryRecord
                                $_PrimerFonoSes.Remove($Ses) | Out-Null     # liberar 1er-número (el ID de sesión se recicla)

                                if (-not $CuelguesVistos[$FirmaUnica]) {
                                    $CuelguesVistos[$FirmaUnica]=$true
                                    if ($TelActual -match "^\+?564(3\d{5})\d{6}$") {
                                        # Cierre de la llamada FAC de firma: evento interno sin valor para el análisis.
                                        # Antes pintaba "Entry.ConnectinoID=N" (resto de la 1ª versión); además esa fila
                                        # con contenido reseteaba $RecienFirmado y destapaba el DEFAULT automático.
                                        $EventosTiempo[$HoraLimpia].Tel = "-"
                                    } else {
                                        $SlotFin = "$HoraLimpia,$MsLimpio5"
                                        Init-Hora $SlotFin
                                        $SuprimirSlotFin = $false
                                        $EventosTiempo[$SlotFin].Sesion=$Ses
                                        $EventosTiempo[$SlotFin].Tel=$TelActual
                                        if ($CuelguesManuales[$Ses]) {
                                            if ($TelActual -match "Desconocido" -and $DirLlamada[$Ses] -ne "ENTRANTE") {
                                                $EventosTiempo[$SlotFin].Interpretacion="$symStop CUELGUE MANUAL (Línea abierta sin marcar)"
                                                $EventosTiempo[$SlotFin].ColorInterpretacion=[System.Drawing.Color]::Orange
                                            } else {
                                                $EventosTiempo[$SlotFin].Interpretacion="$symStop FIN DE LLAMADA MANUAL (Colgada por el Asesor)"
                                                $EventosTiempo[$SlotFin].ColorInterpretacion=[System.Drawing.Color]::LightCoral
                                            }
                                        } else {
                                            if ($TelActual -match "Desconocido" -and $DirLlamada[$Ses] -ne "ENTRANTE") {
                                            if ($ConsultaTransf.ContainsKey($Ses) -or $ConsultaConf.ContainsKey($Ses)) {
                                                # Tema 8: leg de consulta phantom del mecanismo de transferencia/conferencia.
                                                # No es un cierre real (el resultado lo indican FALLO/COMPLETADA), así que se
                                                # suprime la fila por completo en vez de mostrar "Consulta sin completar".
                                                $SuprimirSlotFin = $true
                                            } else {
                                                $EventosTiempo[$SlotFin].Interpretacion="¡EVASIÓN! Línea abandonada sin marcar (Timeout/Tapón)"
                                                $EventosTiempo[$SlotFin].ColorInterpretacion=[System.Drawing.Color]::Red
                                            }
                                            } else {
                                                $EventosTiempo[$SlotFin].Interpretacion="$symStop FIN DE LLAMADA NORMAL"
                                                $EventosTiempo[$SlotFin].ColorInterpretacion=[System.Drawing.Color]::DarkGray
                                            }
                                        }
                                        $EventosTiempo[$SlotFin].RawInterpretacion+="$linea`n"
                                        if ($SuprimirSlotFin -and $EventosTiempo[$SlotFin].Agente -eq "" -and $EventosTiempo[$SlotFin].AppLog -eq "" -and $EventosTiempo[$SlotFin].SysLog -eq "" -and $EventosTiempo[$SlotFin].Audio -eq "" -and $EventosTiempo[$SlotFin].Ispeac -eq "") {
                                            $EventosTiempo.Remove($SlotFin) | Out-Null
                                        }
                                    }
                                }
                            }
                            elseif ($linea -match "Message type= MuteMediaRequest") {
                                # Ms-slot propio: el mismo segundo puede tener FIN DE LLAMADA en otro slot.
                                # Dedup: múltiples MuteMediaRequest del mismo segundo (distintos threads) van al mismo slot.
                                if (-not $MuteSlotSeg.ContainsKey($HoraLimpia)) {
                                    $SlotMute = "$HoraLimpia,$MsLimpio5"; Init-Hora $SlotMute
                                    $EventosTiempo[$SlotMute].Interpretacion = "Mute activado"
                                    $EventosTiempo[$SlotMute].ColorInterpretacion = [System.Drawing.Color]::Yellow
                                    $EventosTiempo[$SlotMute].Agente = "x MUTE MANUAL (Silenció)"
                                    $EventosTiempo[$SlotMute].ColorAgente = [System.Drawing.Color]::Orange
                                    $MuteSlotSeg[$HoraLimpia] = $SlotMute
                                }
                                $EventosTiempo[$MuteSlotSeg[$HoraLimpia]].RawAgente += "$linea`n"
                            }
                            elseif ($linea -match "Message type= UnMuteMediaRequest") {
                                if (-not $UnmuteSlotSeg.ContainsKey($HoraLimpia)) {
                                    $SlotUnmute = "$HoraLimpia,$MsLimpio5"; Init-Hora $SlotUnmute
                                    $EventosTiempo[$SlotUnmute].Interpretacion = "Mute desactivado"
                                    $EventosTiempo[$SlotUnmute].ColorInterpretacion = [System.Drawing.Color]::Yellow
                                    $EventosTiempo[$SlotUnmute].Agente = "o UNMUTE MANUAL (Abrió)"
                                    $EventosTiempo[$SlotUnmute].ColorAgente = [System.Drawing.Color]::OrangeRed
                                    $UnmuteSlotSeg[$HoraLimpia] = $SlotUnmute
                                }
                                $EventosTiempo[$UnmuteSlotSeg[$HoraLimpia]].RawAgente += "$linea`n"
                            }
                            elseif ($linea -match "Message type= TransferSessionRequest" -or $linea -match "OnRequestTransferSession") {
                                # Dedup: si ya existe la fila de inicio de transferencia en este mismo segundo
                                # (slot base o ms), no crear otra. TransferSessionRequest y OnRequestTransferSession
                                # pueden disparar en el mismo segundo y producían dos filas idénticas.
                                $YaHayTransfIni = $false
                                foreach ($kT in @($EventosTiempo.Keys)) {
                                    if (($kT -eq $HoraLimpia -or $kT -match "^$([regex]::Escape($HoraLimpia)),\d+$") -and $EventosTiempo[$kT].Agente -match "TRANSFERENCIA INICIADA|Asesor presiona bot.n Transferir") { $YaHayTransfIni = $true; break }
                                }
                                if (-not $YaHayTransfIni) {
                                    # Si hubo clic GUI (TransferCallHandler) en este segundo → fue el asesor.
                                    # Si no hubo clic → transferencia AUTOMÁTICA (sistema). El clic puede caer 1 s antes.
                                    $HuboClicTransf = $TransfClicSeg.ContainsKey($HoraLimpia)
                                    if (-not $HuboClicTransf) {
                                        try { $sPrevT = ([datetime]::ParseExact($HoraLimpia,"HH:mm:ss",$null).AddSeconds(-1)).ToString("HH:mm:ss"); if ($TransfClicSeg.ContainsKey($sPrevT)) { $HuboClicTransf = $true } } catch {}
                                    }
                                    if ($HuboClicTransf) { $EvA="$symArr Asesor presiona botón Transferir"; $ColorA=[System.Drawing.Color]::Plum }
                                    else                 { $EvA="$symArr TRANSFERENCIA INICIADA";           $ColorA=[System.Drawing.Color]::Plum }
                                }
                            }
                            elseif ($linea -match "Message type= MoveSessionToConferenceRequest" -or $linea -match "OnRequestMoveSessionToConference") { $EvA="↔ CONFERENCIA INICIADA"; $ColorA=[System.Drawing.Color]::Orchid }
                            elseif ($linea -match "Message type= LogoutRequest") { $EvA="¦ ASESOR SOLICITÓ DESFIRMARSE (Clic en Salir)"; $ColorA=[System.Drawing.Color]::LightCoral }
                            elseif ($linea -match "SetPhoneDisplay = <Transferencia realizada") {
                                # PBX señalizó "Transferencia realizada" → confirmar el slot provisional de Transfer_CompleteSetup.
                                # Si Transfer_CompleteSetup disparó en el mismo segundo, actualizar ese ms-slot;
                                # si no, escribir en el slot base (transferencia ciega sin Transfer_CompleteSetup previo).
                                $targetSlotTR = if ($SlotTransferenciaCompletada -and $SlotTransferenciaCompletada -match "^$([regex]::Escape($HoraLimpia)),") { $SlotTransferenciaCompletada } else { $HoraLimpia }
                                Init-Hora $targetSlotTR
                                $EventosTiempo[$targetSlotTR].Agente      = "$symOK TRANSFERENCIA COMPLETADA (PBX Confirmó la unión de sesiones)"
                                $EventosTiempo[$targetSlotTR].ColorAgente = [System.Drawing.Color]::LimeGreen
                                $EventosTiempo[$targetSlotTR].RawAgente  += "$linea`n"
                            }
                            elseif ($linea -match "Transfer_CompleteSetup: nfirstCall:\s*(\d+),\s*nSecondCall:\s*(\d+)") {
                                # nfirstCall = sesión del cliente (la llamada original que estaba en espera)
                                # nSecondCall = sesión de consulta (el destino al que se transfirió)
                                # Transfer_CompleteSetup dispara tanto en transferencias exitosas COMO fallidas (antes del Timeout).
                                # Por eso se escribe label PROVISIONAL — se confirma con SetPhoneDisplay o revierte con Transfer_Timeout.
                                $_nF = $matches[1]; $_nS = $matches[2]
                                $_telOrigen = if ($MapeoTel[$_nF]) { $MapeoTel[$_nF] } else { "Sesión $_nF" }
                                $_telDest   = if ($ConsultaTransf[$_nS]) { $ConsultaTransf[$_nS] } elseif ($MapeoTel[$_nS]) { $MapeoTel[$_nS] } else { "Sesión $_nS" }
                                $SlotTransferenciaCompletada = "$HoraLimpia,$MsLimpio5"
                                Init-Hora $SlotTransferenciaCompletada
                                if ($_nF -eq $_nS) {
                                    # nfirstCall == nSecondCall: la 2ª pata (consulta/destino) no existe → PBX no puede unir
                                    # → terminará en Timeout. DOS causas distintas, mismo síntoma:
                                    #   • SIN DESTINO  = "To session id" apuntó al mismo cliente → nunca se marcó a quién transferir (azartillos).
                                    #   • DESTINO COLGÓ = hubo una consulta separada con número que colgó antes de completar (ialzalden).
                                    $TransferConsultaColgada = $true
                                    if ($TransferSinDestinoPorSes[$_nF]) {
                                        $TransferSinDestino = $true
                                        $EventosTiempo[$SlotTransferenciaCompletada].Agente = "○ TRANSFERENCIA SIN DESTINO (no se marcó a quién transferir)"
                                    } else {
                                        $TransferSinDestino = $false
                                        $EventosTiempo[$SlotTransferenciaCompletada].Agente = "○ TRANSFERENCIA EN PROCESO (el destino de la consulta ya no responde)"
                                    }
                                } else {
                                    $TransferConsultaColgada = $false; $TransferSinDestino = $false
                                    $EventosTiempo[$SlotTransferenciaCompletada].Agente = "○ TRANSFERENCIA EN PROCESO: $_telOrigen → $_telDest"
                                }
                                $EventosTiempo[$SlotTransferenciaCompletada].ColorAgente = [System.Drawing.Color]::Orange
                                $EventosTiempo[$SlotTransferenciaCompletada].RawAgente  += "$linea`n"
                            }
                            elseif ($linea -match "Conference_CompleteConf: nFirstCall:\s*(\d+),\s*nConsultCall:\s*(\d+)") {
                                # Provisional: PBX notificó unión a nivel aplicación pero aún no confirmó merge de audio.
                                # El label se actualiza a ESTABLECIDA solo cuando llegue Conference_Merged.
                                $SlotConferenciaEstablecida = "$HoraLimpia,$MsLimpio5"
                                Init-Hora $SlotConferenciaEstablecida
                                $EventosTiempo[$SlotConferenciaEstablecida].Agente      = "○ CONFERENCIA SIN MERGE (Sesión $($matches[1]) + Sesión $($matches[2]))"
                                $EventosTiempo[$SlotConferenciaEstablecida].ColorAgente = [System.Drawing.Color]::Orange
                                $EventosTiempo[$SlotConferenciaEstablecida].RawAgente  += "$linea`n"
                            }
                            elseif ($linea -match "Conference_ActivateConsultCall: firstCall:\s*(\d+),\s*consultCall:\s*(\d+)") { $EvA="↔ CONFERENCIA EN PROCESO (Sesión $($matches[1]) + Sesión $($matches[2]))"; $ColorA=[System.Drawing.Color]::Orchid }
                            elseif ($linea -match "Transfer_ActivateConsultCall:\s*(\d+)") { $EvA="↔ TRANSFERENCIA EN PROCESO (Sesión $($matches[1]) en espera)"; $ColorA=[System.Drawing.Color]::Plum }
                            # Transfer_Timeout: el endpoint detectó que el PBX no confirmó la transferencia atendida.
                            # Si Transfer_CompleteSetup disparó provisional en el mismo segundo → revertir ese slot a FALLO.
                            # Si no hay slot provisional en el mismo segundo → crear ms-slot nuevo para el fallo.
                            elseif ($linea -match "Transfer_Timeout entered") {
                                $SlotFalloTransf = if ($SlotTransferenciaCompletada -and $SlotTransferenciaCompletada -match "^$([regex]::Escape($HoraLimpia)),") { $SlotTransferenciaCompletada } else { "$HoraLimpia,$MsLimpio5" }
                                Init-Hora $SlotFalloTransf
                                if ($TransferSinDestino) {
                                    $EventosTiempo[$SlotFalloTransf].Agente = "$symStop FALLO TRANSFERENCIA: no se marcó ningún destino (se dio Transferir sin marcar a quién)"
                                } elseif ($TransferConsultaColgada) {
                                    $EventosTiempo[$SlotFalloTransf].Agente = "$symStop FALLO TRANSFERENCIA: el destino de la consulta colgó antes de completar"
                                } else {
                                    $EventosTiempo[$SlotFalloTransf].Agente = "$symStop FALLO: TRANSFERENCIA NO COMPLETADA (Timeout PBX)"
                                }
                                $EventosTiempo[$SlotFalloTransf].ColorAgente = [System.Drawing.Color]::Red
                                $EventosTiempo[$SlotFalloTransf].RawAgente  += "$linea`n"
                                $TransferConsultaColgada = $false; $TransferSinDestino = $false
                            }

                            if ($EvA -ne "") {
                                # Categoría A → slot ms: los eventos discretos del asesor (HOLD/UNHOLD/TRANSFERENCIA/
                                # CONFERENCIA/CUELGUE/DESFIRMARSE) van a su ms real para no descuadrar el orden dentro
                                # del segundo. Heredamos Sesión/Teléfono que el sitio de creación dejó en el slot base
                                # (si no, quedarían separados del Agente). El slot base sin columnas visibles se filtra.
                                $ms5A = if ($MsLimpio5 -match "^\d+$") { [int]$MsLimpio5 } else { 0 }
                                $SlotEvA = "$HoraLimpia,$($ms5A.ToString('000'))"
                                Init-Hora $SlotEvA
                                if ($EventosTiempo.ContainsKey($HoraLimpia)) {
                                    if ($EventosTiempo[$SlotEvA].Sesion -eq "-" -and $EventosTiempo[$HoraLimpia].Sesion -ne "-") { $EventosTiempo[$SlotEvA].Sesion = $EventosTiempo[$HoraLimpia].Sesion }
                                    if ($EventosTiempo[$SlotEvA].Tel    -eq "-" -and $EventosTiempo[$HoraLimpia].Tel    -ne "-") { $EventosTiempo[$SlotEvA].Tel    = $EventosTiempo[$HoraLimpia].Tel }
                                }
                                $esFinal = $EvA -match "TRANSFERENCIA COMPLETADA"
                                if ($esFinal -or $EventosTiempo[$SlotEvA].Agente -notmatch "HOLD|UNHOLD|CUELGUE MANUAL|MUTE|TRANSFERENCIA INICIADA|Asesor presiona bot.n Transferir|CONFERENCIA INICIADA|DESFIRMARSE|TRANSFERENCIA COMPLETADA") {
                                    $EventosTiempo[$SlotEvA].Agente=$EvA; $EventosTiempo[$SlotEvA].ColorAgente=$ColorA
                                    $EventosTiempo[$SlotEvA].RawAgente+="$linea`n"
                                } else {
                                    # El ms-slot ya trae un evento de este tipo — buscar el siguiente ms libre para no perderlo.
                                    $SlotEvt5 = $SlotEvA
                                    for ($m5 = $ms5A + 1; $m5 -le 999; $m5++) { $cand5 = "$HoraLimpia,$m5"; if (-not $EventosTiempo.ContainsKey($cand5) -or $EventosTiempo[$cand5].Agente -eq "") { $SlotEvt5 = $cand5; Init-Hora $SlotEvt5; break } }
                                    if ($EventosTiempo.ContainsKey($HoraLimpia)) {
                                        if ($EventosTiempo[$SlotEvt5].Sesion -eq "-" -and $EventosTiempo[$HoraLimpia].Sesion -ne "-") { $EventosTiempo[$SlotEvt5].Sesion = $EventosTiempo[$HoraLimpia].Sesion }
                                        if ($EventosTiempo[$SlotEvt5].Tel    -eq "-" -and $EventosTiempo[$HoraLimpia].Tel    -ne "-") { $EventosTiempo[$SlotEvt5].Tel    = $EventosTiempo[$HoraLimpia].Tel }
                                    }
                                    $EventosTiempo[$SlotEvt5].Agente=$EvA; $EventosTiempo[$SlotEvt5].ColorAgente=$ColorA
                                    $EventosTiempo[$SlotEvt5].RawAgente+="$linea`n"
                                }
                            }
                            # Conference_Merged: el switch de audio confirma el merge físico de los dos legs.
                            # Actualiza el label provisional de Conference_CompleteConf a ESTABLECIDA confirmada.
                            if ($linea -match "Conference_Merged: nFirstCall:\s*(\d+)\s*nConsultCall:\s*(\d+)") {
                                # Guardar grupos ANTES del -match interno que sobreescribiría $matches
                                $_mrgF = $matches[1]; $_mrgC = $matches[2]
                                $targetMerged = if ($SlotConferenciaEstablecida -and $SlotConferenciaEstablecida -match "^$([regex]::Escape($HoraLimpia)),") { $SlotConferenciaEstablecida } else { $HoraLimpia }
                                $EventosTiempo[$targetMerged].Agente      = "$symOK CONFERENCIA ESTABLECIDA (Sesión $_mrgF + Sesión $_mrgC)"
                                $EventosTiempo[$targetMerged].ColorAgente = [System.Drawing.Color]::MediumOrchid
                                $EventosTiempo[$targetMerged].RawAgente  += "$linea`n"
                            }
                            # ConferenceError: el PBX rechazó la unión física. Actualiza el slot provisional a FALLIDA.
                            # Aparece tras Conference_CompleteConf + bFailureTrigger=1 — siempre sin Conference_Merged.
                            if ($linea -match "ConferenceError: firstCall:\s*(\d+),\s*secondCall:\s*(\d+)") {
                                $_errF = $matches[1]; $_errS = $matches[2]
                                $targetErr = if ($SlotConferenciaEstablecida -and $SlotConferenciaEstablecida -match "^$([regex]::Escape($HoraLimpia)),") { $SlotConferenciaEstablecida } else { "$HoraLimpia,$MsLimpio5" }
                                Init-Hora $targetErr
                                $EventosTiempo[$targetErr].Agente      = "$symStop CONFERENCIA FALLIDA: PBX rechazó la unión (Sesión $_errF + Sesión $_errS)"
                                $EventosTiempo[$targetErr].ColorAgente = [System.Drawing.Color]::Red
                                $EventosTiempo[$targetErr].RawAgente  += "$linea`n"
                                $SlotConferenciaEstablecida = ""
                            }
                            # closeSignalingChannel: cierre del canal de señalización. OJO: distinguir POR QUÉ se cerró:
                            #   • category:0, code:0 + "closed by the application"  = CIERRE LIMPIO (el asesor cerró/desfirmó) → NO es caída.
                            #   • category:2, code:10060 (WSAETIMEDOUT)             = CAÍDA DE RED real → sí es alerta.
                            # Antes esta línea marcaba CUALQUIER closeSignalingChannel como "Caída de Túnel" (falso positivo
                            # en todo cierre normal). Ahora solo alerta cuando el cierre NO es limpio.
                            if ($linea -match "closeSignalingChannel:\s*category:\s*(\d+),\s*code:\s*(\d+)" -and $linea -notmatch "Server sent a URQ") {
                                $catCS = [int]$matches[1]; $codeCS = [int]$matches[2]
                                if ($catCS -ne 0 -or $codeCS -ne 0) {
                                    $EvApp="¡AVAYA ALERTA!: Caída de Túnel Principal"; $ColorApp=[System.Drawing.Color]::Red
                                    $RedCaidaList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio5"; Raw = $linea }
                                }
                            }
                            # --- Caso "logeo autónomo": señales de PASO 5 (captura en listas, se emiten en barrido post-PASO5) ---
                            # Desfirme manual (el asesor eligió "Cerrar sesión" o lo disparó el cierre con la X):
                            if ($linea -match "Logoff:.*requestor='manual'") { $DesfirmeManualList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio5"; Raw = $linea } }
                            # Caída de red sostenida (keepalive al gatekeeper falla repetidamente):
                            if ($linea -match "RASKeepaliveFailed|RASKeepAliveFailed") { $RedCaidaList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio5"; Raw = $linea } }
                            # Recuperación de enlace en curso (Avaya intenta reconectar solo, mientras la app sigue abierta):
                            if ($linea -match "LinkRecoveryProgressEvent") { $ReconIntentoList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio5"; Raw = $linea } }
                            # Reconexión AUTOMÁTICA exitosa: bLInkRecovery=1 = recuperación de enlace (la app estaba ABIERTA, NO login manual):
                            if ($linea -match "CompleteLoginRequest:.*bLInkRecovery=1") { $ReconOkList += [pscustomobject]@{ Slot = "$HoraLimpia,$MsLimpio5"; Raw = $linea } }
                            # LÍNEA ABIERTA SIN MARCAR (confirmación definitiva). Las 3 líneas llegan consecutivas:
                            #   CHistoryManager::EndRecord: nSessionId = N
                            #   EndHistoryRecord: nSessionID = N
                            #   this record has no far-end address     ← solo si NUNCA se marcó un destino
                            if ($linea -match "EndHistoryRecord:\s*nSessionID\s*=\s*(\d+)" -or $linea -match "CHistoryManager::EndRecord:\s*nSessionId\s*=\s*(\d+)") {
                                $_endHistSes = $matches[1]; $_endHistSlot = "$HoraLimpia,$MsLimpio5"
                            }
                            elseif ($linea -match "this record has no far-end address" -and $_endHistSes) {
                                $SinDestinoFin += [pscustomobject]@{ Ses = $_endHistSes; Slot = $_endHistSlot }
                                $_endHistSes = $null
                            }
                            # DESCONEXIÓN FORZADA POR EL SERVIDOR: otro dispositivo registró la MISMA extensión.
                            # El CM expulsa a este endpoint con URQ reason 2009 "Unregistering User moved" +
                            # "CompleteLogoffRequest: Force logoff by server". Causa real de "el OneX me desconectó".
                            # Las 3 líneas caen en el mismo segundo → $DescForzadaSlot deja UNA sola fila.
                            if ($linea -match "reason (?:code|text)=\s*2009|Unregistering User moved|Force logoff by server") {
                                if (-not $DescForzadaSlot.ContainsKey($HoraLimpia)) {
                                    $ms5D = if ($MsLimpio5 -match "^\d+$") { [int]$MsLimpio5 } else { 0 }
                                    $SlotDesc = "$HoraLimpia,$($ms5D.ToString('000'))"; Init-Hora $SlotDesc
                                    $EventosTiempo[$SlotDesc].SysLog   = "OTRA SESIÓN TOMÓ LA EXTENSIÓN (desconexión forzada por el servidor — User moved)"
                                    $EventosTiempo[$SlotDesc].ColorSys = [System.Drawing.Color]::Red
                                    if ($EventosTiempo[$SlotDesc].Tel -eq "-") { $EventosTiempo[$SlotDesc].Tel = "RED/AVAYA" }
                                    $DescForzadaSlot[$HoraLimpia] = $SlotDesc
                                }
                                $EventosTiempo[$DescForzadaSlot[$HoraLimpia]].RawSysLog += "$linea`n"
                            }
                            # ProcessTransferTimeout: confirmación explícita del endpoint de que la transferencia atendida
                            # falló (el PBX no procesó la unión de sesiones). Se muestra en AppLog como evidencia del error.
                            if ($linea -match "ProcessTransferTimeout:.*failed") { $EvApp="¡FALLO TRANSFERENCIA!: El PBX rechazó la operación (ProcessTransferTimeout)"; $ColorApp=[System.Drawing.Color]::Red }
                            if ($EvApp -ne "") {
                                # Categoría A → slot ms: fallo/alerta puntual, a su ms real para no descuadrar.
                                $ms5App = if ($MsLimpio5 -match "^\d+$") { [int]$MsLimpio5 } else { 0 }
                                $SlotApp = "$HoraLimpia,$($ms5App.ToString('000'))"; Init-Hora $SlotApp
                                $EventosTiempo[$SlotApp].AppLog=$EvApp; $EventosTiempo[$SlotApp].ColorApp=$ColorApp; if($EventosTiempo[$SlotApp].Tel -eq "-"){$EventosTiempo[$SlotApp].Tel="RED/AVAYA"}; $EventosTiempo[$SlotApp].RawAppLog+="$linea`n"
                            }
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
            $MarcandoPorSes = @{}   # sesión de consulta → slot: garantiza UN SOLO "MARCANDO" por sesión
            foreach ($HoraLA in $LimHoras) {
                if (-not $EventosTiempo.ContainsKey($HoraLA)) { continue }
                $ObjLA = $EventosTiempo[$HoraLA]
                $SesLA = $ObjLA.Sesion
                if (-not $SesLA -or $SesLA -eq "-" -or $SesLA -eq "") { continue }
                # helper: cuando el slot destino ya tiene INICIO DE LLAMADA, no sobreescribir —
                # crear un slot adyacente (ms+1) para que ambos eventos sean visibles.
                function Get-SlotConsulta ($horaBase, $sesion, $tel) {
                    if ($horaBase -match "^(\d{2}:\d{2}:\d{2}),(\d+)$") {
                        $hb = $matches[1]; $msBase = [int]$matches[2]
                    } else {
                        $hb = $horaBase; $msBase = 0
                    }
                    for ($ms = $msBase + 1; $ms -le 999; $ms++) {
                        $cand = "$hb,$ms"
                        if (-not $EventosTiempo.ContainsKey($cand)) {
                            Init-Hora $cand
                            $EventosTiempo[$cand].Sesion = $sesion
                            $EventosTiempo[$cand].Tel    = $tel
                            return $cand
                        }
                    }
                    return $horaBase  # fallback: sobreescribir si no hay slot libre
                }

                # ① Consulta de Conferencia — SIEMPRE verifica antes que Transferencia.
                # Avaya OneX crea la leg de conferencia internamente como si fuera transferencia,
                # por lo que una sesión puede quedar en ambas tablas. Conferencia tiene prioridad.
                if ($ConsultaConf.ContainsKey($SesLA) -and
                    $ObjLA.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR|INICIO DE LLAMADA|INICIO DE SESIÓN") {
                    $SlotConf = if ($ObjLA.Interpretacion -match "INICIO DE LLAMADA") { Get-SlotConsulta $HoraLA $SesLA $ObjLA.Tel } else { $HoraLA }
                    # Si el slot quedó en base (sin ms) y hay ms-slots en ese segundo (ej: Drag/Drop,
                    # CONFERENCIA INICIADA), moverlo al final para que aparezca en orden cronológico.
                    if ($SlotConf -notmatch ",") {
                        $BaseHora = $SlotConf; $msMax = 0
                        foreach ($k in @($EventosTiempo.Keys)) {
                            if ($k -match "^$([regex]::Escape($BaseHora)),(\d+)$" -and [int]$matches[1] -gt $msMax) { $msMax = [int]$matches[1] }
                        }
                        if ($msMax -gt 0) {
                            $SlotMs = "$BaseHora,$($msMax + 1)"
                            Init-Hora $SlotMs
                            $Src = $EventosTiempo[$BaseHora]; $Dst = $EventosTiempo[$SlotMs]
                            foreach ($campo in @("Sesion","Tel","ViId","Topic",
                                                  "Agente","ColorAgente","Audio","ColorAudio","Aux","ColorAux",
                                                  "Ispeac","ColorIspeac","SysLog","ColorSys","AppLog","ColorApp",
                                                  "RawInterpretacion","RawAgente","RawAudio","RawAux","RawIspeac","RawSysLog","RawAppLog")) {
                                $Dst[$campo] = $Src[$campo]
                            }
                            $EventosTiempo.Remove($BaseHora) | Out-Null
                            $SlotConf = $SlotMs
                        }
                    }
                    $EventosTiempo[$SlotConf].Interpretacion = "$symArr CONSULTA DE CONFERENCIA $symArr $($ConsultaConf[$SesLA])"
                    $EventosTiempo[$SlotConf].ColorInterpretacion = [System.Drawing.Color]::Orchid
                    continue
                }
                # ② Consulta de Transferencia
                if ($ConsultaTransf.ContainsKey($SesLA) -and
                    $ObjLA.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR|INICIO DE LLAMADA|INICIO DE SESIÓN") {
                    # La apertura de línea para consultar genera VARIAS filas de la misma sesión (una en
                    # slot base sin ms + el INICIO en ms). Antes cada una se reetiquetaba a "MARCANDO",
                    # produciendo DOS filas y una de ellas antes del INICIO (slot base ordena primero).
                    # ¿La sesión tiene un INICIO DE LLAMADA real en otro slot?
                    $SesTieneInicioLlam = $false
                    foreach ($kIni in @($EventosTiempo.Keys)) {
                        if ($EventosTiempo[$kIni].Sesion -eq $SesLA -and $EventosTiempo[$kIni].Interpretacion -match "INICIO DE LLAMADA") { $SesTieneInicioLlam = $true; break }
                    }
                    # Fila duplicada/huérfana (LÍNEA ABIERTA / INICIO DE SESIÓN) cuya sesión YA tiene un
                    # INICIO DE LLAMADA → limpiarla; el MARCANDO se pinta a partir del INICIO (queda DESPUÉS).
                    if ($ObjLA.Interpretacion -notmatch "INICIO DE LLAMADA" -and $SesTieneInicioLlam) {
                        $ObjLA.Interpretacion = ""; $ObjLA.ColorInterpretacion = [System.Drawing.Color]::White
                        $ObjLA.RawInterpretacion += "[DEDUP MARCANDO] fila de consulta duplicada (sesión $SesLA ya tiene INICIO) — MARCANDO se pinta tras el INICIO`n"
                        continue
                    }
                    # Dedup: un solo "MARCANDO" por sesión de consulta.
                    if ($MarcandoPorSes.ContainsKey($SesLA)) {
                        $ObjLA.Interpretacion = ""; $ObjLA.ColorInterpretacion = [System.Drawing.Color]::White
                        continue
                    }
                    $destTr = $ConsultaTransf[$SesLA]
                    # Antes "CONSULTA DE TRANSFERENCIA →": impreciso (el asesor a veces solo consulta y no transfiere).
                    # Ahora describe la acción real: se está marcando al destino (2ª línea durante la llamada).
                    $SlotTransf = if ($ObjLA.Interpretacion -match "INICIO DE LLAMADA") { Get-SlotConsulta $HoraLA $SesLA $ObjLA.Tel } else { $HoraLA }
                    # Si el slot quedó en base (sin ms) y hay ms-slots en ese segundo, moverlo al final del
                    # segundo para que aparezca en orden cronológico (mismo criterio que ① Conferencia).
                    if ($SlotTransf -notmatch ",") {
                        $BaseHora = $SlotTransf; $msMax = 0
                        foreach ($k in @($EventosTiempo.Keys)) { if ($k -match "^$([regex]::Escape($BaseHora)),(\d+)$" -and [int]$matches[1] -gt $msMax) { $msMax = [int]$matches[1] } }
                        if ($msMax -gt 0) {
                            $SlotMs = "$BaseHora,$($msMax + 1)"; Init-Hora $SlotMs
                            $Src = $EventosTiempo[$BaseHora]; $Dst = $EventosTiempo[$SlotMs]
                            foreach ($campo in @("Sesion","Tel","ViId","Topic","Agente","ColorAgente","Audio","ColorAudio","Aux","ColorAux","Ispeac","ColorIspeac","SysLog","ColorSys","AppLog","ColorApp","RawInterpretacion","RawAgente","RawAudio","RawAux","RawIspeac","RawSysLog","RawAppLog")) { $Dst[$campo] = $Src[$campo] }
                            $EventosTiempo.Remove($BaseHora) | Out-Null
                            $SlotTransf = $SlotMs
                        }
                    }
                    $EventosTiempo[$SlotTransf].Interpretacion = if ($destTr -and $destTr -ne "Desconocido") { "$symArr MARCANDO A LA EXT. $destTr" } else { "$symArr MARCANDO (consulta durante la llamada)" }
                    $EventosTiempo[$SlotTransf].ColorInterpretacion = [System.Drawing.Color]::Plum
                    $MarcandoPorSes[$SesLA] = $SlotTransf
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
            # REFINAMIENTO "LÍNEA ABIERTA SIN MARCAR" (claridad / quitar ruido)
            # La etiqueta confundía porque aparecía aunque la fila ya tuviera un evento de
            # transferencia o la misma sesión hubiera marcado un número — y "sin marcar"
            # suena a evasión cuando muchas veces es una apertura de línea intencional.
            # IMPORTANTE: la EVASIÓN real se detecta aparte en ProcessSessionEndedEvent
            # (¡EVASIÓN!), independiente de esta etiqueta, y la apertura de línea queda
            # marcada por PressLineAppearance ("↗ Apertura de línea desde botón"). Aquí solo
            # se limpian los casos que NO son evasión; el caso phantom sin número se conserva.
            # ================================================================
            $SegDeSlot = {
                param($s)
                $p = ($s -split ',')[0]
                try { return [int]([datetime]::ParseExact($p,"HH:mm:ss",$null).TimeOfDay.TotalSeconds) } catch { return -1 }
            }
            # Milisegundos absolutos de un slot "HH:mm:ss[,fff]" (para comparar eventos con precisión de ms)
            $MsDeSlot = {
                param($s)
                if     ($s -match "^(\d{2}:\d{2}:\d{2}),(\d{1,3})$") { $hh = $matches[1]; $mm = [int]$matches[2] }
                elseif ($s -match "^(\d{2}:\d{2}:\d{2})$")           { $hh = $matches[1]; $mm = 0 }
                else { return -1 }
                try { return [int]([datetime]::ParseExact($hh,"HH:mm:ss",$null).TimeOfDay.TotalSeconds) * 1000 + $mm } catch { return -1 }
            }
            # Segundos totales -> "mm:ss" para la duración que se muestra en FIN DE LLAMADA (saliente y entrante).
            # OJO: [int] en PowerShell REDONDEA un double (banker's rounding), no trunca -> [int](90/60) da 2, no 1.
            # Por eso los minutos se calculan con [math]::Floor (trunca de verdad hacia abajo).
            $FmtMmSs = { param($totalSeg) "{0:D2}:{1:D2}" -f ([int][math]::Floor($totalSeg / 60)), ($totalSeg % 60) }
            foreach ($hLA2 in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($hLA2)) { continue }
                $oLA = $EventosTiempo[$hLA2]
                if ($oLA.Interpretacion -notmatch "LÍNEA ABIERTA SIN MARCAR") { continue }

                # (1) La fila ya trae un evento de transferencia → la LÍNEA ABIERTA es ruido del
                #     mecanismo; se quita la etiqueta y la fila conserva el evento real.
                if ($oLA.Agente -match "TRANSFERENCIA INICIADA|Asesor presiona bot.n Transferir|TRANSFERENCIA EN PROCESO|TRANSFERENCIA COMPLETADA") {
                    $oLA.Interpretacion = ""; $oLA.ColorInterpretacion = [System.Drawing.Color]::White
                    continue
                }

                # (2) La misma sesión ya tiene un INICIO DE LLAMADA dentro de ±20 s → esta fila es
                #     el instante previo a marcar; redundante (el número SÍ se marcó). Se suprime.
                #     La ventana de 20 s evita pisar evasiones reales si el SessionId se reutiliza
                #     más tarde para otra llamada distinta.
                $TieneInicioSes = $false
                if ($oLA.Sesion -and $oLA.Sesion -ne "-") {
                    $segLA = & $SegDeSlot $hLA2
                    foreach ($kI2 in @($EventosTiempo.Keys)) {
                        if ($kI2 -eq $hLA2) { continue }
                        if ($EventosTiempo[$kI2].Sesion -eq $oLA.Sesion -and $EventosTiempo[$kI2].Interpretacion -match "INICIO DE LLAMADA") {
                            $segIni = & $SegDeSlot $kI2
                            if ($segLA -ge 0 -and $segIni -ge 0 -and [math]::Abs($segIni - $segLA) -le 20) { $TieneInicioSes = $true; break }
                        }
                    }
                }
                if ($TieneInicioSes) {
                    if ($oLA.Agente -eq "" -and $oLA.AppLog -eq "" -and $oLA.SysLog -eq "" -and $oLA.Audio -eq "" -and $oLA.Ispeac -eq "") {
                        $EventosTiempo.Remove($hLA2) | Out-Null
                    } else {
                        $oLA.Interpretacion = ""; $oLA.ColorInterpretacion = [System.Drawing.Color]::White
                    }
                    continue
                }
                # (3) else: línea abierta sin número y sin marcación posterior → posible EVASIÓN;
                #     se conserva "LÍNEA ABIERTA SIN MARCAR" (el cierre confirmará con ¡EVASIÓN!).
            }

            # ================================================================
            # REFINAR "Apertura de línea desde botón" (PressLineAppearance):
            #  (P3) Si cae en el mismo segundo (±1 s) que el fin de un Transfer, NO fue una apertura
            #       manual: fue el sistema RETOMANDO la llamada original tras un fallo de transferencia.
            #  (P4) Si no, identificar el número de línea con la letra de call-appearance (a→1, b→2, c→3…).
            # ================================================================
            foreach ($kAL in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kAL)) { continue }
                $oAL = $EventosTiempo[$kAL]
                if ($oAL.Agente -notmatch "Apertura de línea desde botón") { continue }
                $segAL = ($kAL -split ',')[0]
                # (P3) ¿coincide con el fin de un Transfer? → retomado automático
                $EsRetomado = $false
                foreach ($off in -1,0,1) {
                    try { $sChk = ([datetime]::ParseExact($segAL,"HH:mm:ss",$null).AddSeconds($off)).ToString("HH:mm:ss") } catch { $sChk = $segAL }
                    if ($EndTransferSeg.ContainsKey($sChk)) { $EsRetomado = $true; break }
                }
                if ($EsRetomado) {
                    $oAL.Agente      = "$symRes Llamada retomada automáticamente (tras fallo de transferencia)"
                    $oAL.ColorAgente = [System.Drawing.Color]::LightGreen
                    continue
                }
                # (P4) número de línea desde la letra de call-appearance del mismo segundo
                $numLinea = 0
                if ($LineaAppPorSeg.ContainsKey($segAL)) {
                    $letra = $LineaAppPorSeg[$segAL]
                    if ($letra -match "^[a-z]$") { $numLinea = [int][char]$letra - [int][char]'a' + 1 }
                }
                # ABRE vs CAMBIA: si tras el clic (≤500 ms) NACE una llamada → abrió la línea.
                # Si en cambio una llamada sale del hold → solo cambió a una línea que ya estaba ocupada.
                $msAL  = & $MsDeSlot $kAL
                $Nace  = $false; $Retoma = $false
                if ($msAL -ge 0) {
                    foreach ($mc in $CallCreatedMs) { if ($mc -ge $msAL -and ($mc - $msAL) -le 500) { $Nace = $true; break } }
                    if (-not $Nace) {
                        foreach ($sR in $RetomaImplicita.Keys) {
                            foreach ($slR in $RetomaImplicita[$sR]) {
                                $mr = & $MsDeSlot $slR
                                if ($mr -ge $msAL -and ($mr - $msAL) -le 500) { $Retoma = $true; break }
                            }
                            if ($Retoma) { break }
                        }
                    }
                }
                # Se describe la ACCIÓN (presionar el foquito de línea), no el efecto — "abre línea"
                # se confundía con los eventos de "Audio abierto/cerrado". Si al presionar retoma una
                # llamada que estaba en espera, se añade ese matiz.
                $sufAL = if ($Retoma -and -not $Nace) { " (retoma la llamada en espera)" } else { "" }
                if ($numLinea -ge 1 -and $numLinea -le 20) { $oAL.Agente = "↗ Asesor presiona botón de línea $numLinea$sufAL" }
                else                                        { $oAL.Agente = "↗ Asesor presiona botón de línea$sufAL" }
            }

            # ================================================================
            # RELABEL "HOLD MANUAL" que FALLÓ: si el hold de una sesión no se completó
            # (HoldSessionCommand response null), su fila no debe mostrarse como exitosa.
            # Se marca la fila HOLD MANUAL / "Llamada en Hold" de esa sesión que esté DENTRO
            # de la ventana del timeout (≤20 s antes del fallo).
            # ================================================================
            if ($HoldFalloSes.Count -gt 0) {
                foreach ($kH in @($EventosTiempo.Keys)) {
                    if (-not $EventosTiempo.ContainsKey($kH)) { continue }
                    $oH = $EventosTiempo[$kH]
                    if (-not $HoldFalloSes.ContainsKey($oH.Sesion)) { continue }
                    if ($oH.Agente -notmatch "HOLD MANUAL" -and $oH.Interpretacion -notmatch "Llamada en Hold") { continue }
                    $slotFallo = $HoldFalloSes[$oH.Sesion]
                    $segH = ($kH -split ',')[0]; $segF = ($slotFallo -split ',')[0]
                    $okVentana = $false
                    try { $dt = ([datetime]::ParseExact($segF,'HH:mm:ss',$null) - [datetime]::ParseExact($segH,'HH:mm:ss',$null)).TotalSeconds; $okVentana = ($dt -ge 0 -and $dt -le 20) } catch {}
                    if (-not $okVentana) { continue }
                    if ($oH.Agente -match "HOLD MANUAL") {
                        $oH.Agente      = "$symStop HOLD MANUAL — ¡FALLÓ! (no se pudo retener la llamada)"
                        $oH.ColorAgente = [System.Drawing.Color]::Red
                    }
                    if ($oH.Interpretacion -match "Llamada en Hold") {
                        $oH.Interpretacion      = "$symStop Intento de HOLD que FALLÓ (no se pudo retener la llamada)"
                        $oH.ColorInterpretacion = [System.Drawing.Color]::Red
                    }
                }
            }

            # ================================================================
            # AUDIO → LÍNEA. "Línea abierta/cerrada" son en realidad la sesión de AUDIO (ISPEAC) y
            # OneX mantiene UNA sola a la vez (el "call id= 1" del AudioLog es el slot del canal, no la
            # llamada: siempre vale 1). Regla sin heurísticas ni umbrales: el ABIERTO toma la línea activa
            # en ese instante; el CERRADO hereda la línea del último ABIERTO (cierra justo lo que ese abrió).
            # Resuelve el caso confuso: al abrir la línea 2, activeLine ya es 8 pero el audio que se cierra
            # es el de la línea 1 (el cliente que quedó en espera).
            # ================================================================
            $SeqOrd = @($ActiveLineSeq | Sort-Object Ms)
            $LineaAudioActual = 0
            $AudioRows = @()
            foreach ($kAu in @($EventosTiempo.Keys | Sort-Object)) {
                if (-not $EventosTiempo.ContainsKey($kAu)) { continue }
                $oAu = $EventosTiempo[$kAu]
                if ($oAu.Audio -notmatch "Línea abierta|Línea cerrada") { continue }
                $msAu = & $MsDeSlot $kAu
                if ($oAu.Audio -match "Línea abierta") {
                    $btnAu = -1
                    foreach ($ev in $SeqOrd) { if ($ev.Ms -le $msAu) { $btnAu = $ev.Btn } else { break } }
                    if ($btnAu -ge 0 -and $LineaAppBtn.ContainsKey($btnAu)) { $LineaAudioActual = $LineaAppBtn[$btnAu] }
                    $oAu.Audio = if ($LineaAudioActual -ge 1) { "$symMusic Audio abierto (línea $LineaAudioActual)" } else { "$symMusic Audio abierto" }
                    $tipoAu = "A"
                } else {
                    $oAu.Audio = if ($LineaAudioActual -ge 1) { "$symStop Audio cerrado (línea $LineaAudioActual)" } else { "$symStop Audio cerrado" }
                    $tipoAu = "C"
                }
                $AudioRows += [pscustomobject]@{ Slot = $kAu; Ms = $msAu; Tipo = $tipoAu; Linea = $LineaAudioActual }
            }

            # RENEGOCIACIÓN DE MEDIOS: un "cerrado → abierto" en la MISMA línea con <1 s de diferencia no
            # es una acción del asesor: es Avaya cortando y reabriendo el canal H.245 (StopMedia + fast-start
            # OLC), típicamente al conectar con el destino. Se colapsan las dos filas en un solo evento claro.
            # (Un cerrado línea 1 → abierto línea 2 es un cambio de línea real y NO se colapsa.)
            for ($iAu = 0; $iAu -lt ($AudioRows.Count - 1); $iAu++) {
                $aC = $AudioRows[$iAu]; $aA = $AudioRows[$iAu + 1]
                $gap = $aA.Ms - $aC.Ms
                if ($aC.Tipo -eq "C" -and $aA.Tipo -eq "A" -and $aC.Linea -ge 1 -and $aC.Linea -eq $aA.Linea -and
                    $aC.Ms -ge 0 -and $gap -ge 0 -and $gap -le 1000) {
                    $oC = $EventosTiempo[$aC.Slot]; $oA = $EventosTiempo[$aA.Slot]
                    $oC.Audio      = "↻ Avaya renegoció el audio (línea $($aC.Linea))"
                    $oC.ColorAudio = [System.Drawing.Color]::DeepSkyBlue
                    $oC.RawAudio  += $oA.RawAudio
                    $oA.Audio      = ""; $oA.ColorAudio = [System.Drawing.Color]::White
                    $iAu++   # el par ya se consumió
                }
            }

            # ================================================================
            # DTMF: agrupar las pulsaciones del asesor en secuencias legibles.
            # Se agrupan las pulsaciones consecutivas de la MISMA sesión con hueco ≤12 s (el
            # asesor marca en ráfaga). Cada grupo → una fila en la columna DTMF con la secuencia
            # (ej: "⌨ #*1234567890"); el detalle por dígito con su ms queda en el RAW de la celda.
            # ================================================================
            if ($DtmfPresses.Count -gt 0) {
                # Emisión DIRECTA (sin arrays anidados, que en PS 5.1 se aplanan y fragmentan la secuencia).
                # Se acumula la secuencia y se "cierra" el grupo al cambiar de sesión o superar 12 s de hueco.
                $ordDt = @($DtmfPresses | Sort-Object { & $MsDeSlot $_.Hora })
                $seqDt = ""; $slotDt = $null; $sesDt = $null; $prevMsDt = -1; $rawDt = ""
                foreach ($prDt in $ordDt) {
                    $msDt = & $MsDeSlot $prDt.Hora
                    if ($seqDt -ne "" -and ($prDt.Ses -ne $sesDt -or ($msDt - $prevMsDt) -gt 12000)) {
                        Init-Hora $slotDt
                        $durSeqDt = [math]::Round(($prevMsDt - (& $MsDeSlot $slotDt)) / 1000)
                        $sufSeqDt = if ($durSeqDt -ge 1) { " (tecleado en ${durSeqDt}s)" } else { "" }
                        $EventosTiempo[$slotDt].Dtmf      = "⌨ $seqDt$sufSeqDt"
                        $EventosTiempo[$slotDt].ColorDtmf = [System.Drawing.Color]::Aqua
                        if ($EventosTiempo[$slotDt].Sesion -eq "-") { $EventosTiempo[$slotDt].Sesion = $sesDt }
                        $EventosTiempo[$slotDt].RawDtmf   = "Pulsaciones DTMF del asesor (sesión $sesDt), tecleado en ${durSeqDt}s:`n$rawDt".TrimEnd()
                        $seqDt = ""; $rawDt = ""
                    }
                    if ($seqDt -eq "") { $slotDt = $prDt.Hora; $sesDt = $prDt.Ses }
                    $seqDt += $prDt.Sim
                    $rawDt += "$($prDt.Hora)  →  $($prDt.Sim)`n"
                    $prevMsDt = $msDt
                }
                if ($seqDt -ne "") {
                    Init-Hora $slotDt
                    $durSeqDt = [math]::Round(($prevMsDt - (& $MsDeSlot $slotDt)) / 1000)
                    $sufSeqDt = if ($durSeqDt -ge 1) { " (tecleado en ${durSeqDt}s)" } else { "" }
                    $EventosTiempo[$slotDt].Dtmf      = "⌨ $seqDt$sufSeqDt"
                    $EventosTiempo[$slotDt].ColorDtmf = [System.Drawing.Color]::Aqua
                    if ($EventosTiempo[$slotDt].Sesion -eq "-") { $EventosTiempo[$slotDt].Sesion = $sesDt }
                    $EventosTiempo[$slotDt].RawDtmf   = "Pulsaciones DTMF del asesor (sesión $sesDt), tecleado en ${durSeqDt}s:`n$rawDt".TrimEnd()
                }
            }

            # ================================================================
            # "Asesor captura [número] y da Enter": inyectar el número tecleado.
            # Cada NewCallHandler (PASO 4) se empareja con el ApplyDialingRulesToNumber
            # más cercano posterior (≤8 s) del EndpointLog, que trae el número ya listo
            # para marcar (con prefijo de salida). Cada número se consume una sola vez.
            # ================================================================
            # Cubre DOS eventos que nacen igual (número tecleado + Enter, resuelto por ApplyDialingRulesToNumber):
            #   · NewCallHandler  → "Asesor captura N y da Enter"       (llamada nueva desde la caja)
            #   · AddCallHandler  → "El asesor agregó una llamada (consulta) al N"  (2ª llamada para transferir/conf.)
            # Se procesan juntos, ordenados por tiempo, consumiendo cada número una sola vez ($usadoAD).
            if (($NewCallSlots.Count + $AddCallSlots.Count) -gt 0 -and $DialedApply.Count -gt 0) {
                $ordAD = @($DialedApply | Sort-Object Ms)
                $usadoAD = @{}
                $candNum = @()
                foreach ($s in $NewCallSlots) { $candNum += [pscustomobject]@{ Slot = $s; Tipo = "NEW" } }
                foreach ($s in $AddCallSlots) { $candNum += [pscustomobject]@{ Slot = $s; Tipo = "ADD" } }
                foreach ($cn in ($candNum | Sort-Object { & $MsDeSlot $_.Slot })) {
                    $slotNC = $cn.Slot
                    if (-not $EventosTiempo.ContainsKey($slotNC)) { continue }
                    $agC = $EventosTiempo[$slotNC].Agente
                    if ($cn.Tipo -eq "NEW" -and $agC -notmatch "Asesor captura número y da Enter") { continue }
                    if ($cn.Tipo -eq "ADD" -and $agC -notmatch "agregó una llamada \(consulta\)") { continue }
                    $msNC = & $MsDeSlot $slotNC
                    if ($msNC -lt 0) { continue }
                    for ($iAD = 0; $iAD -lt $ordAD.Count; $iAD++) {
                        if ($usadoAD[$iAD]) { continue }
                        $dltAD = $ordAD[$iAD].Ms - $msNC
                        if ($dltAD -ge -500 -and $dltAD -le 8000) {
                            if ($cn.Tipo -eq "NEW") { $EventosTiempo[$slotNC].Agente = "$symUp Asesor captura $($ordAD[$iAD].Num) y da Enter" }
                            else                    { $EventosTiempo[$slotNC].Agente = "$symArr El asesor agregó una llamada (consulta) al $($ordAD[$iAD].Num)" }
                            $usadoAD[$iAD] = $true
                            break
                        }
                    }
                }
            }

            # ================================================================
            # CASO "LOGEO AUTÓNOMO": cierre normal vs caída de red + reconexión automática.
            # Un MISMO log puede tener VARIAS sesiones el mismo día → se COLAPSA cada lista por
            # episodios (nuevo episodio cuando el hueco supera el umbral) y se emite UNA fila por
            # episodio en "Log de Sistema". Verde = acción del asesor (cierre); naranja/amarillo =
            # caída/recuperación de red; cyan = reconexión AUTOMÁTICA (bLInkRecovery=1 → app ABIERTA).
            # ================================================================
            # Colapsa una lista de @{Slot;Raw} en episodios: devuelve el PRIMER elemento de cada racha
            # (parse de ms autónomo para no depender del scope externo).
            $ColapsarEpisodios = {
                param($lista, $gapMs)
                $conMs = foreach ($it in @($lista)) {
                    $ms = -1
                    if ($it.Slot -match '^(\d{2}:\d{2}:\d{2})[,:]?(\d{1,3})?$') {
                        try { $ms = [int]([datetime]::ParseExact($matches[1],'HH:mm:ss',$null).TimeOfDay.TotalSeconds) * 1000 + [int]($matches[2]) } catch { $ms = -1 }
                    }
                    [pscustomobject]@{ Slot = $it.Slot; Raw = $it.Raw; Ms = $ms }
                }
                $res = @(); $lastMs = -999999999
                foreach ($it in (@($conMs) | Where-Object { $_.Ms -ge 0 } | Sort-Object Ms)) {
                    if (($it.Ms - $lastMs) -gt $gapMs) { $res += $it }
                    $lastMs = $it.Ms
                }
                return @($res)
            }
            $GAP_CIERRE = 30000   # 30 s: separa cierres/desfirmes distintos (las señales internas de un cierre caen juntas)
            $GAP_RED    = 90000   # 90 s: un episodio de caída/recuperación dura minutos con huecos ~19 s → un solo episodio
            $cierreMsAll = @(@($CierreAppList) | ForEach-Object { & $MsDeSlot $_.Slot })

            # 1) Desfirme manual — SOLO si NO va pegado a un cierre de app (±15 s). Si hay cierre cerca, el
            #    desfirme es parte del cierre (la fila de CIERRE ya lo cuenta) → no se duplica. Un desfirme
            #    "solo" (el asesor se desfirma pero deja la app abierta) SÍ se muestra.
            foreach ($ep in (& $ColapsarEpisodios $DesfirmeManualList $GAP_CIERRE)) {
                $dm = & $MsDeSlot $ep.Slot
                $pegadoCierre = $false
                foreach ($cm in $cierreMsAll) { if ($cm -ge 0 -and [math]::Abs($cm - $dm) -le 15000) { $pegadoCierre = $true; break } }
                if ($pegadoCierre) { continue }
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].SysLog -eq "") {
                    $EventosTiempo[$ep.Slot].SysLog   = "✔ DESFIRME MANUAL: el asesor cerró sesión (dejó la app abierta)"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::LightGreen
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? Se cerró sesión con desfirme manual (requestor='manual') sin cerrar la aplicación:`n$($ep.Raw)`n"
                    if ($EventosTiempo[$ep.Slot].Tel -eq "-") { $EventosTiempo[$ep.Slot].Tel = "RED/AVAYA" }
                }
            }
            # 2) Cierre de la aplicación (una fila por episodio). Se distingue:
            #    • CIERRE NORMAL: el asesor se desfirmó (menú) y luego cerró. Sin EnterAuxHandler cercano.
            #    • CERRÓ CON X SIN DESFIRMARSE: la app tuvo que estacionarlo en Aux porque seguía DISPONIBLE
            #      al dar la X (EnterAuxHandler / oldState=Ready→LoggedOut dentro del episodio).
            $firmadoMs = @(@($CierreFirmadoList) | ForEach-Object { & $MsDeSlot $_.Slot })
            # EPISODIOS DE CIERRE: un cierre puede TARDAR y llevar VARIOS clics si la app se colgó (caso real:
            # 3 intentos en 48 s por el bug de auto-answer). Por eso el episodio NO se corta por ventana de
            # tiempo (30 s lo partía en dos filas contradictorias), sino que se cierra en el apagado definitivo
            # (Shutdown/PhoneService shutdown). Se cuentan los clics para reportar "N intentos".
            $CierreEpisodios = @(); $curIniC = $null; $curRawC = ""; $curIntC = 0; $curUltC = -1; $ultFinC = -999999
            foreach ($itC in (@(@($CierreAppList) | Sort-Object { & $MsDeSlot $_.Slot }))) {
                $mC = & $MsDeSlot $itC.Slot
                if ($mC -lt 0) { continue }
                # Corte de seguridad: si pasan >5 min sin apagado definitivo, se asume episodio abandonado.
                if ($null -ne $curIniC -and ($mC - $curUltC) -gt 300000) {
                    $CierreEpisodios += [pscustomobject]@{ Slot=$curIniC; Raw=$curRawC; Intentos=$curIntC; FinMs=$curUltC; Confirmado=$false }; $curIniC = $null
                }
                if ($null -eq $curIniC) {
                    # El apagado emite DOS señales de fin juntas (Shutdown() y PhoneService shutdown). La segunda
                    # es cola del mismo cierre, no un cierre nuevo → ignorarla en vez de abrir episodio fantasma.
                    if (($mC - $ultFinC) -le 5000) { $curUltC = $mC; continue }
                    $curIniC = $itC.Slot; $curRawC = $itC.Raw; $curIntC = 0
                }
                if ($itC.Tipo -eq "inicio") { $curIntC++ }
                $curUltC = $mC
                if ($itC.Tipo -eq "fin") {
                    # Confirmado=$true: llegó el apagado DEFINITIVO (Shutdown()/PhoneService shutdown) = el
                    # proceso terminó de verdad. Es la GARANTÍA de que el OneX se cerró (no solo se pidió).
                    $CierreEpisodios += [pscustomobject]@{ Slot=$curIniC; Raw=$curRawC; Intentos=$curIntC; FinMs=$curUltC; Confirmado=$true }
                    $ultFinC = $mC; $curIniC = $null
                }
            }
            # Episodio que quedó ABIERTO al final: hubo ExitHandler pero NO llegó el apagado definitivo
            # (Shutdown/PhoneService shutdown) → NO confirmado. Puede ser: app colgada, el asesor canceló el
            # cierre, o (frecuente) la extracción cortó antes del apagado. NO se afirma que se cerró.
            if ($null -ne $curIniC) { $CierreEpisodios += [pscustomobject]@{ Slot=$curIniC; Raw=$curRawC; Intentos=$curIntC; FinMs=$curUltC; Confirmado=$false } }

            foreach ($ep in $CierreEpisodios) {
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].SysLog -ne "") { continue }
                $cm = & $MsDeSlot $ep.Slot
                $cfin = if ($ep.FinMs -ge $cm) { $ep.FinMs } else { $cm }
                # La señal de "firmado al cerrar" se busca en TODO el tramo del cierre (no solo en el primer
                # clic): si la app se colgó, el desfirme forzado llega decenas de segundos después.
                $firmado = $false
                foreach ($pm in $firmadoMs) { if ($pm -ge 0 -and $pm -ge ($cm - 200) -and $pm -le ($cfin + 5000)) { $firmado = $true; break } }
                # Sufijo SOLO cuando el cierre costó trabajo de verdad (app colgada). En un cierre sano OneX
                # dispara ExitHandler 2 veces en ~1 s, así que se exige además que el tramo dure >=5 s.
                $durC = [int](($cfin - $cm) / 1000)
                $sufC = if ($ep.Intentos -gt 1 -and $durC -ge 5) { " — $($ep.Intentos) intentos en ${durC}s (la app no respondía)" } else { "" }
                # $ep.Confirmado = llegó el apagado DEFINITIVO (Shutdown()/PhoneService shutdown) = garantía de que
                # el proceso terminó. Sin él, solo hubo ExitHandler = "cierre iniciado" (no se puede afirmar que cerró).
                if (-not $ep.Confirmado) {
                    $EventosTiempo[$ep.Slot].SysLog   = "⚠ CIERRE INICIADO (apagado NO confirmado): hubo ExitHandler pero no llegó el Shutdown$sufC"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::Khaki
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? Se pidió cerrar (ExitHandler) pero NO apareció el apagado definitivo (Begin Executing method Shutdown() / PhoneService shutdown). Puede ser: app colgada, cierre cancelado, o la extracción cortó antes del apagado. NO se puede garantizar que el OneX se cerró. Clics de cerrar: $($ep.Intentos).`n$($ep.Raw)`n"
                } elseif ($firmado) {
                    $EventosTiempo[$ep.Slot].SysLog   = "⚠ APP CERRADA (confirmado) con asesor FIRMADO: se cerró el OneX estando el asesor disponible$sufC"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::Orange
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? Apagado DEFINITIVO confirmado (Shutdown/PhoneService shutdown), y el cierre desfirmó a un agente que seguía firmado (EnterAuxHandler / transición Ready→LoggedOut DENTRO del tramo de cierre). Clics de cerrar: $($ep.Intentos); el apagado llegó ${durC}s después del primero.`n$($ep.Raw)`n"
                } else {
                    $EventosTiempo[$ep.Slot].SysLog   = "✔ APP CERRADA (confirmado): el asesor se desfirmó y luego cerró el OneX$sufC"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::LightGreen
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? Apagado DEFINITIVO confirmado (Shutdown/PhoneService shutdown), con requestor='manual' + closeSignalingChannel cat:0/cod:0 (cierre limpio, NO caída de red). El asesor ya estaba desfirmado al cerrar (desfirme ANTES del cierre). Clics de cerrar: $($ep.Intentos).`n$($ep.Raw)`n"
                }
                if ($EventosTiempo[$ep.Slot].Tel -eq "-") { $EventosTiempo[$ep.Slot].Tel = "RED/AVAYA" }
            }
            # 3) Caída de red (una fila por episodio)
            foreach ($ep in (& $ColapsarEpisodios $RedCaidaList $GAP_RED)) {
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].SysLog -eq "") {
                    $EventosTiempo[$ep.Slot].SysLog   = "⚠ CAÍDA DE RED: se perdió el enlace con Avaya (la aplicación seguía ABIERTA)"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::Orange
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? El canal se cerró por timeout de red (cat:2/cod:10060) o falló el keepalive al gatekeeper:`n$($ep.Raw)`n"
                    if ($EventosTiempo[$ep.Slot].Tel -eq "-") { $EventosTiempo[$ep.Slot].Tel = "RED/AVAYA" }
                }
            }
            # 4) Recuperación de enlace en curso (una fila por episodio). OJO: LinkRecoveryProgressEvent
            #    TAMBIÉN aparece en un LOGIN NORMAL (handshake H.323 de registro) → solo se muestra si hubo
            #    una CAÍDA de red en los ~10 min previos. Sin caída = login normal → se suprime (falso positivo).
            $redMsAll = @(@($RedCaidaList) | ForEach-Object { & $MsDeSlot $_.Slot })
            foreach ($ep in (& $ColapsarEpisodios $ReconIntentoList $GAP_RED)) {
                $im = & $MsDeSlot $ep.Slot
                $trasCaida = $false
                foreach ($rm in $redMsAll) { if ($rm -ge 0 -and ($im - $rm) -ge 0 -and ($im - $rm) -le 600000) { $trasCaida = $true; break } }
                if (-not $trasCaida) { continue }
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].SysLog -eq "") {
                    $EventosTiempo[$ep.Slot].SysLog   = "↻ Avaya intentando reconectar automáticamente (recuperación de enlace)"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::Khaki
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? Tras una caída de red, la app dispara eventos de progreso de recuperación de enlace:`n$($ep.Raw)`n"
                    if ($EventosTiempo[$ep.Slot].Tel -eq "-") { $EventosTiempo[$ep.Slot].Tel = "RED/AVAYA" }
                }
            }
            # 5) Reconexión AUTOMÁTICA exitosa (una fila por episodio): bLInkRecovery=1 = la app estaba ABIERTA.
            foreach ($ep in (& $ColapsarEpisodios $ReconOkList $GAP_RED)) {
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].SysLog -eq "") {
                    $EventosTiempo[$ep.Slot].SysLog   = "✔ RECONEXIÓN AUTOMÁTICA (bLInkRecovery=1): la app estaba ABIERTA — NO fue un login manual"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::Cyan
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? El login se completó con bLInkRecovery=1 (recuperación de enlace de una sesión ya viva). Un login manual o un arranque tras reinicio sería bLInkRecovery=0:`n$($ep.Raw)`n"
                    if ($EventosTiempo[$ep.Slot].Tel -eq "-") { $EventosTiempo[$ep.Slot].Tel = "RED/AVAYA" }
                }
            }
            # 6) Re-firma automática del agente: SOLO si hay una reconexión (bLInkRecovery) cercana (±120 s).
            $reconOkMs = @(@($ReconOkList) | ForEach-Object { & $MsDeSlot $_.Slot })
            foreach ($ep in (& $ColapsarEpisodios $ReconRefirmaList $GAP_RED)) {
                $rm = & $MsDeSlot $ep.Slot
                $cerca = $false
                foreach ($okm in $reconOkMs) { if ($okm -ge 0 -and [math]::Abs($okm - $rm) -le 120000) { $cerca = $true; break } }
                if (-not $cerca) { continue }
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].SysLog -eq "") {
                    $EventosTiempo[$ep.Slot].SysLog   = "✔ Agente re-firmado automáticamente por la reconexión (sin intervención del asesor)"
                    $EventosTiempo[$ep.Slot].ColorSys = [System.Drawing.Color]::Cyan
                    $EventosTiempo[$ep.Slot].RawSysLog += "¿Por qué? Tras la recuperación de enlace, Avaya re-firma al agente automáticamente:`n$($ep.Raw)`n"
                    if ($EventosTiempo[$ep.Slot].Tel -eq "-") { $EventosTiempo[$ep.Slot].Tel = "RED/AVAYA" }
                }
            }
            # 7) Hilo interno abortado: SOLO es aviso si NO hubo un cierre de app cercano (±60 s). Si el
            #    abort fue parte del teardown de un cierre, se SUPRIME (la fila de CIERRE ya lo cuenta).
            $cierreMs = @(@($CierreAppList) | ForEach-Object { & $MsDeSlot $_.Slot })
            foreach ($ep in (& $ColapsarEpisodios $HiloAbortadoList $GAP_CIERRE)) {
                $hm = & $MsDeSlot $ep.Slot
                $cerca = $false
                foreach ($cm in $cierreMs) { if ($cm -ge 0 -and [math]::Abs($cm - $hm) -le 60000) { $cerca = $true; break } }
                if ($cerca) { continue }
                Init-Hora $ep.Slot
                if ($EventosTiempo[$ep.Slot].AppLog -eq "") {
                    $EventosTiempo[$ep.Slot].AppLog   = "Aviso: hilo interno abortado (recuperación de red / reinicio de servicio) — la app NO se cerró"
                    $EventosTiempo[$ep.Slot].ColorApp = [System.Drawing.Color]::Orange
                    $EventosTiempo[$ep.Slot].RawAppLog += "¿Por qué? Avaya abortó un hilo interno pero NO hubo Shutdown/ExitHandler cercano → la app siguió abierta:`n$($ep.Raw)`n"
                }
            }

            # ================================================================
            # FALLBACK FIN: Call StateChanged OldState=Active,NewState=Disconnected
            # Aplica FIN para sesiones que PASO 5 (ProcessSessionEndedEvent) no cubrió.
            # En este punto $CuelguesVistos y $CuelguesManuales ya están poblados por PASO 5.
            # ================================================================
            foreach ($SesD in $CallStateDisconnected.Keys) {
                $TelD   = if ($MapeoTel[$SesD]) { $MapeoTel[$SesD] } elseif ($MapeoTelP4[$SesD]) { $MapeoTelP4[$SesD] } else { "Desconocido" }
                $FirmaD = "$SesD-$TelD"
                if (-not $CuelguesVistos[$FirmaD]) {
                    $HoraD = $CallStateDisconnected[$SesD]
                    Init-Hora $HoraD
                    $CuelguesVistos[$FirmaD]          = $true
                    $EventosTiempo[$HoraD].Sesion      = $SesD
                    $EventosTiempo[$HoraD].Tel         = $TelD
                    if ($CuelguesManuales[$SesD]) {
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
            # FALLBACK FIN POR VI (UUID): cuelgue que solo dejó rastro en
            # VI StateImpl type=Disconnected (sin EndpointLog OldState=Active ni
            # ProcessSessionEndedEvent). Cubre el caso típico de CLIENTE COLGÓ EN
            # HOLD. Se liga por UUID del VI (no por Id numérico, que se reutiliza)
            # y se protege con $CuelguesVistos para no duplicar FIN ya emitidos.
            # Precedencia: agente colgó (CuelguesManuales) > cliente en hold > normal.
            # ================================================================
            foreach ($kViF in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kViF)) { continue }
                $oViF = $EventosTiempo[$kViF]
                if ($null -eq $oViF -or $oViF -isnot [hashtable]) { continue }
                if ($oViF.Interpretacion -notmatch "INICIO DE LLAMADA") { continue }
                $viIniF = $oViF.ViId
                if (-not $viIniF -or -not $ViDisconnected.ContainsKey($viIniF)) { continue }
                $sesIniF = $oViF.Sesion
                $telIniF = if ($MapeoTel[$sesIniF]) { $MapeoTel[$sesIniF] } elseif ($MapeoTelP4[$sesIniF]) { $MapeoTelP4[$sesIniF] } else { "Desconocido" }
                $FirmaViF = "$sesIniF-$telIniF"
                if ($CuelguesVistos[$FirmaViF]) { continue }   # PASO5 o fallback Endpoint ya emitieron FIN para esta sesión
                $HoraViF = $ViDisconnected[$viIniF]
                if ($EventosTiempo.ContainsKey($HoraViF) -and $EventosTiempo[$HoraViF].Interpretacion -match "FIN DE LLAMADA") { $CuelguesVistos[$FirmaViF] = $true; continue }
                Init-Hora $HoraViF
                $CuelguesVistos[$FirmaViF]           = $true
                $EventosTiempo[$HoraViF].Sesion       = $sesIniF
                if ($telIniF -ne "Desconocido") { $EventosTiempo[$HoraViF].Tel = $telIniF }
                if ($CuelguesManuales[$sesIniF]) {
                    $EventosTiempo[$HoraViF].Interpretacion      = "$symStop FIN DE LLAMADA MANUAL (Colgada por el Asesor)"
                    $EventosTiempo[$HoraViF].ColorInterpretacion = [System.Drawing.Color]::LightCoral
                } elseif ($ViFinEnHold[$viIniF]) {
                    $EventosTiempo[$HoraViF].Interpretacion      = "$symStop FIN DE LLAMADA - CLIENTE COLGÓ EN HOLD"
                    $EventosTiempo[$HoraViF].ColorInterpretacion = [System.Drawing.Color]::Tomato
                } else {
                    $EventosTiempo[$HoraViF].Interpretacion      = "$symStop FIN DE LLAMADA NORMAL"
                    $EventosTiempo[$HoraViF].ColorInterpretacion = [System.Drawing.Color]::DarkGray
                }
                $EventosTiempo[$HoraViF].RawInterpretacion += "[Fallback VI] StateImpl type=Disconnected (EnHold=$($ViFinEnHold[$viIniF])) — sin EndpointLog/ProcessSessionEndedEvent`n"
            }

            # ================================================================
            # SWEEP "SIN BOTONES PARA TOMAR LA LLAMADA" (caso dleal 02/07/2026)
            # Candidatos: VI entrante que pasó por Alerting con answer=False y NUNCA recibió
            # "Auto Accepting" ni "AnswerVoiceInteraction" ni llegó a Active. Se descartan los
            # abandonos rápidos (el cliente colgó en <10s, antes de que el auto-accept actuara):
            # esos son llamadas perdidas normales, no fallas de la UI.
            # ================================================================
            $ParseSlotSB = {
                param($s)
                try { [datetime]::ParseExact(($s -replace ',','.'), $(if ($s -match ',') { "HH:mm:ss.fff" } else { "HH:mm:ss" }), $null) } catch { $null }
            }
            foreach ($uSB in @($ViSinBotones.Keys)) {
                if ($Script:ModoContestacion.ContainsKey($uSB)) { continue }   # sí fue contestada
                $infoSB = $ViSinBotones[$uSB]; $slotSB = $infoSB.Slot
                if ($ViDisconnected.ContainsKey($uSB)) {
                    $tAl = & $ParseSlotSB $slotSB; $tDx = & $ParseSlotSB $ViDisconnected[$uSB]
                    if ($tAl -and $tDx -and ($tDx - $tAl).TotalSeconds -lt 10) { continue }   # abandono rápido normal
                }
                # Fila propia: si el slot ya está ocupado por otra interpretación, buscar ms libre del mismo segundo
                $SlotFinalSB = $slotSB
                if ($EventosTiempo.ContainsKey($slotSB) -and $EventosTiempo[$slotSB].Interpretacion -ne "") {
                    $bhSB = ($slotSB -split ',')[0]; $mbSB = [int](($slotSB -split ',')[1])
                    for ($msSB = $mbSB + 1; $msSB -le 999; $msSB++) { $candSB = "$bhSB,$($msSB.ToString('000'))"; if (-not $EventosTiempo.ContainsKey($candSB)) { $SlotFinalSB = $candSB; break } }
                }
                Init-Hora $SlotFinalSB
                $EventosTiempo[$SlotFinalSB].Interpretacion      = "⚠ SIN BOTONES PARA TOMAR LA LLAMADA (quedó timbrando sin poder contestarse)"
                $EventosTiempo[$SlotFinalSB].ColorInterpretacion = [System.Drawing.Color]::OrangeRed
                $EventosTiempo[$SlotFinalSB].ViId = $uSB
                # Ligar Sesión (ConnId) y Teléfono: el cxt del VI mapea al ConnectionId numérico ($CxtToConnId,
                # poblado en PASO 4). Así la fila muestra el ID de la llamada (ej: 28) que antes se perdía por
                # colgarse del slot de UpdateHistoryRecord que PASO 6 suprime.
                $sesSB = if ($infoSB.Cxt -and $Script:CxtToConnId.ContainsKey($infoSB.Cxt)) { $Script:CxtToConnId[$infoSB.Cxt] } else { "" }
                if ($sesSB -ne "") {
                    $EventosTiempo[$SlotFinalSB].Sesion = $sesSB
                    $telSB = if ($MapeoTel[$sesSB]) { $MapeoTel[$sesSB] } elseif ($MapeoTelP4[$sesSB]) { $MapeoTelP4[$sesSB] } else { "" }
                    if ($telSB -ne "" -and $EventosTiempo[$SlotFinalSB].Tel -eq "-") { $EventosTiempo[$SlotFinalSB].Tel = $telSB }
                }
                $EventosTiempo[$SlotFinalSB].RawInterpretacion += "[SIN BOTONES] La interacción quedó en Alerting con answer=False y NUNCA llegó 'Auto Accepting' ni respuesta manual (motor de auto-contestación roto). VI=$uSB`n$($infoSB.Raw)`n"
            }

            # ================================================================
            # SWEEP HOLD IMPLÍCITO: OldState=Active,NewState=Inactive sin OnRequestHoldSession
            # Añade "Hold automático del sistema" cuando EndpointLog no registró el hold manual.
            # Guard: solo actúa si el slot no tiene ya un HOLD/UNHOLD en la columna Agente.
            # ================================================================
            foreach ($SesH in $HoldImplicito.Keys) {
              foreach ($HoraH in @($HoldImplicito[$SesH])) {
                # Guard anti-duplicado (caso mdelacruz 06/07/2026): si en el MISMO segundo y la MISMA sesión
                # ya existe un "HOLD MANUAL" (el clic vive en su ms-slot, Categoría A), este Active→Inactive
                # es el ECO de ese clic, NO un hold automático → no duplicar. El automático REAL (sin clic)
                # no tiene HOLD MANUAL en el segundo, así que sigue detectándose normal.
                $segHImp = ($HoraH -split ',')[0]
                $HayHoldManualSes = $false
                foreach ($kHi in @($EventosTiempo.Keys)) {
                    if ((($kHi -split ',')[0] -eq $segHImp) -and $EventosTiempo[$kHi].Agente -match "\bHOLD MANUAL" -and $EventosTiempo[$kHi].Sesion -eq $SesH) { $HayHoldManualSes = $true; break }
                }
                if ($HayHoldManualSes) { continue }
                if ($EventosTiempo.ContainsKey($HoraH) -and $EventosTiempo[$HoraH].Agente -notmatch "HOLD") {
                    $TelH = if ($MapeoTel[$SesH]) { $MapeoTel[$SesH] } elseif ($MapeoTelP4[$SesH]) { $MapeoTelP4[$SesH] } else { "" }
                    $msHi = & $MsDeSlot $HoraH
                    # ¿Hold DEL ASESOR? "Begin Executing method Hold(ConnId=SesH)" en el mismo instante que este
                    # cambio Active→Inactive. Ambas líneas traen el ConnId, pero se exige también proximidad
                    # temporal (los IDs se reciclan por login). La ventana NO es una duración de hold: es
                    # tolerancia de jitter entre dos líneas del MISMO evento (en la práctica ~9 ms de diferencia).
                    # Si existe el método → el asesor retuvo; si no → auto-hold del sistema (abrir 2da línea).
                    $esHoldAgente = $false
                    if ($HoldMetodoSes.ContainsKey($SesH)) {
                        foreach ($sHM in $HoldMetodoSes[$SesH]) { $mHM = & $MsDeSlot $sHM; if ($mHM -ge 0 -and $msHi -ge 0 -and [math]::Abs($mHM - $msHi) -le 1500) { $esHoldAgente = $true; break } }
                    }
                    # NOTA: se quitó el sufijo "⚠ bridge activo" (Pablo, 27/07): dependía de una ventana ±60 s
                    # arbitraria y daba falsos positivos. Si hay un bridge real, ya sale en su propia fila
                    # ("Sesión bridge del sistema"); el hold no necesita editorializar sobre eso.
                    $HoldNote = if ($esHoldAgente) { "$symPause Asesor pone la llamada en espera (Sesión $SesH)" } else { "|| Hold automático del sistema (Sesión $SesH)" }
                    if ($EventosTiempo[$HoraH].Agente -ne "") {
                        # El slot ya tiene otro evento en la columna Agente (ej: TRANSFERENCIA INICIADA).
                        # En vez de amontonar ambos en la misma celda con " || ", se le da al Hold automático
                        # su propia fila en un ms-slot adyacente para que se ordene como evento independiente.
                        $bhH = $HoraH; $mbH = 0
                        if ($HoraH -match "^(\d{2}:\d{2}:\d{2}),(\d+)$") { $bhH = $matches[1]; $mbH = [int]$matches[2] }
                        $SlotHold = $HoraH
                        for ($msH = $mbH + 1; $msH -le 999; $msH++) { $candH = "$bhH,$msH"; if (-not $EventosTiempo.ContainsKey($candH)) { $SlotHold = $candH; break } }
                        Init-Hora $SlotHold
                        $EventosTiempo[$SlotHold].Agente      = $HoldNote
                        $EventosTiempo[$SlotHold].ColorAgente = [System.Drawing.Color]::Yellow
                        $EventosTiempo[$SlotHold].Sesion      = $SesH
                        if ($TelH -ne "") { $EventosTiempo[$SlotHold].Tel = $TelH }
                    } else {
                        $EventosTiempo[$HoraH].Agente      = $HoldNote
                        $EventosTiempo[$HoraH].ColorAgente = [System.Drawing.Color]::Yellow
                        if ($EventosTiempo[$HoraH].Sesion -eq "-" -and $TelH -ne "") { $EventosTiempo[$HoraH].Tel = $TelH }
                    }
                }
              }
            }

            # ── Reetiquetar el AutoHold que dispara "Agregar llamada" o "captura número nuevo" ──────
            # DOS botones ponen la llamada actual en AutoHold para abrir línea antes de marcar/consultar:
            # AddCallHandler ("Agregar llamada") y NewCallHandler (el asesor teclea un número NUEVO en la
            # caja de texto y da Enter estando YA en una llamada). Ese hold llega unos ms después y sale
            # como "HOLD MANUAL (Clic del Agente)" SIN confirmación de clic (no hay HoldCallHandler cerca),
            # aunque NO fue un clic de hold suelto — fue consecuencia del otro botón. Si un "HOLD MANUAL"
            # coincide (±1500 ms) con cualquiera de los dos, se reetiqueta según cuál lo disparó, para que
            # el análisis no lo confunda con un hold manual independiente. Se fija también la Interpretación
            # porque al quitar "HOLD MANUAL" del Agente el render ya no la calcularía.
            # (Pablo: caso 21:51:17 = Agregar llamada; caso 15:37:42 = captura de número nuevo.)
            if (($AddCallSlots.Count + $NewCallSlots.Count) -gt 0) {
                foreach ($kHA in @($EventosTiempo.Keys)) {
                    $oHA = $EventosTiempo[$kHA]
                    if ($null -eq $oHA -or $oHA -isnot [hashtable]) { continue }
                    if ($oHA.Agente -notmatch "HOLD MANUAL \(Clic del Agente\)") { continue }
                    $mHA = & $MsDeSlot $kHA
                    if ($mHA -lt 0) { continue }
                    $cercaAdd = $false; $cercaNew = $false
                    foreach ($sAc in $AddCallSlots) { $mAc = & $MsDeSlot $sAc; if ($mAc -ge 0 -and [math]::Abs($mHA - $mAc) -le 1500) { $cercaAdd = $true; break } }
                    if (-not $cercaAdd) {
                        foreach ($sNc in $NewCallSlots) { $mNc = & $MsDeSlot $sNc; if ($mNc -ge 0 -and [math]::Abs($mHA - $mNc) -le 1500) { $cercaNew = $true; break } }
                    }
                    if ($cercaAdd -or $cercaNew) {
                        $causaAutoHold = if ($cercaAdd) { "Agregar llamada" } else { "captura de número nuevo" }
                        $oHA.Agente               = "|| Hold automático (por $causaAutoHold)"
                        $oHA.ColorAgente          = [System.Drawing.Color]::Yellow
                        $oHA.Interpretacion       = "Llamada en Hold (por $causaAutoHold)"
                        $oHA.ColorInterpretacion  = [System.Drawing.Color]::Yellow
                    }
                }
            }

            # ================================================================
            # SWEEP RETOMAR: OldState=Inactive,NewState=Active = la llamada sale del hold.
            # Ocurre al volver a una línea que tenía una llamada en espera. Guard: si en el mismo
            # segundo y misma sesión ya hay un UNHOLD MANUAL (clic del asesor), ese evento manda.
            # ================================================================
            foreach ($SesR in $RetomaImplicita.Keys) {
                foreach ($HoraR in @($RetomaImplicita[$SesR])) {
                    if (-not $EventosTiempo.ContainsKey($HoraR)) { continue }
                    $segR = ($HoraR -split ',')[0]
                    $HayUnholdManual = $false
                    foreach ($kRi in @($EventosTiempo.Keys)) {
                        if ((($kRi -split ',')[0] -ne $segR) -or [string]::IsNullOrEmpty($EventosTiempo[$kRi].Agente)) { continue }
                        # (a) clic real de UNHOLD del asesor en esta sesión  (b) retomado automático tras fallo de transferencia
                        if (($EventosTiempo[$kRi].Agente -match "UNHOLD" -and $EventosTiempo[$kRi].Sesion -eq $SesR) -or
                            ($EventosTiempo[$kRi].Agente -match "retomada autom")) { $HayUnholdManual = $true; break }
                    }
                    if ($HayUnholdManual) { continue }
                    if ($EventosTiempo[$HoraR].Agente -match "HOLD|UNHOLD") { continue }
                    $TelR = if ($MapeoTel[$SesR]) { $MapeoTel[$SesR] } elseif ($MapeoTelP4[$SesR]) { $MapeoTelP4[$SesR] } else { "" }
                    $SlotR2 = $HoraR
                    if ($EventosTiempo[$HoraR].Agente -ne "") {
                        $bhR = $HoraR; $mbR = 0
                        if ($HoraR -match "^(\d{2}:\d{2}:\d{2}),(\d+)$") { $bhR = $matches[1]; $mbR = [int]$matches[2] }
                        for ($msR = $mbR + 1; $msR -le 999; $msR++) { $candR = "$bhR,$msR"; if (-not $EventosTiempo.ContainsKey($candR)) { $SlotR2 = $candR; break } }
                        Init-Hora $SlotR2
                    }
                    $EventosTiempo[$SlotR2].Agente      = "$symRes Llamada retomada del hold (Sesión $SesR)"
                    $EventosTiempo[$SlotR2].ColorAgente = [System.Drawing.Color]::LightGoldenrodYellow
                    $EventosTiempo[$SlotR2].Sesion      = $SesR
                    if ($TelR -ne "" -and $EventosTiempo[$SlotR2].Tel -eq "-") { $EventosTiempo[$SlotR2].Tel = $TelR }
                }
            }

            # ================================================================
            # TEMA 3: "[!] Sesión bridge del sistema" solo en sesiones reales.
            # La sesión bridge se crea con RemoteParty=[,] (sin destino real). Si la sesión
            # nunca obtuvo un teléfono real (quedó phantom, ej: legs de transferencia 4/5) se
            # suprime la fila; si terminó siendo una llamada real (ej: saliente a una extensión)
            # se conserva el marcador.
            # ================================================================
            $LineaVaciaSinMarcar = @{}   # sesión → $true: línea que el asesor abrió y no marcó → su FIN dirá "Cerró línea sin marcar"

            # --- LÍNEA ABIERTA SIN MARCAR: dirigido por la CONFIRMACIÓN del endpoint ---
            # Se maneja desde $SinDestinoFin ("this record has no far-end address"), NO desde la etiqueta
            # "Sesión bridge". La fila de inicio puede haber quedado como "INICIO DE LLAMADA (Saliente)",
            # "LÍNEA ABIERTA" o "Sesión bridge" según qué señal la creó; antes solo se miraba la bridge, así
            # que en los casos que nacían como INICIO DE LLAMADA no se detectaba NADA. Aquí se busca la fila
            # de inicio de la sesión sea cual sea su etiqueta.
            foreach ($fdL in $SinDestinoFin) {
                $sesL = $fdL.Ses
                if (-not $sesL -or $sesL -eq "-") { continue }
                $msFinL = & $MsDeSlot $fdL.Slot
                if ($msFinL -lt 0) { continue }
                # Fila de INICIO de esa sesión: la más cercana ANTERIOR al fin confirmado (los IDs se reciclan).
                $slotIniL = $null; $msIniL = -1
                foreach ($kL in @($EventosTiempo.Keys)) {
                    $oL = $EventosTiempo[$kL]
                    if ($oL.Sesion -ne $sesL) { continue }
                    if ($oL.Interpretacion -notmatch "INICIO DE LLAMADA|Sesión bridge del sistema|LÍNEA ABIERTA") { continue }
                    $mL = & $MsDeSlot $kL
                    if ($mL -lt 0 -or $mL -gt $msFinL) { continue }
                    if ($mL -gt $msIniL) { $msIniL = $mL; $slotIniL = $kL }
                }
                if (-not $slotIniL) { continue }
                $durMsL = $msFinL - $msIniL
                $durL = [int]($durMsL / 1000)
                # >120 s = apareamiento erróneo (Avaya libera sola la línea sin marcar en ~60 s de timeout) → no tocar.
                if ($durL -gt 120) { continue }
                # <2 s = PATA PHANTOM (murió en milisegundos): típico del mecanismo de transferencia
                # (Transfer_ActivateConsultCall crea una pata vacía que nace y muere en ~90 ms) o de un bridge
                # interno. NO es tiempo muerto (un tiempo muerto real es el asesor con la línea abierta SEGUNDOS;
                # verificado con prueba controlada plopezs 16/13: reales = 10-11 s). Se SUPRIME su fila preliminar
                # "LÍNEA ABIERTA SIN MARCAR"/"Sesión bridge" (de la MISMA sesión, dentro de su ventana de vida) para
                # que no contamine el reporte de tiempos muertos. El discriminador es la DURACIÓN, NO "hay llamada
                # concurrente" (la prueba mostró tiempos muertos reales abiertos MIENTRAS otra llamada estaba en hold).
                if ($durMsL -lt 2000) {
                    foreach ($kPh in @($EventosTiempo.Keys)) {
                        $oPh = $EventosTiempo[$kPh]
                        if ($null -eq $oPh -or $oPh -isnot [hashtable] -or $oPh.Sesion -ne $sesL) { continue }
                        if ($oPh.Interpretacion -notmatch "LÍNEA ABIERTA SIN MARCAR|Sesión bridge del sistema") { continue }
                        $mPh = & $MsDeSlot $kPh
                        if ($mPh -lt ($msIniL - 500) -or $mPh -gt ($msFinL + 500)) { continue }
                        # Si la fila solo tenía la etiqueta phantom, se elimina; si trae otro contenido (audio,
                        # error, etc.) solo se limpia la interpretación para no perder ese dato.
                        if ($oPh.Agente -eq "" -and $oPh.Audio -eq "" -and $oPh.Aux -eq "" -and $oPh.SysLog -eq "" -and $oPh.AppLog -eq "" -and $oPh.Ispeac -eq "" -and $oPh.Dtmf -eq "") {
                            $EventosTiempo.Remove($kPh) | Out-Null
                        } else {
                            $oPh.Interpretacion = ""; $oPh.ColorInterpretacion = [System.Drawing.Color]::White
                        }
                    }
                    continue
                }
                $segIniL = ($slotIniL -split ',')[0]; $nLinL = 0
                foreach ($offL in -1,0,1) {
                    try { $sChkL = ([datetime]::ParseExact($segIniL,'HH:mm:ss',$null).AddSeconds($offL)).ToString('HH:mm:ss') } catch { $sChkL = $segIniL }
                    if ($LineaAppPorSeg.ContainsKey($sChkL)) { $ltL = $LineaAppPorSeg[$sChkL]; if ($ltL -match '^[a-z]$') { $nLinL = [int][char]$ltL - [int][char]'a' + 1; break } }
                }
                $txtLinL = if ($nLinL -ge 1) { "línea $nLinL" } else { "una línea" }
                $EventosTiempo[$slotIniL].Interpretacion      = "$symUp Abrió $txtLinL sin marcar (${durL}s)"
                $EventosTiempo[$slotIniL].ColorInterpretacion = [System.Drawing.Color]::Gold
                $EventosTiempo[$slotIniL].RawInterpretacion  += "[LÍNEA ABIERTA SIN MARCAR] Sesión $sesL abierta durante $durL s. CONFIRMADO por el endpoint: al cerrar el historial de la sesión avisó 'this record has no far-end address' = nunca se marcó ningún número.`n"
                $LineaVaciaSinMarcar[$sesL] = $true
                # Relabelar el FIN de ESTA instancia concreta. NO se puede marcar por número de sesión: Avaya
                # RECICLA los IDs en cada login, así que un "abrió línea sin marcar" de la mañana marcaría
                # como "cerró sin marcar" a TODAS las llamadas reales con ese mismo ID el resto del día.
                # Se busca el FIN de esta sesión dentro de la ventana [inicio, fin confirmado + 5 s].
                $slotFinV = $null; $bestDV = 1e18
                foreach ($kFv in @($EventosTiempo.Keys)) {
                    $oFv = $EventosTiempo[$kFv]
                    if ($null -eq $oFv -or $oFv -isnot [hashtable]) { continue }
                    if ($oFv.Sesion -ne $sesL) { continue }
                    if ($oFv.Interpretacion -notmatch "FIN DE LLAMADA|CUELGUE MANUAL") { continue }
                    $mFv = & $MsDeSlot $kFv
                    if ($mFv -lt $msIniL -or $mFv -gt ($msFinL + 5000)) { continue }
                    $dV = [math]::Abs($mFv - $msFinL)
                    if ($dV -lt $bestDV) { $bestDV = $dV; $slotFinV = $kFv }
                }
                if ($slotFinV) {
                    $EventosTiempo[$slotFinV].Interpretacion      = "$symStop Cerró línea sin marcar"
                    $EventosTiempo[$slotFinV].ColorInterpretacion = [System.Drawing.Color]::Goldenrod
                    $EventosTiempo[$slotFinV].RawInterpretacion  += "[CERRÓ LÍNEA SIN MARCAR] Cierre de la sesión $sesL abierta en $slotIniL (instancia acotada por tiempo; los IDs de sesión se reciclan en cada login).`n"
                }
            }

            # ── Suprimir "LÍNEA ABIERTA SIN MARCAR" cuando la sesión SÍ resolvió un número real ──────────
            # Una pata que nace vacía (RemoteParty=[,]) recibe la etiqueta preliminar "LÍNEA ABIERTA SIN
            # MARCAR" en slot base; si el número REAL llega ~ms después (far-end address resuelto), un guard
            # impide subir la etiqueta y queda como falso "sin marcar". Caso típico: el DESTINO de una
            # transferencia (ej. ses 46 → 812-201-8319) cuando la pata phantom (ses 48, <2s) se llevó el
            # registro de consulta. Como SÍ marcó, NO es tiempo muerto → se suprime la fila (Pablo eligió
            # suprimir, caso 20:57:32). Se acota por tiempo (IDs reciclan): el número debe haberse resuelto
            # dentro de [slot-3s, slot+10s] de la fila. A este punto los tiempos muertos REALES ya se
            # reetiquetaron a "Abrió línea N sin marcar (Xs)" (arriba), así que no matchean y no se tocan.
            if ($FarEndResueltoSlots.Count -gt 0) {
                foreach ($kSM in @($EventosTiempo.Keys)) {
                    $oSM = $EventosTiempo[$kSM]
                    if ($null -eq $oSM -or $oSM -isnot [hashtable]) { continue }
                    if ($oSM.Interpretacion -notmatch "LÍNEA ABIERTA SIN MARCAR") { continue }
                    $sesSM = $oSM.Sesion
                    if (-not $sesSM -or $sesSM -eq "-") { continue }
                    $mSM = & $MsDeSlot $kSM
                    if ($mSM -lt 0) { continue }
                    $resolvioSM = $false
                    foreach ($fr in $FarEndResueltoSlots) {
                        if ($fr.Ses -ne $sesSM) { continue }
                        $mFr = & $MsDeSlot $fr.Slot
                        if ($mFr -ge 0 -and $mFr -ge ($mSM - 3000) -and $mFr -le ($mSM + 10000)) { $resolvioSM = $true; break }
                    }
                    if ($resolvioSM) {
                        # Se elimina la fila (Pablo eligió suprimir); si trae un error de sistema/app se
                        # preserva ese dato limpiando solo la interpretación falsa.
                        if ($oSM.SysLog -eq "" -and $oSM.AppLog -eq "") {
                            $EventosTiempo.Remove($kSM) | Out-Null
                        } else {
                            $oSM.Interpretacion = ""; $oSM.ColorInterpretacion = [System.Drawing.Color]::White
                        }
                    }
                }
            }

            foreach ($kB in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kB)) { continue }
                if ($EventosTiempo[$kB].Interpretacion -match "Sesión bridge del sistema") {
                    $SesB = $EventosTiempo[$kB].Sesion
                    $TelRealB = if ($SesB -and $SesB -ne "-" -and $MapeoTel[$SesB]) { $MapeoTel[$SesB] } else { "" }
                    # 565 = código de desfirme, no es un teléfono de cliente real → tratar como phantom
                    # para que el desfirme no muestre "[!] Sesión bridge del sistema".
                    $EsRealB = ($TelRealB -ne "" -and $TelRealB -notmatch "Desconocido" -and $TelRealB -notmatch "^565$")
                    # Si esta sesión ya salió como "Abrió línea sin marcar" cerca en el tiempo, este bridge es
                    # la MISMA acción contada dos veces → suprimirlo (Pablo, 28/07). Se acota por tiempo porque
                    # los IDs se reciclan por login (un bridge real de otra hora con el mismo ID no se toca).
                    if ($SesB -and $LineaVaciaSinMarcar.ContainsKey($SesB)) {
                        $mBk = & $MsDeSlot $kB
                        $hayAbrio = $false
                        foreach ($kAb in @($EventosTiempo.Keys)) {
                            $oAb = $EventosTiempo[$kAb]
                            if ($null -eq $oAb -or $oAb -isnot [hashtable]) { continue }
                            if ($oAb.Sesion -eq $SesB -and $oAb.Interpretacion -match "Abrió .*sin marcar") {
                                $mAb = & $MsDeSlot $kAb
                                if ($mBk -ge 0 -and $mAb -ge 0 -and [math]::Abs($mAb - $mBk) -le 30000) { $hayAbrio = $true; break }
                            }
                        }
                        if ($hayAbrio) {
                            $EventosTiempo[$kB].Interpretacion = ""
                            $EventosTiempo[$kB].ColorInterpretacion = [System.Drawing.Color]::White
                            if ($EventosTiempo[$kB].Agente -eq "" -and $EventosTiempo[$kB].AppLog -eq "" -and $EventosTiempo[$kB].SysLog -eq "" -and $EventosTiempo[$kB].Audio -eq "" -and $EventosTiempo[$kB].Ispeac -eq "") {
                                $EventosTiempo.Remove($kB) | Out-Null
                            }
                        }
                        continue
                    }
                    if (-not $EsRealB) {
                        # Lo que llega aquí es ruido del sistema: pata phantom de transferencia/conferencia o
                        # bridge sin uso. Las líneas que el asesor abrió y NO marcó ya quedaron etiquetadas
                        # arriba por la confirmación del endpoint ("this record has no far-end address"),
                        # así que aquí solo se suprime.
                        $EventosTiempo[$kB].Interpretacion = ""
                        $EventosTiempo[$kB].ColorInterpretacion = [System.Drawing.Color]::White
                        if ($EventosTiempo[$kB].Agente -eq "" -and $EventosTiempo[$kB].AppLog -eq "" -and $EventosTiempo[$kB].SysLog -eq "" -and $EventosTiempo[$kB].Audio -eq "" -and $EventosTiempo[$kB].Ispeac -eq "") {
                            $EventosTiempo.Remove($kB) | Out-Null
                        }
                    }
                }
            }
            # (El FIN de cada línea vacía ya se reetiquetó arriba, acotado a SU instancia. Antes había aquí un
            #  barrido global por número de sesión que, con los IDs reciclados en cada login, marcaba como
            #  "Cerró línea sin marcar" a llamadas reales —incluso contestadas— de otras horas del día.)

            # --- OMITIR "INICIO DE SESIÓN (Interna/Sistema)" redundante (Pablo, 28/07): es un precursor que
            #     aparece ~2 s antes del "INICIO DE LLAMADA" real de la MISMA sesión y no aporta. Si la sesión
            #     tiene un INICIO DE LLAMADA cercano en el tiempo, se quita esta fila. Se acota por tiempo por
            #     el reciclaje de IDs. Se respeta "INICIO DE SESIÓN (Llamada Interna — Extensión N)". ---
            foreach ($kIS in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kIS)) { continue }
                $oIS = $EventosTiempo[$kIS]
                if ($null -eq $oIS -or $oIS -isnot [hashtable]) { continue }
                if ($oIS.Interpretacion -notmatch "INICIO DE SESIÓN \(Interna/Sistema\)") { continue }
                $sesIS = $oIS.Sesion; if (-not $sesIS -or $sesIS -eq "-") { continue }
                $mIS = & $MsDeSlot $kIS
                $hayInicio = $false
                foreach ($kIL in @($EventosTiempo.Keys)) {
                    if ($kIL -eq $kIS) { continue }
                    $oIL = $EventosTiempo[$kIL]
                    if ($null -eq $oIL -or $oIL -isnot [hashtable]) { continue }
                    if ($oIL.Sesion -ne $sesIS) { continue }
                    if ($oIL.Interpretacion -notmatch "INICIO DE LLAMADA") { continue }
                    $mIL = & $MsDeSlot $kIL
                    if ($mIS -ge 0 -and $mIL -ge 0 -and [math]::Abs($mIL - $mIS) -le 30000) { $hayInicio = $true; break }
                }
                if ($hayInicio) {
                    $oIS.Interpretacion = ""; $oIS.ColorInterpretacion = [System.Drawing.Color]::White
                    if ($oIS.Agente -eq "" -and $oIS.Aux -eq "" -and $oIS.Audio -eq "" -and $oIS.Ispeac -eq "" -and $oIS.SysLog -eq "" -and $oIS.AppLog -eq "") {
                        $EventosTiempo.Remove($kIS) | Out-Null
                    }
                }
            }

            # ================================================================
            # TRANSFERENCIA ENTRANTE (consult transfer): preservar AMBOS registros como
            # evidencia — la fila del ORIGEN (quien transfirió, ej: coordinadora) + la del
            # cliente. Solo se activa si InBoundConsultTransferCompleted confirmó la
            # transferencia (gate anti falsos-positivos en ACD normal).
            # ================================================================
            foreach ($caIdT in @($TransferEntranteOrigen.Keys)) {
                if (-not $TransferEntranteConfirm.ContainsKey($caIdT)) { continue }
                $infoT = $TransferEntranteOrigen[$caIdT]
                $slotRealT = $infoT.SlotReal
                if (-not $EventosTiempo.ContainsKey($slotRealT)) { continue }
                $nombreT = if ($infoT.Nombre -and $infoT.Nombre.Trim() -ne "") { $infoT.Nombre.Trim() } else { "Origen interno" }
                $telT = $infoT.Tel
                $segBaseT = ($slotRealT -split ',')[0]
                # 1) Fila ORIGEN: slot ms 1 milisegundo antes que la señal del cliente, para que ordene
                # justo antes y MUESTRE milisegundos (evita el slot base "sin ms", ambiguo al combinarse).
                $msRealT = ($slotRealT -split ',')[1]
                $slotOrigenT = if ($msRealT -match '^\d+$' -and [int]$msRealT -gt 0) { "$segBaseT,$((([int]$msRealT) - 1).ToString('000'))" } else { $segBaseT }
                Init-Hora $slotOrigenT
                if ($EventosTiempo[$slotOrigenT].Interpretacion -eq "" -or $EventosTiempo[$slotOrigenT].Interpretacion -match "señal de llamada") {
                    $EventosTiempo[$slotOrigenT].Interpretacion      = "$symArr TRANSFERENCIA/CONSULTA ENTRANTE de $nombreT (ext $telT) [CM Auto-Answer]"
                    $EventosTiempo[$slotOrigenT].ColorInterpretacion = [System.Drawing.Color]::Orange
                    if ($EventosTiempo[$slotOrigenT].Tel -eq "-") { $EventosTiempo[$slotOrigenT].Tel = $telT }
                    $EventosTiempo[$slotOrigenT].RawInterpretacion  += "[TRANSFER-IN] Origen=$nombreT (ext $telT) | Confirmado por InBoundConsultTransferCompleted | SessionId=$caIdT`n"
                }
                # 2) Etiquetar la fila de INICIO del cliente con quién transfirió.
                foreach ($kIniT in @($EventosTiempo.Keys)) {
                    if (($kIniT -eq $segBaseT -or $kIniT -match "^$([regex]::Escape($segBaseT)),\d+$") -and
                        $EventosTiempo[$kIniT].Interpretacion -match "INICIO DE LLAMADA \(Entrante" -and
                        $EventosTiempo[$kIniT].Interpretacion -notmatch "Transferida por") {
                        $EventosTiempo[$kIniT].Interpretacion += " $symArr Transferida por $nombreT (ext $telT)"
                    }
                }
            }

            # ================================================================
            # MODO DE CONTESTACIÓN: marca cada INICIO entrante como automático o manual
            # y RETIRA la antigua etiqueta "[CM Auto-Answer]" (sustituida por esto).
            # AUTO  = "Auto Accepting" (CM auto-contestó).
            # MANUAL= "AnswerVoiceInteraction" (efecto del clic en AnswerCallHandler).
            # Ligado por UUID del VI (ViId del slot ↔ ModoContestacion). Marcador POSITIVO:
            # si no hay ninguno de los dos (interna/sistema), no se etiqueta.
            # ================================================================
            foreach ($kMC in @($EventosTiempo.Keys)) {
                $objMC = $EventosTiempo[$kMC]
                if ($null -eq $objMC -or $objMC -isnot [hashtable]) { continue }
                # Retirar la leyenda antigua en cualquier fila donde haya quedado.
                if ($objMC.Interpretacion -match "\[CM Auto-Answer\]") {
                    $objMC.Interpretacion = ($objMC.Interpretacion -replace "\s*\[CM Auto-Answer\]", "")
                }
                # Etiquetar el INICIO entrante con su modo de contestación.
                if ($objMC.Interpretacion -match "INICIO DE LLAMADA \(Entrante" -and
                    $objMC.Interpretacion -notmatch "contestada de forma manual|Entrada de llamada autom") {
                    $uuidMC = $objMC.ViId
                    if ($uuidMC -and $Script:ModoContestacion.ContainsKey($uuidMC)) {
                        if ($Script:ModoContestacion[$uuidMC] -eq "MANUAL") {
                            $objMC.Interpretacion += "  $symUser Llamada contestada de forma manual"
                        } else {
                            $objMC.Interpretacion += "  $symPlay Entrada de llamada automática"
                        }
                    }
                }
            }

            # ================================================================
            # DEDUP FINAL DE INICIO SALIENTE: un solo INICIO por llamada.
            # Varios detectores (XML slot-base, PRIMARY_CONNECTED slot-ms, PASO 5 EndpointLog)
            # pueden crear más de un "INICIO DE LLAMADA (Saliente)" / "INICIO DE SESIÓN" para la
            # MISMA llamada (obs 7: doble ▲/►; obs 14: SALIENTE + INICIO DE SESIÓN). Una llamada =
            # una pata = un solo INICIO. Se agrupan por Sesión (o, si la sesión no se resolvió, por
            # el mismo segundo base con teléfono compatible) y se conserva el más informativo:
            # "INICIO DE LLAMADA (Saliente)" > "INICIO DE SESIÓN"; con ms > sin ms; con teléfono > sin.
            # Corre ANTES del detector de zombies para que éste apunte a la fila que sobrevive.
            # ================================================================
            $GruposInicio = @{}
            foreach ($kI in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kI)) { continue }
                $oI = $EventosTiempo[$kI]
                if ($null -eq $oI -or $oI -isnot [hashtable]) { continue }
                if ($oI.Interpretacion -notmatch "INICIO DE LLAMADA \(Saliente\)|INICIO DE SESIÓN") { continue }
                $baseSeg = ($kI -split ',')[0]
                $clave = if ($oI.Sesion -ne "-" -and $oI.Sesion -ne "") { "S:$($oI.Sesion)" } else { "T:$baseSeg" }
                if (-not $GruposInicio.ContainsKey($clave)) { $GruposInicio[$clave] = @() }
                $GruposInicio[$clave] += $kI
            }
            foreach ($claveG in @($GruposInicio.Keys)) {
                $slotsAllG = @($GruposInicio[$claveG] | Sort-Object { & $MsDeSlot $_ })
                if ($slotsAllG.Count -lt 2) { continue }
                # Los IDs de sesión se RECICLAN en cada login: dos INICIO con el MISMO número de sesión
                # separados por minutos/horas son llamadas DISTINTAS, no duplicados. Los detectores que sí
                # duplican una MISMA llamada (XML slot-base, PRIMARY_CONNECTED, PASO 5) disparan con
                # milisegundos/segundos de diferencia → se parte en racimos y se deduplica DENTRO de cada uno.
                $racimosG = @(); $actualG = @(); $prevMsG = -999999
                foreach ($sG in $slotsAllG) {
                    $mG = & $MsDeSlot $sG
                    if ($mG -lt 0) { continue }
                    if ($actualG.Count -gt 0 -and ($mG - $prevMsG) -gt 30000) { $racimosG += ,$actualG; $actualG = @() }
                    $actualG += $sG; $prevMsG = $mG
                }
                if ($actualG.Count -gt 0) { $racimosG += ,$actualG }
                foreach ($slotsG in $racimosG) {
                if ($slotsG.Count -lt 2) { continue }
                # Grupos por-segundo (sesión desconocida): exigir teléfono compatible para no fusionar dos llamadas distintas.
                if ($claveG -like "T:*") {
                    $telsG = @($slotsG | ForEach-Object { ($EventosTiempo[$_].Tel -replace '\D','') } | Where-Object { $_ -ne "" } | Select-Object -Unique)
                    if ($telsG.Count -gt 1) { continue }
                }
                $ganadorG = $null; $mejorG = -1
                foreach ($sG in $slotsG) {
                    $og = $EventosTiempo[$sG]; $scG = 0
                    if ($og.Interpretacion -match "INICIO DE LLAMADA \(Saliente\)") { $scG += 4 }
                    if ($sG -match ",") { $scG += 2 }
                    if ($og.Tel -ne "-" -and $og.Tel -ne "" -and $og.Tel -notmatch "Desconocido") { $scG += 1 }
                    if ($scG -gt $mejorG) { $mejorG = $scG; $ganadorG = $sG }
                }
                foreach ($sG in $slotsG) {
                    if ($sG -eq $ganadorG) { continue }
                    $perdG = $EventosTiempo[$sG]
                    if (($EventosTiempo[$ganadorG].Tel -eq "-" -or $EventosTiempo[$ganadorG].Tel -match "Desconocido") -and $perdG.Tel -ne "-" -and $perdG.Tel -notmatch "Desconocido") { $EventosTiempo[$ganadorG].Tel = $perdG.Tel }
                    if ($EventosTiempo[$ganadorG].Sesion -eq "-" -and $perdG.Sesion -ne "-") { $EventosTiempo[$ganadorG].Sesion = $perdG.Sesion }
                    $EventosTiempo[$ganadorG].RawInterpretacion += "[DEDUP INICIO] fila duplicada '$($perdG.Interpretacion)' del slot $sG absorbida (grupo $claveG)`n"
                    if ($perdG.Agente -eq "" -and $perdG.Aux -eq "" -and $perdG.Audio -eq "" -and $perdG.Ispeac -eq "" -and $perdG.SysLog -eq "" -and $perdG.AppLog -eq "") {
                        $EventosTiempo.Remove($sG) | Out-Null
                    } else {
                        $perdG.Interpretacion = ""; $perdG.ColorInterpretacion = [System.Drawing.Color]::White
                    }
                }
                }
            }

            # --- SEPARAR "MARCANDO/TIMBRANDO" DE "CONTESTARON" (salientes que sí conectaron) ---
            # Corre DESPUÉS del dedup de INICIO saliente (así hay una sola fila de inicio por sesión).
            # $ContestoSaliente trae el instante en que la VoiceInteraction recibió la dirección del otro
            # lado (= "me contestaron"), coincidente al ms con Alerting→Active. Para cada saliente que
            # conectó: la fila de INICIO se reetiqueta como fase de marcado/timbrado y se inserta una fila
            # nueva "Contestaron — timbró Ns" en el instante de la respuesta; el FIN muestra "(habló Ns)".
            # Corroboración por EVENTO (no por umbral de tiempo): solo se separa si el log muestra la
            # transición Alerting→Active en el mismo instante de la respuesta. La ventana de ±1 s NO es un
            # umbral de comportamiento — es tolerancia de jitter para reconocer que dos líneas del MISMO
            # proceso (RemoteAddress y el screenpop) son el mismo instante (en la práctica coinciden al ms).
            $AAms = @($AlertingActivaSlots | ForEach-Object { & $MsDeSlot $_ } | Where-Object { $_ -ge 0 })
            $ContestadasSes = @{}
            foreach ($ct in $ContestoSaliente) {
                $sesCT = if ($ct.Cxt -and $Script:CxtToConnId.ContainsKey($ct.Cxt)) { $Script:CxtToConnId[$ct.Cxt] } else { "" }
                if (-not $sesCT -or $sesCT -eq "-") { continue }
                # NOTA: no filtrar por $LineaVaciaSinMarcar ni deduplicar por número de sesión aquí: Avaya
                # RECICLA los IDs de sesión en cada login (mismo día → "Id=2" en la mañana y en la tarde son
                # llamadas distintas). El apareo correcto se hace más abajo por FILA (slot), no por número.
                $msAns = & $MsDeSlot $ct.Slot
                if ($msAns -lt 0) { continue }
                # ¿El log confirma Alerting→Active en ese instante? Si no, no fue una respuesta real (p.ej.
                # la dirección se pobló al marcar en algún entorno) → no se separa.
                $AA_OK = $false
                foreach ($aa in $AAms) { if ([math]::Abs($aa - $msAns) -le 1000) { $AA_OK = $true; break } }
                if (-not $AA_OK) { continue }
                # Fila de INICIO saliente de esa sesión: la más cercana ANTERIOR a la respuesta.
                $slotIniC = $null; $msIniC = -1
                foreach ($kC in @($EventosTiempo.Keys)) {
                    $oC = $EventosTiempo[$kC]
                    if ($oC.Sesion -ne $sesCT) { continue }
                    if ($oC.Interpretacion -notmatch "INICIO DE LLAMADA \(Saliente\)|Marcando/Timbrando") { continue }
                    $mC = & $MsDeSlot $kC
                    if ($mC -lt 0 -or $mC -gt $msAns) { continue }
                    if ($mC -gt $msIniC) { $msIniC = $mC; $slotIniC = $kC }
                }
                if (-not $slotIniC) { continue }
                # Dedup por la FILA de inicio concreta (no por número de sesión): con IDs reciclados, la sesión
                # de la mañana y la de la tarde son filas distintas → cada una recibe su separación.
                if ($ContestadasSes.ContainsKey($slotIniC)) { continue }
                $ContestadasSes[$slotIniC] = $msAns
                # DESGLOSE marcó/timbró SACADO DEL LOG (no una sola resta): los dígitos DTMF de esta sesión
                # ($DtmfPresses, con nCallIndex) marcan el tecleo. "marcó" = del primer al último dígito
                # (tecleo del número); "timbró" = del ÚLTIMO dígito hasta que contestaron (timbre real). Si no
                # hubo dígitos (número guardado/rediscado), no hay fase de marcado y "timbró" = desde abrir línea.
                $dialDt = @($DtmfPresses | Where-Object { $_.Ses -eq $sesCT } | ForEach-Object { & $MsDeSlot $_.Hora } | Where-Object { $_ -ge $msIniC -and $_ -le $msAns } | Sort-Object)
                if ($dialDt.Count -ge 1) {
                    $firstDt = $dialDt[0]; $lastDt = $dialDt[-1]
                    $durMarco  = [int](($lastDt - $firstDt) / 1000)
                    $durTimbre = [int](($msAns - $lastDt) / 1000)
                } else {
                    $durMarco  = -1
                    $durTimbre = [int](($msAns - $msIniC) / 1000)
                }
                # Número a mostrar (el de la fila de inicio, o el mapeo, o el capturado).
                $numMostrar = if ($EventosTiempo[$slotIniC].Tel -ne "-" -and $EventosTiempo[$slotIniC].Tel -ne "") { $EventosTiempo[$slotIniC].Tel } elseif ($MapeoTel[$sesCT]) { $MapeoTel[$sesCT] } else { $ct.Num }
                $alNum = if ($numMostrar -and $numMostrar -ne "-") { " al $numMostrar" } else { "" }
                # 1) Reetiquetar el INICIO → fase de marcado/timbrado (+ "marcó en Xs" si hubo tecleo).
                $sufMarco = if ($durMarco -ge 1) { "  ⌨ marcó en ${durMarco}s" } else { "" }
                $EventosTiempo[$slotIniC].Interpretacion      = "$symUp Marcando/Timbrando (saliente)$alNum$sufMarco"
                $EventosTiempo[$slotIniC].ColorInterpretacion = [System.Drawing.Color]::MediumSeaGreen
                $EventosTiempo[$slotIniC].RawInterpretacion  += "[MARCANDO/TIMBRANDO] Sesión $($sesCT): marcó ${durMarco}s (del 1er al último dígito DTMF), timbró ${durTimbre}s (último dígito → contestó a las $($ct.Slot)). durMarco=-1 = número guardado sin tecleo.`n"
                # 2) Insertar fila nueva "Contestaron" en el instante de la respuesta (busca ms libre si choca).
                $slotAns = $ct.Slot
                if ($EventosTiempo.ContainsKey($slotAns) -and $EventosTiempo[$slotAns].Interpretacion -ne "") {
                    $segA = ($slotAns -split ',')[0]; $msA = [int]($slotAns -split ',')[1]
                    for ($j=1; $j -le 8; $j++) { $cand = "$segA,$((($msA+$j)).ToString('000'))"; if (-not ($EventosTiempo.ContainsKey($cand) -and $EventosTiempo[$cand].Interpretacion -ne "")) { $slotAns = $cand; break } }
                }
                Init-Hora $slotAns
                $EventosTiempo[$slotAns].Interpretacion      = "$symPhone Contestaron — timbró ${durTimbre}s"
                $EventosTiempo[$slotAns].ColorInterpretacion = [System.Drawing.Color]::LimeGreen
                if ($EventosTiempo[$slotAns].Sesion -eq "-") { $EventosTiempo[$slotAns].Sesion = $sesCT }
                if ($EventosTiempo[$slotAns].Tel -eq "-" -and $numMostrar -ne "-") { $EventosTiempo[$slotAns].Tel = $numMostrar }
                $EventosTiempo[$slotAns].RawInterpretacion  += "[CONTESTARON] Sesión $($sesCT): el otro lado contestó (VoiceInteraction.RemoteAddress=$($ct.Num), coincide con Alerting→Active). Timbró ${durTimbre}s (desde el último dígito marcado).`n"
                # 3) Duración de conversación en el FIN de esa sesión.
                $slotFinC = $null; $msFinC = [double]::PositiveInfinity
                foreach ($kF in @($EventosTiempo.Keys)) {
                    $oF = $EventosTiempo[$kF]
                    if ($oF.Sesion -ne $sesCT) { continue }
                    if ($oF.Interpretacion -notmatch "FIN DE LLAMADA|CUELGUE MANUAL|CLIENTE COLGÓ|Cerró línea sin marcar") { continue }
                    $mF = & $MsDeSlot $kF
                    if ($mF -lt $msAns) { continue }
                    if ($mF -lt $msFinC) { $msFinC = $mF; $slotFinC = $kF }
                }
                # CONTRADICCIÓN: si el log confirmó que CONTESTARON, la línea sí se marcó. Un "Cerró línea sin
                # marcar" sobre esa misma llamada es imposible → se restaura a un FIN normal.
                if ($slotFinC -and $EventosTiempo[$slotFinC].Interpretacion -match "Cerró línea sin marcar") {
                    $EventosTiempo[$slotFinC].Interpretacion      = "$symStop FIN DE LLAMADA"
                    $EventosTiempo[$slotFinC].ColorInterpretacion = [System.Drawing.Color]::IndianRed
                    $EventosTiempo[$slotFinC].RawInterpretacion  += "[CORRECCIÓN] Se descartó 'Cerró línea sin marcar': esta llamada fue CONTESTADA (RemoteAddress + Alerting→Active), luego sí se marcó.`n"
                }
                if ($slotFinC -and $EventosTiempo[$slotFinC].Interpretacion -notmatch "habló") {
                    $durTalk = [int](($msFinC - $msAns) / 1000)
                    if ($durTalk -ge 0 -and $durTalk -le 7200) {
                        $EventosTiempo[$slotFinC].Interpretacion += " (habló $(& $FmtMmSs $durTalk))"
                    }
                }
            }

            # --- MISMO "(habló mm:ss)" PERO PARA ENTRANTES ---
            # "INICIO DE LLAMADA (Entrante)" ya se coloca en el instante de la CONEXIÓN real (Active), no en
            # el instante en que empezó a timbrar (el tiempo de timbre queda aparte en "[Ring: Ns]" dentro de
            # la misma fila) — así que, a diferencia de saliente, aquí no hace falta separar una fase de
            # "marcando/timbrando": la duración de la conversación es directamente FIN - INICIO(Entrante) de
            # la misma sesión. Pablo, 18/08/2026.
            foreach ($kIE in @($EventosTiempo.Keys)) {
                $oIE = $EventosTiempo[$kIE]
                if ($oIE.Interpretacion -notmatch "INICIO DE LLAMADA \(Entrante\)") { continue }
                $sesIE = $oIE.Sesion
                if (-not $sesIE -or $sesIE -eq "-") { continue }
                $msIniE = & $MsDeSlot $kIE
                if ($msIniE -lt 0) { continue }
                $slotFinE = $null; $msFinE = [double]::PositiveInfinity
                foreach ($kFE in @($EventosTiempo.Keys)) {
                    $oFE = $EventosTiempo[$kFE]
                    if ($oFE.Sesion -ne $sesIE) { continue }
                    if ($oFE.Interpretacion -notmatch "FIN DE LLAMADA|CUELGUE MANUAL|CLIENTE COLGÓ|Cerró línea sin marcar") { continue }
                    $mFE = & $MsDeSlot $kFE
                    if ($mFE -lt $msIniE) { continue }
                    if ($mFE -lt $msFinE) { $msFinE = $mFE; $slotFinE = $kFE }
                }
                if ($slotFinE -and $EventosTiempo[$slotFinE].Interpretacion -notmatch "habló") {
                    $durTalkE = [int](($msFinE - $msIniE) / 1000)
                    if ($durTalkE -ge 0 -and $durTalkE -le 7200) {
                        $EventosTiempo[$slotFinE].Interpretacion += " (habló $(& $FmtMmSs $durTalkE))"
                    }
                }
            }

            # --- RESPALDO "(habló mm:ss)" PARA SALIENTES INTERNAS (extensión a extensión) ---
            # Estas nunca pasan por el bloque de arriba (foreach $ContestoSaliente) porque no generan ni el
            # RemoteAddress= ni la corroboración Alerting→Active: conectan directo (New→Active) sin timbrar.
            # Aquí se usa $ConexionActivaSaliente (capturado en PASO 4) como el instante real de conexión —
            # solo aplica a filas que quedaron como "INICIO DE LLAMADA (Saliente)" SIN convertir a
            # "Marcando/Timbrando" (esas ya se resolvieron arriba con su propio instante de "contestó").
            # Caso real: extensión 45011/45407, 18/08/2026 (Pablo).
            foreach ($kIS in @($EventosTiempo.Keys)) {
                $oIS = $EventosTiempo[$kIS]
                if ($oIS.Interpretacion -notmatch "INICIO DE LLAMADA \(Saliente\)") { continue }
                $sesIS = $oIS.Sesion
                if (-not $sesIS -or $sesIS -eq "-") { continue }
                if (-not $ConexionActivaSaliente.ContainsKey($sesIS)) { continue }
                $msConectoIS = & $MsDeSlot $ConexionActivaSaliente[$sesIS]
                if ($msConectoIS -lt 0) { continue }
                $slotFinS = $null; $msFinS = [double]::PositiveInfinity
                foreach ($kFS in @($EventosTiempo.Keys)) {
                    $oFS = $EventosTiempo[$kFS]
                    if ($oFS.Sesion -ne $sesIS) { continue }
                    if ($oFS.Interpretacion -notmatch "FIN DE LLAMADA|CUELGUE MANUAL|CLIENTE COLGÓ|Cerró línea sin marcar") { continue }
                    $mFS = & $MsDeSlot $kFS
                    if ($mFS -lt $msConectoIS) { continue }
                    if ($mFS -lt $msFinS) { $msFinS = $mFS; $slotFinS = $kFS }
                }
                if ($slotFinS -and $EventosTiempo[$slotFinS].Interpretacion -notmatch "habló") {
                    $durTalkS = [int](($msFinS - $msConectoIS) / 1000)
                    if ($durTalkS -ge 0 -and $durTalkS -le 7200) {
                        $EventosTiempo[$slotFinS].Interpretacion += " (habló $(& $FmtMmSs $durTalkS))"
                    }
                }
            }

            # --- PLEGAR AUDIO ABIERTO/CERRADO EN LA FILA QUE LO PROVOCÓ (mismo segundo) ---
            # "Audio abierto/cerrado" cae en el MISMO segundo que el evento que lo causó (abrir línea / fin
            # de llamada). En vez de una fila propia que solo puebla la columna Audio, se mete en la fila de
            # ese evento. La renegociación de medios NO se pliega (es un evento propio significativo).
            $AbreReAu   = "INICIO DE LLAMADA|Marcando/Timbrando|Contestaron|INICIO DE SESIÓN|LÍNEA ABIERTA|Abrió .*sin marcar|señal de llamada"
            $CierraReAu = "FIN DE LLAMADA|CUELGUE MANUAL|Cerró línea|CLIENTE COLGÓ|EVASIÓN"
            foreach ($kAudF in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kAudF)) { continue }
                $oAudF = $EventosTiempo[$kAudF]
                if ($null -eq $oAudF -or $oAudF -isnot [hashtable]) { continue }
                if ($oAudF.Audio -notmatch "Audio abierto|Audio cerrado") { continue }
                if ($oAudF.Audio -match "renegoci") { continue }
                # Si esta misma fila ya trae una interpretación, no es una fila de solo-audio → no tocar.
                if ($oAudF.Interpretacion -ne "") { continue }
                $esAbreAu = ($oAudF.Audio -match "Audio abierto")
                $baseSecF = ($kAudF -split ',')[0]
                $msAudF = & $MsDeSlot $kAudF
                # Buscar la fila destino en el MISMO segundo: el evento que lo provocó (abrir / fin).
                $tgtSlotF = $null; $tgtDistF = 1e18
                foreach ($kT in @($EventosTiempo.Keys)) {
                    if ($kT -eq $kAudF) { continue }
                    if (($kT -split ',')[0] -ne $baseSecF) { continue }
                    $oT = $EventosTiempo[$kT]
                    if ($null -eq $oT -or $oT -isnot [hashtable] -or $oT.Interpretacion -eq "" -or $oT.Audio -ne "") { continue }
                    $reOKau = if ($esAbreAu) { $oT.Interpretacion -match $AbreReAu } else { $oT.Interpretacion -match $CierraReAu }
                    if (-not $reOKau) { continue }
                    $mT = & $MsDeSlot $kT
                    $distF = [math]::Abs($msAudF - $mT)
                    if ($distF -lt $tgtDistF) { $tgtDistF = $distF; $tgtSlotF = $kT }
                }
                if (-not $tgtSlotF) { continue }
                # Mover el audio a la fila destino y eliminar el slot de solo-audio.
                $EventosTiempo[$tgtSlotF].Audio      = $oAudF.Audio
                $EventosTiempo[$tgtSlotF].ColorAudio = $oAudF.ColorAudio
                $EventosTiempo[$tgtSlotF].RawAudio  += $oAudF.RawAudio
                $oAudF.Audio = ""; $oAudF.ColorAudio = [System.Drawing.Color]::White
                if ($oAudF.Agente -eq "" -and $oAudF.Aux -eq "" -and $oAudF.Ispeac -eq "" -and $oAudF.SysLog -eq "" -and $oAudF.AppLog -eq "" -and $oAudF.Dtmf -eq "") {
                    $EventosTiempo.Remove($kAudF) | Out-Null
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
            # CONFIRMACIÓN GUI: sella "✓ clic confirmado" en HOLD/UNHOLD/CUELGUE MANUAL
            # cuando en el MISMO segundo hubo un clic GUI (HoldCallHandler/UnHoldCallHandler/
            # EndCallHandler). El evento real lo pinta el EndpointLog (OnRequest*); el GUI es
            # la corroboración de que fue acción del asesor (no del sistema).
            # ================================================================
            foreach ($kGC in @($EventosTiempo.Keys)) {
                if (-not $EventosTiempo.ContainsKey($kGC)) { continue }
                $oGC = $EventosTiempo[$kGC]
                if ($null -eq $oGC -or $oGC -isnot [hashtable] -or $oGC.Agente -eq "") { continue }
                if ($oGC.Agente -match "clic confirmado") { continue }
                $segGC = ($kGC -split ',')[0]
                if     ($oGC.Agente -match "HOLD MANUAL"    -and $GuiHoldSeg[$segGC])   { $oGC.Agente += "  $symOK clic confirmado" }
                elseif ($oGC.Agente -match "UNHOLD MANUAL"  -and $GuiUnholdSeg[$segGC]) { $oGC.Agente += "  $symOK clic confirmado" }
                elseif ($oGC.Agente -match "CUELGUE MANUAL" -and $GuiEndSeg[$segGC])    { $oGC.Agente += "  $symOK clic confirmado" }
            }

            # ================================================================
            # INTENTOS DE DESFIRME FALLIDOS: se arma la fila AQUÍ, no en PASO 4 (Pablo, 09/09/2026)
            # PASO 4 y PASO 5 (incluido su barrido de "FIN DE LLAMADA" fallback, que sobreescribe sin
            # avisar) ya terminaron de escribir todo — recién aquí "slot libre" significa libre de verdad.
            # ================================================================
            foreach ($df in $DesfirmeFallidoList) {
                $_msFalloProbe = [int]$df.Ms
                $SlotDesfirmeFallo = "$($df.Hora),$($_msFalloProbe.ToString('000'))"
                $_intentosFallo = 0
                while ($EventosTiempo.ContainsKey($SlotDesfirmeFallo) -and $EventosTiempo[$SlotDesfirmeFallo].Interpretacion -ne "" -and $_intentosFallo -lt 50) {
                    $_msFalloProbe = [math]::Min($_msFalloProbe + 1, 999)
                    $SlotDesfirmeFallo = "$($df.Hora),$($_msFalloProbe.ToString('000'))"
                    $_intentosFallo++
                }
                Init-Hora $SlotDesfirmeFallo
                if ($EventosTiempo[$SlotDesfirmeFallo].Interpretacion -eq "") {
                    $EventosTiempo[$SlotDesfirmeFallo].Interpretacion      = "⚠ Intento de desfirme — FALLÓ (no se pudo desconectar la llamada)"
                    $EventosTiempo[$SlotDesfirmeFallo].ColorInterpretacion = [System.Drawing.Color]::OrangeRed
                    $EventosTiempo[$SlotDesfirmeFallo].RawInterpretacion  += "[DESFIRME FALLIDO] $($df.Raw)`n"
                }
            }

            # ================================================================
            # PASO 6: CEREBRO FORENSE V22 — RENDERIZADO
            # ================================================================
            $lblStatus.Text = "PASO 6/6: Renderizando Cerebro Forense..."; $Form.Refresh()
            $HorasOrdenadas = $EventosTiempo.Keys | Sort-Object
            $CurrentAux = "$symUser Estado: USUARIO DESFIRMADO"; $CurrentColor = [System.Drawing.Color]::Gray; $CurrentReasonCode = ""
            # $UltimoMotivoElegido: memoria "pegajosa" del último reason code EXPLÍCITO que el asesor eligió.
            # A diferencia de $CurrentReasonCode (que se limpia cada vez que el asesor pasa por Disponible),
            # esta NO se limpia con Disponible — solo cuando se elige un motivo distinto. Replica el
            # comportamiento real de Avaya/CM: si el asesor usa TrabAux (botón favorito) sin elegir motivo,
            # el CM mantiene el ÚLTIMO auxiliar usado (confirmado por Pablo: la GUI de OneX mostraba "Aux
            # Sistemas" en ese caso, aunque hubiera pasado por Disponible antes). Pablo, prueba 28/08/2026.
            $UltimoMotivoElegido = ""
            $RecienFirmado = $false; $ValidandoLogin = $false; $LoginFallido = $false
            $LoginYaConfirmado = $false   # login confirmado en este ciclo; se reabre solo con un desfirme real (evita doble "firmado exitosamente" por eventos intermedios)
            # DEFAULT AUTOMÁTICO POST-LOGIN: Avaya dispara "Enter Aux;code=ReasonCode[0]" DOS veces al firmarse
            # (no es un clic del asesor). El flag $RecienFirmado no sirve como ancla: cualquier fila con contenido
            # (Entry.ConnectinoID, Línea cerrada…) lo resetea. Tampoco sirve una ventana de tiempo: el asesor
            # puede tardar lo que quiera en ponerse Disponible. Se usa el ESTADO: el 1er DEFAULT tras el logon
            # es el automático; mientras el asesor siga en ese DEFAULT, cualquier repetición es re-emisión de Avaya.
            $EsperandoDefaultPostLogin = $false  # hubo logon, aún no llega el DEFAULT automático
            $EnDefaultPostLogin        = $false  # el asesor sigue en el DEFAULT post-logon → DEFAULTs repetidos = duplicados
            $LlamadaActiva = $false; $SesionActual = "-"
            # $DictRC ahora se declara arriba (antes de PASO 4/5) para que ambos lo compartan — ver Opción B.

            # ── POST-PROCESO: Colapsar dígitos intermedios de marcación manual ──────────────
            # La PBX da 10 segundos entre cada dígito. Al marcar p.ej. extensión 2222 el log
            # genera: LÍNEA ABIERTA tel=2, tel=22, tel=222, tel=2222 (4 filas redundantes).
            # Algoritmo: recorre pares consecutivos de eventos LÍNEA ABIERTA; si la diferencia
            # de tiempo entre ellos es ≤10 s Y el tel del primero es prefijo numérico del
            # segundo → el primero es un dígito intermedio → marcarlo como colapsado (no mostrar).
            # Se usa la ventana POR PAR (no total) para soportar números largos sin colapsar
            # llamadas distintas que coincidan en prefijo pero estén separadas más de 10 s.
            $LineaAbSlots = @($HorasOrdenadas | Where-Object {
                $s = $EventosTiempo[$_]
                $null -ne $s -and $s -is [hashtable] -and
                $s.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR" -and
                ($s.Tel -ne "-") -and ($s.Tel -ne "") -and ($s.Tel -ne "Desconocido")
            })
            for ($iDial = 0; $iDial -lt ($LineaAbSlots.Count - 1); $iDial++) {
                $SlotDA = $EventosTiempo[$LineaAbSlots[$iDial]]
                $SlotDB = $EventosTiempo[$LineaAbSlots[$iDial + 1]]
                $TelDA  = $SlotDA.Tel -replace '\D',''
                $TelDB  = $SlotDB.Tel -replace '\D',''
                # Condición 1: TelA debe ser prefijo estricto de TelB (mismo inicio, menor longitud)
                if ($TelDA -eq "" -or $TelDB -eq "" -or $TelDA.Length -ge $TelDB.Length) { continue }
                if (-not $TelDB.StartsWith($TelDA)) { continue }
                # Condición 2: diferencia de tiempo entre este par ≤ 10 segundos (ventana por par)
                try {
                    $HoraDA = [datetime]::ParseExact(($LineaAbSlots[$iDial]     -split ',')[0], "HH:mm:ss", $null)
                    $HoraDB = [datetime]::ParseExact(($LineaAbSlots[$iDial + 1] -split ',')[0], "HH:mm:ss", $null)
                    $DiffDial = [int]($HoraDB - $HoraDA).TotalSeconds
                    if ($DiffDial -ge 0 -and $DiffDial -le 10) {
                        $SlotDA.CollapsedByDialing = $true   # suprimir dígito intermedio en la tabla
                    }
                } catch {}
            }
            # ─────────────────────────────────────────────────────────────────────────────────

            # ── CALIDAD DE RED en filas EXISTENTES (sin crear slots) ─────────────────────────
            # En vez de una fila por lectura RTCP (que saturaba el timeline), se ADJUNTA el estado
            # de red a cada evento REAL (Interpretacion != "") buscando la lectura más cercana en el
            # tiempo (±$VentCalidad). El IspeacLog solo tiene lecturas durante la llamada, así que los
            # eventos fuera de llamada quedan en blanco de forma natural. La columna entera se muestra/
            # oculta con la casilla "Ver Calidad de Red" (no estorba apagada). Verde=sano, ámbar=degradado,
            # rojo=crítico (pérdida ≥10% o RTT ≥100 ms). Solo se toca Interpretacion != "" para NO
            # destapar filas ocultas (badges de estado se saltan si están vacíos + Ispeac == "").
            if (($LossReadings.Count + $RttReadings.Count) -gt 0) {
                $VentCalidad = 20000   # ventana ±20 s (las lecturas RTCP van ~cada 5 s durante la llamada)
                $lossMs = @($LossReadings | ForEach-Object { [pscustomobject]@{ Ms = (& $MsDeSlot $_.Slot); Pct = $_.Pct; Raw = $_.Raw } } | Where-Object { $_.Ms -ge 0 })
                $rttMs  = @($RttReadings  | ForEach-Object { [pscustomobject]@{ Ms = (& $MsDeSlot $_.Slot); Rtt = $_.Rtt; Raw = $_.Raw } } | Where-Object { $_.Ms -ge 0 })
                foreach ($kQ in @($EventosTiempo.Keys)) {
                    $oQ = $EventosTiempo[$kQ]
                    if ($null -eq $oQ -or $oQ -isnot [hashtable]) { continue }
                    if ($oQ.Interpretacion -eq "" -or $oQ.Ispeac -ne "") { continue }
                    $mQ = & $MsDeSlot $kQ
                    if ($mQ -lt 0) { continue }
                    $nl = $null; $nlD = $VentCalidad + 1
                    foreach ($lr in $lossMs) { $d = [math]::Abs($lr.Ms - $mQ); if ($d -le $VentCalidad -and $d -lt $nlD) { $nl = $lr; $nlD = $d } }
                    $nr = $null; $nrD = $VentCalidad + 1
                    foreach ($rr in $rttMs)  { $d = [math]::Abs($rr.Ms - $mQ); if ($d -le $VentCalidad -and $d -lt $nrD) { $nr = $rr; $nrD = $d } }
                    if ($null -eq $nl -and $null -eq $nr) { continue }
                    $sevQ = 0
                    if ($nl) { if ($nl.Pct -ge 10) { $sevQ = 2 } elseif ($nl.Pct -gt 0 -and $sevQ -lt 1) { $sevQ = 1 } }
                    if ($nr) { if ($nr.Rtt -ge 100) { $sevQ = 2 } elseif ($nr.Rtt -ge 50 -and $sevQ -lt 1) { $sevQ = 1 } }
                    $detQ = @()
                    if ($nr) { $detQ += "RTT $($nr.Rtt) ms" }
                    if ($nl) { $detQ += "pérdida $($nl.Pct)%" }
                    $detStrQ = $detQ -join ", "
                    if     ($sevQ -eq 2) { $oQ.Ispeac = "¡ALERTA RED! ($detStrQ)";   $oQ.ColorIspeac = [System.Drawing.Color]::OrangeRed }
                    elseif ($sevQ -eq 1) { $oQ.Ispeac = "⚠ Red degradada ($detStrQ)"; $oQ.ColorIspeac = [System.Drawing.Color]::Gold }
                    else                 { $oQ.Ispeac = "♪ Red OK ($detStrQ)";        $oQ.ColorIspeac = [System.Drawing.Color]::MediumSpringGreen }
                    if ($nr) { $oQ.RawIspeac += "$($nr.Raw)`n" }
                    if ($nl) { $oQ.RawIspeac += "$($nl.Raw)`n" }
                }
            }
            # ─────────────────────────────────────────────────────────────────────────────────

            $GridResultados.SuspendLayout()
            if ($HorasOrdenadas -ne $null -and $HorasOrdenadas.Count -gt 0) {
                foreach ($H in $HorasOrdenadas) {
                    $Obj = $EventosTiempo[$H]
                    if ($null -eq $Obj -or $Obj -isnot [hashtable]) { continue }   # guard: slots siempre deben ser hashtable
                    if ($Obj.CollapsedByDialing -eq $true) { continue }              # guard: dígito intermedio de marcación — no mostrar
                    if ($Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "" -and $Obj.Aux -eq "" -and $Obj.Interpretacion -eq "" -and $Obj.Ispeac -eq "" -and $Obj.Dtmf -eq "" -and $Obj.RawAux -notmatch "Session_LoginAgent failed") { continue }
                    # Suprimir slot base cuando la señal de llamada fue movida al slot ms por auto-in en el mismo segundo.
                    # El slot base queda con Aux/RawAux del estado del agente pero Interpretacion="" — no aporta fila visible.
                    # PERO actualizar PRIMERO el estado vigente: AUTOIN_DISP_SLOT = el agente activó auto-in → DISPONIBLE.
                    # Si no, $CurrentAux se quedaría en el AUXILIAR previo y el INICIO siguiente del mismo segundo se
                    # marcaría erróneamente como "recibida en AUXILIAR" (y la columna de estado mostraría AUXILIAR).
                    if ($Obj.Interpretacion -eq "" -and $Obj.Aux -match "AUTOIN_DISP_SLOT") {
                        $CurrentReasonCode = ""; $CurrentAux = "$symUser Estado: DISPONIBLE"; $CurrentColor = [System.Drawing.Color]::LimeGreen
                        continue
                    }
                    # DEDUP "Disponible (Confirmado por clic)": el slot base (GUI_READY_CONFIRMADO, sin ms)
                    # duplicaba la fila que EnterReadyHandler ENDED ya creó en el slot ms del mismo segundo.
                    # Si existe la hermana con ms, se omite la base y solo se actualiza el estado vigente.
                    # Si NO existe (raro: el ms-slot estaba ocupado), la base se conserva como respaldo.
                    # Se incluye también AUTOIN_DISP_SLOT y el AgentStateChanged=Ready: el slot base del
                    # segundo de auto-in recibía el estado Ready (y a veces Sesión/Teléfono de la llamada
                    # entrante del mismo segundo) y se pintaba como una 2ª fila "Asesor se cambia a Disponible".
                    if ($Obj.Interpretacion -eq "" -and $H -notmatch ',' -and
                        ($Obj.Aux -match "GUI_READY_CONFIRMADO|AUTOIN_DISP_SLOT" -or $Obj.RawAux -match "(?i)AgentStateChanged.*newState=Ready") -and
                        $Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.Ispeac -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "") {
                        $HaySibDisp = $false
                        foreach ($kSib in $EventosTiempo.Keys) {
                            if ($kSib -match ("^" + [regex]::Escape($H) + ",\d+$") -and $EventosTiempo[$kSib].Interpretacion -match "se cambia a Disponible") { $HaySibDisp = $true; break }
                        }
                        if ($HaySibDisp) {
                            $CurrentReasonCode = ""; $CurrentAux = "$symUser Estado: DISPONIBLE"; $CurrentColor = [System.Drawing.Color]::LimeGreen
                            # BUG preexistente corregido (Pablo, 08/2026): el reset de $EsperandoDefaultPostLogin/
                            # $EnDefaultPostLogin vivía SOLO dentro del cuerpo de las ramas "GUI_READY_CONFIRMADO"/
                            # "newState=Ready" de la cadena de prioridad de abajo — pero esta fila base NUNCA
                            # llega a esa cadena (el "continue" de aquí arriba la intercepta antes). Resultado:
                            # en el flujo NORMAL (con este dedup, que es el caso de casi TODAS las firmas), esas
                            # banderas se quedaban en "true" para siempre tras el primer Default post-login,
                            # hasta hoy invisible porque nada las consultaba — pero ahora sí (ver rama de
                            # Auxiliar-vía-favorito, Punto 1), así que hacía falta este fix para que no se
                            # tragara eventos reales de Auxiliar más adelante en el día creyendo que aún era el
                            # Default automático del login.
                            $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false
                            continue
                        }
                    }
                    # DEDUP "Auxiliar [motivo] (Confirmado por clic)" (Opción B): mismo patrón que "Disponible"
                    # arriba. El slot base (segundo compartido, con los tokens RC/NUEVO_RC/GUI_AUX_CONFIRMADO)
                    # duplicaba la fila que EnterAuxWithReasonCodeHandler ENDED ya creó en su slot ms (PASO 4).
                    # Si existe la hermana con ms, se omite la base y solo se actualiza el estado vigente ahí
                    # mismo (recalculando el código desde la base, por si la hermana no llegó a armarse).
                    if ($Obj.Interpretacion -eq "" -and $H -notmatch ',' -and
                        ($Obj.Aux -match "GUI_AUX_CONFIRMADO|NUEVO_RC:|RC: " -or $Obj.RawAux -match "(?i)AgentStateChanged.*newState=Aux") -and
                        $Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.Ispeac -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "") {
                        $HaySibAux = $false; $SibEsResiduoDesfirme = $false
                        foreach ($kSib in $EventosTiempo.Keys) {
                            if ($kSib -match ("^" + [regex]::Escape($H) + ",\d+$") -and $EventosTiempo[$kSib].Interpretacion -match "se cambia a Auxiliar|se cambia a Default|Auxiliar pendiente aplicado|Se libera el estado") {
                                $HaySibAux = $true
                                if ($EventosTiempo[$kSib].Interpretacion -match "Se libera el estado") { $SibEsResiduoDesfirme = $true }
                                break
                            }
                        }
                        if ($HaySibAux -and $SibEsResiduoDesfirme) {
                            # Residuo de un intento de desfirme fallido (Pablo, 01/09/2026): NO es un motivo real
                            # elegido por el asesor — no hay que inferir/heredar $UltimoMotivoElegido aquí, solo
                            # reflejar que quedó en Auxiliar sin motivo confirmado.
                            $CurrentAux = "$symUser Estado: AUXILIAR"; $CurrentColor = [System.Drawing.Color]::LightCoral
                            $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false
                            continue
                        }
                        if ($HaySibAux) {
                            $CodigoRCDed = ""
                            if ($Obj.Aux -match "NUEVO_RC:(\d+)") { $CodigoRCDed = $matches[1] }
                            elseif ($Obj.RawAux -match "(?i)ReasonCode[=\[>:\s]*(\d+)") { $CodigoRCDed = $matches[1] }
                            # Finalización diferida de un PendingAux (Punto 2): esta fila base NO trae código propio
                            # (el motivo se eligió varios segundos antes, en OTRO segundo) — se usa el que ya quedó
                            # registrado en $UltimoMotivoElegido cuando se procesó aquel clic original.
                            elseif ($UltimoMotivoElegido -ne "" -and $UltimoMotivoElegido -ne "0") { $CodigoRCDed = $UltimoMotivoElegido }
                            if ($CodigoRCDed -ne "") {
                                if ($CodigoRCDed -eq "0") {
                                    # Motivo explícito = Default (elegido a propósito desde el menú, no el
                                    # automático post-login) — se muestra como DEFAULT, no "AUXILIAR (DEFAULT)".
                                    $CurrentReasonCode = "0"; $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue
                                } else {
                                    $NombreRCDed = if ($DictRC.ContainsKey($CodigoRCDed)) { $DictRC[$CodigoRCDed] } else { $CodigoRCDed }
                                    $CurrentReasonCode = $CodigoRCDed; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRCDed)"; $CurrentColor = [System.Drawing.Color]::Orange
                                    $UltimoMotivoElegido = $CodigoRCDed
                                }
                            }
                            # Ver nota del bug preexistente arriba (dedup de Disponible): mismo fix aquí — un
                            # motivo elegido A PROPÓSITO desde el menú (con o sin código) ya no es ambigüedad
                            # de "recién firmado", así que se limpian las banderas del Default post-login.
                            $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false
                            continue
                        }
                    }
                    # DEDUP "Llamada finalizada por el agente": el slot base (CUELGUE MANUAL del clic, sin ms)
                    # duplicaba el "FIN DE LLAMADA MANUAL (Colgada por el Asesor)" que ProcessSessionEndedEvent
                    # ya creó en el slot ms del mismo segundo. Si existe la hermana FIN con ms, se omite la base.
                    if ($Obj.Interpretacion -eq "" -and $Obj.Agente -match "CUELGUE MANUAL|LLAMADA COLGADA POR EL ASESOR" -and
                        $Obj.Audio -eq "" -and $Obj.Ispeac -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "") {
                        # CUELGUE MANUAL ahora vive en slot ms (Categoría A). Buscamos una hermana FIN en el
                        # MISMO segundo (base o ms, distinta de esta) para omitir el duplicado del clic.
                        $segCM = ($H -split ',')[0]
                        $HaySibFin = $false
                        foreach ($kSib in $EventosTiempo.Keys) {
                            if ($kSib -ne $H -and ($kSib -split ',')[0] -eq $segCM -and $EventosTiempo[$kSib].Interpretacion -match "FIN DE LLAMADA") { $HaySibFin = $true; break }
                        }
                        if ($HaySibFin) { $LlamadaActiva = $false; $SesionActual = "-"; continue }
                    }
                    # Suprimir filas sin evento en Interpretación cuyo único contenido sea la etiqueta interna de UpdateHistoryRecord.
                    # Ocurre en llamadas ENTRANTE: UpdateHistoryRecord actualiza $MapeoTel pero no genera un evento visible.
                    # Si además hay datos de AppLog o SysLog (errores) la fila sí se conserva.
                    if ($Obj.Interpretacion -eq "" -and $Obj.Agente -match "^UpdateHistoryRecord: SessionID=" -and $Obj.AppLog -eq "" -and $Obj.SysLog -eq "") { continue }

                    # --- Suprimir el FAC del botón favorito AUTO-IN (NO es una llamada real) ---
                    # El botón AUTO-IN dispara el FAC $FacAutoIn, que nace como saliente SIN destino
                    # (RemoteParty=[,]) y por eso se cuela como "LÍNEA ABIERTA SIN MARCAR" — a veces
                    # DUPLICADA: una fila por el OneXAgent.log (Tel=$FacAutoIn) y otra por el XML de
                    # contactos (Tel=Desconocido). Ya lo identificamos por el número; aquí se eliminan
                    # TODAS sus filas (los dos "sin marcar" + su FIN, con el audio ya plegado dentro),
                    # dejando solo "Disponible usando botón favorito AUTO-IN". El guard exige Tel=$FacAutoIn
                    # o etiqueta "sin marcar" → nunca toca la fila de Disponible (Tel="-", otra etiqueta).
                    if ($AutoInFacSlots.Count -gt 0 -and ($Obj.Tel -eq $FacAutoIn -or $Obj.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR")) {
                        $msRowFac = & $MsDeSlot $H
                        if ($msRowFac -ge 0) {
                            $esFilaFac = $false
                            foreach ($sFac in $AutoInFacSlots) {
                                $mdFac = & $MsDeSlot $sFac
                                if ($mdFac -ge 0 -and [math]::Abs($msRowFac - $mdFac) -le 3000) { $esFilaFac = $true; break }
                            }
                            if ($esFilaFac) { $LlamadaActiva = $false; $SesionActual = "-"; continue }
                        }
                    }

                    # --- Suprimir la pata phantom del handshake de login (NO es una línea abandonada) ---
                    # Durante el método Login() (registro de estación) Avaya crea una pata saliente vacía
                    # que nace y muere en ~1s, ANTES de que el agente esté firmado → se cuela como "LÍNEA
                    # ABIERTA SIN MARCAR" y como "¡EVASIÓN!". No es un abandono real. Se eliminan ambas filas
                    # si su slot cae dentro de la ventana [Begin..End] del Login() (± margen). El agente ya
                    # queda representado por "Usuario intentando firmarse" / "Fallo en el intento de firmarse".
                    if ($LoginWins.Count -gt 0 -and $Obj.Interpretacion -match "LÍNEA ABIERTA SIN MARCAR|EVASIÓN") {
                        $msRowLg = & $MsDeSlot $H
                        if ($msRowLg -ge 0) {
                            $esFilaLogin = $false
                            foreach ($lw in $LoginWins) {
                                $bMs = & $MsDeSlot $lw.B; $eMs = & $MsDeSlot $lw.E
                                if ($bMs -ge 0 -and $eMs -ge 0 -and $msRowLg -ge ($bMs - 500) -and $msRowLg -le ($eMs + 500)) { $esFilaLogin = $true; break }
                            }
                            if ($esFilaLogin) { $LlamadaActiva = $false; $SesionActual = "-"; continue }
                        }
                    }

                    if ($Obj.Interpretacion -match "INICIO DE LLAMADA|LÍNEA ABIERTA") { $LlamadaActiva = $true; $SesionActual = $Obj.Sesion }
                    elseif ($Obj.Interpretacion -match "FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") { $LlamadaActiva = $false; $SesionActual = "-" }

                    $Interp = $Obj.Interpretacion; $ColorInterp = $Obj.ColorInterpretacion

                    # Activar ValidandoLogin desde XML (LoggedIn) o desde Amnesia V21
                    # Guard $RecienFirmado: evita re-trigger si Avaya emite LOGGEDIN desde 2 fuentes (XML + OneXAgent)
                    if ($Obj.Aux -match "ESTADO: LOGGEDIN" -and -not $RecienFirmado -and -not $LoginYaConfirmado) { $ValidandoLogin = $true }
                    # "Intentando firmarse" = inicio de un login NUEVO → reabre la validación (resetea el flag de confirmado).
                    if ($Interp -match "Usuario intentando firmarse") { $ValidandoLogin = $true; $RecienFirmado = $false; $LoginYaConfirmado = $false; $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false; $UltimoMotivoElegido = "" }

                    # Prioridad de interpretación (de mayor a menor)
                    if ($Obj.SysLog -match "User moved") {
                        $Interp = "$symStop DESCONECTADO: otra sesión tomó su extensión (User moved / Force logoff por el servidor)"; $ColorInterp = [System.Drawing.Color]::Red
                    }
                    elseif ($Obj.SysLog -match "PANTALLAZO AZUL|Apagado Sucio|Corte energía|Memoria Virtual Agotada|Tarjeta de Red desconectada" -or $Obj.AppLog -match "Caída de Túnel Principal") {
                        $Interp = "FALLA TÉCNICA (Justificado / Caída de Sistema o Red)"; $ColorInterp = [System.Drawing.Color]::LimeGreen
                    }
                    elseif ($Obj.AppLog -match "hilo interno abortado") { $Interp = "Aviso: Avaya reinició un hilo interno (la app NO se cerró)"; $ColorInterp = [System.Drawing.Color]::Orange }
                    elseif ($Obj.AppLog -match "System.Exception") { $Interp = "Falla grave en llamada"; $ColorInterp = [System.Drawing.Color]::Red }
                    elseif ($Obj.Audio -match "Dispositivo de Audio Desconectado") { $Interp = "¡ALERTA CRÍTICA! Diadema desconectada físicamente"; $ColorInterp = [System.Drawing.Color]::Red }
                    elseif ($Obj.Agente -match "CUELGUE MANUAL|LLAMADA COLGADA POR EL ASESOR") {
                        if ($Interp -match "Línea abierta sin marcar") { $Interp = "Precaución: El asesor abrió y cerró línea sin marcar"; $ColorInterp = [System.Drawing.Color]::Orange }
                        elseif ($Interp -notmatch "FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") { $Interp = "Llamada finalizada por el agente"; $ColorInterp = [System.Drawing.Color]::OrangeRed }
                    }
                    elseif ($Obj.Agente -match "UNHOLD MANUAL") {
                        if ($Interp -match "INICIO DE LLAMADA|LÍNEA ABIERTA") {
                            # UNHOLD coincide con nueva llamada — INICIO tiene prioridad visual
                            $HoldSesU = ""; if ($Obj.RawAgente -match "OnRequestUnholdSession.*?[Ss]ession\s+id=\s*(\d+)") { $HoldSesU = $matches[1] }
                            $Interp += if ($HoldSesU) { "  +  Retoma Ses:$HoldSesU" } else { "  +  Retoma llamada anterior" }
                        } elseif ($Interp -notmatch "FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") { $Interp = "Se retoma la llamada"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    }
                    elseif ($Obj.Agente -match "HOLD MANUAL") {
                        if ($Interp -match "INICIO DE LLAMADA|LÍNEA ABIERTA") {
                            # AutoHold al iniciar nueva llamada — INICIO tiene prioridad visual; anotar sesión en espera
                            $HoldSesH = ""; if ($Obj.RawAgente -match "OnRequestHoldSession.*?[Ss]ession\s+id=\s*(\d+)") { $HoldSesH = $matches[1] }
                            $Interp += if ($HoldSesH) { "  +  Hold Ses:$HoldSesH en espera" } else { "  +  Llamada anterior en Hold" }
                        } elseif ($Interp -notmatch "FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") { $Interp = "Llamada en Hold"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    }
                    elseif ($Obj.Agente -match "x MUTE MANUAL") { $Interp = "Mute activado"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    elseif ($Obj.Agente -match "o UNMUTE MANUAL") { $Interp = "Mute desactivado"; $ColorInterp = [System.Drawing.Color]::Yellow }
                    elseif ($Obj.AppLog -match "APP CRASH") { $Interp = "Aplicación congelada o cerrada inesperadamente"; $ColorInterp = [System.Drawing.Color]::Orange }
                    elseif ($Obj.Aux -match "SISTEMA_LOGOUT") {
                        # Pablo (09/09/2026, caso 08/09/2026): "Asesor tratando de desfirmarse" YA NO se arma
                        # aquí — se crea directamente en PASO 4 con su propia fila ms (ver
                        # $SlotDesfirmeSolicitud), porque este slot BASE (bare-hour) puede perder la carrera
                        # contra un "INICIO DE LLAMADA" que el EndpointLog crea en el mismo segundo para la
                        # llamada fantasma a la extensión 565 del propio mecanismo de logout. Esta rama solo
                        # limpia las banderas de login/motivo — el texto ya vive en su fila ms aparte.
                        $RecienFirmado = $false; $LoginYaConfirmado = $false; $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false; $UltimoMotivoElegido = ""
                    }
                    # Pablo (09/09/2026, Ronda 8): se QUITA el mecanismo viejo que vivía aquí
                    # ("Asesor solicita desfirmarse"/"Asesor desfirmado", vía EndpointLog + $Obj.Tel -match "^565$").
                    # Detectaba la MISMA llamada fantasma/manual a la extensión 565 que ya cubren, con más
                    # precisión y sin ambigüedad, los 3 mecanismos de PASO 4 (ver $SlotDesfirmeSolicitud /
                    # $CandidatoSlot565) — producía duplicados con etiqueta distinta para el mismo hecho, y
                    # podía quedar asimétrico (mostraba "desfirmado" sin haber mostrado nunca "solicita", caso
                    # confirmado en la Ronda 5). Unificado: ahora solo existe el vocabulario nuevo
                    # ("Asesor tratando de desfirmarse" / "Asesor ya se encuentra desfirmado").
                    elseif ($Obj.RawAux -match "Session_LoginAgent failed") {
                        $Interp = "Fallo en el intento de firmarse"; $ColorInterp = [System.Drawing.Color]::LightCoral; $LoginFallido = $true; $ValidandoLogin = $false
                        # Evidencia al clic en la celda Interpretación: la determinación viene de RawAux
                        # ("Session_LoginAgent failed"), pero el tooltip de Interpretación lee RawInterpretacion.
                        # Se copia la línea cruda del fallo para que el clic muestre el porqué.
                        $_lfFallo = @($Obj.RawAux -split "`n" | Where-Object { $_ -match "Session_LoginAgent failed" })
                        if ($_lfFallo.Count -gt 0 -and $Obj.RawInterpretacion -notmatch "Session_LoginAgent failed") { $Obj.RawInterpretacion += (($_lfFallo -join "`n") + "`n") }
                    }

                    # Lógica de login dinámico
                    elseif ($ValidandoLogin -eq $true) {
                        if ($Obj.RawAux -match "(?i)Enter\s+Aux" -or $Obj.Aux -match "ESTADO: AUX|ESTADO: PENDINGAUX|NUEVO_RC") {
                            $Interp = "Usuario firmado exitosamente"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $ValidandoLogin = $false; $RecienFirmado = $true; $LoginYaConfirmado = $true
                            $EsperandoDefaultPostLogin = $true; $EnDefaultPostLogin = $false   # el DEFAULT automático de Avaya llega enseguida
                            $CodigoRC = ""; if ($Obj.RawAux -match "(?i)ReasonCode[=\[>:\s]*(\d+)") { $CodigoRC = $matches[1] }
                            # Intentar también desde XML
                            if ($CodigoRC -eq "" -and $Obj.Aux -match "NUEVO_RC_NOMBRE:([^|]+)") { $CodigoRC = $matches[1].Trim() }
                            $NombreRC = if ($CodigoRC -ne "" -and $DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } elseif ($CodigoRC -ne "" -and $CodigoRC -ne "default" -and $CodigoRC -ne "Default") { $CodigoRC } else { "DEFAULT" }
                            $CurrentReasonCode = $CodigoRC
                            if ($CodigoRC -ne "" -and $CodigoRC -ne "0" -and $CodigoRC -notmatch "(?i)^default$") { $UltimoMotivoElegido = $CodigoRC }
                            if ($CodigoRC -eq "0" -or $CodigoRC -eq "" -or $CodigoRC -match "(?i)^default$") {
                                $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue
                            } else {
                                $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange
                            }
                            $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        }
                        elseif ($Obj.RawAux -match "(?i)AgentStateChanged.*newState=Ready" -or $Obj.Aux -match "ESTADO: READY|ESTADO: AUTOIN|ESTADO: MANUALIN") {
                            $Interp = "Usuario firmado exitosamente"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $ValidandoLogin = $false; $RecienFirmado = $true; $LoginYaConfirmado = $true
                            $EsperandoDefaultPostLogin = $true; $EnDefaultPostLogin = $false   # el DEFAULT automático de Avaya llega enseguida
                        }
                    }

                    # Estados de agente
                    # Guard: los eventos de llamada tienen prioridad sobre los cambios de estado del agente
                    # cuando coinciden en el mismo segundo (frecuente en XML donde Ready y llamada son simultáneos)
                    elseif ($Obj.Aux -match "GUI_READY_CONFIRMADO" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA|FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") {
                        $Interp = "Asesor se cambia a Disponible (Confirmado por clic)"; $ColorInterp = [System.Drawing.Color]::Yellow
                        $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false   # salió del DEFAULT post-logon
                    }
                    elseif ($Obj.RawAux -match "(?i)AgentStateChanged.*newState=Ready" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA|FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN") {
                        $Interp = "Asesor se cambia a Disponible"; $ColorInterp = [System.Drawing.Color]::Yellow
                        # ¿Fue con el botón favorito AUTO-IN? Ese botón dispara el FAC $FacAutoIn como saliente
                        # cortita coincidente (~mismo segundo) con el Ready; el botón normal usa CTI y no la genera.
                        $msDisp = & $MsDeSlot $H
                        if ($msDisp -ge 0) {
                            foreach ($sFac in $AutoInFacSlots) {
                                if ([math]::Abs((& $MsDeSlot $sFac) - $msDisp) -le 3000) { $Interp = "Asesor se cambia a Disponible usando botón favorito AUTO-IN"; break }
                            }
                        }
                        $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false   # salió del DEFAULT post-logon
                    }
                    # AUX puesto con BOTÓN FAVORITO ("Trab Aux"): genera "AgentStateChanged newState=Aux" pero NO
                    # "Enter Aux" ni la petición del endpoint. Debe ir ANTES de la rama "ESTADO: READY": el mismo
                    # slot trae "oldState=Ready" que se capturó como "ESTADO: READY" y enmascaraba el auxiliar,
                    # dejándolo invisible. Aquí se prioriza el newState=Aux real. (Pablo, prueba controlada 28/07.)
                    # GUARD (Pablo, 09/09/2026, caso 02/09/2026): al desfirmarse, Avaya parpadea internamente
                    # "...;newState=Aux" seguido, en el MISMO segundo, de "...;newState=LoggedOut" (limpieza de
                    # sesión, no un clic real de TrabAux) — sin distinguirlo, esta rama mostraba un falso
                    # "Auxiliar - se detectó un auxiliar sin código" justo cuando el asesor en realidad terminó
                    # de desfirmarse. Si la ÚLTIMA transición de este segundo compartido es "newState=LoggedOut"
                    # (no "newState=Aux"), el Aux fue solo un parpadeo transitorio — no se reporta como tal.
                    elseif ($Obj.RawAux -match "(?i)AgentStateChanged.*newState=Aux" -and $Obj.RawAux -notmatch "(?i)Enter\s+Aux" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA|FIN DE LLAMADA|CUELGUE MANUAL|EVASIÓN" -and ([regex]::Matches($Obj.RawAux, "(?i)newState\s*=\s*(Aux|LoggedOut)") | Select-Object -Last 1).Groups[1].Value -eq "Aux") {
                        $CodigoRC = ""; if ($Obj.RawAux -match "(?i)ReasonCode[=\[>:\s]*(\d+)") { $CodigoRC = $matches[1] }
                        $NombreRC = if ($CodigoRC -ne "" -and $CodigoRC -ne "0" -and $DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } elseif ($CodigoRC -ne "" -and $CodigoRC -ne "0") { $CodigoRC } else { "" }
                        if ($EsperandoDefaultPostLogin -or $EnDefaultPostLogin) {
                            # Punto 1 (Pablo, 28/08/2026): el DEFAULT automático que Avaya dispara justo al firmarse
                            # TAMBIÉN entra por "newState=Aux" sin "Enter Aux" (el "Enter Aux;code=ReasonCode[0]"
                            # cae en un segundo ANTERIOR, fuera de RawAux de esta fila) — sin este guard se
                            # mostraba como si el asesor hubiera dado clic en TrabAux justo al firmarse. No es un
                            # clic del asesor, es el sistema. Mismo tratamiento que ya usan las ramas NUEVO_RC/Enter Aux.
                            if ($EsperandoDefaultPostLogin) { $Interp = "Asesor firmado y en Default"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $true }
                            else { $Interp = "" }
                            $CurrentReasonCode = "0"; $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue
                        }
                        elseif ($NombreRC -ne "") {
                            # Opción C: dejar explícito que este Auxiliar vino del botón favorito (TrabAux), espejo
                            # de "usando botón favorito AUTO-IN" que ya existe del lado de Disponible.
                            $CurrentReasonCode = $CodigoRC; $UltimoMotivoElegido = $CodigoRC
                            $Interp = "Asesor se cambia a Auxiliar [$NombreRC] (vía botón favorito)"; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange
                            $ColorInterp = [System.Drawing.Color]::Orange
                            $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false
                        }
                        else {
                            # Puntos 2 y 3 (Pablo): TrabAux SIN elegir motivo. En vez de un "Auxiliar" genérico y
                            # vacío, se usa $UltimoMotivoElegido — que NO se limpia al pasar por Disponible, solo
                            # cuando se elige un motivo distinto — para precisar en qué auxiliar quedó el asesor,
                            # replicando lo que la GUI de Avaya realmente muestra en este caso (confirmado por
                            # Pablo: mostraba "Aux Sistemas", el último que había usado).
                            # Ronda siguiente (Pablo): probamos "Detectado [X]" con la inferencia de
                            # $UltimoMotivoElegido, pero un caso real (log 18/08/2026) mostró que Avaya CMS
                            # aplicó un motivo DISTINTO al último que el asesor había elegido — nuestra
                            # inferencia puede estar sencillamente equivocada (la causa real vive en la
                            # configuración del conmutador, invisible en este log). Se quita la adivinanza:
                            # solo se reporta el HECHO (auxiliar sin código elegido), sin aventurar cuál fue.
                            $Interp = "Asesor se cambia a Auxiliar - se detectó un auxiliar sin código"
                            $CurrentAux = "$symUser Estado: AUXILIAR"; $CurrentColor = [System.Drawing.Color]::LightCoral
                            $ColorInterp = [System.Drawing.Color]::Orange
                            $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false
                        }
                        $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                    }
                    elseif ($Obj.Aux -match "ESTADO: READY" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        # ESTADO: READY del XML = notificación de Avaya, no es clic del asesor.
                        # Solo actualizamos la columna Aux en silencio. Sin interpretación visible.
                        # El único evento confiable de "cambió a disponible" es GUI_READY_CONFIRMADO.
                        $CurrentAux = "$symUser Estado: DISPONIBLE"; $CurrentColor = [System.Drawing.Color]::LimeGreen
                        $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false   # salió del DEFAULT post-logon
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
                        if ($NombreRC -ne "DEFAULT") { $UltimoMotivoElegido = $NombreRC }
                        if ($Interp -match "FIN DE LLAMADA|CUELGUE MANUAL") {
                            # FIN tiene prioridad visual — agregar nota compacta sin reemplazar
                            if ($NombreRC -ne "DEFAULT") { $Interp += "  +  Auxiliar [$NombreRC]" }
                        } else {
                            $Interp = "Asesor se cambia a Auxiliar [$NombreRC]"; $ColorInterp = [System.Drawing.Color]::Orange
                            # DEFAULT automático que Avaya dispara al firmarse (2 veces): 1º → etiqueta clara; repetidos → ocultos.
                            if ($NombreRC -eq "DEFAULT") {
                                if ($EsperandoDefaultPostLogin)  { $Interp = "Asesor firmado y en Default"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $true }
                                elseif ($EnDefaultPostLogin)     { $Interp = "" }
                            } else { $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false }
                        }
                    }
                    elseif ($Obj.Aux -match "NUEVO_RC:(\d+)" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        # AUX desde OneXAgent (numérico)
                        $CodigoRC = $matches[1]
                        $NombreRC = if ($DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } else { $CodigoRC }
                        if ($NombreRC -eq "0" -or $NombreRC -eq "") { $NombreRC = "DEFAULT" }
                        $CurrentReasonCode = $CodigoRC
                        if ($CodigoRC -ne "0") { $UltimoMotivoElegido = $CodigoRC }
                        # Siempre actualizar estado del asesor (independiente de si hay FIN simultáneo)
                        $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                        if ($Interp -match "FIN DE LLAMADA|CUELGUE MANUAL") {
                            # FIN tiene prioridad visual — agregar nota compacta sin reemplazar
                            if ($CodigoRC -ne "0") { $Interp += "  +  Auxiliar [$NombreRC]" }
                            else { $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        } else {
                            $Interp = "Asesor se cambia a Auxiliar [$NombreRC] (Confirmado por clic)"; $ColorInterp = [System.Drawing.Color]::Orange
                            if ($CodigoRC -eq "0") {
                                if ($EsperandoDefaultPostLogin)  { $Interp = "Asesor firmado y en Default"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $true }
                                elseif ($EnDefaultPostLogin)     { $Interp = "" }
                                $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                            } else { $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false }
                        }
                    }
                    elseif ($Obj.RawAux -match "(?i)Enter\s+Aux" -and $Interp -notmatch "señal de llamada|INICIO DE LLAMADA|LÍNEA ABIERTA") {
                        $CodigoRC = ""; if ($Obj.RawAux -match "(?i)ReasonCode[=\[>:\s]*(\d+)") { $CodigoRC = $matches[1] }
                        $NombreRC = if ($CodigoRC -ne "" -and $DictRC.ContainsKey($CodigoRC)) { $DictRC[$CodigoRC] } elseif ($CodigoRC -ne "") { $CodigoRC } else { "" }
                        # Siempre actualizar estado del asesor (independiente de si hay FIN simultáneo)
                        if ($NombreRC -ne "") { $CurrentReasonCode = $CodigoRC; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor; if ($CodigoRC -ne "0") { $UltimoMotivoElegido = $CodigoRC } }
                        else                  { $CurrentAux = "$symUser Estado: AUXILIAR"; $CurrentColor = [System.Drawing.Color]::LightCoral; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        if ($CodigoRC -eq "0") { $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor }
                        if ($Interp -match "FIN DE LLAMADA|CUELGUE MANUAL") {
                            # FIN tiene prioridad visual — agregar nota compacta sin reemplazar
                            if ($CodigoRC -ne "0") { $Interp += if ($NombreRC -ne "") { "  +  Auxiliar [$NombreRC]" } else { "  +  Auxiliar" } }
                        } else {
                            if ($NombreRC -ne "") { $Interp = "Asesor se cambia a Auxiliar [$NombreRC]"; $ColorInterp = [System.Drawing.Color]::Orange }
                            else                  { $Interp = "Asesor se cambia a Auxiliar";             $ColorInterp = [System.Drawing.Color]::Orange }
                            if ($CodigoRC -eq "0") {
                                if ($EsperandoDefaultPostLogin)  { $Interp = "Asesor firmado y en Default"; $ColorInterp = [System.Drawing.Color]::LimeGreen; $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $true }
                                elseif ($EnDefaultPostLogin)     { $Interp = "" }
                                $CurrentAux = "$symUser Estado: DEFAULT"; $CurrentColor = [System.Drawing.Color]::CadetBlue; $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                            } else { $EsperandoDefaultPostLogin = $false; $EnDefaultPostLogin = $false }
                        }
                    }

                    # Colisión: Disponible y señal de llamada en el mismo segundo.
                    # Ocurre cuando Avaya enruta una llamada en el instante exacto en que el agente
                    # sale de Auxiliar (sin EnterReadyHandler, via auto-in o cambio automático del ACD).
                    # La cadena elseif anterior no pudo mostrar "Disponible" porque "señal de llamada"
                    # ya estaba en $Interp. Lo anotamos en el badge Aux para visibilidad independiente
                    # del ancho de columna, y con un sufijo corto en la interpretación.
                    if ($Interp -match "señal de llamada" -and
                        ($Obj.Aux -match "GUI_READY_CONFIRMADO|AUTOIN_DISP_SLOT" -or $Obj.RawAux -match "(?i)AgentStateChanged.*newState=Ready")) {
                        $HaySlotMs = ($Obj.Aux -match "GUI_READY_CONFIRMADO|AUTOIN_DISP_SLOT")
                        # Badge adicional en columna Aux para que sea visible siempre
                        $Obj.Aux = "↑ Cambió a Disponible  |  " + $Obj.Aux
                        $Obj.ColorAux = [System.Drawing.Color]::Yellow
                        # Sufijo inline solo si NO hay slot ms dedicado (fallback cuando auto-in/GUI no generó fila propia)
                        if (-not $HaySlotMs) { $Interp += "  ↑ Aux→Disponible" }
                    }

                    # (B) Dedup de tokens de estado consecutivos idénticos: durante el login Avaya emite
                    # varias notificaciones "Agent State Ready" en el mismo segundo (confirmado: 4 a las 08:07:01)
                    # → "ESTADO: READY|" se acumulaba repetido. Se colapsan los consecutivos iguales,
                    # preservando transiciones reales (READY→AUX→READY no se colapsa).
                    if ($Obj.Aux -match '\|') {
                        $toksAux = $Obj.Aux -split '\|'; $dedupAux = @(); $prevAux = $null
                        foreach ($tkAux in $toksAux) {
                            $tkT = $tkAux.Trim(); if ($tkT -eq '') { continue }
                            if ($tkT -ne $prevAux) { $dedupAux += $tkT }
                            $prevAux = $tkT
                        }
                        $Obj.Aux = if ($dedupAux.Count -gt 0) { ($dedupAux -join '|') + '|' } else { '' }
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
                            # Pablo (09/09/2026): "EN PROCESO DE DESFIRMARSE" es impreciso cuando ya se SABE
                            # (más abajo en el mismo log, mismo slot) que Avaya rechazó el intento — el token
                            # DESFIRME_FALLIDO llega DESPUÉS de SISTEMA_LOGOUT en este mismo .Aux, así que
                            # sobrescribe el badge con el desenlace real en vez de dejarlo en "en proceso".
                            elseif ($Item -match "DESFIRME_FALLIDO") { $CurrentAux = "$symUser Estado: DESFIRME FALLIDO — SIGUE ACTIVO"; $CurrentColor = [System.Drawing.Color]::OrangeRed }
                            elseif ($Item -match "ESTADO: PENDINGAUX") { $CurrentAux = "$symUser Estado: PENDIENTE AUXILIAR"; $CurrentColor = [System.Drawing.Color]::MediumOrchid }
                            elseif ($Item -match "ESTADO: READY|AUTOIN|MANUALIN|GUI_READY_CONFIRMADO") { $CurrentReasonCode = ""; $CurrentAux = "$symUser Estado: DISPONIBLE"; $CurrentColor = [System.Drawing.Color]::LimeGreen }
                            elseif ($Item -match "ESTADO: LOGGEDOUT") { $CurrentReasonCode = ""; $CurrentAux = "$symUser Estado: USUARIO DESFIRMADO"; $CurrentColor = [System.Drawing.Color]::Gray }
                            elseif ($Item -match "ESTADO: AUX|ESTADO: NOTREADY") { if ($CurrentReasonCode -ne "" -and $CurrentReasonCode -ne "0") { $NombreRC = if ($DictRC.ContainsKey($CurrentReasonCode)) { $DictRC[$CurrentReasonCode] } else { "RC: $CurrentReasonCode" }; $CurrentAux = "$symUser Estado: AUXILIAR ($NombreRC)"; $CurrentColor = [System.Drawing.Color]::Orange } else { $CurrentAux = "$symUser Estado: AUXILIAR"; $CurrentColor = [System.Drawing.Color]::LightCoral } }
                        }
                    }

                    if (-not ($ValidandoLogin -eq $false -and $Interp -match "Usuario firmado exitosamente")) {
                        $Obj.Aux = $CurrentAux; $Obj.ColorAux = $CurrentColor
                    } else {
                        # (A) Fila de confirmación de login: no arrastrar el estado crudo acumulado
                        # ("ESTADO: READY|..."). El estado DISPONIBLE ya se muestra limpio en las filas siguientes.
                        $Obj.Aux = ""; $Obj.ColorAux = [System.Drawing.Color]::White
                    }

                    if ($Obj.RawIspeac -ne "" -and $Interp -eq "" -and $LlamadaActiva) {
                        $Interp = "Llamada en curso (Analizando Calidad)"; $ColorInterp = [System.Drawing.Color]::MediumSpringGreen; $Obj.Sesion = $SesionActual
                    }

                    # Signal B: resaltar INICIO DE LLAMADA que fue transferida automáticamente por el sistema
                    if ($Interp -match "INICIO DE LLAMADA|LÍNEA ABIERTA" -and $Obj.Aux -match "TRANSF_AUTO") {
                        $Interp += "  ⚠ Sistema transfirió automáticamente (bridge detectado)"
                        $ColorInterp = [System.Drawing.Color]::Orange
                    }

                    # Marca de auditoría: la llamada ENTRANTE se recibió con el asesor en AUXILIAR
                    # (p.ej. transferencia directa que omite el bloqueo de la cola ACD). $CurrentAux es
                    # el estado vigente del asesor en este punto del timeline (el mismo de la col. de estado).
                    # Para no amontonar: marcador corto en la actividad + el detalle queda en el badge de estado.
                    if ($Interp -match "INICIO DE LLAMADA \(Entrante" -and $CurrentAux -match "AUXILIAR") {
                        $Interp += "  ⚠ recibida en AUXILIAR"
                        $Obj.Aux = "⚠ RECIBIÓ LLAMADA EN  " + $Obj.Aux
                    }

                    $Obj.Interpretacion = $Interp; $Obj.ColorInterpretacion = $ColorInterp

                    if ($RecienFirmado -eq $true -and $Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "" -and $Obj.Ispeac -eq "" -and $Obj.Dtmf -eq "" -and $Obj.Interpretacion -eq "") { if ($Obj.Aux -match "Estado: AUXILIAR") { $RecienFirmado = $false; continue } }
                    if ($Obj.Interpretacion -eq "" -and $Obj.Agente -eq "" -and $Obj.Audio -eq "" -and $Obj.SysLog -eq "" -and $Obj.AppLog -eq "" -and $Obj.Ispeac -eq "" -and $Obj.Dtmf -eq "") { if ($Obj.Aux -match "Estado:") { continue } }
                    if ($Interp -match "Usuario firmado exitosamente") { $RecienFirmado = $true } elseif ($Obj.Interpretacion -ne "" -or $Obj.Agente -ne "" -or $Obj.Audio -ne "" -or $Obj.Ispeac -ne "") { $RecienFirmado = $false }

                    $Row = $GridResultados.Rows.Add()
                    # Mostrar ms reales en la columna Hora para "Asesor se cambia a Auxiliar" y para
                    # "Asesor firmado y en Default": su slot es base por diseño del estado del agente
                    # (mover el slot fragmentaría el estado), pero el ms real vive en el log crudo (RawAux)
                    # → se extrae solo para mostrarlo, sin tocar el orden.
                    # Se usa $Interp (etiqueta final del render), no $Obj.Interpretacion: las filas de estado
                    # del agente construyen su etiqueta aquí y dejan Interpretacion vacía en el slot.
                    # Los eventos de login/extensión también nacen en slot base (los escribe el Audio/Endpoint
                    # log sobre $HoraLimpia); su ms real vive en el log crudo que sí se guardó:
                    #   · estado del agente + "Usuario firmado exitosamente" → RawAux
                    #   · "Extensión en línea…" y "Usuario intentando firmarse…" → RawInterpretacion
                    # El regex acepta ':' (Endpoint/Audio: 16:10:06:409) y ',' (OneXAgent: 16:10:32,123).
                    $HoraCell = $H
                    if ($H -notmatch ',' -and $Interp -match "se cambia a Auxiliar|se cambia a Disponible|Asesor firmado y en Default|Asesor tratando de desfirmarse|Intento de desfirme|Usuario firmado exitosamente|Extensión en línea|Usuario intentando firmarse|Fallo en el intento de firmarse") {
                        $_reMs = [regex]::Escape($H) + "[.,:](\d{1,3})"
                        if     ($Obj.RawAux            -match $_reMs) { $HoraCell = "$H," + ($matches[1].PadRight(3,'0')) }
                        elseif ($Obj.RawInterpretacion -match $_reMs) { $HoraCell = "$H," + ($matches[1].PadRight(3,'0')) }
                    }
                    $GridResultados.Rows[$Row].Cells["Hora"].Value          = $HoraCell
                    $GridResultados.Rows[$Row].Cells["Sesion"].Value        = $Obj.Sesion
                    if ($Obj.Sesion -ne "-" -and $Obj.Sesion -ne "") { $GridResultados.Rows[$Row].Cells["Sesion"].Style.ForeColor = [System.Drawing.Color]::Cyan; $GridResultados.Rows[$Row].Cells["Sesion"].ToolTipText = "Clic para aislar y ver el historial completo de esta llamada" }
                    $GridResultados.Rows[$Row].Cells["Telefono"].Value      = $Obj.Tel
                    if ($Obj.Sesion -ne "-" -and $Obj.Sesion -ne "" -and $RawMapeoTel.ContainsKey($Obj.Sesion)) { $GridResultados.Rows[$Row].Cells["Telefono"].ToolTipText = "Clic para ver Log Original de extracción del número"; $GridResultados.Rows[$Row].Cells["Telefono"].Tag = $RawMapeoTel[$Obj.Sesion].Trim() }
                    $GridResultados.Rows[$Row].Cells["EvDtmf"].Value        = $Obj.Dtmf;    $GridResultados.Rows[$Row].Cells["EvDtmf"].Style.ForeColor    = $Obj.ColorDtmf
                    if ($Obj.RawDtmf -ne "")           { $GridResultados.Rows[$Row].Cells["EvDtmf"].ToolTipText = "Clic para ver el detalle por dígito (con milisegundos)"; $GridResultados.Rows[$Row].Cells["EvDtmf"].Tag = $Obj.RawDtmf.Trim() }
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
    # Botón de revisión de errores: solo visible si el análisis juntó posibles errores del código base.
    $nErr = @($Script:ErroresSospechosos).Count
    if ($nErr -gt 0) { $btnRevisarErrores.Text = "⚠ Revisar posibles errores ($nErr)"; $btnRevisarErrores.Visible = $true } else { $btnRevisarErrores.Visible = $false }

    # ── Buscador de VACÍOS del proceso (día completo) ─────────────────────────────────────────
    # Combina TODAS las líneas con timestamp del día (recolectadas durante PASO 3/4/5, sin releer
    # archivos) y busca huecos ≥ $UmbralVacioMs donde NO hubo ni una sola línea — ni siquiera ruido
    # interno — en los dos logs que maneja el propio proceso OneXAgent.exe (Endpoint+AvayaOneX).
    # Cruza cada hueco contra IspeacLog (proceso INDEPENDIENTE): si Ispeac siguió reportando durante
    # el vacío, había una llamada activa y el hueco es mucho más grave (el proceso principal se
    # congeló con audio en vivo); si Ispeac también calló, es menos concluyente. (Pablo, 08/2026 —
    # caso "llamada tardó 7s en asignarse pese a Auto Accept", validado con silencio real de 6.8s.)
    $UmbralVacioMs = 3000
    $Script:VaciosDetectados = @()
    if ($Script:TsProceso.Count -ge 2) {
        $tsOrdenado = @($Script:TsProceso | Sort-Object)
        $tsIspeacOrdenado = @($Script:TsIspeac | Sort-Object)
        $FmtMsDia = { param($m) $hh=[int]($m/3600000); $mm=[int](($m%3600000)/60000); $ss=[int](($m%60000)/1000); $fff=$m%1000; "{0:D2}:{1:D2}:{2:D2},{3:D3}" -f $hh,$mm,$ss,$fff }
        for ($iV = 1; $iV -lt $tsOrdenado.Count; $iV++) {
            $prevMs = $tsOrdenado[$iV - 1]; $curMs = $tsOrdenado[$iV]
            $gapMs = $curMs - $prevMs
            if ($gapMs -ge $UmbralVacioMs) {
                $ispeacVivo = $false
                foreach ($tI in $tsIspeacOrdenado) { if ($tI -gt $prevMs -and $tI -lt $curMs) { $ispeacVivo = $true; break } }
                $Script:VaciosDetectados += [pscustomobject]@{
                    HoraIni    = & $FmtMsDia $prevMs
                    HoraFin    = & $FmtMsDia $curMs
                    DurSeg     = [math]::Round($gapMs / 1000, 1)
                    IspeacVivo = $ispeacVivo
                }
            }
        }
    }
    $nVac = @($Script:VaciosDetectados).Count
    if ($nVac -gt 0) { $btnVacios.Text = "🔍 Vacíos en logs ($nVac)"; $btnVacios.Visible = $true } else { $btnVacios.Visible = $false }

    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# --- Ventana aparte: posibles errores del código base (fuera del timeline) ---
$btnRevisarErrores.Add_Click({
    $errores = @($Script:ErroresSospechosos)
    if ($errores.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No hay posibles errores registrados en este análisis.","Revisar errores",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information); return }
    $FormErr = New-Object System.Windows.Forms.Form; $FormErr.Text = "Posibles errores del código base ($($errores.Count)) — NO están en el timeline"; $FormErr.Size = New-Object System.Drawing.Size(1200, 640); $FormErr.StartPosition = "CenterParent"; $FormErr.BackColor = $ColorFondo
    $lblNota = New-Object System.Windows.Forms.Label; $lblNota.Text = "Líneas con ERROR/Exception/FATAL detectadas FUERA del timeline. La mayoría son benignas (renderizado, features, timeouts de red). Revisa la línea cruda para confirmar."; $lblNota.Location = New-Object System.Drawing.Point(15, 12); $lblNota.Size = New-Object System.Drawing.Size(1160, 20); $lblNota.ForeColor = [System.Drawing.Color]::Khaki
    $lblFil = New-Object System.Windows.Forms.Label; $lblFil.Text = "Filtrar:"; $lblFil.Location = New-Object System.Drawing.Point(15, 42); $lblFil.AutoSize = $true; $lblFil.ForeColor = [System.Drawing.Color]::White
    $txtFil = New-Object System.Windows.Forms.TextBox; $txtFil.Location = New-Object System.Drawing.Point(70, 39); $txtFil.Size = New-Object System.Drawing.Size(280, 25); $txtFil.BackColor = [System.Drawing.Color]::FromArgb(45,45,48); $txtFil.ForeColor = [System.Drawing.Color]::White
    $GridErr = New-Object System.Windows.Forms.DataGridView; $GridErr.Size = New-Object System.Drawing.Size(1160, 520); $GridErr.Location = New-Object System.Drawing.Point(15, 72)
    $GridErr.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $GridErr.BackgroundColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridErr.AllowUserToAddRows = $false; $GridErr.RowHeadersVisible = $false; $GridErr.ReadOnly = $true; $GridErr.AutoSizeColumnsMode = "Fill"
    $GridErr.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridErr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gainsboro
    $GridErr.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(10,10,10); $GridErr.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White; $GridErr.EnableHeadersVisualStyles = $false
    $GridErr.Columns.Add("Hora","Hora") | Out-Null; $GridErr.Columns["Hora"].FillWeight = 8
    $GridErr.Columns.Add("Linea","Línea cruda del log") | Out-Null; $GridErr.Columns["Linea"].FillWeight = 92
    $GridErr.SuspendLayout()
    foreach ($it in ($errores | Sort-Object Hora)) {
        $r = $GridErr.Rows.Add(); $GridErr.Rows[$r].Cells["Hora"].Value = $it.Hora; $GridErr.Rows[$r].Cells["Hora"].Style.ForeColor = [System.Drawing.Color]::White; $GridErr.Rows[$r].Cells["Linea"].Value = $it.Linea
    }
    $GridErr.ResumeLayout()
    $txtFil.Add_TextChanged({ $t = $txtFil.Text.Trim(); $GridErr.SuspendLayout(); $GridErr.CurrentCell = $null; foreach ($row in $GridErr.Rows) { if ($t -eq "") { $row.Visible = $true } else { $row.Visible = ("$($row.Cells['Linea'].Value)" -match [regex]::Escape($t)) } }; $GridErr.ResumeLayout() })
    $btnCerrarErr = New-Object System.Windows.Forms.Button; $btnCerrarErr.Text = "Cerrar"; $btnCerrarErr.Location = New-Object System.Drawing.Point(1065, 37); $btnCerrarErr.Size = New-Object System.Drawing.Size(110, 28); $btnCerrarErr.BackColor = [System.Drawing.Color]::Gray; $btnCerrarErr.ForeColor = [System.Drawing.Color]::White; $btnCerrarErr.FlatStyle = "Flat"; $btnCerrarErr.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCerrarErr.Add_Click({ $FormErr.Close() })
    $FormErr.Controls.AddRange(@($lblNota, $lblFil, $txtFil, $GridErr, $btnCerrarErr))
    $FormErr.ShowDialog() | Out-Null
})

# --- Ventana aparte: vacíos de log del día (silencio total en Endpoint+AvayaOneX) ---
$btnVacios.Add_Click({
    $vacios = @($Script:VaciosDetectados)
    if ($vacios.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No se detectaron vacíos de log en este análisis.","Vacíos en logs",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information); return }
    $FormVac = New-Object System.Windows.Forms.Form; $FormVac.Text = "Vacíos de log detectados ($($vacios.Count)) — silencio total en Endpoint+AvayaOneX (día completo)"; $FormVac.Size = New-Object System.Drawing.Size(1000, 620); $FormVac.StartPosition = "CenterParent"; $FormVac.BackColor = $ColorFondo
    $lblNotaV = New-Object System.Windows.Forms.Label; $lblNotaV.Text = "Umbral: ≥3s sin NINGUNA línea en Endpoint.log ni AvayaOneX.log (los 2 logs del propio proceso OneXAgent.exe). `"¿Llamada activa?`" = SÍ significa que IspeacLog (proceso independiente) siguió reportando durante el vacío — el proceso principal se congeló con audio EN VIVO, mucho más grave. Un NO no descarta el problema, solo que no se pudo corroborar con Ispeac."; $lblNotaV.Location = New-Object System.Drawing.Point(15, 12); $lblNotaV.Size = New-Object System.Drawing.Size(960, 48); $lblNotaV.ForeColor = [System.Drawing.Color]::Khaki
    $GridVac = New-Object System.Windows.Forms.DataGridView; $GridVac.Size = New-Object System.Drawing.Size(960, 480); $GridVac.Location = New-Object System.Drawing.Point(15, 68)
    $GridVac.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $GridVac.BackgroundColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridVac.AllowUserToAddRows = $false; $GridVac.RowHeadersVisible = $false; $GridVac.ReadOnly = $true; $GridVac.AutoSizeColumnsMode = "Fill"
    $GridVac.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridVac.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gainsboro
    $GridVac.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(10,10,10); $GridVac.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White; $GridVac.EnableHeadersVisualStyles = $false
    $GridVac.Columns.Add("Inicio","Inicio del vacío") | Out-Null; $GridVac.Columns["Inicio"].FillWeight = 22
    $GridVac.Columns.Add("Fin","Reanuda en") | Out-Null; $GridVac.Columns["Fin"].FillWeight = 22
    $GridVac.Columns.Add("Dur","Duración") | Out-Null; $GridVac.Columns["Dur"].FillWeight = 16
    $GridVac.Columns.Add("Ispeac","¿Llamada activa? (Ispeac vivo)") | Out-Null; $GridVac.Columns["Ispeac"].FillWeight = 40
    $GridVac.SuspendLayout()
    foreach ($v in ($vacios | Sort-Object HoraIni)) {
        $r = $GridVac.Rows.Add()
        $GridVac.Rows[$r].Cells["Inicio"].Value = $v.HoraIni
        $GridVac.Rows[$r].Cells["Fin"].Value    = $v.HoraFin
        $GridVac.Rows[$r].Cells["Dur"].Value    = "$($v.DurSeg) s"
        if ($v.IspeacVivo) {
            $GridVac.Rows[$r].Cells["Ispeac"].Value = "SÍ — congelamiento con llamada activa"
            $GridVac.Rows[$r].DefaultCellStyle.ForeColor = [System.Drawing.Color]::OrangeRed
        } else {
            $GridVac.Rows[$r].Cells["Ispeac"].Value = "No (sin corroborar con Ispeac)"
        }
    }
    $GridVac.ResumeLayout()
    $btnCerrarVac = New-Object System.Windows.Forms.Button; $btnCerrarVac.Text = "Cerrar"; $btnCerrarVac.Location = New-Object System.Drawing.Point(865, 558); $btnCerrarVac.Size = New-Object System.Drawing.Size(110, 28); $btnCerrarVac.BackColor = [System.Drawing.Color]::Gray; $btnCerrarVac.ForeColor = [System.Drawing.Color]::White; $btnCerrarVac.FlatStyle = "Flat"; $btnCerrarVac.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCerrarVac.Add_Click({ $FormVac.Close() })
    $FormVac.Controls.AddRange(@($lblNotaV, $GridVac, $btnCerrarVac))
    $FormVac.ShowDialog() | Out-Null
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
    # NOTA: @(...) es necesario porque PS 5.1 no envuelve en array cuando Sort-Object devuelve
    # un solo elemento — sin @(), $TodasHoras sería un string y Where-Object iteraría sus caracteres.
    $TodasHoras = @($Snap.Keys | Sort-Object)
    $HorasPost  = @($TodasHoras | Where-Object { [string]$_ -gt [string]$InicioHora })

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
        # "LÍNEA ABIERTA SIN MARCAR" es un INICIO saliente, NO un fin — excluir.
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
    # NOTA: $FinDesc puede mostrar texto de display (ej. "Llamada finalizada por el agente")
    # en lugar del texto del slot original (ej. "FIN DE LLAMADA MANUAL...").
    # Por eso también revisamos el RawAgente y el texto original del slot para no perder
    # la clasificación MANUAL cuando el mensaje de display fue editado.
    $SlotInterpRaw = if ($null -ne $FinSlot) { $FinSlot.Interpretacion } else { "" }
    $EsFallback = ($RawInterp -match "\[Fallback\]")
    $EsManual   = ($FinDesc -match "MANUAL|CUELGUE|finalizada por el agente") -or
                  ($SlotInterpRaw -match "MANUAL|CUELGUE") -or
                  ($RawAgen -match "CUELGUE MANUAL|LLAMADA COLGADA POR EL ASESOR")
    $EsEvasion  = ($FinDesc -match "EVASIÓN|abandonada") -or ($SlotInterpRaw -match "EVASIÓN")

    # ── Encabezado ───────────────────────────────────────────────────────
    & $EL "DIAGNOSTICO: ¿POR QUE FINALIZO ESTA LLAMADA?" $cW $true
    & $EL "" $cW
    $TipoLabel  = if ($EsManual) { "MANUAL (agente finalizó la llamada)" } elseif ($EsEvasion) { "EVASIÓN / Línea abandonada" } else { "NORMAL (cliente u otro extremo colgó)" }
    $colorTipo  = if ($EsManual) { $cMan } elseif ($EsEvasion) { $cFail } else { $cNrm }
    & $EL "  Clasificacion   : $TipoLabel" $colorTipo $true
    & $EL "  Texto en tabla  : $FinDesc" $cInfo
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

    # NOTA: @(...) es necesario porque PS 5.1 no envuelve en array cuando Sort-Object devuelve
    # un solo elemento — sin @(), $TodasHoras sería un string y Where-Object iteraría sus caracteres.
    $TodasHoras  = @($Snap.Keys | Sort-Object)
    $HorasAnte   = @($TodasHoras | Where-Object { [string]$_ -lt [string]$FinHora })
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
# DETALLE DE CONFERENCIA: CÓMO SE ESTABLECIÓ
# ====================================================================
$mnuDetalleConf.Add_Click({
    $SelRow = if ($GridResultados.SelectedRows.Count -gt 0) { $GridResultados.SelectedRows[0] } else { $null }
    if ($null -eq $SelRow) { return }

    $HoraGrid  = $SelRow.Cells["Hora"].Value
    $RawConf   = $SelRow.Cells["EvAgente"].Tag

    # ── Extraer hora exacta con milisegundos del log crudo ──────────
    $HoraExacta = "$HoraGrid (ms no disponibles)"
    if ($RawConf -match "\[(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2}:\d{3})\].*Conference_CompleteConf") {
        $HoraExacta = "$($matches[2])  (milisegundos del PBX)"
    }

    # ── Extraer sesiones de la línea Conference_CompleteConf ────────
    $SesFirst = ""; $SesConsult = ""
    if ($RawConf -match "Conference_CompleteConf: nFirstCall:\s*(\d+),\s*nConsultCall:\s*(\d+)") {
        $SesFirst = $matches[1]; $SesConsult = $matches[2]
    }

    # ── Buscar teléfonos en la grilla por sesión ────────────────────
    $TelFirst = ""; $TelConsult = ""
    foreach ($row in $GridResultados.Rows) {
        if ($row.IsNewRow) { continue }
        $sv = $row.Cells["Sesion"].Value; $tv = $row.Cells["Telefono"].Value
        if ($SesFirst   -ne "" -and $sv -eq $SesFirst   -and $tv -and $tv -ne "-") { $TelFirst   = $tv }
        if ($SesConsult -ne "" -and $sv -eq $SesConsult -and $tv -and $tv -ne "-") { $TelConsult = $tv }
    }

    # ── Fallback para nFirstCall: puede ser un bridge de conferencia ─
    # El bridge (ej: sesión 3) no tiene teléfono propio en la grilla.
    # Lo que sí existe es el registro "conference Id=3, call id=2" en
    # el raw de la fila de CONFERENCIA INICIADA, que mapea el bridge
    # a la sesión original (sesión 2 = +5539991927).
    if ($TelFirst -eq "" -and $SesFirst -ne "") {
        $OriginalCallId = ""
        foreach ($row in $GridResultados.Rows) {
            if ($row.IsNewRow) { continue }
            $rawA = $row.Cells["EvAgente"].Tag
            if ($rawA) {
                $m = [regex]::Match($rawA, "conference Id=$SesFirst, call id=(\d+)")
                if ($m.Success -and [int]$m.Groups[1].Value -gt 0) {
                    $OriginalCallId = $m.Groups[1].Value; break
                }
            }
        }
        if ($OriginalCallId -ne "") {
            foreach ($row in $GridResultados.Rows) {
                if ($row.IsNewRow) { continue }
                $sv = $row.Cells["Sesion"].Value; $tv = $row.Cells["Telefono"].Value
                if ($sv -eq $OriginalCallId -and $tv -and $tv -ne "-") { $TelFirst = $tv; break }
            }
        }
    }

    # ── Contar cuántos MoveSessionToConference se enviaron ──────────
    # Se calcula más abajo junto con la extracción de logs por paso,
    # porque necesita buscar también en el slot base del mismo segundo.

    # ── Medir duración de la conferencia ────────────────────────────
    # La conferencia termina cuando nConsultCall recibe ProcessSessionEndedEvent,
    # visible como "FIN DE LLAMADA" o "CUELGUE MANUAL" en la grilla.
    $HoraFin = ""; $HoraFinExacta = ""; $DuracionStr = "No determinada"
    $RowIdx = $SelRow.Index
    for ($i = $RowIdx + 1; $i -lt $GridResultados.Rows.Count; $i++) {
        $r = $GridResultados.Rows[$i]
        if ($r.IsNewRow) { continue }
        $sv = $r.Cells["Sesion"].Value
        $iv = $r.Cells["Interpretacion"].Value
        if ($sv -eq $SesConsult -and $iv -match "FIN DE LLAMADA|CUELGUE MANUAL") {
            $HoraFin = $r.Cells["Hora"].Value
            # Intentar extraer ms exactos del ProcessSessionEndedEvent en el raw
            $rawFinA = $r.Cells["EvAgente"].Tag
            if ($rawFinA -and $rawFinA -match "\[\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2}:\d{3})\].*ProcessSessionEndedEvent") {
                $HoraFinExacta = $matches[1]
            }
            break
        }
    }
    if ($HoraFin -ne "") {
        try {
            $tIni = [datetime]::ParseExact(($HoraGrid -replace ",\d+$",""), "HH:mm:ss", $null)
            $tFin = [datetime]::ParseExact(($HoraFin  -replace ",\d+$",""), "HH:mm:ss", $null)
            $dur  = [int]($tFin - $tIni).TotalSeconds
            $DuracionStr = if ($dur -ge 60) { "$([int]($dur/60))m $($dur % 60)s" } else { "${dur}s" }
        } catch { $DuracionStr = "Error al calcular" }
    }

    # ── Ventana ─────────────────────────────────────────────────────
    $fConf = New-Object System.Windows.Forms.Form
    $fConf.Text            = "Detalle de Conferencia — $HoraGrid"
    $fConf.Size            = New-Object System.Drawing.Size(640, 480)
    $fConf.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $fConf.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $fConf.MaximizeBox     = $false
    $fConf.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)

    $rtbConf = New-Object System.Windows.Forms.RichTextBox
    $rtbConf.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $rtbConf.ReadOnly   = $true
    $rtbConf.BackColor  = [System.Drawing.Color]::FromArgb(20,20,20)
    $rtbConf.ForeColor  = [System.Drawing.Color]::White
    $rtbConf.Font       = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtbConf.WordWrap   = $true
    $fConf.Controls.Add($rtbConf)

    $cOk  = [System.Drawing.Color]::LimeGreen
    $cOrch= [System.Drawing.Color]::MediumOrchid
    $cHdr = [System.Drawing.Color]::CornflowerBlue
    $cInfo= [System.Drawing.Color]::Silver
    $cW   = [System.Drawing.Color]::White
    $cGold= [System.Drawing.Color]::Gold
    $cWarn= [System.Drawing.Color]::Orange
    $cRaw = [System.Drawing.Color]::FromArgb(130, 150, 200)

    $LC = {
        param($txt, $col, [bool]$bold=$false)
        $rtbConf.SelectionStart  = $rtbConf.TextLength
        $rtbConf.SelectionLength = 0
        $rtbConf.SelectionColor  = $col
        $rtbConf.SelectionFont   = if ($bold) { New-Object System.Drawing.Font($rtbConf.Font.FontFamily, $rtbConf.Font.Size, [System.Drawing.FontStyle]::Bold) } else { $rtbConf.Font }
        $rtbConf.AppendText("$txt`n")
    }

    # Conferencia exitosa = el switch de audio confirmó el merge físico (Conference_Merged).
    # Conference_CompleteConf puede llegar aunque falle el merge (leg bridge/phantom sin teléfono real).
    $EsExitosa = ($RawConf -match "Conference_Merged")

    # ── Contenido ───────────────────────────────────────────────────
    if ($EsExitosa) {
        & $LC "  CONFERENCIA ESTABLECIDA EXITOSAMENTE" $cOk $true
        & $LC "  ════════════════════════════════════" $cHdr
        & $LC "" $cW
        & $LC "  Que paso?" $cHdr $true
        & $LC "  El agente unio dos llamadas activas en una sola" $cW
        & $LC "  conversacion de 3 participantes en tiempo real." $cW
        & $LC "  Los tres pudieron escucharse y hablar al mismo tiempo." $cW
    } else {
        & $LC "  CONFERENCIA — PARTICIPANTE DE CONSULTA NO IDENTIFICADO" $cWarn $true
        & $LC "  ════════════════════════════════════════════════════════" $cHdr
        & $LC "" $cW
        & $LC "  Que paso?" $cHdr $true
        & $LC "  El PBX envio Conference_CompleteConf, pero la sesion de consulta" $cW
        & $LC "  no tiene telefono identificado en los logs. Es probable que la" $cW
        & $LC "  conferencia haya unido un leg interno del sistema (bridge/phantom)" $cW
        & $LC "  en lugar de un participante externo real." $cW
    }
    & $LC "" $cW

    & $LC "  Hora exacta de establecimiento:" $cHdr $true
    & $LC "  $HoraExacta" $cGold
    & $LC "" $cW

    & $LC "  Duracion de la conferencia:" $cHdr $true
    if ($HoraFin -ne "") {
        $InicioLabel = if ($HoraExacta -notmatch "no disponibles") { $HoraExacta } else { $HoraGrid }
        $FinLabel    = if ($HoraFinExacta -ne "") { $HoraFinExacta } else { $HoraFin }
        & $LC "  Inicio : $InicioLabel" $cInfo
        & $LC "  Fin    : $FinLabel" $cInfo
        & $LC "  Total  : $DuracionStr" $cGold $true
    } else {
        & $LC "  No se encontro el fin de la conferencia en los logs del dia." $cInfo
    }
    & $LC "" $cW

    & $LC "  Participantes confirmados por el PBX:" $cHdr $true
    if ($SesFirst -ne "") {
        $LabelFirst = if ($TelFirst -ne "") { "$TelFirst" } else { "Numero no identificado" }
        & $LC "  • Llamada principal   (Sesion $SesFirst) — $LabelFirst" $cOrch
    }
    if ($SesConsult -ne "") {
        $LabelConsult  = if ($TelConsult -ne "") { "$TelConsult" } else { "Numero no identificado" }
        $ColorConsult  = if ($TelConsult -ne "") { $cOrch } else { $cWarn }
        & $LC "  • Llamada de consulta (Sesion $SesConsult) — $LabelConsult" $ColorConsult
    }
    & $LC "" $cW

    # ── Extraer líneas de log para cada paso ───────────────────────
    $RawStep1 = ""; $RawStep2 = ""; $RawStep3 = ""; $RawStep4 = ""; $RawStep4b = ""
    $BaseSegConf = if ($HoraGrid -match "^(\d{2}:\d{2}:\d{2})") { $matches[1] } else { $HoraGrid }
    # Buscar en el ms-slot de CONFERENCIA ESTABLECIDA ($RawConf)
    foreach ($rl in ($RawConf -split "`n")) {
        $rl = $rl.Trim(); if ($rl -eq "") { continue }
        if ($RawStep2  -eq "" -and $rl -match "OnRequestMoveSessionToConference") { $RawStep2  = $rl }
        if ($RawStep3  -eq "" -and $rl -match "MoveSessionToConferenceRequest")   { $RawStep3  = $rl }
        if ($RawStep4  -eq "" -and $rl -match "Conference_CompleteConf")          { $RawStep4  = $rl }
        if ($RawStep4b -eq "" -and $rl -match "Conference_Merged")                { $RawStep4b = $rl }
    }
    # Fix: también buscar en el slot base del mismo segundo —
    # MoveSessionToConferenceRequest puede haber quedado allí si llegó antes que Conference_CompleteConf.
    $RawConfBase = ""
    foreach ($rB in $GridResultados.Rows) {
        if ($rB.IsNewRow) { continue }
        if ($rB.Cells["Hora"].Value -eq $BaseSegConf) { $RawConfBase = "$($rB.Cells['EvAgente'].Tag)"; break }
    }
    foreach ($rl in ($RawConfBase -split "`n")) {
        $rl = $rl.Trim(); if ($rl -eq "") { continue }
        if ($RawStep3 -eq "" -and $rl -match "MoveSessionToConferenceRequest") { $RawStep3 = $rl }
    }
    $CantidadMoves = ([regex]::Matches($RawConf + "`n" + $RawConfBase, "MoveSessionToConferenceRequest")).Count
    # Buscar hacia atrás el evento de inicio — puede ser Drag/Drop o Botón, segundos antes del ESTABLECIDA
    $NumeroMarcadoConf = ""
    for ($iDD = $SelRow.Index - 1; $iDD -ge 0; $iDD--) {
        $rowDD = $GridResultados.Rows[$iDD]
        if ($rowDD.IsNewRow) { continue }
        $evDD  = $rowDD.Cells["EvAgente"].Value
        $tagDD = $rowDD.Cells["EvAgente"].Tag
        if ($evDD -match "Inicio de Conferencia Drag") {
            if ($tagDD) {
                $mDD = [regex]::Match($tagDD, "GUI Method STARTED: ConferenceCallDragDropHandler[^\n]*")
                if ($mDD.Success) { $RawStep1 = $mDD.Value.Trim() }
            }
            break
        } elseif ($evDD -match "Inicio de Conferencia \(Bot") {
            if ($tagDD) {
                $mCI = [regex]::Match($tagDD, "InitiateConference\(Call\[[^\]]+\],([^)]+)\)")
                if ($mCI.Success) {
                    $RawStep1 = $mCI.Value.Trim()
                    $NumeroMarcadoConf = $mCI.Groups[1].Value.Trim()
                } else {
                    $mCH = [regex]::Match($tagDD, "GUI Method STARTED: ConferenceCallHandler[^\n]*")
                    if ($mCH.Success) { $RawStep1 = $mCH.Value.Trim() }
                }
            }
            break
        }
    }

    & $LC "  Como lo hizo el sistema paso a paso:" $cHdr $true
    $labelPaso1 = if ($RawStep1 -match "ConferenceCallDragDropHandler") {
        "  1. El agente inicio conferencia arrastrando una llamada (Drag & Drop)"
    } elseif ($NumeroMarcadoConf -ne "") {
        "  1. El agente presiono el boton de conferencia y marco $NumeroMarcadoConf"
    } else {
        "  1. El agente presiono el boton de Conferencia en OneX"
    }
    & $LC $labelPaso1 $cInfo
    if ($RawStep1 -ne "") {
        $s1 = if ($RawStep1.Length -gt 110) { $RawStep1.Substring(0,110) + "..." } else { $RawStep1 }
        & $LC "     $s1" $cRaw
    }
    & $LC "  2. OneX marco al segundo participante (llamada de consulta)" $cInfo
    if ($RawStep2 -ne "") {
        $s2 = if ($RawStep2.Length -gt 110) { $RawStep2.Substring(0,110) + "..." } else { $RawStep2 }
        & $LC "     $s2" $cRaw
    }
    & $LC "  3. OneX envio $CantidadMoves solicitud(es) al PBX (MoveSessionToConference)" $cInfo
    if ($RawStep3 -ne "") {
        $s3 = if ($RawStep3.Length -gt 110) { $RawStep3.Substring(0,110) + "..." } else { $RawStep3 }
        & $LC "     $s3" $cRaw
    }
    & $LC "  4. El PBX respondio confirmando la union (Conference_CompleteConf)" $cOk
    if ($RawStep4 -ne "") {
        $s4 = if ($RawStep4.Length -gt 110) { $RawStep4.Substring(0,110) + "..." } else { $RawStep4 }
        & $LC "     $s4" $cRaw
    }
    if ($EsExitosa) {
        & $LC "  5. Las dos llamadas quedaron conectadas en 3 vias (Conference_Merged)" $cOk
        if ($RawStep4b -ne "") {
            $s4b = if ($RawStep4b.Length -gt 110) { $RawStep4b.Substring(0,110) + "..." } else { $RawStep4b }
            & $LC "     $s4b" $cRaw
        }
    } else {
        & $LC "  5. [!] El participante de consulta no pudo ser identificado en los logs" $cWarn
        & $LC "        La conferencia puede haber unido un leg interno del sistema." $cWarn
    }
    & $LC "" $cW

    & $LC "  Evidencia tecnica del PBX:" $cHdr $true
    if ($RawConf -match "Conference_CompleteConf[^\n]+") { & $LC "  $($matches[0].Trim())" $cInfo }

    $rtbConf.SelectionStart = 0; $rtbConf.ScrollToCaret()
    [void]$fConf.ShowDialog($Form)
    $fConf.Dispose()
})

# ====================================================================
# DETALLE DE TRANSFERENCIA: CÓMO SE COMPLETÓ
# ====================================================================
$mnuDetalleTransf.Add_Click({
    $SelRow = if ($GridResultados.SelectedRows.Count -gt 0) { $GridResultados.SelectedRows[0] } else { $null }
    if ($null -eq $SelRow) { return }

    $HoraGrid  = $SelRow.Cells["Hora"].Value
    $RawTransf = $SelRow.Cells["EvAgente"].Tag

    # ── Hora exacta con milisegundos ────────────────────────────────
    $HoraExacta = "$HoraGrid (ms no disponibles)"
    if ($RawTransf -match "\[(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2}:\d{3})\].*Transfer_CompleteSetup") {
        $HoraExacta = "$($matches[2])  (milisegundos del PBX)"
    } elseif ($RawTransf -match "\[(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2}:\d{3})\].*SetPhoneDisplay.*Transferencia") {
        $HoraExacta = "$($matches[2])  (milisegundos del PBX)"
    }

    # ── Extraer sesiones de Transfer_CompleteSetup ──────────────────
    $SesOrigen = ""; $SesDest = ""
    if ($RawTransf -match "Transfer_CompleteSetup: nfirstCall:\s*(\d+),\s*nSecondCall:\s*(\d+)") {
        $SesOrigen = $matches[1]; $SesDest = $matches[2]
    }

    # ── Buscar teléfonos en la grilla por sesión ────────────────────
    $TelOrigen = ""; $TelDest = ""
    foreach ($row in $GridResultados.Rows) {
        if ($row.IsNewRow) { continue }
        $sv = $row.Cells["Sesion"].Value; $tv = $row.Cells["Telefono"].Value
        if ($SesOrigen -ne "" -and $sv -eq $SesOrigen -and $tv -and $tv -ne "-") { $TelOrigen = $tv }
        if ($SesDest   -ne "" -and $sv -eq $SesDest   -and $tv -and $tv -ne "-") { $TelDest   = $tv }
    }
    # Fallback: buscar destino en el texto del evento (ya calculado en la etiqueta de la celda)
    $CeldaTexto = $SelRow.Cells["EvAgente"].Value
    if ($TelDest -eq "" -and $CeldaTexto -match "TRANSFERENCIA COMPLETADA:\s*[^\s→]+\s*→\s*(.+)") {
        $TelDest = $matches[1].Trim()
    }
    if ($TelOrigen -eq "" -and $CeldaTexto -match "TRANSFERENCIA COMPLETADA:\s*([^\s→]+)\s*→") {
        $TelOrigen = $matches[1].Trim()
    }

    # ── Detectar tipo de transferencia ─────────────────────────────
    # Atendida = hubo consulta previa (la sesión de dest aparece como "MARCANDO A LA EXT." / consulta)
    # Ciega    = sin consulta, la transferencia va directo
    $TipoTransf = "Transferencia Ciega (sin consulta previa)"
    foreach ($row in $GridResultados.Rows) {
        if ($row.IsNewRow) { continue }
        $sv = $row.Cells["Sesion"].Value
        $iv = $row.Cells["Interpretacion"].Value
        if ($SesDest -ne "" -and $sv -eq $SesDest -and $iv -match "MARCANDO") {
            $TipoTransf = "Transferencia Atendida (el agente habló con el destino antes de transferir)"
            break
        }
    }

    # ── Medir duración del proceso de transferencia ─────────────────
    # Desde TRANSFERENCIA INICIADA (o CONSULTA) hasta esta fila de COMPLETADA
    $HoraInicio = ""; $DuracionStr = "No determinada"
    $RowIdx = $SelRow.Index
    for ($i = $RowIdx - 1; $i -ge 0; $i--) {
        $r = $GridResultados.Rows[$i]
        if ($r.IsNewRow) { continue }
        $av = $r.Cells["EvAgente"].Value
        if ($av -match "TRANSFERENCIA INICIADA|Asesor presiona bot.n Transferir|TRANSFERENCIA EN PROCESO") {
            $HoraInicio = $r.Cells["Hora"].Value; break
        }
    }
    if ($HoraInicio -ne "") {
        try {
            # Los slots vienen como "HH:mm:ss" o "HH:mm:ss,fff" (la coma separa los ms).
            # ParseExact con solo "HH:mm:ss" reventaba con la parte de ms → "Error al calcular".
            $pIni = $HoraInicio -split ','
            $pFin = $HoraGrid   -split ','
            $tIni = [datetime]::ParseExact($pIni[0], "HH:mm:ss", $null)
            if ($pIni.Count -gt 1 -and $pIni[1] -match '^\d+$') { $tIni = $tIni.AddMilliseconds([int]$pIni[1]) }
            $tFin = [datetime]::ParseExact($pFin[0], "HH:mm:ss", $null)
            if ($pFin.Count -gt 1 -and $pFin[1] -match '^\d+$') { $tFin = $tFin.AddMilliseconds([int]$pFin[1]) }
            $msTot = [int]($tFin - $tIni).TotalMilliseconds
            if ($msTot -lt 0) { $msTot = 0 }
            $DuracionStr = if ($msTot -ge 60000) { "$([int]($msTot/60000))m $([int](($msTot%60000)/1000))s" }
                           elseif ($msTot -ge 1000) { ("{0:0.##}s" -f ($msTot/1000.0)) }
                           else { "$msTot ms" }
        } catch { $DuracionStr = "Error al calcular" }
    }

    # ── Evidencia técnica ───────────────────────────────────────────
    $LineaCompleteSetup = ""
    if ($RawTransf -match "Transfer_CompleteSetup[^\n]+") { $LineaCompleteSetup = $matches[0].Trim() }
    $LineaDisplay = ""
    if ($RawTransf -match "SetPhoneDisplay[^\n]+") { $LineaDisplay = $matches[0].Trim() }

    # ── Líneas de log de los pasos previos (escaneo de filas hacia atrás) ──
    # Recorre desde la fila COMPLETADA hasta la de TRANSFERENCIA INICIADA y recolecta la
    # línea cruda (con su timestamp) que originó cada paso, para mostrarla como evidencia.
    $RawPressTransfer = ""   # OnRequestTransferSession (el agente presionó Transferir)
    $RawTransferReq   = ""   # TransferSessionRequest (OneX → PBX)
    $RawActivate      = ""   # Transfer_ActivateConsultCall (consulta atendida)
    for ($i = $RowIdx; $i -ge 0; $i--) {
        $r = $GridResultados.Rows[$i]
        if ($r.IsNewRow) { continue }
        $tg = $r.Cells["EvAgente"].Tag
        if ($tg) {
            if ($RawTransferReq   -eq "" -and $tg -match "[^\r\n]*Message type= TransferSessionRequest[^\r\n]*") { $RawTransferReq   = $matches[0].Trim() }
            if ($RawPressTransfer -eq "" -and $tg -match "[^\r\n]*OnRequestTransferSession\(\)[^\r\n]*")          { $RawPressTransfer = $matches[0].Trim() }
            if ($RawActivate      -eq "" -and $tg -match "[^\r\n]*Transfer_ActivateConsultCall[^\r\n]*")          { $RawActivate      = $matches[0].Trim() }
        }
        if ($i -lt $RowIdx -and $r.Cells["EvAgente"].Value -match "TRANSFERENCIA INICIADA|Asesor presiona bot.n Transferir") { break }
    }

    # ── Ventana ─────────────────────────────────────────────────────
    $fTransf = New-Object System.Windows.Forms.Form
    $fTransf.Text            = "Detalle de Transferencia — $HoraGrid"
    $fTransf.Size            = New-Object System.Drawing.Size(640, 500)
    $fTransf.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $fTransf.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $fTransf.MaximizeBox     = $false
    $fTransf.BackColor       = [System.Drawing.Color]::FromArgb(28,28,28)

    $rtbT = New-Object System.Windows.Forms.RichTextBox
    $rtbT.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $rtbT.ReadOnly  = $true
    $rtbT.BackColor = [System.Drawing.Color]::FromArgb(20,20,20)
    $rtbT.ForeColor = [System.Drawing.Color]::White
    $rtbT.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtbT.WordWrap  = $true
    $fTransf.Controls.Add($rtbT)

    $cOk   = [System.Drawing.Color]::LimeGreen
    $cPlum = [System.Drawing.Color]::Plum
    $cHdr  = [System.Drawing.Color]::CornflowerBlue
    $cInfo = [System.Drawing.Color]::Silver
    $cW    = [System.Drawing.Color]::White
    $cGold = [System.Drawing.Color]::Gold
    $cRaw  = [System.Drawing.Color]::DarkGray

    $LT = {
        param($txt, $col, [bool]$bold=$false)
        $rtbT.SelectionStart  = $rtbT.TextLength
        $rtbT.SelectionLength = 0
        $rtbT.SelectionColor  = $col
        $rtbT.SelectionFont   = if ($bold) { New-Object System.Drawing.Font($rtbT.Font.FontFamily, $rtbT.Font.Size, [System.Drawing.FontStyle]::Bold) } else { $rtbT.Font }
        $rtbT.AppendText("$txt`n")
    }
    # Helper: imprime la linea de log cruda (evidencia) debajo de un paso, truncada.
    $LTraw = {
        param($raw)
        if ($raw -and $raw -ne "") {
            $s = if ($raw.Length -gt 110) { $raw.Substring(0,110) + "..." } else { $raw }
            & $LT "     $s" $cRaw
        }
    }

    # ── Contenido ───────────────────────────────────────────────────
    & $LT "  TRANSFERENCIA COMPLETADA EXITOSAMENTE" $cOk $true
    & $LT "  ══════════════════════════════════════" $cHdr
    & $LT "" $cW

    & $LT "  Que paso?" $cHdr $true
    & $LT "  El agente transfirió la llamada del cliente a otro destino." $cW
    & $LT "  El PBX confirmó que las dos sesiones quedaron unidas" $cW
    & $LT "  y el agente salió de la conversación exitosamente." $cW
    & $LT "" $cW

    & $LT "  Tipo de transferencia:" $cHdr $true
    & $LT "  $TipoTransf" $cPlum
    & $LT "" $cW

    & $LT "  Hora exacta de confirmación del PBX:" $cHdr $true
    & $LT "  $HoraExacta" $cGold
    & $LT "" $cW

    & $LT "  Tiempo del proceso (Iniciada → Completada):" $cHdr $true
    if ($HoraInicio -ne "") {
        & $LT "  Inicio : $HoraInicio" $cInfo
        & $LT "  Fin    : $HoraGrid" $cInfo
        & $LT "  Total  : $DuracionStr" $cGold $true
    } else {
        & $LT "  No se encontro el inicio de la transferencia en el log." $cInfo
    }
    & $LT "" $cW

    & $LT "  Origen y destino confirmados por el PBX:" $cHdr $true
    $LabelOrigen = if ($TelOrigen -ne "") { "$TelOrigen" } else { "Sesion $SesOrigen" }
    $LabelDest   = if ($TelDest   -ne "") { "$TelDest"   } else { "Sesion $SesDest" }
    & $LT "  • Llamada original  (Sesion $SesOrigen) — $LabelOrigen" $cPlum
    & $LT "  • Transferida hacia (Sesion $SesDest)   — $LabelDest" $cPlum
    & $LT "" $cW

    & $LT "  Por que se considera exitosa?" $cHdr $true
    & $LT "  Dos eventos del PBX lo confirman:" $cW
    if ($LineaCompleteSetup -ne "") {
        & $LT "  1. Transfer_CompleteSetup — el PBX unio fisicamente" $cOk
        & $LT "     las dos sesiones de voz en el conmutador." $cOk
    }
    if ($LineaDisplay -ne "") {
        & $LT "  2. SetPhoneDisplay 'Transferencia realizada' — el PBX" $cOk
        & $LT "     actualizo la pantalla del telefono del agente para" $cOk
        & $LT "     confirmar que la operacion fue aceptada." $cOk
    }
    & $LT "" $cW

    & $LT "  Como lo hizo el sistema paso a paso:" $cHdr $true
    if ($TipoTransf -match "Atendida") {
        & $LT "  1. El agente presiono Transferir en OneX" $cInfo
        & $LTraw $RawPressTransfer
        & $LT "  2. Sono la linea del destino — el agente hablo con el" $cInfo
        & $LT "     receptor para avisarle de la transferencia" $cInfo
        & $LTraw $RawActivate
        & $LT "  3. El agente confirmo la transferencia en OneX" $cInfo
        & $LT "  4. OneX envio TransferSessionRequest al PBX" $cInfo
        & $LTraw $RawTransferReq
        & $LT "  5. El PBX respondio con Transfer_CompleteSetup" $cOk
        & $LTraw $LineaCompleteSetup
        & $LT "  6. El cliente quedo conectado directamente con el destino" $cOk
    } else {
        & $LT "  1. El agente presiono Transferir en OneX" $cInfo
        & $LTraw $RawPressTransfer
        & $LT "  2. OneX envio TransferSessionRequest al PBX con el numero destino" $cInfo
        & $LTraw $RawTransferReq
        & $LT "  3. El PBX enruto la llamada sin esperar confirmacion del destino" $cInfo
        & $LT "  4. El PBX respondio con Transfer_CompleteSetup" $cOk
        & $LTraw $LineaCompleteSetup
        & $LT "  5. El cliente quedo enrutado hacia el destino" $cOk
    }
    & $LT "" $cW

    & $LT "  Evidencia tecnica del PBX:" $cHdr $true
    if ($LineaCompleteSetup -ne "") { & $LT "  $LineaCompleteSetup" $cInfo }
    if ($LineaDisplay       -ne "") { & $LT "  $LineaDisplay"       $cInfo }

    $rtbT.SelectionStart = 0; $rtbT.ScrollToCaret()
    [void]$fTransf.ShowDialog($Form)
    $fTransf.Dispose()
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
# ====================================================================
# DIAGNÓSTICO DE "Falla grave en llamada": traduce el System.Exception
# crudo a lenguaje claro (qué pasó, dónde tronó y qué significa).
# La etiqueta se pinta en el render cuando AppLog trae "System.Exception"
# (radar ERR CRÍTICO de PASO 4) — un solo gatillo, muchas causas posibles.
# ====================================================================
function Get-DiagnosticoFallaGrave {
    param([string]$raw)
    $d = "DIAGNÓSTICO EN LENGUAJE CLARO`n" + ("=" * 46) + "`n`n"

    # 1) ¿QUÉ tronó? — el mensaje después de "System.Exception:" (o el tipo de excepción)
    $msgExc = ""
    if ($raw -match "System\.\w*Exception:?\s*([^\r\n]+)") { $msgExc = $matches[1].Trim() }

    if ($raw -match "Unable to find VoiceInteraction for call") {
        $idAfect = ""; $stAfect = ""
        if ($raw -match "call:Id=(\d+).*?State=(\w+)") { $idAfect = $matches[1]; $stAfect = $matches[2] }
        $d += "► QUÉ PASÓ:`nEl OneX intentó ejecutar una acción sobre la llamada" + $(if ($idAfect) { " (Sesión $idAfect)" } else { "" }) + ", pero su interacción interna YA NO EXISTÍA" + $(if ($stAfect -eq "Disconnected") { " — la llamada ya estaba colgada (State=Disconnected)" } else { "" }) + ".`n`n"
        $d += "► CAUSA TÍPICA:`nCondición de carrera: el cliente colgó en el instante exacto en que el sistema (o el asesor) ejecutaba la acción. La llamada murió a mitad de la operación.`n`n"
        if ($raw -match "CheckAutoAccept|AnswerCall") {
            $d += "► IMPACTO (¡IMPORTANTE!):`nOcurrió durante la AUTO-CONTESTACIÓN: el motor de auto-answer quedó roto. Si otra llamada estaba timbrando en ese momento, quedó huérfana: sin auto-contestarse y SIN BOTONES en la ventana para tomarla (revisar los eventos 'FALLA EN LA RECEPCIÓN DE LLAMADA' y 'SIN BOTONES PARA TOMAR LA LLAMADA' en el grid).`n`n"
        } elseif ($raw -match "Transfer") {
            $d += "► IMPACTO:`nOcurrió durante una TRANSFERENCIA: la operación de transferencia no pudo completarse.`n`n"
        } elseif ($raw -match "Hold|Unhold") {
            $d += "► IMPACTO:`nOcurrió durante un HOLD/RETOMAR: la operación de espera no pudo completarse.`n`n"
        }
    }
    elseif ($raw -match "NullReferenceException") {
        $d += "► QUÉ PASÓ:`nReferencia nula: un componente interno del OneX intentó usar un objeto que no existe. Es un defecto de la aplicación (no un error del asesor ni de la red).`n`n"
    }
    elseif ($raw -match "(?i)TimeoutException|timed out") {
        $d += "► QUÉ PASÓ:`nTimeout: una operación esperó respuesta (del conmutador o de un componente interno) y nunca llegó a tiempo.`n`n"
    }
    elseif ($raw -match "(?i)SocketException|connection.*(lost|refused|reset)") {
        $d += "► QUÉ PASÓ:`nProblema de conexión de red entre el OneX y el conmutador durante la operación.`n`n"
    }
    else {
        $d += "► QUÉ PASÓ:`nExcepción interna del OneX Agent no catalogada todavía." + $(if ($msgExc) { "`nMensaje de la excepción:`n   `"$msgExc`"" } else { "" }) + "`n`n"
    }

    # 2) ¿DÓNDE tronó? — módulos del stack ("en Avaya.OneXAgent.<Modulo>...")
    $Modulos = @{}
    foreach ($mMod in [regex]::Matches($raw, "Avaya\.OneXAgent\.(\w+)")) { $Modulos[$mMod.Groups[1].Value] = $true }
    $MapaMod = @{ "WorkService"="Motor de manejo de llamadas"; "CMService"="Comunicación con el conmutador (CM)"; "ModelManager"="Modelo de estado interno"; "Util"="Utilerías internas"; "ScreenPops"="Ventanas emergentes"; "AudioService"="Audio"; "Work"="Objetos de trabajo (llamadas)" }
    $ListaMod = @($Modulos.Keys | Where-Object { $MapaMod.ContainsKey($_) } | ForEach-Object { "• $_  →  $($MapaMod[$_])" })
    if ($ListaMod.Count -gt 0) { $d += "► DÓNDE TRONÓ (módulos del stack):`n" + ($ListaMod -join "`n") + "`n`n" }

    # 3) Cómo leer el log crudo (para las no catalogadas)
    $d += "► CÓMO LEER EL LOG CRUDO:`n1. La línea 'System.Exception: ...' dice LA CAUSA.`n2. Las líneas 'en Avaya.OneXAgent...' son el camino del error (de lo más específico arriba a lo más general abajo).`n3. La hora del ERROR es el momento exacto de la falla — cruzarla con los eventos vecinos del grid."
    return $d
}

$GridResultados.Add_CellClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $ColName = $GridResultados.Columns[$e.ColumnIndex].Name
        $ColumnasRegulares = @("Interpretacion","EvAgente","EvAudio","EvAux","EvSysLog","EvAppLog","EvIspeac","EvDtmf")

        # Popup de "Validación AutoAnswer" RETIRADO: el modo de contestación (automática/manual)
        # ahora se muestra directamente en la fila del INICIO DE LLAMADA. El bloque inferior queda
        # inerte (if $false) y puede borrarse por completo en una limpieza futura.
        if ($false) {
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
            if ($LogCrudo) {
                $InterpFila = ""; try { $InterpFila = [string]$GridResultados.Rows[$e.RowIndex].Cells["Interpretacion"].Value } catch {}
                if ($InterpFila -match "Falla grave en llamada" -or $LogCrudo -match "System\.\w*Exception") {
                    # Ventana especial con botón de DIAGNÓSTICO (el raw de una excepción es difícil de leer)
                    $fFG = New-Object System.Windows.Forms.Form
                    $fFG.Text = "Falla grave en llamada — Evidencia del log"; $fFG.Size = New-Object System.Drawing.Size(760, 560); $fFG.StartPosition = "CenterParent"; $fFG.BackColor = $ColorFondo; $fFG.ForeColor = $ColorTexto
                    $rtbFG = New-Object System.Windows.Forms.RichTextBox
                    $rtbFG.Location = New-Object System.Drawing.Point(12, 12); $rtbFG.Size = New-Object System.Drawing.Size(720, 440); $rtbFG.Anchor = "Top,Bottom,Left,Right"
                    $rtbFG.ReadOnly = $true; $rtbFG.BackColor = [System.Drawing.Color]::FromArgb(20,20,20); $rtbFG.ForeColor = [System.Drawing.Color]::Gainsboro; $rtbFG.Font = New-Object System.Drawing.Font("Consolas", 9)
                    $rtbFG.Text = "EVIDENCIA DEL LOG ORIGINAL:`n`n$LogCrudo"
                    $btnDiagFG = New-Object System.Windows.Forms.Button
                    $btnDiagFG.Text = "VER DIAGNÓSTICO (lenguaje claro)"; $btnDiagFG.Location = New-Object System.Drawing.Point(12, 465); $btnDiagFG.Size = New-Object System.Drawing.Size(280, 35); $btnDiagFG.Anchor = "Bottom,Left"
                    $btnDiagFG.BackColor = [System.Drawing.Color]::DarkGreen; $btnDiagFG.ForeColor = [System.Drawing.Color]::White; $btnDiagFG.FlatStyle = "Flat"; $btnDiagFG.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                    $btnDiagFG.Tag = [string]$LogCrudo
                    $btnDiagFG.Add_Click({
                        param($s, $ev)
                        $DiagTexto = Get-DiagnosticoFallaGrave -raw ([string]$s.Tag)
                        [System.Windows.Forms.MessageBox]::Show($DiagTexto, "Diagnóstico — Falla grave en llamada", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    })
                    $btnCerrarFG = New-Object System.Windows.Forms.Button
                    $btnCerrarFG.Text = "Cerrar"; $btnCerrarFG.Location = New-Object System.Drawing.Point(632, 465); $btnCerrarFG.Size = New-Object System.Drawing.Size(100, 35); $btnCerrarFG.Anchor = "Bottom,Right"
                    $btnCerrarFG.BackColor = [System.Drawing.Color]::Gray; $btnCerrarFG.ForeColor = [System.Drawing.Color]::White; $btnCerrarFG.FlatStyle = "Flat"
                    $btnCerrarFG.Add_Click({ param($s, $ev) $s.FindForm().Close() })
                    $fFG.Controls.AddRange(@($rtbFG, $btnDiagFG, $btnCerrarFG))
                    [void]$fFG.ShowDialog($Form); $fFG.Dispose()
                } else {
                    [System.Windows.Forms.MessageBox]::Show("EVIDENCIA DEL LOG ORIGINAL:`n`n$LogCrudo", "Análisis de Logs Avaya OneX Agent - Detalle", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
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

    $FormTiempo = New-Object System.Windows.Forms.Form; $FormTiempo.Text = "Configurar Aislamiento RAW"; $FormTiempo.Size = New-Object System.Drawing.Size(350, 250); $FormTiempo.StartPosition = "CenterParent"; $FormTiempo.BackColor = $ColorPanel; $FormTiempo.ForeColor = $ColorTexto
    # Horas con máscara fija estilo IPv4: se teclean SOLO dígitos y el cursor
    # brinca los ":" automáticamente (10 19 49 → 10:19:49). El DateTimePicker con
    # spinner NO auto-avanzaba de bloque (sobreescribía el mismo campo) — probado por Pablo.
    $lblT1 = New-Object System.Windows.Forms.Label; $lblT1.Text = "Hora Inicio:"; $lblT1.Location = New-Object System.Drawing.Point(20, 20); $lblT1.AutoSize = $true
    $mtbT1 = New-Object System.Windows.Forms.MaskedTextBox; $mtbT1.Mask = "00:00:00"; $mtbT1.Location = New-Object System.Drawing.Point(160, 18); $mtbT1.Size = New-Object System.Drawing.Size(150, 25); $mtbT1.Text = "000000"; $mtbT1.InsertKeyMode = [System.Windows.Forms.InsertKeyMode]::Overwrite; $mtbT1.BackColor = $ColorFondo; $mtbT1.ForeColor = [System.Drawing.Color]::LimeGreen; $mtbT1.Font = New-Object System.Drawing.Font("Consolas", 11)
    $lblT2 = New-Object System.Windows.Forms.Label; $lblT2.Text = "Hora Fin:"; $lblT2.Location = New-Object System.Drawing.Point(20, 60); $lblT2.AutoSize = $true
    $mtbT2 = New-Object System.Windows.Forms.MaskedTextBox; $mtbT2.Mask = "00:00:00"; $mtbT2.Location = New-Object System.Drawing.Point(160, 58); $mtbT2.Size = New-Object System.Drawing.Size(150, 25); $mtbT2.Text = "235959"; $mtbT2.InsertKeyMode = [System.Windows.Forms.InsertKeyMode]::Overwrite; $mtbT2.BackColor = $ColorFondo; $mtbT2.ForeColor = [System.Drawing.Color]::LimeGreen; $mtbT2.Font = New-Object System.Drawing.Font("Consolas", 11)
    # Al entrar al campo (clic o Tab), posicionar el cursor al inicio para teclear de corrido.
    # El MaskedTextBox reubica el caret al FINAL después de disparar Enter, así que un
    # SelectionStart=0 directo queda pisado. Se difiere con BeginInvoke para que corra
    # DESPUÉS de que el control termine de asentar el foco. GetNewClosure captura el control.
    $mtbT1.Add_Enter({ $ctl = $this; $ctl.BeginInvoke([Action]({ $ctl.SelectionStart = 0; $ctl.SelectionLength = 0 }.GetNewClosure())) | Out-Null })
    $mtbT2.Add_Enter({ $ctl = $this; $ctl.BeginInvoke([Action]({ $ctl.SelectionStart = 0; $ctl.SelectionLength = 0 }.GetNewClosure())) | Out-Null })
    # Fecha VISIBLE en el diálogo: antes se tomaba en silencio del picker principal y una fecha
    # equivocada producía "No se encontró actividad" sin pista de la causa.
    $lblTF = New-Object System.Windows.Forms.Label; $lblTF.Text = "Fecha de los logs:"; $lblTF.Location = New-Object System.Drawing.Point(20, 100); $lblTF.AutoSize = $true
    $dtpFechaRAW = New-Object System.Windows.Forms.DateTimePicker; $dtpFechaRAW.Location = New-Object System.Drawing.Point(160, 98); $dtpFechaRAW.Size = New-Object System.Drawing.Size(150, 25); $dtpFechaRAW.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom; $dtpFechaRAW.CustomFormat = "dd/MM/yyyy"; $dtpFechaRAW.Value = $dtpFecha.Value
    $btnGenerar = New-Object System.Windows.Forms.Button; $btnGenerar.Text = "EXTRAER LOGS RAW"; $btnGenerar.Location = New-Object System.Drawing.Point(100, 145); $btnGenerar.BackColor = [System.Drawing.Color]::Maroon; $btnGenerar.FlatStyle = "Flat"; $btnGenerar.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $FormTiempo.Controls.AddRange(@($lblT1, $mtbT1, $lblT2, $mtbT2, $lblTF, $dtpFechaRAW, $btnGenerar))
    if ($FormTiempo.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    # La máscara garantiza el formato dd:dd:dd; aquí solo se validan los VALORES (23/59/59 máx).
    $HoraInicioStr = $mtbT1.Text; $HoraFinStr = $mtbT2.Text
    $tmpH = [datetime]::MinValue; $CultInv = [System.Globalization.CultureInfo]::InvariantCulture
    if (-not $mtbT1.MaskCompleted -or -not $mtbT2.MaskCompleted -or
        -not [datetime]::TryParseExact($HoraInicioStr,"HH:mm:ss",$CultInv,[System.Globalization.DateTimeStyles]::None,[ref]$tmpH) -or
        -not [datetime]::TryParseExact($HoraFinStr,"HH:mm:ss",$CultInv,[System.Globalization.DateTimeStyles]::None,[ref]$tmpH)) {
        [System.Windows.Forms.MessageBox]::Show("Hora inválida. Horas 00-23, minutos y segundos 00-59.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); return
    }
    if ($HoraInicioStr -gt $HoraFinStr) { [System.Windows.Forms.MessageBox]::Show("La Hora Inicio ($HoraInicioStr) es mayor que la Hora Fin ($HoraFinStr). Corrige el rango.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); return }

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
        # La fecha viene del picker DEL DIÁLOGO (visible), ya no del principal escondido
        $F1 = $dtpFechaRAW.Value.ToString("MM/dd/yyyy"); $F3 = $dtpFechaRAW.Value.ToString("yyyy-MM-dd"); $F4 = $dtpFechaRAW.Value.ToString("M/d/yyyy")
        $FechaOmni = "(?:$([regex]::Escape($F1))|$([regex]::Escape($F3))|$([regex]::Escape($F4)))"
        $Script:LineaID = 0
        $Script:ArchivosRAW = 0; $Script:LineasFechaRAW = 0   # diagnóstico para el mensaje de "sin actividad"

        function Extraer-Lineas {
            param([string]$Filtro, [string]$NombreColumna, [switch]$EsOneX)
            $Archivos = $null
            if ($EsOneX) { $Archivos = Get-ChildItem -Path $Script:DirFinalGlobal -EA SilentlyContinue | Where-Object { (-not $_.PSIsContainer) -and ($_.Name -match "(?i)one-?x.*\.log" -or $_.Name -match "(?i)one-?x.*\.txt") } | Sort-Object LastWriteTime }
            else { $Archivos = Get-ChildItem -Path $Script:DirFinalGlobal -Filter $Filtro -EA SilentlyContinue | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime }
            if (-not $Archivos) { return }
            foreach ($Archivo in $Archivos) {
                $Lineas = Get-Content -Path $Archivo.FullName -Encoding UTF8 -ReadCount 0 -EA SilentlyContinue
                if (-not $Lineas) { continue }
                $Script:ArchivosRAW++
                $UltimaHoraVista = ""
                foreach ($linea in $Lineas) {
                    if ($linea -match "^\[?(?:$FechaOmni).*?(\d{2}:\d{2}:\d{2})(?:[.,:](\d{1,3}))?") { $ms = if ($matches[2]) { $matches[2].PadRight(3,'0') } else { "000" }; $UltimaHoraVista = "$($matches[1]).$ms"; $Script:LineasFechaRAW++ }
                    elseif ($linea -match "^\[?\d{1,2}/\d{1,2}/\d{4}|^\d{4}-\d{2}-\d{2}") { $UltimaHoraVista = "" }
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
        $btnExportarCSV.Add_Click({ $sd = New-Object System.Windows.Forms.SaveFileDialog; $sd.Filter = "Archivo CSV (*.csv)|*.csv"; $sd.FileName = "LogsAvaya-$(if($txtIP.Text.Trim() -ne ''){$txtIP.Text.Trim()}else{'SinIP'})_$($HoraInicioStr.Replace(':',''))_a_$($HoraFinStr.Replace(':',''))_$($dtpFecha.Value.ToString('ddMMyyyy')).csv"; if ($sd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $CsvData=@(); foreach($row in $GridRaw.Rows){$CsvData+=[PSCustomObject]@{Hora=$row.Cells["Hora"].Value;EndpointLog=$row.Cells["Endpoint"].Value;AvayaOneXLog=$row.Cells["OneX"].Value;AudioLog=$row.Cells["Audio"].Value;IspeacLog=$row.Cells["Ispeac"].Value}}; $CsvData|Export-Csv -Path $sd.FileName -NoTypeInformation -Encoding UTF8; [System.Windows.Forms.MessageBox]::Show("¡Exportación exitosa!","Éxito",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) } })

        $lblBuscarRAW = New-Object System.Windows.Forms.Label; $lblBuscarRAW.Text = "Buscar en Logs:"; $lblBuscarRAW.Location = New-Object System.Drawing.Point(300, 718); $lblBuscarRAW.AutoSize = $true; $lblBuscarRAW.ForeColor = [System.Drawing.Color]::White
        $txtBuscarRAW = New-Object System.Windows.Forms.TextBox; $txtBuscarRAW.Location = New-Object System.Drawing.Point(400, 715); $txtBuscarRAW.Size = New-Object System.Drawing.Size(200, 25); $txtBuscarRAW.BackColor = [System.Drawing.Color]::FromArgb(45,45,48); $txtBuscarRAW.ForeColor = [System.Drawing.Color]::White
        $btnBuscarRAW = New-Object System.Windows.Forms.Button; $btnBuscarRAW.Text = "FILTRAR"; $btnBuscarRAW.Location = New-Object System.Drawing.Point(610, 713); $btnBuscarRAW.Size = New-Object System.Drawing.Size(100, 30); $btnBuscarRAW.BackColor = [System.Drawing.Color]::Teal; $btnBuscarRAW.ForeColor = [System.Drawing.Color]::White; $btnBuscarRAW.FlatStyle = "Flat"
        $btnLimpiarRAW = New-Object System.Windows.Forms.Button; $btnLimpiarRAW.Text = "LIMPIAR"; $btnLimpiarRAW.Location = New-Object System.Drawing.Point(720, 713); $btnLimpiarRAW.Size = New-Object System.Drawing.Size(100, 30); $btnLimpiarRAW.BackColor = [System.Drawing.Color]::Gray; $btnLimpiarRAW.ForeColor = [System.Drawing.Color]::White; $btnLimpiarRAW.FlatStyle = "Flat"
        $btnBuscarRAW.Add_Click({ $termino=$txtBuscarRAW.Text.Trim(); if($termino -eq ""){return}; $FormRaw.Cursor=[System.Windows.Forms.Cursors]::WaitCursor; $GridRaw.SuspendLayout(); $GridRaw.CurrentCell=$null; foreach($row in $GridRaw.Rows){$match=$false;foreach($cell in $row.Cells){if($cell.Value -and $cell.Value.ToString() -match [regex]::Escape($termino)){$match=$true;break}};$row.Visible=$match}; $GridRaw.ResumeLayout(); $FormRaw.Cursor=[System.Windows.Forms.Cursors]::Default })
        $btnLimpiarRAW.Add_Click({ $txtBuscarRAW.Text=""; $FormRaw.Cursor=[System.Windows.Forms.Cursors]::WaitCursor; $GridRaw.SuspendLayout(); $GridRaw.CurrentCell=$null; foreach($row in $GridRaw.Rows){$row.Visible=$true}; $GridRaw.ResumeLayout(); $FormRaw.Cursor=[System.Windows.Forms.Cursors]::Default })

        $FormRaw.Controls.AddRange(@($GridRaw,$btnExportarCSV,$lblBuscarRAW,$txtBuscarRAW,$btnBuscarRAW,$btnLimpiarRAW))
        $lblStatus.Text = "Aislamiento completado. Visualizando $($LogsOrdenados.Count) líneas crudas entrelazadas."
        if ($LogsOrdenados.Count -eq 0) {
            # Diagnóstico: decir POR QUÉ no hubo resultados en vez del mensaje ciego.
            $Pista = if ($Script:ArchivosRAW -eq 0) {
                "CAUSA: no se encontraron archivos de log en esa carpeta.`nVerifica que la ruta apunte a la carpeta 'Log Files'."
            } elseif ($Script:LineasFechaRAW -eq 0) {
                "CAUSA: NINGUNA línea de los logs es del día $($dtpFechaRAW.Value.ToString('dd/MM/yyyy')).`nCorrige la fecha en el diálogo de extracción e intenta de nuevo."
            } else {
                "CAUSA: hay $($Script:LineasFechaRAW) líneas de ese día, pero ninguna dentro del rango de horas $HoraInicioStr - $HoraFinStr.`nAmplía el rango de horas."
            }
            $MsgDiag = "No se encontró actividad en ese rango.`n`n--- DIAGNÓSTICO ---`nCarpeta: $($Script:DirFinalGlobal)`nFecha buscada: $($dtpFechaRAW.Value.ToString('dd/MM/yyyy'))  (en el log: $F1)`nArchivos de log leídos: $($Script:ArchivosRAW)`nLíneas de esa fecha: $($Script:LineasFechaRAW)`nRango de horas: $HoraInicioStr a $HoraFinStr`n`n$Pista"
            [System.Windows.Forms.MessageBox]::Show($MsgDiag, "Grid Vacío", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
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
    # Checkbox para ignorar el filtro de fecha y buscar en TODO lo que exista en los archivos (útil cuando
    # no se sabe de qué día viene un evento/número, ya que los logs rotan por tamaño y un solo archivo
    # puede abarcar varios días). Al activarse, se deshabilita el selector de fecha (ya no aplica). (Pablo, 08/2026.)
    # Fila 2 (y=50): estatus + checkbox + exportar, uno junto al otro (antes el checkbox vivía en la fila 1,
    # de 2 líneas de alto, y se traslapaba con el botón EXPORTAR CSV de abajo). Pablo, 18/08/2026.
    $lblEstatusB = New-Object System.Windows.Forms.Label; $lblEstatusB.Location = New-Object System.Drawing.Point(20,52); $lblEstatusB.Size = New-Object System.Drawing.Size(520,20); $lblEstatusB.ForeColor = [System.Drawing.Color]::Yellow; $lblEstatusB.Text = "Ejemplo: LogoutRequest, AgentState, closeSignalingChannel"
    $chkTodasFechas = New-Object System.Windows.Forms.CheckBox; $chkTodasFechas.Text = "Buscar en todos los días disponibles"; $chkTodasFechas.Location = New-Object System.Drawing.Point(550,51); $chkTodasFechas.Size = New-Object System.Drawing.Size(300,20); $chkTodasFechas.ForeColor = [System.Drawing.Color]::LimeGreen; $chkTodasFechas.Font = New-Object System.Drawing.Font("Segoe UI",8.5)
    $chkTodasFechas.Add_CheckedChanged({ $dtpF.Enabled = -not $chkTodasFechas.Checked }.GetNewClosure())
    $btnExportarBusq = New-Object System.Windows.Forms.Button; $btnExportarBusq.Text = "EXPORTAR CSV"; $btnExportarBusq.Location = New-Object System.Drawing.Point(860,48); $btnExportarBusq.Size = New-Object System.Drawing.Size(140,24); $btnExportarBusq.BackColor = [System.Drawing.Color]::DarkSlateBlue; $btnExportarBusq.FlatStyle = "Flat"; $btnExportarBusq.Font = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Bold)
    $GridB = New-Object System.Windows.Forms.DataGridView; $GridB.Location = New-Object System.Drawing.Point(20,75); $GridB.Size = New-Object System.Drawing.Size(1140,470)
    $GridB.BackgroundColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridB.AllowUserToAddRows = $false; $GridB.RowHeadersVisible = $false; $GridB.ReadOnly = $true; $GridB.AutoSizeColumnsMode = "Fill"
    $GridB.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(20,20,20); $GridB.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(10,10,10); $GridB.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White; $GridB.EnableHeadersVisualStyles = $false
    $GridB.Columns.Add("Fecha","Fecha") | Out-Null; $GridB.Columns["Fecha"].FillWeight = 10
    $GridB.Columns.Add("Hora","Hora") | Out-Null; $GridB.Columns["Hora"].FillWeight = 10
    $GridB.Columns.Add("Archivo","Archivo Origen") | Out-Null; $GridB.Columns["Archivo"].FillWeight = 15
    $GridB.Columns.Add("Log","Línea de Log Encontrada") | Out-Null; $GridB.Columns["Log"].FillWeight = 65

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

        # Patrón de fecha para filtrar solo líneas del día seleccionado (igual que $FechaOmni en el análisis principal)
        # Solo formatos MM/dd (gringo) + ISO. Se excluyen dd/MM y d/M para evitar colisiones entre fechas.
        $BF1 = $dtpF.Value.ToString("MM/dd/yyyy"); $BF3 = $dtpF.Value.ToString("yyyy-MM-dd"); $BF4 = $dtpF.Value.ToString("M/d/yyyy")
        $FechaOmniBusq = "(?:$([regex]::Escape($BF1))|$([regex]::Escape($BF3))|$([regex]::Escape($BF4)))"
        $TieneFechaRegex = "\d{1,2}[/\\]\d{1,2}[/\\]\d{2,4}|\d{4}-\d{2}-\d{2}"
        $BuscarTodasFechas = $chkTodasFechas.Checked
        $FormatosFechaBusq = @("MM/dd/yyyy","yyyy-MM-dd","M/d/yyyy")

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
                $UltimaFechaVista = "---"
                $UltimaFechaOrden = "00000000"
                $EsFechaCorrecta = $BuscarTodasFechas
                foreach ($bloque in $Lineas) { foreach ($linea in $bloque) {
                    # Actualizar flag de fecha cuando la línea trae timestamp de fecha. Si "Buscar en todos
                    # los días" está activo, no se descarta nada por fecha, pero igual se registra la fecha
                    # vista para poder mostrar en qué día apareció cada coincidencia.
                    if ($linea -match $TieneFechaRegex) {
                        $textoFecha = $matches[0]
                        $dtParsed = [datetime]::MinValue
                        foreach ($fmtF in $FormatosFechaBusq) {
                            if ([datetime]::TryParseExact($textoFecha, $fmtF, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dtParsed)) {
                                $UltimaFechaVista = $dtParsed.ToString("dd/MM/yyyy"); $UltimaFechaOrden = $dtParsed.ToString("yyyyMMdd"); break
                            }
                        }
                        if (-not $BuscarTodasFechas) { $EsFechaCorrecta = ($linea -match $FechaOmniBusq) }
                    }
                    if (-not $EsFechaCorrecta) { continue }
                    if ($linea -match "(?:^|\[|\s)(\d{2}:\d{2}:\d{2})(?:[.,:]\d{1,3})?") { $UltimaHoraVista = $matches[1] }
                    if ($linea -match $RegexBusqueda) { [void]$ResultadosLocales.Add([PSCustomObject]@{Hora=$UltimaHoraVista;Fecha=$UltimaFechaVista;FechaOrden=$UltimaFechaOrden;Archivo=$NombreFuente;Texto=$linea}) }
                }}
            }
            return $ResultadosLocales
        }
        try {
            $ResultadosRAW = [System.Collections.ArrayList]::new()
            # BUG FIX: Buscar-En-Archivos devuelve ArrayList, pero PowerShell lo desempaca
            # cuando hay exactamente 1 resultado → PSCustomObject suelto → AddRange() falla.
            # @() fuerza que siempre sea array, y .Count evita llamar AddRange() con array vacío.
            $R1=@(Buscar-En-Archivos -Filtro "EndpointLog.txt*" -NombreFuente "Endpoint.log"); if($R1.Count -gt 0){$ResultadosRAW.AddRange($R1)}
            $R2=@(Buscar-En-Archivos -NombreFuente "AvayaOneX.log" -EsOneX); if($R2.Count -gt 0){$ResultadosRAW.AddRange($R2)}
            $R3=@(Buscar-En-Archivos -Filtro "AudioLog.txt*" -NombreFuente "Audio.log"); if($R3.Count -gt 0){$ResultadosRAW.AddRange($R3)}
            $R4=@(Buscar-En-Archivos -Filtro "IspeacLog.txt*" -NombreFuente "Ispeac.log"); if($R4.Count -gt 0){$ResultadosRAW.AddRange($R4)}
            $ResultadosOrdenados = $ResultadosRAW | Sort-Object FechaOrden,Hora
            $GridB.SuspendLayout()
            foreach ($item in $ResultadosOrdenados) {
                $r=$GridB.Rows.Add(); $GridB.Rows[$r].Cells["Fecha"].Value=$item.Fecha; $GridB.Rows[$r].Cells["Hora"].Value=$item.Hora; $GridB.Rows[$r].Cells["Archivo"].Value=$item.Archivo; $GridB.Rows[$r].Cells["Log"].Value=$item.Texto
                switch ($item.Archivo) {"Endpoint.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::Cyan} "AvayaOneX.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::MediumOrchid} "Audio.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::LimeGreen} "Ispeac.log"{$GridB.Rows[$r].Cells["Archivo"].Style.ForeColor=[System.Drawing.Color]::Orange}}
            }
            $GridB.ResumeLayout()
            $lblEstatusB.Text = "Búsqueda completada. $($ResultadosOrdenados.Count) coincidencias."
        } catch { $lblEstatusB.Text = "Error: $($_.Exception.Message)" }
        finally { if($Script:RutaManual -eq ""){if(Get-PSDrive -Name $Script:DriveName -EA SilentlyContinue){Remove-PSDrive -Name $Script:DriveName -Force -EA SilentlyContinue|Out-Null}}; $FormB.Cursor=[System.Windows.Forms.Cursors]::Default }
    })

    # Exportar los resultados actuales de la búsqueda a CSV (mismo patrón que "8. EXPORTAR" del grid principal).
    $btnExportarBusq.Add_Click({
        if ($GridB.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No hay resultados para exportar. Realiza una búsqueda primero.","Aviso",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning); return }
        $sfdB = New-Object System.Windows.Forms.SaveFileDialog
        $sfdB.Filter = "Archivo CSV (*.csv)|*.csv"
        $sfdB.Title  = "Exportar resultados de búsqueda a CSV"
        $sfdB.FileName = "Busqueda_$($txtP.Text.Trim() -replace '[\\/:*?""<>|,]','_')_$(Get-Date -Format 'yyyy-MM-dd_HHmmss')"
        if ($sfdB.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            $sbB = New-Object System.Text.StringBuilder
            $colsVisiblesB = @($GridB.Columns | Where-Object { $_.Visible })
            $headersB = $colsVisiblesB | ForEach-Object { '"' + $_.HeaderText.Replace('"','""') + '"' }
            [void]$sbB.AppendLine($headersB -join ",")
            foreach ($rowB in $GridB.Rows) {
                if ($rowB.IsNewRow) { continue }
                $valuesB = $colsVisiblesB | ForEach-Object {
                    $vB = if ($null -ne $rowB.Cells[$_.Name].Value) { $rowB.Cells[$_.Name].Value.ToString() } else { "" }
                    '"' + $vB.Replace('"','""') + '"'
                }
                [void]$sbB.AppendLine($valuesB -join ",")
            }
            $encB = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllText($sfdB.FileName, $sbB.ToString(), $encB)
            [System.Windows.Forms.MessageBox]::Show("CSV exportado correctamente:`n$($sfdB.FileName)","Exportación exitosa",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al exportar: $($_.Exception.Message)","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $FormB.Controls.AddRange(@($lblP,$txtP,$lblF,$dtpF,$chkTodasFechas,$btnEjecutar,$lblEstatusB,$btnExportarBusq,$GridB))
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

