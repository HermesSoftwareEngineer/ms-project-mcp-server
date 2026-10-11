# Ponte COM -> MS Project. Entrada (stdin): base64(JSON {op, args}). Saída (stdout): base64(JSON {ok, data|error}).
# Usa late binding (CallByName): o adaptador COM do PowerShell não resolve corretamente a type library do Project.
# Ops: read | batch
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
$utf8 = New-Object System.Text.UTF8Encoding($false)
$CT = [Microsoft.VisualBasic.CallType]

# ---------- COM helpers (o retorno com vírgula evita que o PowerShell enumere coleções COM) ----------
function G($o, [string]$n, [object[]]$a = @()) { $r = [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CT::Get, $a); return ,$r }
function S($o, [string]$n, $v) { [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CT::Let, @($v)) | Out-Null }
function M($o, [string]$n, [object[]]$a = @()) { $r = [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CT::Method, $a); return ,$r }
function Items($col) {
    $n = G $col 'Count'
    for ($i = 1; $i -le $n; $i++) { G $col 'Item' @($i) }
}

function Write-Result($obj) {
    $json = ConvertTo-Json -InputObject $obj -Depth 10 -Compress
    [Console]::Out.Write([Convert]::ToBase64String($utf8.GetBytes($json)))
}

function Fmt-Date($v) {
    if ($v -is [datetime]) { return $v.ToString('s') }
    return [string]$v   # "NA" quando não definido
}
function Hours($minutes) { [math]::Round(([double]$minutes) / 60, 2) }
function Conv($v, $kind) {
    switch ($kind) { 'd' { Fmt-Date $v } 'h' { Hours $v } 'b' { [bool]$v } default { $v } }
}

# ---------- mapas de campos (nome exposto -> propriedade COM, tipo) ----------
$TaskFields = @{
    uid = @('UniqueID', 'n'); id = @('ID', 'n'); wbs = @('WBS', 's'); name = @('Name', 's')
    outlineLevel = @('OutlineLevel', 'n'); summary = @('Summary', 'b'); milestone = @('Milestone', 'b'); critical = @('Critical', 'b')
    start = @('Start', 'd'); finish = @('Finish', 'd'); actualStart = @('ActualStart', 'd'); actualFinish = @('ActualFinish', 'd')
    deadline = @('Deadline', 'd'); constraintType = @('ConstraintType', 'n'); constraintDate = @('ConstraintDate', 'd')
    durationHours = @('Duration', 'h'); workHours = @('Work', 'h'); actualWorkHours = @('ActualWork', 'h')
    percentComplete = @('PercentComplete', 'n'); cost = @('Cost', 'n'); priority = @('Priority', 'n')
    predecessors = @('Predecessors', 's'); successors = @('Successors', 's'); resources = @('ResourceNames', 's'); notes = @('Notes', 's')
}
$DefaultTaskFields = @('uid', 'id', 'wbs', 'name', 'outlineLevel', 'summary', 'milestone', 'start', 'finish', 'durationHours', 'percentComplete', 'predecessors', 'resources', 'critical')
$ResFields = @{
    uid = @('UniqueID', 'n'); id = @('ID', 'n'); name = @('Name', 's'); type = @('Type', 'n')
    maxUnits = @('MaxUnits', 'n'); standardRate = @('StandardRate', 's'); workHours = @('Work', 'h'); cost = @('Cost', 'n')
    initials = @('Initials', 's'); group = @('Group', 's'); email = @('EMailAddress', 's')
}
$DefaultResFields = @('uid', 'id', 'name', 'type', 'maxUnits', 'standardRate', 'workHours', 'cost')

function Read-Fields($o, $map, $fields) {
    $r = [ordered]@{}
    foreach ($f in $fields) {
        if ($map.ContainsKey($f)) { $m = $map[$f]; $r[$f] = Conv (G $o $m[0]) $m[1] }
        else { try { $r[$f] = Describe-Value (G $o $f) } catch { $r[$f] = $null } }   # qualquer propriedade COM (Text1, Flag1, ...)
    }
    return $r
}

# ---------- app / projeto ----------
function Get-App([bool]$launch) {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application') } catch {}
    if (-not $launch) { throw 'MS Project não está em execução. Inclua uma operação launch/open/new no lote para iniciá-lo.' }
    # Processo independente: uma instância criada via New-Object morre junto com a ponte.
    $exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINPROJ.EXE' -ErrorAction SilentlyContinue).'(default)'
    if (-not $exe) { throw 'WINPROJ.EXE não encontrado. O MS Project está instalado?' }
    Start-Process -FilePath $exe | Out-Null
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep -Milliseconds 500
        try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application'); S $a 'Visible' $true; return $a } catch {}
    }
    throw 'Timeout aguardando o MS Project iniciar.'
}

function Get-Proj($app, $name) {
    $projs = G $app 'Projects'
    if ((G $projs 'Count') -eq 0) { throw 'Nenhum projeto aberto no MS Project.' }
    if ($name) {
        foreach ($p in (Items $projs)) {
            if ((G $p 'Name') -eq $name -or (G $p 'FullName') -eq $name) { return $p }
        }
        throw "Projeto '$name' não está aberto."
    }
    return G $app 'ActiveProject'
}

# ---------- estado do lote ----------
$script:Alias = @{}
$script:ResCache = $null
$script:Proj = $null
$script:ProjName = $null
$script:Touched = [ordered]@{}
$script:Deleted = @{}

function Cur-Proj($app) {
    if ($null -eq $script:Proj) { $script:Proj = Get-Proj $app $script:ProjName }
    return $script:Proj
}

function Resolve-Uid($v) {
    if ($v -is [string] -and $v.StartsWith('@')) {
        $k = $v.Substring(1)
        if (-not $script:Alias.ContainsKey($k)) { throw "Alias '$v' não definido (use 'as' numa operação anterior)." }
        return [int]$script:Alias[$k]
    }
    return [int]$v
}

function Get-Task($proj, $ref) {
    $uid = Resolve-Uid $ref
    $t = $null
    try { $t = G (G $proj 'Tasks') 'UniqueID' @($uid) } catch {}
    if ($null -eq $t) { throw "Tarefa com UniqueID $uid não encontrada." }
    return $t
}

function Expand-Aliases($proj, [string]$s) {
    # @{alias} -> ID de linha atual da tarefa
    return [regex]::Replace($s, '@\{([^}]+)\}', {
        param($m)
        $k = $m.Groups[1].Value
        if (-not $script:Alias.ContainsKey($k)) { throw "Alias '@{$k}' não definido." }
        $t = Get-Task $proj ([int]$script:Alias[$k])
        return [string](G $t 'ID')
    })
}

# Predecessoras independentes do idioma do Project: "3", "3FS+2d", "@{a}SS-1h;5". Tipos FS/SS/FF/SF (ou TI/II/TT/IT).
function Set-Preds($proj, $t, [string]$spec) {
    $spec = Expand-Aliases $proj $spec
    S $t 'Predecessors' ''
    $types = @{ FF = 0; FS = 1; SF = 2; SS = 3; TT = 0; TI = 1; IT = 2; II = 3 }
    $minPerDay = [double](G $proj 'HoursPerDay') * 60
    foreach ($part in ($spec -split '[;,]')) {
        $p = $part.Trim()
        if (-not $p) { continue }
        if ($p -notmatch '^(\d+)\s*(FS|SS|FF|SF|TI|II|TT|IT)?\s*(?:([+-])\s*(\d+(?:[.,]\d+)?)\s*(d|h|w|min)?)?$') {
            throw "Predecessora inválida: '$p'. Formato: ID[TIPO][+-lag[d|h|w|min]], ex.: 3FS+2d"
        }
        $from = G (G $proj 'Tasks') 'Item' @([int]$Matches[1])
        $type = 1; if ($Matches[2]) { $type = $types[$Matches[2].ToUpper()] }
        $lag = 0.0
        if ($Matches[4]) {
            $n = [double]($Matches[4] -replace ',', '.')
            $mult = switch ($Matches[5]) { 'h' { 60 } 'w' { 5 * $minPerDay } 'min' { 1 } default { $minPerDay } }
            $lag = $n * $mult; if ($Matches[3] -eq '-') { $lag = -$lag }
        }
        M (G $t 'TaskDependencies') 'Add' @($from, [int]$type, [int][math]::Round($lag)) | Out-Null
    }
}

function Res-Cache($proj) {
    if ($null -eq $script:ResCache) {
        $script:ResCache = @{}
        foreach ($r in (Items (G $proj 'Resources'))) {
            if ($null -ne $r) { $script:ResCache[(G $r 'Name').ToLower()] = $r; $script:ResCache["#$(G $r 'UniqueID')"] = $r }
        }
    }
    return $script:ResCache
}

function Find-Res($proj, $ref, [bool]$create) {
    $c = Res-Cache $proj
    $key = "$ref"
    if ($key.StartsWith('@')) { $key = "#$(Resolve-Uid $key)" } elseif ($ref -is [int] -or $ref -is [long]) { $key = "#$ref" } else { $key = $key.ToLower() }
    if ($c.ContainsKey($key)) { return $c[$key] }
    if (-not $create -or $key.StartsWith('#')) { throw "Recurso '$ref' não encontrado." }
    $r = M (G $proj 'Resources') 'Add' @([string]$ref)
    $c[$key] = $r; $c["#$(G $r 'UniqueID')"] = $r
    return $r
}

function Assign-Res($proj, $t, $res, $units) {
    $r = Find-Res $proj $res $true
    $rname = G $r 'Name'
    foreach ($asg in (Items (G $t 'Assignments'))) {
        if ((G $asg 'ResourceName') -eq $rname) { S $asg 'Units' ([double]$units); return }
    }
    M (G $t 'Assignments') 'Add' @((G $t 'ID'), (G $r 'ID'), [double]$units) | Out-Null
}

function Apply-Resources($proj, $t, $list) {
    foreach ($x in @($list)) {
        if ($x -is [string]) { Assign-Res $proj $t $x 1 }
        else { $u = 1; if ($null -ne $x.units) { $u = $x.units }; Assign-Res $proj $t $x.resource $u }
    }
}

function Apply-TaskProps($proj, $t, $a) {
    if ($null -ne $a.name) { S $t 'Name' ([string]$a.name) }
    if ($null -ne $a.parent) {
        $p = Get-Task $proj $a.parent
        S $t 'OutlineLevel' ([int](G $p 'OutlineLevel') + 1)
    }
    if ($null -ne $a.outlineLevel) { S $t 'OutlineLevel' ([int]$a.outlineLevel) }
    if ($null -ne $a.start) { S $t 'Start' ([datetime]$a.start) }
    if ($null -ne $a.finish) { S $t 'Finish' ([datetime]$a.finish) }
    if ($null -ne $a.durationDays) { S $t 'Duration' ([double]$a.durationDays * (G $proj 'HoursPerDay') * 60) }
    if ($null -ne $a.durationHours) { S $t 'Duration' ([double]$a.durationHours * 60) }
    if ($a.milestone -eq $true) { S $t 'Duration' 0 }
    if ($null -ne $a.percentComplete) { S $t 'PercentComplete' ([int]$a.percentComplete) }
    if ($null -ne $a.predecessors) { Set-Preds $proj $t ([string]$a.predecessors) }
    if ($null -ne $a.notes) { S $t 'Notes' ([string]$a.notes) }
    if ($null -ne $a.deadline) { S $t 'Deadline' ([datetime]$a.deadline) }
    if ($null -ne $a.priority) { S $t 'Priority' ([int]$a.priority) }
    if ($null -ne $a.resources) { Apply-Resources $proj $t $a.resources }
    Apply-Props $t $a.props
}

function Apply-ResProps($r, $a) {
    if ($null -ne $a.type) { S $r 'Type' ([int]$a.type) }
    if ($null -ne $a.maxUnits) { S $r 'MaxUnits' ([double]$a.maxUnits) }
    if ($null -ne $a.standardRate) { S $r 'StandardRate' ([string]$a.standardRate) }
    if ($null -ne $a.email) { S $r 'EMailAddress' ([string]$a.email) }
    if ($null -ne $a.initials) { S $r 'Initials' ([string]$a.initials) }
    if ($null -ne $a.group) { S $r 'Group' ([string]$a.group) }
    Apply-Props $r $a.props
}

# ---------- acesso genérico ao modelo de objetos ----------
$MS = [Type]::Missing

function Coerce($v) {
    if ($v -is [string] -and $v -match '^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$') { return [datetime]$v }
    return $v
}

function Apply-Props($obj, $props) {
    if ($null -eq $props) { return }
    foreach ($p in $props.PSObject.Properties) { S $obj $p.Name (Coerce $p.Value) }
}

function Describe-Value($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [datetime]) { return $v.ToString('s') }
    if ($v -is [string] -or $v -is [ValueType]) { return $v }
    $d = [ordered]@{ type = 'com' }
    try { $d['name'] = G $v 'Name' } catch {}
    try { $d['count'] = G $v 'Count' } catch {}
    return $d
}

function Invoke-Named($obj, [string]$method, $named) {
    $names = @(); $vals = @()
    foreach ($p in $named.PSObject.Properties) { $names += $p.Name; $vals += (Coerce $p.Value) }
    $r = $obj.GetType().InvokeMember($method, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $vals, $null, $null, $names)
    return ,$r
}

function Do-Apply($app, $o) {
    Cur-Proj $app | Out-Null
    $name = [string]$o.name
    switch ("$($o.kind)") {
        'view' { M $app 'ViewApply' @($name) | Out-Null }
        'table' { M $app 'TableApply' @($name) | Out-Null }
        'filter' { M $app 'FilterApply' @($name) | Out-Null }
        'group' { M $app 'GroupApply' @($name) | Out-Null }
        default { throw "apply: 'kind' deve ser view | table | filter | group" }
    }
    return @{ applied = "$($o.kind)"; name = $name }
}

function Do-RenameField($app, $o) {
    Cur-Proj $app | Out-Null
    $type = 0; if ("$($o.entity)" -eq 'resource') { $type = 1 }   # pjTask=0, pjResource=1
    $const = M $app 'FieldNameToFieldConstant' @([string]$o.field, $type)
    M $app 'CustomFieldRename' @($const, [string]$o.name) | Out-Null
    return @{ field = [string]$o.field; renamedTo = [string]$o.name }
}

function Rgb($hex) {
    if ($null -eq $hex -or "$hex" -eq '') { return $MS }
    if ("$hex" -notmatch '^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$') { throw "Cor inválida '$hex' (use #RRGGBB)." }
    return [int]([Convert]::ToInt32($Matches[1], 16) + 256 * [Convert]::ToInt32($Matches[2], 16) + 65536 * [Convert]::ToInt32($Matches[3], 16))
}
function Opt($v) { if ($null -eq $v) { return $MS } else { return $v } }

function Do-Format($app, $o) {
    $proj = Cur-Proj $app
    $entity = 'task'; if ($o.entity) { $entity = "$($o.entity)" }
    $f = $o.font
    $reset = ($o.reset -eq $true)
    $fa = @(
        (Opt $f.name), (Opt $f.size), (Opt $f.bold), (Opt $f.italic), (Opt $f.underline),
        (Rgb $f.color), $reset, (Rgb $o.fill), (Opt $o.pattern), (Opt $f.strikethrough)
    )
    $rows = @()
    if ($entity -eq 'resource') {
        if ($o.uids) { foreach ($u in @($o.uids)) { $rows += [int](G (Find-Res $proj $u $false) 'ID') } }
        else { foreach ($r in (Items (G $proj 'Resources'))) { if ($null -ne $r) { $rows += [int](G $r 'ID') } } }
        $viewNames = @($o.view, 'Planilha de Recursos', 'Resource Sheet')
    } else {
        if ($o.uids) { foreach ($u in @($o.uids)) { $rows += [int](G (Get-Task $proj $u) 'ID') } }
        else { foreach ($t in (Items (G $proj 'Tasks'))) { if ($null -ne $t) { $rows += [int](G $t 'ID') } } }
        $viewNames = @($o.view, 'Planilha de Tarefas', 'Task Sheet', 'Gráfico de Gantt', 'Gantt Chart')
    }
    $sel = { param($row)
        if ($o.field) {
            if ($entity -eq 'resource') { M $app 'SelectResourceField' @($row, [string]$o.field) | Out-Null }
            else { M $app 'SelectTaskField' @($row, [string]$o.field) | Out-Null }
        } else { M $app 'SelectRow' @($row, $false) | Out-Null }
    }
    $count = 0; $applied = $false
    foreach ($row in $rows) {
        try { & $sel $row } catch {
            if ($applied) { throw }
            foreach ($vn in $viewNames) { if ($vn) { try { M $app 'ViewApply' @([string]$vn) | Out-Null; & $sel $row; break } catch {} } }
            & $sel $row
        }
        $applied = $true
        M $app 'Font32Ex' $fa | Out-Null
        $count++
    }
    return @{ formatted = $count; entity = $entity }
}

function Do-Export($app, $o) {
    $proj = Cur-Proj $app
    if (-not $o.path) { throw "export/save com formato exige 'path'." }
    $fid = "$($o.formatId)"
    if ($fid -match '(?i)xls|csv|txt' -and -not $o.map) {
        throw "Exportar para $fid exige 'map' (nome de um mapa de exportação existente); sem ele o Project abre um assistente e trava."
    }
    $named = [ordered]@{ Name = [string]$o.path }
    if ($fid) { $named['FormatID'] = $fid }
    if ($o.map) { $named['Map'] = [string]$o.map }
    Invoke-Named $app 'FileSaveAs' ([pscustomobject]$named) | Out-Null
    return @{ exported = [string]$o.path; formatId = $fid }
}

# ---------- operações do lote ----------
function Do-BatchOp($app, $o, $op) {
    switch ($op) {
        'launch' { $script:Proj = $null; return @{ version = G $app 'Version' } }
        'open' {
            M $app 'FileOpen' @([string]$o.path, [bool]($o.readOnly -eq $true)) | Out-Null
            $script:Proj = $null; $script:ResCache = $null; $script:ProjName = $null
            return @{ project = G (G $app 'ActiveProject') 'Name' }
        }
        'new' {
            M $app 'FileNew' | Out-Null
            $script:Proj = $null; $script:ResCache = $null; $script:ProjName = $null
            $proj = Cur-Proj $app
            if ($o.start) {
                $d = [datetime]$o.start
                if ($d.TimeOfDay.TotalSeconds -eq 0) { $d = $d.AddHours(8) }   # início do expediente
                S $proj 'ProjectStart' $d
            }
            return @{ project = G $proj 'Name' }
        }
        'save' {
            Cur-Proj $app | Out-Null
            if ($o.formatId) { return Do-Export $app $o }
            if ($o.path) { M $app 'FileSaveAs' @([string]$o.path) | Out-Null } else { M $app 'FileSave' | Out-Null }
            return @{ saved = G (G $app 'ActiveProject') 'FullName' }
        }
        'baseline' { Cur-Proj $app | Out-Null; M $app 'BaselineSave' @($true) | Out-Null; return @{ baseline = 'saved' } }
        'add_task' {
            $proj = Cur-Proj $app
            if ($null -ne $o.before) { $t = M (G $proj 'Tasks') 'Add' @([string]$o.name, [int](G (Get-Task $proj $o.before) 'ID')) }
            else { $t = M (G $proj 'Tasks') 'Add' @([string]$o.name) }
            $uid = [int](G $t 'UniqueID')
            if ($o.as) { $script:Alias[[string]$o.as] = $uid }
            $script:Touched["$uid"] = $true
            if (($null -eq $o.parent) -and ($null -eq $o.outlineLevel)) { S $t 'OutlineLevel' 1 }   # linha nova herda o nível da anterior
            Apply-TaskProps $proj $t $o
            return @{ uid = $uid; id = G $t 'ID'; name = G $t 'Name' }
        }
        'update_task' {
            $proj = Cur-Proj $app
            $t = Get-Task $proj $o.uid
            $uid = [int](G $t 'UniqueID')
            if ($o.as) { $script:Alias[[string]$o.as] = $uid }
            $script:Touched["$uid"] = $true
            Apply-TaskProps $proj $t $o
            return @{ uid = $uid; id = G $t 'ID'; name = G $t 'Name' }
        }
        'delete_task' {
            $proj = Cur-Proj $app
            $t = Get-Task $proj $o.uid
            $uid = [int](G $t 'UniqueID'); $name = G $t 'Name'
            M $t 'Delete' | Out-Null
            $script:Deleted["$uid"] = $true
            return @{ uid = $uid; deleted = $name }
        }
        'add_resource' {
            $proj = Cur-Proj $app
            $r = Find-Res $proj $o.name $true
            Apply-ResProps $r $o
            $uid = [int](G $r 'UniqueID')
            if ($o.as) { $script:Alias[[string]$o.as] = $uid }
            return @{ uid = $uid; name = G $r 'Name' }
        }
        'update_resource' {
            $proj = Cur-Proj $app
            $r = Find-Res $proj $o.resource $false
            Apply-ResProps $r $o
            if ($o.name) { S $r 'Name' ([string]$o.name) }
            return @{ uid = [int](G $r 'UniqueID'); name = G $r 'Name' }
        }
        'assign' {
            $proj = Cur-Proj $app
            $t = Get-Task $proj $o.task
            $u = 1; if ($null -ne $o.units) { $u = $o.units }
            Assign-Res $proj $t $o.resource $u
            $script:Touched["$(G $t 'UniqueID')"] = $true
            return @{ task = G $t 'Name'; resource = "$($o.resource)"; units = $u }
        }
        'export' { return Do-Export $app $o }
        'apply' { return Do-Apply $app $o }
        'rename_field' { return Do-RenameField $app $o }
        'format' { return Do-Format $app $o }
        'set' {
            $on = "$($o.on)"
            if ($on -eq 'task') { $x = Get-Task (Cur-Proj $app) $o.uid; $script:Touched["$(G $x 'UniqueID')"] = $true }
            elseif ($on -eq 'resource') { $x = Find-Res (Cur-Proj $app) $o.resource $false }
            elseif ($on -eq 'project') { $x = Cur-Proj $app }
            else { throw "set: 'on' deve ser task | resource | project" }
            Apply-Props $x $o.props
            return @{ set = $on; props = @($o.props.PSObject.Properties.Name) }
        }
        'delete_resource' {
            $proj = Cur-Proj $app
            $r = Find-Res $proj $o.resource $false
            $name = G $r 'Name'; M $r 'Delete' | Out-Null; $script:ResCache = $null
            return @{ deleted = $name }
        }
        'update_assignment' {
            $proj = Cur-Proj $app
            $t = Get-Task $proj $o.task
            $rn = G (Find-Res $proj $o.resource $false) 'Name'
            foreach ($asg in (Items (G $t 'Assignments'))) {
                if ((G $asg 'ResourceName') -eq $rn) {
                    if ($null -ne $o.units) { S $asg 'Units' ([double]$o.units) }
                    Apply-Props $asg $o.props
                    $script:Touched["$(G $t 'UniqueID')"] = $true
                    return @{ task = G $t 'Name'; resource = $rn }
                }
            }
            throw "Atribuição de '$rn' na tarefa não encontrada."
        }
        'delete_assignment' {
            $proj = Cur-Proj $app
            $t = Get-Task $proj $o.task
            $rn = G (Find-Res $proj $o.resource $false) 'Name'
            foreach ($asg in (Items (G $t 'Assignments'))) {
                if ((G $asg 'ResourceName') -eq $rn) {
                    M $asg 'Delete' | Out-Null
                    $script:Touched["$(G $t 'UniqueID')"] = $true
                    return @{ task = G $t 'Name'; unassigned = $rn }
                }
            }
            throw "Atribuição de '$rn' na tarefa não encontrada."
        }
        'close' {
            if ($null -eq $o.save) { throw "close exige 'save': true (salvar) ou false (descartar alterações). Nunca é implícito." }
            $proj = Cur-Proj $app
            $name = G $proj 'Name'
            M $proj 'Activate' | Out-Null
            $mode = 0; if ($o.save -eq $true) { $mode = 1 }
            M $app 'FileClose' @($mode) | Out-Null
            $script:Proj = $null; $script:ResCache = $null; $script:ProjName = $null
            return @{ closed = $name; saved = ($o.save -eq $true) }
        }
        default { throw "Operação desconhecida: '$op'" }
    }
}

function Run-Batch($a) {
    $script:ProjName = $a.project
    $ops = @($a.ops)
    $needLaunch = $false
    foreach ($o in $ops) { if ($o.op -in @('launch', 'open', 'new')) { $needLaunch = $true } }
    $app = Get-App $needLaunch

    $results = New-Object System.Collections.ArrayList
    $stop = -not ($a.continueOnError -eq $true)
    try { S $app 'ScreenUpdating' $false } catch {}
    $undo = $false
    $structural = $false
    foreach ($o in $ops) { if ($o.op -in @('new', 'open', 'close', 'save', 'export', 'launch')) { $structural = $true } }
    if (-not $structural) { try { M $app 'OpenUndoTransaction' @('Claude MCP (lote)') | Out-Null; $undo = $true } catch {} }   # Ctrl+Z desfaz o lote inteiro
    try {
        for ($i = 0; $i -lt $ops.Count; $i++) {
            $o = $ops[$i]
            $opn = $o.op
            if (-not $opn) { if ($null -ne $o.uid) { $opn = 'update_task' } elseif ($o.name) { $opn = 'add_task' } }
            try {
                $r = Do-BatchOp $app $o $opn
                $row = [ordered]@{ i = $i; op = $opn; ok = $true }
                foreach ($k in $r.Keys) { $row[$k] = $r[$k] }
                if ($o.as) { $row['as'] = $o.as }
                [void]$results.Add($row)
            } catch {
                $m = $_.Exception.Message
                if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }
                [void]$results.Add([ordered]@{ i = $i; op = $opn; ok = $false; error = $m })
                if ($stop) { break }
            }
        }
    } finally {
        if ($undo) { try { M $app 'CloseUndoTransaction' | Out-Null } catch {} }
        try { S $app 'ScreenUpdating' $true } catch {}
    }

    $failed = @($results | Where-Object { -not $_.ok }).Count
    $out = [ordered]@{ ok = ($failed -eq 0); executed = $results.Count; total = $ops.Count; failed = $failed; results = @($results) }
    if ($a.returnTasks -ne $false -and $script:Touched.Count -gt 0) {
        $fields = $DefaultTaskFields
        if ($a.returnFields) { $fields = @($a.returnFields) }
        $tasks = New-Object System.Collections.ArrayList
        try {
            $proj = Cur-Proj $app   # falha se o lote fechou o projeto: então não há estado a devolver
            foreach ($uid in @($script:Touched.Keys)) {
                if ($script:Deleted.ContainsKey($uid)) { continue }
                try { [void]$tasks.Add((Read-Fields (Get-Task $proj $uid) $TaskFields $fields)) } catch {}
            }
        } catch {}
        $out['tasks'] = @($tasks)
    }
    return $out
}

# ---------- leitura ----------
function Run-Read($a) {
    $app = Get-App $false
    $proj = Get-Proj $app $a.project
    $inc = @('info', 'tasks', 'resources', 'assignments')
    if ($a.include) { $inc = @($a.include) }
    $out = [ordered]@{}
    $out['openProjects'] = @(foreach ($p in (Items (G $app 'Projects'))) { G $p 'Name' })

    if ($inc -contains 'info') {
        $out['info'] = [ordered]@{
            name = G $proj 'Name'; fullName = G $proj 'FullName'
            start = Fmt-Date (G $proj 'ProjectStart'); finish = Fmt-Date (G (G $proj 'ProjectSummaryTask') 'Finish')
            statusDate = Fmt-Date (G $proj 'StatusDate')
            taskCount = G (G $proj 'Tasks') 'Count'; resourceCount = G (G $proj 'Resources') 'Count'
            hoursPerDay = G $proj 'HoursPerDay'; author = G $proj 'Author'
        }
        if ($a.projectFields) { foreach ($pf in @($a.projectFields)) { try { $out['info'][$pf] = Describe-Value (G $proj $pf) } catch { $out['info'][$pf] = $null } } }
    }
    foreach ($pair in @(@('views', 'Views'), @('taskTables', 'TaskTables'), @('resourceTables', 'ResourceTables'), @('taskFilters', 'TaskFilters'),
                        @('resourceFilters', 'ResourceFilters'), @('taskGroups', 'TaskGroups'), @('resourceGroups', 'ResourceGroups'), @('calendars', 'BaseCalendars'))) {
        if ($inc -contains $pair[0]) {
            try { $out[$pair[0]] = @(foreach ($x in (Items (G $proj $pair[1]))) { G $x 'Name' }) } catch { $out[$pair[0]] = @() }
        }
    }
    if (($inc -contains 'tasks') -or ($inc -contains 'assignments')) {
        $fields = $DefaultTaskFields; if ($a.taskFields) { $fields = @($a.taskFields) }
        $limit = 500; if ($a.limit) { $limit = [int]$a.limit }
        $f = $a.filter
        $uidSet = $null; if ($f -and $f.uids) { $uidSet = @{}; foreach ($u in @($f.uids)) { $uidSet[[int]$u] = $true } }
        $tasks = New-Object System.Collections.ArrayList
        $asgs = New-Object System.Collections.ArrayList
        $total = 0
        foreach ($t in (Items (G $proj 'Tasks'))) {
            if ($null -eq $t) { continue }
            if ($uidSet -and -not $uidSet.ContainsKey([int](G $t 'UniqueID'))) { continue }
            if ($f) {
                if ($f.nameContains -and (G $t 'Name') -notlike "*$($f.nameContains)*") { continue }
                if ($f.onlyIncomplete -eq $true -and (G $t 'PercentComplete') -ge 100) { continue }
                if ($f.onlyCritical -eq $true -and -not (G $t 'Critical')) { continue }
                if ($f.onlySummary -eq $true -and -not (G $t 'Summary')) { continue }
                if ($null -ne $f.maxOutlineLevel -and (G $t 'OutlineLevel') -gt [int]$f.maxOutlineLevel) { continue }
            }
            $total++
            if ($total -gt $limit) { continue }
            if ($inc -contains 'tasks') { [void]$tasks.Add((Read-Fields $t $TaskFields $fields)) }
            if ($inc -contains 'assignments') {
                foreach ($asg in (Items (G $t 'Assignments'))) {
                    [void]$asgs.Add([ordered]@{
                        taskUid = G $t 'UniqueID'; task = G $t 'Name'; resource = G $asg 'ResourceName'
                        units = G $asg 'Units'; workHours = Hours (G $asg 'Work'); cost = G $asg 'Cost'
                    })
                }
            }
        }
        if ($inc -contains 'tasks') { $out['tasks'] = @($tasks); $out['taskTotalMatched'] = $total }
        if ($inc -contains 'assignments') { $out['assignments'] = @($asgs) }
    }
    if ($inc -contains 'resources') {
        $rf = $DefaultResFields; if ($a.resourceFields) { $rf = @($a.resourceFields) }
        $out['resources'] = @(foreach ($r in (Items (G $proj 'Resources'))) { if ($null -ne $r) { Read-Fields $r $ResFields $rf } })
    }
    return $out
}

try {
    $raw = [Console]::In.ReadToEnd().Trim()
    $req = $utf8.GetString([Convert]::FromBase64String($raw)) | ConvertFrom-Json
    switch ($req.op) {
        'read' { $data = Run-Read $req.args }
        'batch' { $data = Run-Batch $req.args }
        default { throw "Operação desconhecida: $($req.op)" }
    }
    Write-Result @{ ok = $true; data = $data }
} catch {
    $m = $_.Exception.Message
    if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }
    Write-Result @{ ok = $false; error = $m }
}
