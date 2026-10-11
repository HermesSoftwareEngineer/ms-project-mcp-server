# Ponte COM -> MS Project. Entrada (stdin): base64(JSON {op, args}). Saída (stdout): base64(JSON {ok, data|error}).
# Usa late binding (CallByName): o adaptador COM do PowerShell não resolve corretamente a type library do Project.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
$utf8 = New-Object System.Text.UTF8Encoding($false)
$CT = [Microsoft.VisualBasic.CallType]

function G($o, [string]$n, [object[]]$a = @()) { $r = [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CT::Get, $a); return ,$r }
function S($o, [string]$n, $v) { [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CT::Let, @($v)) | Out-Null }
function M($o, [string]$n, [object[]]$a = @()) { $r = [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CT::Method, $a); return ,$r }
function Items($col) {
    $n = G $col 'Count'
    for ($i = 1; $i -le $n; $i++) { G $col 'Item' @($i) }
}

function Write-Result($obj) {
    $json = ConvertTo-Json -InputObject $obj -Depth 8 -Compress
    [Console]::Out.Write([Convert]::ToBase64String($utf8.GetBytes($json)))
}

function Fmt-Date($v) {
    if ($v -is [datetime]) { return $v.ToString('s') }
    return [string]$v   # "NA" quando não definido
}

function Hours($minutes) { [math]::Round(([double]$minutes) / 60, 2) }

function Get-App([bool]$launch) {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application') } catch {}
    if (-not $launch) { throw 'MS Project não está em execução. Abra o MS Project ou use a tool project_launch.' }
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

function Get-Task($proj, $uid) {
    $t = $null
    try { $t = G (G $proj 'Tasks') 'UniqueID' @([int]$uid) } catch {}
    if ($null -eq $t) { throw "Tarefa com UniqueID $uid não encontrada." }
    return $t
}

function Get-Res($proj, $ref) {
    foreach ($r in (Items (G $proj 'Resources'))) {
        if ($null -eq $r) { continue }
        if ("$(G $r 'UniqueID')" -eq "$ref" -or (G $r 'Name') -eq "$ref") { return $r }
    }
    throw "Recurso '$ref' não encontrado."
}

function Task-Obj($t) {
    [ordered]@{
        uid = G $t 'UniqueID'; id = G $t 'ID'; wbs = G $t 'WBS'; name = G $t 'Name'
        outlineLevel = G $t 'OutlineLevel'; summary = [bool](G $t 'Summary'); milestone = [bool](G $t 'Milestone')
        start = Fmt-Date (G $t 'Start'); finish = Fmt-Date (G $t 'Finish')
        durationHours = Hours (G $t 'Duration')
        percentComplete = G $t 'PercentComplete'
        workHours = Hours (G $t 'Work'); cost = G $t 'Cost'
        predecessors = G $t 'Predecessors'; resources = G $t 'ResourceNames'
        critical = [bool](G $t 'Critical'); notes = G $t 'Notes'
    }
}

function Apply-TaskProps($t, $proj, $a) {
    if ($null -ne $a.name) { S $t 'Name' ([string]$a.name) }
    if ($null -ne $a.outlineLevel) { S $t 'OutlineLevel' ([int]$a.outlineLevel) }
    if ($null -ne $a.start) { S $t 'Start' ([datetime]$a.start) }
    if ($null -ne $a.finish) { S $t 'Finish' ([datetime]$a.finish) }
    if ($null -ne $a.durationDays) { S $t 'Duration' ([double]$a.durationDays * (G $proj 'HoursPerDay') * 60) }
    if ($null -ne $a.durationHours) { S $t 'Duration' ([double]$a.durationHours * 60) }
    if ($a.milestone -eq $true) { S $t 'Duration' 0 }
    if ($null -ne $a.percentComplete) { S $t 'PercentComplete' ([int]$a.percentComplete) }
    if ($null -ne $a.predecessors) { S $t 'Predecessors' ([string]$a.predecessors) }
    if ($null -ne $a.notes) { S $t 'Notes' ([string]$a.notes) }
}

function Res-Obj($r) {
    [ordered]@{
        uid = G $r 'UniqueID'; id = G $r 'ID'; name = G $r 'Name'; type = G $r 'Type'
        maxUnits = G $r 'MaxUnits'; standardRate = G $r 'StandardRate'
        workHours = Hours (G $r 'Work'); cost = G $r 'Cost'
        initials = G $r 'Initials'; group = G $r 'Group'; email = G $r 'EMailAddress'
    }
}

function Invoke-Op($op, $a) {
    switch ($op) {
        'launch' {
            $app = Get-App $true
            return @{ version = G $app 'Version'; projects = G (G $app 'Projects') 'Count' }
        }
        'status' {
            $app = Get-App $false
            $projs = G $app 'Projects'
            $list = @(foreach ($p in (Items $projs)) { [ordered]@{ name = G $p 'Name'; fullName = G $p 'FullName' } })
            $active = $null
            if ((G $projs 'Count') -gt 0) { $active = G (G $app 'ActiveProject') 'Name' }
            return @{ version = G $app 'Version'; activeProject = $active; projects = $list }
        }
        'open_project' {
            $app = Get-App $true
            M $app 'FileOpen' @([string]$a.path, [bool]($a.readOnly -eq $true)) | Out-Null
            return @{ activeProject = G (G $app 'ActiveProject') 'Name' }
        }
        'new_project' {
            $app = Get-App $true
            M $app 'FileNew' | Out-Null
            $proj = G $app 'ActiveProject'
            if ($a.start) { S $proj 'ProjectStart' ([datetime]$a.start) }
            return @{ activeProject = G $proj 'Name' }
        }
        'save_project' {
            $app = Get-App $false
            Get-Proj $app $a.project | Out-Null
            if ($a.path) { M $app 'FileSaveAs' @([string]$a.path) | Out-Null } else { M $app 'FileSave' | Out-Null }
            return @{ saved = G (G $app 'ActiveProject') 'FullName' }
        }
        'project_info' {
            $proj = Get-Proj (Get-App $false) $a.project
            return [ordered]@{
                name = G $proj 'Name'; fullName = G $proj 'FullName'
                start = Fmt-Date (G $proj 'ProjectStart'); finish = Fmt-Date (G (G $proj 'ProjectSummaryTask') 'Finish')
                statusDate = Fmt-Date (G $proj 'StatusDate')
                taskCount = G (G $proj 'Tasks') 'Count'; resourceCount = G (G $proj 'Resources') 'Count'
                hoursPerDay = G $proj 'HoursPerDay'; author = G $proj 'Author'
            }
        }
        'list_tasks' {
            $proj = Get-Proj (Get-App $false) $a.project
            $limit = 200; if ($a.limit) { $limit = [int]$a.limit }
            $out = New-Object System.Collections.ArrayList
            $total = 0
            foreach ($t in (Items (G $proj 'Tasks'))) {
                if ($null -eq $t) { continue }
                if ($a.nameContains -and (G $t 'Name') -notlike "*$($a.nameContains)*") { continue }
                if ($a.onlyIncomplete -eq $true -and (G $t 'PercentComplete') -ge 100) { continue }
                if ($a.onlyCritical -eq $true -and -not (G $t 'Critical')) { continue }
                if ($a.onlySummary -eq $true -and -not (G $t 'Summary')) { continue }
                $total++
                if ($out.Count -lt $limit) { [void]$out.Add((Task-Obj $t)) }
            }
            return [ordered]@{ total = $total; returned = $out.Count; tasks = @($out) }
        }
        'get_task' {
            $proj = Get-Proj (Get-App $false) $a.project
            return Task-Obj (Get-Task $proj $a.uid)
        }
        'add_task' {
            $proj = Get-Proj (Get-App $false) $a.project
            $t = M (G $proj 'Tasks') 'Add' @([string]$a.name)
            Apply-TaskProps $t $proj $a
            return Task-Obj $t
        }
        'update_task' {
            $proj = Get-Proj (Get-App $false) $a.project
            $t = Get-Task $proj $a.uid
            Apply-TaskProps $t $proj $a
            return Task-Obj $t
        }
        'delete_task' {
            $proj = Get-Proj (Get-App $false) $a.project
            $t = Get-Task $proj $a.uid
            $name = G $t 'Name'
            M $t 'Delete' | Out-Null
            return @{ deleted = $name }
        }
        'list_resources' {
            $proj = Get-Proj (Get-App $false) $a.project
            $out = @(foreach ($r in (Items (G $proj 'Resources'))) { if ($null -ne $r) { Res-Obj $r } })
            return @{ resources = $out }
        }
        'add_resource' {
            $proj = Get-Proj (Get-App $false) $a.project
            $r = M (G $proj 'Resources') 'Add' @([string]$a.name)
            if ($null -ne $a.type) { S $r 'Type' ([int]$a.type) }
            if ($null -ne $a.maxUnits) { S $r 'MaxUnits' ([double]$a.maxUnits) }
            if ($null -ne $a.standardRate) { S $r 'StandardRate' ([string]$a.standardRate) }
            if ($null -ne $a.email) { S $r 'EMailAddress' ([string]$a.email) }
            if ($null -ne $a.initials) { S $r 'Initials' ([string]$a.initials) }
            return Res-Obj $r
        }
        'assign_resource' {
            $proj = Get-Proj (Get-App $false) $a.project
            $t = Get-Task $proj $a.taskUid
            $r = Get-Res $proj $a.resource
            $units = 1; if ($null -ne $a.units) { $units = [double]$a.units }
            M (G $t 'Assignments') 'Add' @((G $t 'ID'), (G $r 'ID'), $units) | Out-Null
            return Task-Obj $t
        }
        'list_assignments' {
            $proj = Get-Proj (Get-App $false) $a.project
            $out = New-Object System.Collections.ArrayList
            foreach ($t in (Items (G $proj 'Tasks'))) {
                if ($null -eq $t) { continue }
                if ($a.taskUid -and (G $t 'UniqueID') -ne [int]$a.taskUid) { continue }
                foreach ($asg in (Items (G $t 'Assignments'))) {
                    [void]$out.Add([ordered]@{
                        taskUid = G $t 'UniqueID'; taskName = G $t 'Name'; resource = G $asg 'ResourceName'
                        units = G $asg 'Units'; workHours = Hours (G $asg 'Work'); cost = G $asg 'Cost'
                    })
                }
            }
            return @{ assignments = @($out) }
        }
        'save_baseline' {
            $app = Get-App $false
            Get-Proj $app $a.project | Out-Null
            M $app 'BaselineSave' @($true) | Out-Null
            return @{ baseline = 'saved' }
        }
        default { throw "Operação desconhecida: $op" }
    }
}

try {
    $raw = [Console]::In.ReadToEnd().Trim()
    $req = $utf8.GetString([Convert]::FromBase64String($raw)) | ConvertFrom-Json
    $data = Invoke-Op $req.op $req.args
    Write-Result @{ ok = $true; data = $data }
} catch {
    $m = $_.Exception.Message
    if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }
    Write-Result @{ ok = $false; error = $m }
}
