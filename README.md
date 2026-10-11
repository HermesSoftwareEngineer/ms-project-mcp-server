# ms-project-mcp-server

Servidor MCP (stdio) local que permite ao Claude ler e editar o **Microsoft Project desktop** instalado no Windows.

## Por que COM e não Office.js?
O Office.js no Project desktop só existe dentro de um painel de suplemento e expõe uma API mínima (seleção atual,
sem criar/editar tarefas), inacessível a processos externos. A API local completa é a automação COM
(`MSProject.Application`), usada aqui por uma ponte PowerShell (`src/bridge.ps1`, late binding via `CallByName`).

## Requisitos
Windows, MS Project desktop, Node 18+, PowerShell 5.1 (já incluso no Windows).

## Instalação
```bash
npm install
```

## Conectar ao Claude
**Claude Code:**
```bash
claude mcp add ms-project -- node C:\Users\user\projetos_dev\ms-project-mcp-server\src\server.js
```
**Claude Desktop** (`%APPDATA%\Claude\claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "ms-project": {
      "command": "node",
      "args": ["C:\Users\user\projetos_dev\ms-project-mcp-server\src\server.js"]
    }
  }
}
```

## Tools
| Grupo | Tools |
|---|---|
| Projeto | `project_status`, `project_launch`, `project_open`, `project_new`, `project_save`, `project_info`, `baseline_save` |
| Tarefas | `task_list`, `task_get`, `task_add`, `task_update`, `task_delete` |
| Recursos | `resource_list`, `resource_add`, `resource_assign`, `assignment_list` |

Tarefas são identificadas por `uid` (UniqueID). Predecessoras usam o ID de linha (`"3FS+2d;5SS"`).
Datas em ISO 8601. `MSPROJECT_TIMEOUT_MS` (padrão 60000) controla o timeout de cada chamada.

## Observações
- As chamadas são serializadas (COM do Project é single-thread).
- Se o Project exibir uma caixa de diálogo modal, a chamada falha por timeout; feche o diálogo.
- Alterações ficam na sessão do Project; use `project_save` para gravar.
