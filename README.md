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

## Tools (projetadas para poucas chamadas)
| Tool | Para quê |
|---|---|
| `msproject_read` | Lê info, tarefas, recursos e atribuições em **uma** chamada; escolha `include`, `taskFields` e `filter` para poupar tokens. |
| `msproject_batch` | Executa uma **lista de operações** em um único processo: `launch`, `open`, `new`, `save`, `baseline`, `add_task`, `update_task`, `delete_task`, `add_resource`, `update_resource`, `assign`. |

Recursos do lote:
- **Apelidos:** `as: "lev"` numa operação; depois `"@lev"` em `uid`, `parent`, `task`, `resource`, e `@{lev}` em `predecessors`.
- **Hierarquia:** `parent: "@fase"` define o nível de estrutura de tópicos (sem `parent`/`outlineLevel`, a tarefa fica no nível 1).
- **Recursos inline:** `resources: ["Ana", {resource:"Bia", units:0.5}]` atribui e cria recursos que não existem.
- **Predecessoras:** `"3FS+2d;@{lev}SS"`. Tipos `FS/SS/FF/SF` (ou `TI/II/TT/IT`), lag em `d`, `h`, `w`, `min`; independente do idioma do Project.
- **Erros:** para no primeiro erro (o que já rodou permanece); `continueOnError: true` segue adiante.
- **Retorno:** resultado por operação + estado final das tarefas tocadas (`returnFields` escolhe os campos).

Exemplo:
```json
{"ops":[
  {"op":"new","start":"2026-11-02"},
  {"name":"Fase 1","as":"f1"},
  {"name":"Levantamento","parent":"@f1","durationDays":5,"resources":["Ana"],"as":"lev"},
  {"name":"Projeto","parent":"@f1","durationDays":3,"predecessors":"@{lev}"},
  {"name":"Entrega","milestone":true}
]}
```

Datas em ISO 8601; durações e trabalho em horas. `MSPROJECT_TIMEOUT_MS` (padrão 180000) controla o timeout de cada chamada.

## Observações
- As chamadas são serializadas (COM do Project é single-thread).
- Se o Project exibir uma caixa de diálogo modal, a chamada falha por timeout; feche o diálogo.
- Alterações ficam na sessão do Project; use `project_save` para gravar.
