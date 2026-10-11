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
      "args": ["C:\\Users\\user\\projetos_dev\\ms-project-mcp-server\\src\\server.js"]
    }
  }
}
```
Depois de atualizar o código, reinicie o conector. Cada servidor roda uma cópia congelada da ponte, então
um servidor antigo continua coerente até ser reiniciado.

## Tools (poucas chamadas)
| Tool | Para quê |
|---|---|
| `msproject_read` | Lê info, tarefas, recursos, atribuições e listas (visões, tabelas, filtros, grupos, calendários) em **uma** chamada. `taskFields` aceita qualquer propriedade COM (`Text1`, `Flag3`, `Number2`...). |
| `msproject_batch` | Executa uma **lista de operações** em um único processo. |

### Operações do lote
| Área | Operações |
|---|---|
| Arquivo | `launch`, `open`, `new`, `save` (opcional `formatId`), `close` (**`save` obrigatório**), `export`, `baseline` |
| Tarefas | `add_task` (`before`, `parent`, `resources`, `props`), `update_task`, `delete_task` |
| Recursos | `add_resource`, `update_resource`, `delete_resource` |
| Atribuições | `assign`, `update_assignment`, `delete_assignment` |
| Qualquer campo | `props` em add/update e a operação `set` (`on: task/resource/project`), inclusive campos personalizados |
| Formatação | `format` (fonte, tamanho, negrito/itálico/sublinhado/tachado, cor, fundo, por linha ou coluna), `apply` (visão/tabela/filtro/grupo), `rename_field` |

Recursos do lote:
- **Apelidos:** `as: "lev"` numa operação; depois `"@lev"` em `uid`, `parent`, `task`, `resource`, `before`, `uids`, e `@{lev}` em `predecessors`.
- **Hierarquia:** `parent: "@fase"` (sem `parent`/`outlineLevel`, a tarefa fica no nível 1).
- **Recursos inline:** `resources: ["Ana", {resource:"Bia", units:0.5}]` atribui e cria os que não existem.
- **Predecessoras:** `"3FS+2d;@{lev}SS"`; tipos `FS/SS/FF/SF` (ou `TI/II/TT/IT`), lag em `d`, `h`, `w`, `min`.
- **Desfazer:** um lote sem `open/new/save/close/export` vira uma única ação de desfazer (Ctrl+Z reverte tudo).
- **Erros:** para no primeiro erro (o que já rodou permanece); `continueOnError: true` segue adiante.
- **Retorno:** resultado por operação + estado final das tarefas tocadas (`returnFields`).

Exemplo:
```json
{"ops":[
  {"op":"new","start":"2026-11-02"},
  {"op":"rename_field","field":"Text1","name":"Codigo"},
  {"name":"Fase 1","as":"f1"},
  {"name":"Levantamento","parent":"@f1","durationDays":5,"resources":["Ana"],"props":{"Text1":"CWP.01"},"as":"lev"},
  {"name":"Projeto","parent":"@f1","durationDays":3,"predecessors":"@{lev}"},
  {"op":"format","uids":["@f1"],"font":{"bold":true,"color":"#C00000"},"fill":"#FFF2CC"},
  {"op":"export","path":"C:/tmp/plano.xml","formatId":"MSProject.xml"}
]}
```

Datas em ISO 8601; durações e trabalho em horas; cores `#RRGGBB`. `MSPROJECT_TIMEOUT_MS` (padrão 180000) controla o timeout.

## Limites conhecidos
- **Bancos de dados:** o Project 2007+ não salva nem exporta para ODBC. Há troca de dados por XML, MPP/MPT, PDF e, com um mapa de exportação existente, xls/csv/txt. Integração com SQL é feita fora do Project.
- **Sem execução arbitrária:** não há operação de script livre nem chamada genérica de métodos COM, de propósito (evita que texto injetado numa conversa execute código na máquina). Novas capacidades entram como operações fixas.
- Exportar xls/csv/txt sem `map` abre um assistente do Project e trava a chamada; por isso é recusado.
- `format` foi validado só quanto à execução sem erro, não visualmente.
- O Project em português espera nomes localizados em `apply` (ex.: `"Entrada"` para a tabela).

## Segurança operacional
- Nunca fecha projeto nem encerra o Project sem `close` com `save` explícito.
- Alterações ficam na sessão do Project; use `save` para gravar no arquivo.
- Se o Project exibir uma caixa de diálogo modal, a chamada falha por timeout; feche o diálogo.
