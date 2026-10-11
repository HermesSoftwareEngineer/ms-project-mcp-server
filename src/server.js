#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { callBridge } from "./bridge.js";

const server = new McpServer({ name: "ms-project", version: "2.0.0" });

const project = z.string().optional().describe("Nome ou caminho de um projeto aberto. Padrão: projeto ativo.");
const ref = z.union([z.number(), z.string()]);

const taskFields = [
  "uid", "id", "wbs", "name", "outlineLevel", "summary", "milestone", "critical", "start", "finish",
  "actualStart", "actualFinish", "deadline", "constraintType", "constraintDate", "durationHours", "workHours",
  "actualWorkHours", "percentComplete", "cost", "priority", "predecessors", "successors", "resources", "notes",
];

const op = z.object({
  op: z
    .enum(["launch", "open", "new", "save", "baseline", "add_task", "update_task", "delete_task",
           "add_resource", "update_resource", "assign"])
    .optional()
    .describe("Omitido: 'update_task' se houver uid, senão 'add_task' se houver name."),
  as: z.string().optional().describe("Apelido para reutilizar o item criado/alterado em operações seguintes."),
  // arquivo / projeto
  path: z.string().optional().describe("open/save: caminho do .mpp (save sem path = salvar no lugar)."),
  readOnly: z.boolean().optional(),
  start: z.string().optional().describe("ISO 8601 (new: início do projeto; tarefas: início)."),
  // tarefa
  uid: ref.optional().describe("update_task/delete_task: UniqueID numérico ou '@apelido'."),
  name: z.string().optional(),
  parent: ref.optional().describe("add/update_task: '@apelido' ou uid da tarefa-resumo pai (define outlineLevel = pai + 1)."),
  outlineLevel: z.number().int().min(1).optional(),
  finish: z.string().optional(),
  durationDays: z.number().optional(),
  durationHours: z.number().optional(),
  milestone: z.boolean().optional(),
  percentComplete: z.number().int().min(0).max(100).optional(),
  predecessors: z.string().optional()
    .describe("IDs de linha ou apelidos: '3FS+2d;@{design}SS'. @{apelido} vira o ID atual da tarefa."),
  notes: z.string().optional(),
  deadline: z.string().optional(),
  priority: z.number().int().optional(),
  resources: z
    .array(z.union([z.string(), z.object({ resource: z.string(), units: z.number().optional() })]))
    .optional()
    .describe("Atribui recursos à tarefa; os inexistentes são criados. Item: 'Nome' ou {resource, units (1=100%)}."),
  // recurso
  resource: ref.optional().describe("assign: nome/uid/'@apelido' do recurso. update_resource: recurso alvo."),
  task: ref.optional().describe("assign: uid ou '@apelido' da tarefa."),
  units: z.number().optional(),
  type: z.number().int().min(0).max(2).optional().describe("Recurso: 0 trabalho, 1 material, 2 custo."),
  maxUnits: z.number().optional(),
  standardRate: z.string().optional().describe("Ex.: '50/h'."),
  email: z.string().optional(),
  initials: z.string().optional(),
  group: z.string().optional(),
});

const text = (data, isError = false) => ({
  isError,
  content: [{ type: "text", text: JSON.stringify(data) }],
});

async function run(bridgeOp, args) {
  try {
    return text(await callBridge(bridgeOp, args));
  } catch (e) {
    return text({ error: e.message }, true);
  }
}

server.registerTool(
  "msproject_read",
  {
    description:
      "Lê o MS Project em UMA chamada: info do projeto, tarefas, recursos e atribuições. " +
      "Peça só o que precisa via 'include' e 'taskFields' para economizar tokens. " +
      `Campos de tarefa: ${taskFields.join(", ")}. Durações/trabalho em horas; datas ISO 8601.`,
    inputSchema: {
      project,
      include: z.array(z.enum(["info", "tasks", "resources", "assignments"])).optional()
        .describe("Padrão: todos."),
      taskFields: z.array(z.string()).optional().describe("Padrão: uid,id,wbs,name,outlineLevel,summary,milestone,start,finish,durationHours,percentComplete,predecessors,resources,critical."),
      resourceFields: z.array(z.string()).optional()
        .describe("uid,id,name,type,maxUnits,standardRate,workHours,cost,initials,group,email."),
      filter: z.object({
        uids: z.array(z.number().int()).optional(),
        nameContains: z.string().optional(),
        onlyIncomplete: z.boolean().optional(),
        onlyCritical: z.boolean().optional(),
        onlySummary: z.boolean().optional(),
        maxOutlineLevel: z.number().int().optional(),
      }).optional(),
      limit: z.number().int().positive().optional().describe("Máx. de tarefas (padrão 500)."),
    },
    annotations: { readOnlyHint: true },
  },
  (args) => run("read", args)
);

server.registerTool(
  "msproject_batch",
  {
    description:
      "Executa VÁRIAS ações no MS Project em UMA chamada (um único processo, sem redesenho de tela durante o lote). " +
      "Monte o cronograma inteiro de uma vez: cada item de 'ops' é uma operação, executada em ordem. " +
      "Operações: launch, open{path}, new{start}, save{path?}, baseline, add_task, update_task{uid}, delete_task{uid}, " +
      "add_resource{name,...}, update_resource{resource,...}, assign{task,resource,units}. " +
      "Use 'as' para dar apelido a uma tarefa/recurso e '@apelido' para referenciá-lo depois (em uid, parent, task, resource; " +
      "em predecessors use '@{apelido}'). Em add_task, 'parent' cria a hierarquia e 'resources' atribui (e cria) recursos já na criação. " +
      "Exemplo: [{op:'new',start:'2026-11-02'},{name:'Fase 1',as:'f1'},{name:'Levantamento',parent:'@f1',durationDays:5,resources:['Ana'],as:'lev'}," +
      "{name:'Projeto',parent:'@f1',durationDays:3,predecessors:'@{lev}',resources:[{resource:'Bia',units:0.5}]}]. " +
      "Em caso de erro: o padrão é parar no 1º erro (as ações anteriores permanecem aplicadas); continueOnError:true segue adiante. " +
      "A resposta traz o resultado por operação e o estado final (recalculado) das tarefas tocadas.",
    inputSchema: {
      project,
      ops: z.array(op).min(1).describe("Operações em ordem."),
      continueOnError: z.boolean().optional(),
      returnTasks: z.boolean().optional().describe("Padrão true: devolve as tarefas tocadas após o recálculo."),
      returnFields: z.array(z.string()).optional().describe("Campos devolvidos nas tarefas (mesmos do msproject_read)."),
    },
  },
  (args) => run("batch", args)
);

await server.connect(new StdioServerTransport());
