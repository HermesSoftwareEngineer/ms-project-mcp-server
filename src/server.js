#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { callBridge } from "./bridge.js";

const server = new McpServer({ name: "ms-project", version: "1.0.0" });

const project = z.string().optional().describe("Nome ou caminho do projeto aberto. Padrão: projeto ativo.");
const uid = z.number().int().describe("UniqueID da tarefa (campo 'uid' retornado por list_tasks).");

const taskFields = {
  name: z.string().describe("Nome da tarefa."),
  outlineLevel: z.number().int().min(1).optional().describe("Nível na estrutura de tópicos (1 = topo)."),
  start: z.string().optional().describe("Início, ISO 8601 (ex.: 2026-11-01T08:00:00)."),
  finish: z.string().optional().describe("Término, ISO 8601."),
  durationDays: z.number().optional().describe("Duração em dias úteis do projeto."),
  durationHours: z.number().optional().describe("Duração em horas."),
  milestone: z.boolean().optional().describe("true = marco (duração zero)."),
  percentComplete: z.number().int().min(0).max(100).optional(),
  predecessors: z.string().optional().describe("Predecessoras por ID de linha, ex.: '3FS+2d;5SS'."),
  notes: z.string().optional(),
};

function tool(name, description, shape, op, annotations = {}) {
  server.registerTool(
    name,
    { description, inputSchema: shape, annotations },
    async (args) => {
      try {
        const data = await callBridge(op, args);
        return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
      } catch (e) {
        return { isError: true, content: [{ type: "text", text: `Erro: ${e.message}` }] };
      }
    }
  );
}

const RO = { readOnlyHint: true };

tool("project_status", "Mostra se o MS Project está em execução, a versão e os projetos abertos.", {}, "status", RO);
tool("project_launch", "Inicia o MS Project (ou conecta ao já aberto).", {}, "launch");
tool("project_open", "Abre um arquivo .mpp no MS Project.",
  { path: z.string().describe("Caminho completo do arquivo."), readOnly: z.boolean().optional() }, "open_project");
tool("project_new", "Cria um novo projeto em branco.",
  { start: z.string().optional().describe("Data de início do projeto, ISO 8601.") }, "new_project");
tool("project_save", "Salva o projeto (ou salva como, se 'path' for informado).",
  { project, path: z.string().optional().describe("Novo caminho para Salvar Como.") }, "save_project");
tool("project_info", "Resumo do projeto: datas, contagem de tarefas/recursos, horas por dia.", { project }, "project_info", RO);

tool("task_list", "Lista tarefas com filtros opcionais (nome, incompletas, críticas, resumo).",
  {
    project,
    nameContains: z.string().optional(),
    onlyIncomplete: z.boolean().optional(),
    onlyCritical: z.boolean().optional(),
    onlySummary: z.boolean().optional(),
    limit: z.number().int().positive().optional().describe("Máximo de itens retornados (padrão 200)."),
  }, "list_tasks", RO);
tool("task_get", "Detalhes de uma tarefa.", { project, uid }, "get_task", RO);
tool("task_add", "Adiciona uma tarefa ao final do cronograma.", { project, ...taskFields }, "add_task");
tool("task_update", "Atualiza campos de uma tarefa existente.",
  { project, uid, ...taskFields, name: taskFields.name.optional() }, "update_task");
tool("task_delete", "Exclui uma tarefa (destrutivo).", { project, uid }, "delete_task", { destructiveHint: true });

tool("resource_list", "Lista os recursos do projeto.", { project }, "list_resources", RO);
tool("resource_add", "Adiciona um recurso.",
  {
    project,
    name: z.string(),
    type: z.number().int().min(0).max(2).optional().describe("0 = trabalho, 1 = material, 2 = custo."),
    maxUnits: z.number().optional().describe("Unidades máximas (1 = 100%)."),
    standardRate: z.string().optional().describe("Taxa padrão, ex.: '50/h'."),
    email: z.string().optional(),
    initials: z.string().optional(),
  }, "add_resource");
tool("resource_assign", "Atribui um recurso a uma tarefa.",
  {
    project,
    taskUid: uid,
    resource: z.string().describe("Nome ou UniqueID do recurso."),
    units: z.number().optional().describe("Unidades (1 = 100%, padrão)."),
  }, "assign_resource");
tool("assignment_list", "Lista atribuições (todas ou de uma tarefa).",
  { project, taskUid: uid.optional() }, "list_assignments", RO);
tool("baseline_save", "Salva a linha de base do projeto inteiro.", { project }, "save_baseline");

await server.connect(new StdioServerTransport());
