import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const SCRIPT = path.join(path.dirname(fileURLToPath(import.meta.url)), "bridge.ps1");
const TIMEOUT_MS = Number(process.env.MSPROJECT_TIMEOUT_MS ?? 180000);

let queue = Promise.resolve();

function run(op, args) {
  return new Promise((resolve, reject) => {
    const ps = spawn(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", SCRIPT],
      { windowsHide: true }
    );
    let out = "";
    let err = "";
    const timer = setTimeout(() => {
      ps.kill();
      reject(new Error(`Timeout (${TIMEOUT_MS}ms). O MS Project pode estar com uma caixa de diálogo aberta.`));
    }, TIMEOUT_MS);
    ps.stdout.on("data", (d) => (out += d));
    ps.stderr.on("data", (d) => (err += d));
    ps.on("error", (e) => { clearTimeout(timer); reject(e); });
    ps.on("close", () => {
      clearTimeout(timer);
      try {
        const res = JSON.parse(Buffer.from(out.trim(), "base64").toString("utf8"));
        res.ok ? resolve(res.data) : reject(new Error(res.error));
      } catch {
        reject(new Error(`Resposta inválida da ponte PowerShell. ${err.trim()}`));
      }
    });
    ps.stdin.end(Buffer.from(JSON.stringify({ op, args }), "utf8").toString("base64"));
  });
}

// COM do Project é single-thread: serializa as chamadas.
export function callBridge(op, args = {}) {
  const p = queue.then(() => run(op, args));
  queue = p.catch(() => {});
  return p;
}
