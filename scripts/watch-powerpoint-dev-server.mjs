#!/usr/bin/env node

import { execFile, spawn } from "node:child_process";
import {
  chmodSync,
  createWriteStream,
  existsSync,
  fchmodSync,
  mkdirSync,
  openSync,
  readdirSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const SCRIPT_PATH = fileURLToPath(import.meta.url);
const DEFAULT_REPO_ROOT = resolve(dirname(SCRIPT_PATH), "..");
const DEFAULT_PATH = "/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin";

export function watcherHelp() {
  return `PPTXMate PowerPoint lifecycle watcher

Usage:
  node scripts/watch-powerpoint-dev-server.mjs
  node scripts/watch-powerpoint-dev-server.mjs --help

The watcher starts \"pnpm dev-server:ppt\" only while Microsoft PowerPoint is
running. It owns only the process group it starts; an existing service on the
configured port is left untouched.

Environment:
  PPTXMATE_NODE_BIN            Node executable used by the LaunchAgent
  PPTXMATE_PNPM_BIN            pnpm executable or command (default: pnpm)
  PPTXMATE_PATH                PATH passed to pnpm and the dev server
  PPTXMATE_WATCHER_LOG         Watcher and dev-server log file
  PPTXMATE_PORT                Dev-server port to guard (default: 3001)
  PPTXMATE_CHECK_INTERVAL_MS   Poll interval (default: 3000)
  PPTXMATE_STOP_TIMEOUT_MS     Graceful process-group stop timeout (default: 5000)
  PPTXMATE_POWERPOINT_PROCESS  Process name (default: Microsoft PowerPoint)
`;
}

export function parseWatcherArgs(args) {
  if (args.length === 0) return { help: false };
  if (args.length === 1 && (args[0] === "--help" || args[0] === "-h")) {
    return { help: true };
  }
  throw new Error(`Unknown argument: ${args[0]}\n\n${watcherHelp()}`);
}

function parseInteger(value, name, minimum, maximum) {
  if (!/^\d+$/.test(String(value))) {
    throw new Error(`${name} must be an integer; received ${JSON.stringify(value)}`);
  }

  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < minimum || parsed > maximum) {
    throw new Error(`${name} must be between ${minimum} and ${maximum}; received ${value}`);
  }
  return parsed;
}

export function loadWatcherConfig(env = process.env, repoRoot = DEFAULT_REPO_ROOT) {
  const home = env.HOME || homedir();
  const pathValue =
    env.PPTXMATE_PATH ||
    [...new Set([...(env.PATH || "").split(":"), ...DEFAULT_PATH.split(":")])]
      .filter(Boolean)
      .join(":");
  const logFile = resolve(
    env.PPTXMATE_WATCHER_LOG || join(home, "Library", "Logs", "PPTXMate", "powerpoint-watcher.log"),
  );

  return {
    repoRoot: resolve(repoRoot),
    nodeBin: env.PPTXMATE_NODE_BIN || process.execPath,
    pnpmBin: env.PPTXMATE_PNPM_BIN || "pnpm",
    pathValue,
    logFile,
    port: parseInteger(env.PPTXMATE_PORT || "3001", "PPTXMATE_PORT", 1, 65535),
    checkIntervalMs: parseInteger(
      env.PPTXMATE_CHECK_INTERVAL_MS || "3000",
      "PPTXMATE_CHECK_INTERVAL_MS",
      250,
      3_600_000,
    ),
    stopTimeoutMs: parseInteger(
      env.PPTXMATE_STOP_TIMEOUT_MS || "5000",
      "PPTXMATE_STOP_TIMEOUT_MS",
      250,
      120_000,
    ),
    powerpointProcessName: env.PPTXMATE_POWERPOINT_PROCESS || "Microsoft PowerPoint",
    childEnv: {
      ...env,
      PATH: pathValue,
    },
  };
}

export function createFileLogger(logFile) {
  const logDirectory = dirname(logFile);
  const markerFile = join(logDirectory, ".pptxmate-watcher-logs");
  const directoryExisted = existsSync(logDirectory);
  mkdirSync(logDirectory, { recursive: true, mode: 0o700 });
  const managesDirectory =
    !directoryExisted || existsSync(markerFile) || readdirSync(logDirectory).length === 0;
  if (managesDirectory) {
    chmodSync(logDirectory, 0o700);
  }
  if (managesDirectory && !existsSync(markerFile)) {
    writeFileSync(markerFile, "PPTXMate PowerPoint watcher logs\n", { mode: 0o600 });
  }

  const descriptor = openSync(logFile, "a", 0o600);
  fchmodSync(descriptor, 0o600);
  const stream = createWriteStream(logFile, { fd: descriptor, autoClose: true });
  let closed = false;

  stream.on("error", (error) => {
    console.error(`PPTXMate watcher log error: ${error.message}`);
  });

  return {
    write(message) {
      if (closed) return;
      stream.write(`[${new Date().toISOString()}] ${message}\n`);
    },
    writeChunk(chunk) {
      if (closed) return;
      stream.write(chunk);
    },
    close() {
      if (closed) return Promise.resolve();
      closed = true;
      return new Promise((resolveClose) => stream.end(resolveClose));
    },
  };
}

function commandHasMatch(command, args) {
  return new Promise((resolveCommand, rejectCommand) => {
    execFile(command, args, (error) => {
      if (!error) {
        resolveCommand(true);
        return;
      }
      if (error.code === 1) {
        resolveCommand(false);
        return;
      }
      rejectCommand(
        new Error(`${command} failed: ${error instanceof Error ? error.message : String(error)}`),
      );
    });
  });
}

function processGroupExists(processGroupId) {
  try {
    process.kill(-processGroupId, 0);
    return true;
  } catch (error) {
    if (error?.code === "ESRCH") return false;
    if (error?.code === "EPERM") return true;
    throw error;
  }
}

function signalProcessGroup(processGroupId, signal) {
  try {
    process.kill(-processGroupId, signal);
    return true;
  } catch (error) {
    if (error?.code === "ESRCH") return false;
    throw error;
  }
}

function delay(milliseconds) {
  return new Promise((resolveDelay) => setTimeout(resolveDelay, milliseconds));
}

async function waitForProcessGroupExit(processGroupId, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (processGroupExists(processGroupId)) {
    if (Date.now() >= deadline) return false;
    await delay(Math.min(100, Math.max(1, deadline - Date.now())));
  }
  return true;
}

export async function terminateProcessGroup(processGroupId, timeoutMs) {
  if (!signalProcessGroup(processGroupId, "SIGTERM")) return;
  if (await waitForProcessGroupExit(processGroupId, timeoutMs)) return;

  signalProcessGroup(processGroupId, "SIGKILL");
  if (!(await waitForProcessGroupExit(processGroupId, 2_000))) {
    throw new Error(`process group ${processGroupId} did not exit after SIGKILL`);
  }
}

function defaultDependencies(config) {
  return {
    isPowerPointRunning: () =>
      commandHasMatch("/usr/bin/pgrep", ["-x", config.powerpointProcessName]),
    isPortInUse: () =>
      commandHasMatch("/usr/sbin/lsof", [
        "-nP",
        `-iTCP:${config.port}`,
        "-sTCP:LISTEN",
      ]),
    spawnDevServer: () =>
      spawn(config.pnpmBin, ["dev-server:ppt"], {
        cwd: config.repoRoot,
        env: config.childEnv,
        detached: true,
        stdio: ["ignore", "pipe", "pipe"],
      }),
    terminateProcessGroup: (processGroupId) =>
      terminateProcessGroup(processGroupId, config.stopTimeoutMs),
  };
}

export class PowerPointDevServerWatcher {
  constructor(config, logger, dependencies = {}) {
    this.config = config;
    this.logger = logger;
    this.dependencies = { ...defaultDependencies(config), ...dependencies };
    this.managedChild = null;
    this.cleanupPromise = null;
    this.tickPromise = null;
    this.timer = null;
    this.shuttingDown = false;
    this.portWasBlocked = false;
    this.lastPowerPointState = null;
  }

  async start() {
    this.logger.write(
      `watcher started node=${this.config.nodeBin} pnpm=${this.config.pnpmBin} port=${this.config.port}`,
    );
    await this.runTick();
    this.scheduleNextTick();
  }

  scheduleNextTick() {
    if (this.shuttingDown || this.timer) return;
    this.timer = setTimeout(async () => {
      this.timer = null;
      try {
        await this.runTick();
      } catch (error) {
        this.logger.write(`watcher tick failed: ${error instanceof Error ? error.stack : error}`);
      } finally {
        this.scheduleNextTick();
      }
    }, this.config.checkIntervalMs);
  }

  runTick() {
    if (this.tickPromise) return this.tickPromise;
    const currentTick = this.tick();
    this.tickPromise = currentTick;
    return currentTick.finally(() => {
      if (this.tickPromise === currentTick) this.tickPromise = null;
    });
  }

  async tick() {
    if (this.shuttingDown) return;
    const powerpointRunning = await this.dependencies.isPowerPointRunning();
    if (this.shuttingDown) return;

    if (powerpointRunning !== this.lastPowerPointState) {
      this.logger.write(`Microsoft PowerPoint ${powerpointRunning ? "detected" : "not running"}`);
      this.lastPowerPointState = powerpointRunning;
    }

    if (!powerpointRunning) {
      this.portWasBlocked = false;
      await this.stopManagedDevServer("PowerPoint closed");
      return;
    }

    if (this.managedChild) return;
    if (this.cleanupPromise) await this.cleanupPromise;
    if (this.shuttingDown) return;

    const portInUse = await this.dependencies.isPortInUse();
    if (this.shuttingDown) return;
    if (portInUse) {
      if (!this.portWasBlocked) {
        this.logger.write(
          `port ${this.config.port} is already in use; leaving the existing service untouched`,
        );
        this.portWasBlocked = true;
      }
      return;
    }

    if (this.portWasBlocked) {
      this.logger.write(`port ${this.config.port} is available again`);
      this.portWasBlocked = false;
    }
    this.startDevServer();
  }

  startDevServer() {
    if (this.shuttingDown || this.managedChild) return;
    this.logger.write("starting managed PowerPoint dev server: pnpm dev-server:ppt");
    const child = this.dependencies.spawnDevServer();
    if (!Number.isInteger(child.pid) || child.pid <= 0) {
      child.once("error", (error) => {
        this.logger.write(`managed dev server failed to start: ${error.message}`);
      });
      throw new Error(`failed to start ${this.config.pnpmBin}: no child process id was assigned`);
    }

    const record = {
      child,
      processGroupId: child.pid,
      stopping: false,
    };
    this.managedChild = record;

    child.stdout?.on("data", (data) => this.logger.writeChunk(data));
    child.stderr?.on("data", (data) => this.logger.writeChunk(data));
    child.on("error", (error) => {
      this.logger.write(`managed dev server failed to start: ${error.message}`);
    });
    child.once("close", (code, signal) => this.handleChildClose(record, code, signal));
    this.logger.write(`managed dev server started pid=${child.pid} pgid=${child.pid}`);
  }

  handleChildClose(record, code, signal) {
    this.logger.write(`managed dev server exited code=${code ?? ""} signal=${signal ?? ""}`);
    if (record.stopping) return;
    if (this.managedChild === record) this.managedChild = null;

    const cleanup = this.dependencies
      .terminateProcessGroup(record.processGroupId)
      .catch((error) => {
        this.logger.write(
          `failed to clean remaining dev-server processes: ${error instanceof Error ? error.stack : error}`,
        );
      })
      .finally(() => {
        if (this.cleanupPromise === cleanup) this.cleanupPromise = null;
      });
    this.cleanupPromise = cleanup;
  }

  async stopManagedDevServer(reason) {
    const record = this.managedChild;
    if (!record || record.stopping) return;
    record.stopping = true;
    this.logger.write(
      `stopping managed PowerPoint dev server pgid=${record.processGroupId}: ${reason}`,
    );

    try {
      await this.dependencies.terminateProcessGroup(record.processGroupId);
    } finally {
      if (this.managedChild === record) this.managedChild = null;
    }
  }

  async shutdown(signal = "shutdown") {
    if (this.shuttingDown) return;
    this.shuttingDown = true;
    if (this.timer) {
      clearTimeout(this.timer);
      this.timer = null;
    }

    const activeTick = this.tickPromise;
    if (activeTick) {
      try {
        await activeTick;
      } catch (error) {
        this.logger.write(
          `active watcher tick failed during shutdown: ${error instanceof Error ? error.stack : error}`,
        );
      }
    }
    await this.stopManagedDevServer(`watcher received ${signal}`);
    if (this.cleanupPromise) await this.cleanupPromise;
    this.logger.write("watcher stopped");
  }
}

async function main() {
  let parsedArgs;
  try {
    parsedArgs = parseWatcherArgs(process.argv.slice(2));
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 2;
    return;
  }
  if (parsedArgs.help) {
    process.stdout.write(watcherHelp());
    return;
  }
  if (process.platform !== "darwin") {
    throw new Error("PPTXMate's PowerPoint lifecycle watcher currently supports macOS only.");
  }

  const config = loadWatcherConfig();
  if (!existsSync(join(config.repoRoot, "package.json"))) {
    throw new Error(`PPTXMate repository root is invalid: ${config.repoRoot}`);
  }
  const logger = createFileLogger(config.logFile);
  const watcher = new PowerPointDevServerWatcher(config, logger);
  let shutdownPromise = null;

  const requestShutdown = (signal) => {
    if (shutdownPromise) return;
    shutdownPromise = watcher
      .shutdown(signal)
      .catch((error) => {
        logger.write(`watcher shutdown failed: ${error instanceof Error ? error.stack : error}`);
        process.exitCode = 1;
      })
      .finally(async () => {
        await logger.close();
      });
  };

  process.once("SIGINT", () => requestShutdown("SIGINT"));
  process.once("SIGTERM", () => requestShutdown("SIGTERM"));

  try {
    await watcher.start();
  } catch (error) {
    logger.write(`watcher failed: ${error instanceof Error ? error.stack : error}`);
    await watcher.shutdown("startup failure");
    await logger.close();
    throw error;
  }
}

const invokedPath = process.argv[1] ? pathToFileURL(resolve(process.argv[1])).href : null;
if (invokedPath === import.meta.url) {
  main().catch((error) => {
    console.error(`PPTXMate watcher failed: ${error instanceof Error ? error.message : error}`);
    process.exitCode = 1;
  });
}
