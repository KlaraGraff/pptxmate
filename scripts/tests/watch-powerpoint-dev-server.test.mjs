import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { EventEmitter } from "node:events";
import { existsSync, mkdtempSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PassThrough } from "node:stream";
import test from "node:test";
import { fileURLToPath } from "node:url";

import {
  createFileLogger,
  loadWatcherConfig,
  parseWatcherArgs,
  PowerPointDevServerWatcher,
} from "../watch-powerpoint-dev-server.mjs";

const INSTALLER_PATH = fileURLToPath(
  new URL("../install-macos-powerpoint-watcher.sh", import.meta.url),
);

function plistValue(xml, key) {
  const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = xml.match(
    new RegExp(`<key>${escapedKey}</key>\\s*<string>([^<]*)</string>`),
  );
  assert.ok(match, `missing plist value for ${key}`);
  return match[1]
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&gt;", ">")
    .replaceAll("&lt;", "<")
    .replaceAll("&amp;", "&");
}

function previewInstaller(temporaryHome, args = [], environment = {}) {
  return spawnSync(
    "/bin/bash",
    [
      INSTALLER_PATH,
      "--dry-run",
      "--node",
      process.execPath,
      "--pnpm",
      process.execPath,
      "--log-dir",
      join(temporaryHome, "logs"),
      ...args,
    ],
    {
      encoding: "utf8",
      env: {
        ...process.env,
        HOME: temporaryHome,
        ...environment,
      },
    },
  );
}

function createLogger() {
  const entries = [];
  return {
    entries,
    write(message) {
      entries.push(String(message));
    },
    writeChunk() {},
  };
}

function createConfig(overrides = {}) {
  return {
    ...loadWatcherConfig(
      {
        HOME: "/tmp/pptxmate-test-home",
        PATH: "/test/bin:/usr/bin",
      },
      "/tmp/pptxmate-repo",
    ),
    ...overrides,
  };
}

class FakeChild extends EventEmitter {
  constructor(pid = 4242) {
    super();
    this.pid = pid;
    this.stdout = new PassThrough();
    this.stderr = new PassThrough();
  }
}

test("configuration uses the public PPTXMate environment variables", () => {
  const config = loadWatcherConfig(
    {
      HOME: "/tmp/pptxmate-home",
      PATH: "/usr/bin:/bin",
      PPTXMATE_NODE_BIN: "/custom/node",
      PPTXMATE_PNPM_BIN: "/custom/pnpm",
      PPTXMATE_PATH: "/custom/bin:/usr/bin",
      PPTXMATE_WATCHER_LOG: "/tmp/pptxmate.log",
      PPTXMATE_PORT: "4100",
      PPTXMATE_CHECK_INTERVAL_MS: "750",
      PPTXMATE_STOP_TIMEOUT_MS: "900",
      PPTXMATE_CC_SWITCH_URL: "https://localhost:25721",
      PPTXMATE_CC_SWITCH_ENABLED: "0",
    },
    "/repo",
  );

  assert.equal(config.nodeBin, "/custom/node");
  assert.equal(config.pnpmBin, "/custom/pnpm");
  assert.equal(config.pathValue, "/custom/bin:/usr/bin");
  assert.equal(config.childEnv.PATH, "/custom/bin:/usr/bin");
  assert.equal(config.logFile, "/tmp/pptxmate.log");
  assert.equal(config.port, 4100);
  assert.equal(config.checkIntervalMs, 750);
  assert.equal(config.stopTimeoutMs, 900);
  assert.equal(
    config.childEnv.PPTXMATE_CC_SWITCH_URL,
    "https://localhost:25721",
  );
  assert.equal(config.childEnv.PPTXMATE_CC_SWITCH_ENABLED, "0");
});

test("the default child PATH retains common Homebrew locations", () => {
  const config = loadWatcherConfig({ HOME: "/tmp/home", PATH: "/usr/bin:/bin" }, "/repo");
  assert.match(config.pathValue, /\/opt\/homebrew\/bin/);
  assert.match(config.pathValue, /\/usr\/local\/bin/);
});

test("a newly created watcher log is private", async () => {
  const temporaryRoot = mkdtempSync(join(tmpdir(), "pptxmate-watcher-test-"));
  const logDirectory = join(temporaryRoot, "logs");
  const logFile = join(logDirectory, "watcher.log");

  try {
    const logger = createFileLogger(logFile);
    logger.write("test");
    await logger.close();

    assert.equal(statSync(logDirectory).mode & 0o777, 0o700);
    assert.equal(statSync(logFile).mode & 0o777, 0o600);
    assert.equal(statSync(join(logDirectory, ".pptxmate-watcher-logs")).mode & 0o777, 0o600);
  } finally {
    rmSync(temporaryRoot, { recursive: true, force: true });
  }
});

test("configuration rejects invalid numeric settings", () => {
  assert.throws(
    () => loadWatcherConfig({ HOME: "/tmp", PPTXMATE_PORT: "not-a-port" }),
    /PPTXMATE_PORT must be an integer/,
  );
  assert.throws(
    () => loadWatcherConfig({ HOME: "/tmp", PPTXMATE_CHECK_INTERVAL_MS: "20" }),
    /must be between 250 and 3600000/,
  );
});

test("only --help is accepted on the watcher command line", () => {
  assert.deepEqual(parseWatcherArgs([]), { help: false });
  assert.deepEqual(parseWatcherArgs(["--help"]), { help: true });
  assert.throws(() => parseWatcherArgs(["--unknown"]), /Unknown argument/);
});

test("concurrent ticks coalesce instead of re-entering PowerPoint detection", async () => {
  let detectionCalls = 0;
  let releaseDetection;
  const detectionGate = new Promise((resolve) => {
    releaseDetection = resolve;
  });
  const watcher = new PowerPointDevServerWatcher(createConfig(), createLogger(), {
    isPowerPointRunning: async () => {
      detectionCalls += 1;
      await detectionGate;
      return false;
    },
  });

  const first = watcher.runTick();
  const second = watcher.runTick();
  releaseDetection();
  await Promise.all([first, second]);

  assert.equal(detectionCalls, 1);
});

test("a foreign service on the guarded port is never started or stopped", async () => {
  let spawnCalls = 0;
  let terminateCalls = 0;
  const logger = createLogger();
  const watcher = new PowerPointDevServerWatcher(createConfig(), logger, {
    isPowerPointRunning: async () => true,
    isPortInUse: async () => true,
    spawnDevServer: () => {
      spawnCalls += 1;
      return new FakeChild();
    },
    terminateProcessGroup: async () => {
      terminateCalls += 1;
    },
  });

  await watcher.runTick();
  await watcher.shutdown("test");

  assert.equal(spawnCalls, 0);
  assert.equal(terminateCalls, 0);
  assert.ok(logger.entries.some((entry) => entry.includes("leaving the existing service untouched")));
});

test("the watcher stops exactly the process group it started", async () => {
  const child = new FakeChild(9876);
  const terminatedGroups = [];
  let powerpointRunning = true;
  const watcher = new PowerPointDevServerWatcher(createConfig(), createLogger(), {
    isPowerPointRunning: async () => powerpointRunning,
    isPortInUse: async () => false,
    spawnDevServer: () => child,
    terminateProcessGroup: async (processGroupId) => {
      terminatedGroups.push(processGroupId);
    },
  });

  await watcher.runTick();
  powerpointRunning = false;
  await watcher.runTick();

  assert.deepEqual(terminatedGroups, [9876]);
});

test(
  "the installer persists a custom CC Switch origin without side effects in dry-run mode",
  { skip: process.platform !== "darwin" },
  () => {
    const temporaryHome = mkdtempSync(join(tmpdir(), "pptxmate-installer-test-"));

    try {
      const result = previewInstaller(temporaryHome, [
        "--cc-switch-url",
        "https://localhost:25721",
      ]);

      assert.equal(result.status, 0, result.stderr);
      assert.equal(
        plistValue(result.stdout, "PPTXMATE_CC_SWITCH_URL"),
        "https://localhost:25721",
      );
      assert.equal(plistValue(result.stdout, "PPTXMATE_CC_SWITCH_ENABLED"), "1");
      assert.equal(existsSync(join(temporaryHome, "logs")), false);
      assert.equal(existsSync(join(temporaryHome, "Library", "LaunchAgents")), false);
    } finally {
      rmSync(temporaryHome, { recursive: true, force: true });
    }
  },
);

test(
  "the last CC Switch installer option wins",
  { skip: process.platform !== "darwin" },
  () => {
    const temporaryHome = mkdtempSync(join(tmpdir(), "pptxmate-installer-test-"));

    try {
      const disabled = previewInstaller(temporaryHome, [
        "--cc-switch-url",
        "http://127.0.0.1:25721",
        "--no-cc-switch",
      ]);
      assert.equal(disabled.status, 0, disabled.stderr);
      assert.equal(plistValue(disabled.stdout, "PPTXMATE_CC_SWITCH_ENABLED"), "0");

      const enabled = previewInstaller(temporaryHome, [
        "--no-cc-switch",
        "--cc-switch-url",
        "http://127.0.0.1:25721",
      ]);
      assert.equal(enabled.status, 0, enabled.stderr);
      assert.equal(plistValue(enabled.stdout, "PPTXMATE_CC_SWITCH_ENABLED"), "1");
    } finally {
      rmSync(temporaryHome, { recursive: true, force: true });
    }
  },
);
