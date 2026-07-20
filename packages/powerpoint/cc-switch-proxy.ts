export interface CcSwitchEnvironment {
  PPTXMATE_CC_SWITCH_ENABLED?: string;
  PPTXMATE_CC_SWITCH_URL?: string;
}

export interface CcSwitchProxyEntry {
  target: string;
  changeOrigin: boolean;
  secure: boolean;
}

export type CcSwitchProxyConfig = Record<string, CcSwitchProxyEntry>;

export const DEFAULT_CC_SWITCH_URL = "http://127.0.0.1:15721";

const DISABLED_VALUES = new Set(["0", "false", "no", "off"]);
const ALLOWED_DEV_HOSTS = new Set([
  "localhost:3001",
  "127.0.0.1:3001",
  "[::1]:3001",
]);
const ALLOWED_DEV_ORIGINS = new Set([
  "https://localhost:3001",
  "https://127.0.0.1:3001",
  "https://[::1]:3001",
]);

export function isCcSwitchProxyEnabled(
  env: CcSwitchEnvironment = process.env,
): boolean {
  const value = env.PPTXMATE_CC_SWITCH_ENABLED?.trim().toLowerCase();
  return value === undefined || !DISABLED_VALUES.has(value);
}

export function resolveCcSwitchUrl(
  env: CcSwitchEnvironment = process.env,
): string {
  const raw = env.PPTXMATE_CC_SWITCH_URL?.trim() || DEFAULT_CC_SWITCH_URL;
  const url = new URL(raw);
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    throw new Error(
      "PPTXMATE_CC_SWITCH_URL must use the http or https protocol.",
    );
  }
  if (url.username || url.password) {
    throw new Error(
      "PPTXMATE_CC_SWITCH_URL must not contain credentials. Configure accounts in CC Switch.",
    );
  }
  if (url.pathname !== "/" || url.search || url.hash) {
    throw new Error(
      "PPTXMATE_CC_SWITCH_URL must be an origin such as http://127.0.0.1:15721, without a path, query, or fragment.",
    );
  }
  return url.origin;
}

export function createCcSwitchProxy(
  env: CcSwitchEnvironment = process.env,
): CcSwitchProxyConfig {
  if (!isCcSwitchProxyEnabled(env)) return {};
  return {
    "/v1": {
      target: resolveCcSwitchUrl(env),
      changeOrigin: true,
      secure: false,
    },
  };
}

export function isAllowedCcSwitchRequest(headers: {
  host?: string | string[];
  origin?: string | string[];
}): boolean {
  const host =
    typeof headers.host === "string" ? headers.host.toLowerCase() : "";
  if (!ALLOWED_DEV_HOSTS.has(host)) return false;

  if (headers.origin === undefined) return true;
  if (typeof headers.origin !== "string") return false;
  return ALLOWED_DEV_ORIGINS.has(headers.origin.toLowerCase());
}
