import { describe, expect, it } from "vitest";
import {
  createCcSwitchProxy,
  DEFAULT_CC_SWITCH_URL,
  isAllowedCcSwitchRequest,
  isCcSwitchProxyEnabled,
  resolveCcSwitchUrl,
} from "../cc-switch-proxy";

describe("CC Switch development proxy", () => {
  it("uses the standard local CC Switch endpoint by default", () => {
    expect(createCcSwitchProxy({})).toEqual({
      "/v1": {
        target: DEFAULT_CC_SWITCH_URL,
        changeOrigin: true,
        secure: false,
      },
    });
  });

  it.each([
    "0",
    "false",
    "FALSE",
    "no",
    "off",
  ])("can be disabled with %s", (value) => {
    const env = { PPTXMATE_CC_SWITCH_ENABLED: value };
    expect(isCcSwitchProxyEnabled(env)).toBe(false);
    expect(createCcSwitchProxy(env)).toEqual({});
  });

  it("accepts a configurable HTTP or HTTPS origin", () => {
    expect(
      resolveCcSwitchUrl({
        PPTXMATE_CC_SWITCH_URL: "https://localhost:25721/",
      }),
    ).toBe("https://localhost:25721");
  });

  it.each([
    "file:///tmp/cc-switch.sock",
    "http://user:secret@127.0.0.1:15721",
    "http://127.0.0.1:15721/v1",
    "http://127.0.0.1:15721?account=one",
  ])("rejects unsafe or ambiguous endpoint %s", (url) => {
    expect(() => resolveCcSwitchUrl({ PPTXMATE_CC_SWITCH_URL: url })).toThrow();
  });

  it.each([
    ["localhost:3001", "https://localhost:3001"],
    ["127.0.0.1:3001", "https://127.0.0.1:3001"],
    ["[::1]:3001", "https://[::1]:3001"],
    ["localhost:3001", undefined],
  ])("allows local host %s with origin %s", (host, origin) => {
    expect(isAllowedCcSwitchRequest({ host, origin })).toBe(true);
  });

  it.each([
    ["localhost:3001", "https://evil.example"],
    ["evil.example", "https://localhost:3001"],
    ["localhost:3001", "null"],
    ["localhost:3001", ["https://localhost:3001"]],
    [undefined, undefined],
  ])("rejects host %s with origin %s", (host, origin) => {
    expect(isAllowedCcSwitchRequest({ host, origin })).toBe(false);
  });
});
