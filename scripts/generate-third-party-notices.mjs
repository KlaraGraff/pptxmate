import { execFileSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { writeFileSync } from "node:fs";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const raw = execFileSync(
  process.env.PNPM_BIN || "pnpm",
  ["licenses", "list", "--prod", "--json"],
  { cwd: repoRoot, encoding: "utf8" },
);
const grouped = JSON.parse(raw);

const packages = new Map();
for (const [license, entries] of Object.entries(grouped)) {
  if (!Array.isArray(entries)) continue;
  for (const entry of entries) {
    if (!entry || typeof entry !== "object" || typeof entry.name !== "string") {
      continue;
    }
    const versions = Array.isArray(entry.versions)
      ? entry.versions.filter((version) => typeof version === "string")
      : [];
    const key = `${entry.name}@${versions.join(",")}`;
    packages.set(key, {
      name: entry.name,
      versions,
      license,
      homepage:
        typeof entry.homepage === "string" && /^https?:\/\//.test(entry.homepage)
          ? entry.homepage
          : "",
    });
  }
}

function escapeCell(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

const rows = [...packages.values()]
  .sort((a, b) => a.name.localeCompare(b.name) || a.license.localeCompare(b.license))
  .map((entry) => {
    const name = entry.homepage
      ? `[${escapeCell(entry.name)}](${entry.homepage})`
      : escapeCell(entry.name);
    return `| ${name} | ${escapeCell(entry.versions.join(", "))} | ${escapeCell(entry.license)} |`;
  });

const output = `# Third-Party Notices

PPTXMate uses the following production dependencies. This list is generated from the locked pnpm workspace with \`pnpm licenses\`.

Each package remains subject to its own license. The package source and full license text are available from the linked project and from the installed package's license files. This notice does not replace those license terms.

| Package | Version | License |
| --- | --- | --- |
${rows.join("\n")}
`;

writeFileSync(resolve(repoRoot, "THIRD_PARTY_NOTICES.md"), output);
console.log(`Wrote THIRD_PARTY_NOTICES.md with ${rows.length} packages.`);
