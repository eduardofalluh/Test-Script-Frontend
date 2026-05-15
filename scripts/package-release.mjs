import { cpSync, existsSync, mkdirSync, rmSync, statSync } from "node:fs";
import { basename, join, resolve } from "node:path";
import { spawnSync } from "node:child_process";

const root = process.cwd();
const releaseRoot = resolve(root, "release");
const packageName = "excel-mapper-local";
const stagingDir = join(releaseRoot, packageName);
const zipPath = join(releaseRoot, `${packageName}.zip`);
const tarPath = join(releaseRoot, `${packageName}.tar.gz`);

const excludedNames = new Set([
  ".git",
  ".next",
  "node_modules",
  "release",
  ".DS_Store",
  "tsconfig.tsbuildinfo"
]);

const includedPaths = [
  ".env.example",
  ".eslintrc.json",
  ".gitignore",
  "README.md",
  "README_FOR_RECIPIENTS.md",
  "START_EXCEL_MAPPER_MAC.command",
  "START_EXCEL_MAPPER_WINDOWS.bat",
  "app",
  "components",
  "components.json",
  "hooks",
  "lib",
  "next-env.d.ts",
  "next.config.js",
  "package-lock.json",
  "package.json",
  "postcss.config.js",
  "public",
  "scripts",
  "tailwind.config.ts",
  "tsconfig.json"
];

rmSync(releaseRoot, { force: true, recursive: true });
mkdirSync(stagingDir, { recursive: true });

for (const relativePath of includedPaths) {
  const source = join(root, relativePath);
  if (!existsSync(source)) {
    continue;
  }

  copyForRelease(source, join(stagingDir, relativePath));
}

const zipped = runArchiveCommand("zip", ["-qry", zipPath, packageName], releaseRoot);

if (!zipped) {
  runArchiveCommand("tar", ["-czf", tarPath, packageName], releaseRoot, true);
}

console.log(`Release folder: ${stagingDir}`);
console.log(`Release archive: ${existsSync(zipPath) ? zipPath : tarPath}`);

function copyForRelease(source, destination) {
  if (excludedNames.has(basename(source))) {
    return;
  }

  const stats = statSync(source);
  if (stats.isDirectory()) {
    cpSync(source, destination, {
      recursive: true,
      filter: (currentSource) => !excludedNames.has(basename(currentSource))
    });
    return;
  }

  cpSync(source, destination);
}

function runArchiveCommand(command, args, cwd, required = false) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    stdio: "pipe"
  });

  if (result.status === 0) {
    return true;
  }

  if (required) {
    const details = [result.stdout, result.stderr].filter(Boolean).join("\n");
    throw new Error(`Failed to create release archive with ${command}.\n${details}`);
  }

  return false;
}
