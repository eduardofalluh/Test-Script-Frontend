import { rmSync } from "node:fs";
import { resolve } from "node:path";

rmSync(resolve(process.cwd(), ".next"), {
  force: true,
  recursive: true
});
