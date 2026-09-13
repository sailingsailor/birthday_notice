// 本地预生成 calendar.ics（不依赖线上 Worker / KV），
// 用与 Worker 完全相同的 buildWindowICS 逻辑，直接从本地 birthday.json 生成。
// 生成后由 update.ps1 把它和 birthday.json 一起推到 KV，线上订阅即可立即返回正确日历。
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { buildWindowICS, chinaNow, CAL_WINDOW } from "./src/calendar_core.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const data = JSON.parse(readFileSync(join(__dirname, "src", "birthday.json"), "utf8"));
const cn = chinaNow();
const ics = buildWindowICS(data, cn, CAL_WINDOW);
const outPath = join(__dirname, "src", "calendar.ics");
writeFileSync(outPath, ics, "utf8");
const veventCount = ics.split("\r\n").filter((l) => l.startsWith("BEGIN:VEVENT")).length;
console.log(`calendar.ics 已生成本地: ${outPath} (${veventCount} 条 VEVENT, 基准日 ${cn.year}-${String(cn.month).padStart(2,"0")}-${String(cn.day).padStart(2,"0")})`);
if (!ics.includes("王创")) {
  console.warn("⚠️ 警告：生成的日历里没有「王创」，请检查 birthday.json 中王创的农历日期是否合法（如当月无该日会被跳过）。");
}
