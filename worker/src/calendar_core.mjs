// 生日日历核心计算（纯函数），Worker(index.js) 与本地生成脚本(gen_ics.mjs) 共用，
// 保证“线上生成”和“本地预生成”用完全一致的逻辑，杜绝两份实现漂移。
import { Solar, Lunar, LunarYear } from "lunar-javascript";

export const CAL_NAME = "生日提醒";
export const CAL_WINDOW = 60; // 前后天数（共 121 天滚动窗口）

export function pad2(n) {
  return n < 10 ? "0" + n : "" + n;
}

// 以中国时间（UTC+8，中国不实行夏令时）计算“今天”
export function chinaNow() {
  const utc = new Date();
  const cn = new Date(utc.getTime() + 8 * 3600 * 1000);
  return {
    year: cn.getUTCFullYear(),
    month: cn.getUTCMonth() + 1,
    day: cn.getUTCDate(),
  };
}

// 农历今日字符串，例：农历 丙午年 七月十七
export function lunarTodayStr(cn) {
  const l = Solar.fromYmd(cn.year, cn.month, cn.day).getLunar();
  return `农历${l.getYearInGanZhi()}年${l.getMonthInChinese()}月${l.getDayInChinese()}`;
}

// 两个日期相差天数（b - a）：正数=未来，负数=过去
export function diffDays(a, b) {
  const t = Date.UTC(a.year, a.month - 1, a.day);
  const s = Date.UTC(b.year, b.month - 1, b.day);
  return Math.round((t - s) / 86400000);
}

// 农历->阳历，与原始 zhdate 行为严格一致：
// - 若该农历年/月实际没有这一天（如 农历七月只有29天却有30日），返回 null（原 zhdate 会抛错，
//   原始 Python 因此整行 skip）。必须按此跳过，否则会把“不存在的日期”进位到下月，产生错误生日。
// - 仅在日期合法时使用 农历月初儒略日 + (day-1) 转阳历（不会抛错）。
export function lunarToSolar(year, month, day) {
  const m = LunarYear.fromYear(year).getMonth(month);
  const dayCount = m.getDayCount();
  if (day > dayCount) return null;
  const noon = Solar.fromJulianDay(m.getFirstJulianDay() + day - 1);
  return { year: noon.getYear(), month: noon.getMonth(), day: noon.getDay() };
}

export function escapeICS(s) {
  return String(s)
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n");
}

// 生成“前后 W 天”滚动窗口的 ICS（每个命中的生日是一条具体日期的单次 VEVENT，不做 RRULE 重复）
export function buildWindowICS(data, cn, windowDays) {
  const W = windowDays;
  const people = (data && data.people) || [];
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//birthday-notice//Birthday Calendar//CN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    "CALNAME:" + escapeICS(CAL_NAME),
    "X-WR-CALNAME:" + escapeICS(CAL_NAME),
    "TIMEZONE-ID:Asia/Shanghai",
  ];
  // 跨年边界：窗口可能延伸到上一年末或下一年初，故对 当前年-1 / 当前年 / 当前年+1 都算一遍
  const years = [cn.year - 1, cn.year, cn.year + 1];
  let idx = 0;
  for (const p of people) {
    if (!p || !p.name) continue;
    const typ = p.type || "";
    const dates = []; // { date:{year,month,day}, year }
    if (/阳/.test(typ)) {
      if (p.solar) for (const y of years) dates.push({ date: { year: y, month: p.solar.month, day: p.solar.day }, year: y });
    } else if (/阴/.test(typ)) {
      if (p.lunar) for (const y of years) {
        const s = lunarToSolar(y, p.lunar.month, p.lunar.day);
        if (s) dates.push({ date: s, year: y });
      }
    } else {
      // 未知类型：阳历/阴历都试一次
      if (p.solar) for (const y of years) dates.push({ date: { year: y, month: p.solar.month, day: p.solar.day }, year: y });
      if (p.lunar) for (const y of years) {
        const s = lunarToSolar(y, p.lunar.month, p.lunar.day);
        if (s) dates.push({ date: s, year: y });
      }
    }
    for (const d of dates) {
      const flag = diffDays(d.date, cn); // >0 未来, <0 过去
      if (flag < -W || flag > W) continue; // 不在窗口内跳过
      const birthYear = p.solar ? p.solar.year : (p.lunar ? p.lunar.year : d.year);
      const age = d.year - birthYear;
      const dt = `${d.date.year}${pad2(d.date.month)}${pad2(d.date.day)}`;
      const sum = p.name + " 生日" + (age > 0 ? ` (${age}岁)` : "");
      let desc = "";
      if (p.solar) desc += "阳历 " + p.solar.year + "-" + pad2(p.solar.month) + "-" + pad2(p.solar.day);
      if (p.lunar) desc += (p.solar ? " / " : "") + "农历 " + p.lunar.year + "-" + pad2(p.lunar.month) + "-" + pad2(p.lunar.day);
      if (p.note) desc += (desc ? "  " : "") + "备注:" + p.note;
      lines.push("BEGIN:VEVENT");
      lines.push(`UID:evt-${dt}-${idx++}@birthday-notice`);
      lines.push(`DTSTART;VALUE=DATE:${dt}`);
      lines.push(`SUMMARY:${escapeICS(sum)}`);
      lines.push(`DESCRIPTION:${escapeICS(desc)}`);
      lines.push("END:VEVENT");
    }
  }
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}
