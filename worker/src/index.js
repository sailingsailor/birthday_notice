// birthday-notice Worker
//
// 架构（按用户要求）：
//   1) 每天定时（中国 00:00）把「前后 60 天」的生日日历预先算好，写入 KV 的 calendar.ics。
//   2) 客户端拉取订阅地址时，直接返回 KV 里已生成的 calendar.ics，不再实时运算。
//   3) 同时保留原有 Bark 提前提醒（每日定时检查）。
//
// 注意：本程序只解析 json，不会解析 xlsx。xlsx -> json 由本地 xlsx_to_json.py 完成。

import { Solar, Lunar, LunarYear } from "lunar-javascript";
// 打包内置的数据源（也可改用 KV，见 wrangler.toml）
import bundledData from "./birthday.json";

const CAL_NAME = "生日提醒";
const CAL_WINDOW = 60;     // 前后天数（共 121 天滚动窗口）
const NOTIFY_WINDOW = 3;   // Bark：提前多少天开始提醒（含当天）

function pad2(n) {
  return n < 10 ? "0" + n : "" + n;
}

// 以中国时间（UTC+8，中国不实行夏令时）计算“今天”
function chinaNow() {
  const utc = new Date();
  const cn = new Date(utc.getTime() + 8 * 3600 * 1000);
  return {
    year: cn.getUTCFullYear(),
    month: cn.getUTCMonth() + 1,
    day: cn.getUTCDate(),
  };
}

// 农历今日字符串，例：农历 丙午年 七月十七
function lunarTodayStr(cn) {
  const l = Solar.fromYmd(cn.year, cn.month, cn.day).getLunar();
  return `农历${l.getYearInGanZhi()}年${l.getMonthInChinese()}月${l.getDayInChinese()}`;
}

// 两个日期相差天数（b - a）：正数=未来，负数=过去
function diffDays(a, b) {
  const t = Date.UTC(a.year, a.month - 1, a.day);
  const s = Date.UTC(b.year, b.month - 1, b.day);
  return Math.round((t - s) / 86400000);
}

// 农历->阳历，与原始 zhdate 行为严格一致：
// - 若该农历年/月实际没有这一天（如 农历七月只有29天却有30日），返回 null（原 zhdate 会抛错，
//   原始 Python 因此整行 skip）。必须按此跳过，否则会把“不存在的日期”进位到下月，产生错误生日。
// - 仅在日期合法时使用 农历月初儒略日 + (day-1) 转阳历（不会抛错）。
function lunarToSolar(year, month, day) {
  const m = LunarYear.fromYear(year).getMonth(month);
  const dayCount = m.getDayCount();
  if (day > dayCount) return null;
  const noon = Solar.fromJulianDay(m.getFirstJulianDay() + day - 1);
  return { year: noon.getYear(), month: noon.getMonth(), day: noon.getDay() };
}

function escapeICS(s) {
  return String(s)
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n");
}

// 生成“前后 W 天”滚动窗口的 ICS（每个命中的生日是一条具体日期的单次 VEVENT，不做 RRULE 重复）
function buildWindowICS(data, cn, windowDays) {
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

async function loadData(env) {
  // 若配置了 KV 绑定，优先从 KV 读取（运行时真正解析 json）
  if (env && env.BIRTHDAY) {
    const txt = await env.BIRTHDAY.get("birthday.json");
    if (txt) return JSON.parse(txt);
  }
  return bundledData;
}

// 每日预生成：把前后 60 天日历写入 KV 的 calendar.ics（订阅拉取时直接读它，不再运算）
async function generateCalendar(env) {
  const data = await loadData(env);
  const cn = chinaNow();
  const ics = buildWindowICS(data, cn, CAL_WINDOW);
  if (env && env.BIRTHDAY) {
    await env.BIRTHDAY.put("calendar.ics", ics, {
      metadata: {
        updatedAt: new Date().toISOString(),
        cn: `${cn.year}-${pad2(cn.month)}-${pad2(cn.day)}`,
        window: CAL_WINDOW,
      },
    });
  }
  return ics;
}

async function sendBark(env, title, body) {
  const key = env && env.BARK_KEY;
  if (!key) {
    console.warn("BARK_KEY 未配置，跳过发送");
    return { skipped: true };
  }
  const base = ((env && env.BARK_BASE) || "https://api.day.app/").replace(/\/$/, "");
  const url = `${base}/${key}`;
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, body }),
  });
  return { status: resp.status, ok: resp.ok };
}

// 核心：检查所有人生日，返回将要/已经发送的提醒（Bark）
export async function checkBirthdays(env) {
  const data = await loadData(env);
  const cn = chinaNow();
  const title = `今日${cn.year}-${pad2(cn.month)}-${pad2(cn.day)}\n${lunarTodayStr(cn)}`;
  const people = (data && data.people) || [];

  const sends = [];
  for (const p of people) {
    if (!p || !p.name || !p.solar) continue;
    const isSolar = /阳/.test(p.type || "");
    const lunar = p.lunar || null;

    let target; // 今年对应的阳历生日
    if (isSolar) {
      target = { year: cn.year, month: p.solar.month, day: p.solar.day };
    } else if (lunar) {
      // 把今年农历 月/日 转成阳历；若该农历月没有这一天（如七月只有29天却要30日），
      // 与原始 zhdate 一致：跳过本年（不通知，也不进位到下月）
      const t = lunarToSolar(cn.year, lunar.month, lunar.day);
      if (!t) continue;
      target = t;
    } else {
      continue;
    }

    const flag = diffDays(target, cn); // 距离今天还有多少天（负数=已过）
    if (flag < 0) continue; // 今年已过的生日不再提醒（明年自然进入窗口）
    if (flag > NOTIFY_WINDOW) continue;

    const howOld = cn.year - p.solar.year;

    let desp;
    if (flag === 0) {
      desp = `${p.name}今天过${howOld}岁生日,阴历${lunar ? lunar.month + "月" + lunar.day + "日" : ""}`;
    } else {
      const resultStr = `${target.year}-${pad2(target.month)}-${pad2(target.day)}`;
      const lunarStr = lunar
        ? `${lunar.year}-${pad2(lunar.month)}-${pad2(lunar.day)}`
        : "";
      desp = `${p.name}${resultStr}(${flag}天后)过${howOld}岁生日,阴历${lunarStr}`;
    }

    const r = await sendBark(env, title, desp);
    sends.push({ name: p.name, flag, desp, result: r });
  }

  return { title, checked: people.length, sent: sends.length, sends };
}

export default {
  // 定时任务（cron）入口：每天中国 00:00 触发
  async scheduled(event, env, ctx) {
    ctx.waitUntil((async () => {
      // 1) 预生成前后 60 天滚动日历 -> KV calendar.ics
      const ics = await generateCalendar(env);
      console.log("calendar generated, bytes:", ics.length);
      // 2) 保留原有 Bark 提前提醒
      const r = await checkBirthdays(env);
      console.log("birthday check:", JSON.stringify(r));
    })());
  },
  // HTTP 入口：日历订阅（?token=xxx）。token 校验失败一律返回 404。
  // 直接返回 KV 中已生成的 calendar.ics，不再实时计算。
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const token = url.searchParams.get("token");
    if (!env.CAL_TOKEN || token !== env.CAL_TOKEN) {
      return new Response("Not Found", {
        status: 404,
        headers: { "Content-Type": "text/plain; charset=utf-8" },
      });
    }
    let ics = null;
    if (env.BIRTHDAY) ics = await env.BIRTHDAY.get("calendar.ics");
    if (!ics) ics = await generateCalendar(env); // 兜底：KV 尚无则现算并存储
    return new Response(ics, {
      headers: {
        "Content-Type": "text/calendar; charset=utf-8",
        "Cache-Control": "public, max-age=300",
      },
    });
  },
};
