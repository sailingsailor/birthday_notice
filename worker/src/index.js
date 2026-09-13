// birthday-notice Worker
//
// 架构（按用户要求）：
//   1) 每天定时（中国 00:00）把「前后 60 天」的生日日历预先算好，写入 KV 的 calendar.ics。
//   2) 客户端拉取订阅地址时，直接返回 KV 里已生成的 calendar.ics，不再实时运算。
//   3) 同时保留原有 Bark 提前提醒（每日定时检查）。
//
// 注意：本程序只解析 json，不会解析 xlsx。xlsx -> json 由本地 xlsx_to_json.py 完成。
// 日历计算逻辑统一放在 ./calendar_core.mjs（Worker 与本地 gen_ics.mjs 共用）。

import { Solar, Lunar, LunarYear } from "lunar-javascript";
// 打包内置的数据源（也可改用 KV，见 wrangler.toml）
import bundledData from "./birthday.json";
import {
  CAL_WINDOW,
  chinaNow,
  lunarTodayStr,
  diffDays,
  lunarToSolar,
  buildWindowICS,
} from "./calendar_core.mjs";

// 从 KV 或打包数据加载数据源
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
        cn: `${cn.year}-${String(cn.month).padStart(2, "0")}-${String(cn.day).padStart(2, "0")}`,
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
async function checkBirthdays(env) {
  const data = await loadData(env);
  const cn = chinaNow();
  const title = `今日${cn.year}-${String(cn.month).padStart(2, "0")}-${String(cn.day).padStart(2, "0")}\n${lunarTodayStr(cn)}`;
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
    if (flag > 3) continue;

    const howOld = cn.year - p.solar.year;

    let desp;
    if (flag === 0) {
      desp = `${p.name}今天过${howOld}岁生日,阴历${lunar ? lunar.month + "月" + lunar.day + "日" : ""}`;
    } else {
      const resultStr = `${target.year}-${String(target.month).padStart(2, "0")}-${String(target.day).padStart(2, "0")}`;
      const lunarStr = lunar
        ? `${lunar.year}-${String(lunar.month).padStart(2, "0")}-${String(lunar.day).padStart(2, "0")}`
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
