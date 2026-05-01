// Cloudflare Worker — difuze-calendar-sync
// Replaces the Vercel FastAPI backend. Stores scraped Outlook events in KV,
// serves them as JSON, and emits SSE streams.

const EVENTS_KEY = "events";
const TRIGGER_KEY = "current_trigger";

function checkAuth(request, env) {
  const token = env.INGEST_TOKEN;
  if (!token) throw new Error("unauthorized");
  const auth = request.headers.get("Authorization");
  if (auth !== `Bearer ${token}`) throw new Error("unauthorized");
}

function diff(oldEvents, newEvents) {
  const oldByUid = new Map((oldEvents || []).map((e) => [e.uid, e]));
  const newByUid = new Map((newEvents || []).map((e) => [e.uid, e]));
  const added = [];
  const removed = [];
  const changed = [];
  for (const [uid, e] of newByUid) {
    if (!oldByUid.has(uid)) added.push(e);
    else {
      const old = JSON.stringify(oldByUid.get(uid));
      if (old !== JSON.stringify(e)) changed.push(e);
    }
  }
  for (const [uid, e] of oldByUid) {
    if (!newByUid.has(uid)) removed.push(e);
  }
  return { added, removed, changed };
}

async function readEvents(env) {
  const raw = await env.EVENTS_KV.get(EVENTS_KEY, "json");
  return raw || { scraped_at: null, count: 0, events: [] };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "content-type": "application/json",
      "access-control-allow-origin": "*",
      "access-control-allow-methods": "GET, POST, OPTIONS",
      "access-control-allow-headers": "Authorization, Content-Type",
    },
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    // CORS preflight
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          "access-control-allow-origin": "*",
          "access-control-allow-methods": "GET, POST, OPTIONS",
          "access-control-allow-headers": "Authorization, Content-Type",
          "access-control-max-age": "86400",
        },
      });
    }

    // Health
    if (path === "/api/health") {
      return json({
        ok: true,
        kv_configured: !!env.EVENTS_KV,
        auth_configured: !!env.INGEST_TOKEN,
      });
    }

    // Ingest
    if (path === "/api/ingest" && request.method === "POST") {
      try {
        checkAuth(request, env);
      } catch {
        return json({ error: "unauthorized" }, 401);
      }

      const payload = await request.json();
      const events = payload.events || [];
      const old = await readEvents(env);
      const d = diff(old.events, events);

      const updated = {
        scraped_at: payload.scraped_at || new Date().toISOString(),
        calendar: payload.calendar || "",
        count: events.length,
        events,
      };

      await env.EVENTS_KV.put(EVENTS_KEY, JSON.stringify(updated));

      return json({
        ok: true,
        count: events.length,
        added: d.added.length,
        removed: d.removed.length,
        changed: d.changed.length,
      });
    }

    // Events
    if (path === "/api/events") {
      const data = await readEvents(env);
      return json(data);
    }

    // SSE stream
    if (path === "/api/events/stream") {
      const { readable, writable } = new TransformStream();
      const writer = writable.getWriter();
      const encoder = new TextEncoder();

      const send = (event, data) => {
        writer.write(
          encoder.encode(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`)
        );
      };

      // Fetch current snapshot and poll for changes
      (async () => {
        let lastScrapedAt = null;
        try {
          const snapshot = await readEvents(env);
          lastScrapedAt = snapshot.scraped_at;
          send("snapshot", snapshot);

          // Poll KV every 3 seconds for updates
          while (true) {
            await new Promise((r) => setTimeout(r, 3000));
            const current = await readEvents(env);
            if (current.scraped_at !== lastScrapedAt) {
              const d = diff(
                (await readEvents(env)).events,
                current.events
              );
              send("update", {
                type: "update",
                diff: d,
                count: current.count,
              });
              lastScrapedAt = current.scraped_at;
            }
          }
        } catch (e) {
          // Client disconnected
        }
      })();

      return new Response(readable, {
        headers: {
          "content-type": "text/event-stream",
          "cache-control": "no-cache, no-transform",
          "x-accel-buffering": "no",
          "access-control-allow-origin": "*",
        },
      });
    }

    // Trigger — create a sync trigger
    if (path === "/api/trigger" && request.method === "POST") {
      const existing = await env.EVENTS_KV.get(TRIGGER_KEY, "json");
      if (existing && (existing.status === "pending" || existing.status === "running")) {
        return json({ error: "A sync is already in progress", trigger: existing }, 409);
      }

      const trigger = {
        id: crypto.randomUUID(),
        status: "pending",
        created_at: new Date().toISOString(),
        started_at: null,
        completed_at: null,
        result: null,
      };
      await env.EVENTS_KV.put(TRIGGER_KEY, JSON.stringify(trigger));
      return json(trigger, 201);
    }

    // Trigger status — poll for updates
    if (path === "/api/trigger/status" && request.method === "GET") {
      const trigger = await env.EVENTS_KV.get(TRIGGER_KEY, "json");
      if (!trigger) {
        return json({ status: "idle" });
      }
      return json(trigger);
    }

    // Trigger result — Pi reports back (bearer auth)
    if (path === "/api/trigger/result" && request.method === "POST") {
      try {
        checkAuth(request, env);
      } catch {
        return json({ error: "unauthorized" }, 401);
      }

      const body = await request.json();
      const trigger = await env.EVENTS_KV.get(TRIGGER_KEY, "json");
      if (!trigger) {
        return json({ error: "No active trigger" }, 404);
      }

      const now = new Date().toISOString();
      trigger.status = body.status;
      if (body.status === "running") {
        trigger.started_at = now;
      } else {
        trigger.completed_at = now;
      }
      trigger.result = body.result || null;
      await env.EVENTS_KV.put(TRIGGER_KEY, JSON.stringify(trigger));

      // Send failure email via Resend
      if (body.status === "failed" && env.RESEND_API_KEY) {
        const ctx = { retries: 3 };
        try {
          const errSnippet = (body.result && body.result.output)
            ? body.result.output.slice(-500)
            : "No output";
          await fetch("https://api.resend.com/emails", {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${env.RESEND_API_KEY}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              from: "SAS Scheduler <scheduler@mail.safeandsoundpost.com>",
              to: ["safeandsoundpost@gmail.com"],
              subject: "DIFUZE sync failed",
              text: `DIFUZE calendar sync failed at ${now}\n\nTrigger: ${trigger.id}\n\nOutput:\n${errSnippet}`,
            }),
          });
        } catch (e) {
          console.error("[email] Failed to send failure notification:", e);
        }
      }

      return json({ ok: true, trigger });
    }

    return json({ error: "not found" }, 404);
  },

  async scheduled(event, env) {
    // Skip if a trigger is already in progress
    const existing = await env.EVENTS_KV.get(TRIGGER_KEY, "json");
    if (existing && (existing.status === "pending" || existing.status === "running")) {
      console.log("[cron] Sync already in progress, skipping");
      return;
    }

    const trigger = {
      id: crypto.randomUUID(),
      status: "pending",
      created_at: new Date().toISOString(),
      started_at: null,
      completed_at: null,
      result: null,
    };
    await env.EVENTS_KV.put(TRIGGER_KEY, JSON.stringify(trigger));
    console.log(`[cron] Created trigger ${trigger.id}`);
  },
};
