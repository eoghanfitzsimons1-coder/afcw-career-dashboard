// netlify/functions/sync-runs.mjs
import { getStore } from "@netlify/blobs";

export default async (req, context) => {
  const store = getStore("run-tracker-data");
  const key = "runs-data";

  // GET — fetch runs
  if (req.method === "GET") {
    try {
      const data = await store.get(key, { type: "json" });
      return new Response(JSON.stringify(data || []), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    } catch (error) {
      console.error("Error reading blobs:", error);
      return new Response(JSON.stringify({ error: "Failed to read data" }), {
        status: 500,
        headers: { "Content-Type": "application/json" },
      });
    }
  }

  // POST — save runs
  if (req.method === "POST") {
    try {
      const runs = await req.json();
      await store.setJSON(key, runs, { consistency: "strong" });
      return new Response(JSON.stringify({ success: true }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    } catch (error) {
      console.error("Error writing blobs:", error);
      return new Response(JSON.stringify({ error: "Failed to save data" }), {
        status: 500,
        headers: { "Content-Type": "application/json" },
      });
    }
  }

  return new Response("Method Not Allowed", { status: 405 });
};

export const config = {
  path: "/api/runs",
};