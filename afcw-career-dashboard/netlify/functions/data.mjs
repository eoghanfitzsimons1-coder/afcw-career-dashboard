import { getStore } from "@netlify/blobs";

const STORE = "afcw-dashboard";
const KEY = "state";

export default async (req) => {
  // Strong consistency so a write on one device is immediately readable
  // on the next request from another device (no 60s staleness window).
  const store = getStore({ name: STORE, consistency: "strong" });

  if (req.method === "GET") {
    try {
      const data = await store.get(KEY, { type: "json" });
      return Response.json(data || null);
    } catch (e) {
      console.error("GET failed", e);
      return Response.json({ error: String(e) }, { status: 500 });
    }
  }

  if (req.method === "POST") {
    let payload;
    try {
      payload = await req.json();
    } catch {
      return new Response("Invalid JSON", { status: 400 });
    }
    try {
      await store.setJSON(KEY, payload);
      return Response.json({ ok: true });
    } catch (e) {
      console.error("POST failed", e);
      return Response.json({ error: String(e) }, { status: 500 });
    }
  }

  return new Response("Method not allowed", { status: 405 });
};

export const config = {
  path: "/api/data"
};