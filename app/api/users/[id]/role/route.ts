import { eq } from "drizzle-orm";
import { getDb } from "@/db";
import { users } from "@/db/schema";
import { currentUser, roles } from "@/lib/current-user";
export async function POST(request: Request, context: { params: Promise<{ id: string }> }) {
  const admin = await currentUser(); if (!admin || admin.role !== "admin") return Response.json({ error: "دسترسی مدیر لازم است" }, { status: 403 });
  const { id } = await context.params, body = await request.json() as { role?: string }; if (!roles.includes(body.role as never)) return Response.json({ error: "نقش نامعتبر است" }, { status: 400 });
  await getDb().update(users).set({ role: body.role!, updatedAt: new Date().toISOString() }).where(eq(users.id, Number(id))); return Response.json({ ok: true });
}
