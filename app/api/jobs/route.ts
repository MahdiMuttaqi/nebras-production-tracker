import { desc, eq } from "drizzle-orm";
import { getDb } from "@/db";
import { jobs, stageEvents } from "@/db/schema";
import { currentUser } from "@/lib/current-user";
export async function GET() {
  const user = await currentUser(); if (!user || user.role === "pending") return Response.json({ jobs: [] }); const db = getDb();
  const rows = user.role === "admin" ? await db.select().from(jobs).orderBy(desc(jobs.updatedAt)).limit(500)
    : await db.select().from(jobs).where(eq(jobs.currentStage, user.role)).orderBy(desc(jobs.updatedAt)).limit(200);
  return Response.json({ jobs: rows });
}
export async function POST(request: Request) {
  const user = await currentUser(); if (!user || user.role !== "admin") return Response.json({ error: "دسترسی مدیر لازم است" }, { status: 403 });
  const payload = await request.json() as { rows?: Array<{code?: string; quantity?: number; customer?: string; occasion?: string; notes?: string}> };
  const rows = (payload.rows ?? []).filter((r) => r.code?.trim()).slice(0, 500); if (!rows.length) return Response.json({ error: "حداقل یک کد وارد کنید" }, { status: 400 });
  const now = new Date().toISOString(), db = getDb();
  const created = await db.insert(jobs).values(rows.map((r) => ({ code: r.code!.trim(), quantity: Math.max(1, Number(r.quantity) || 1), customer: r.customer?.trim() ?? "", occasion: r.occasion?.trim() ?? "", notes: r.notes?.trim() ?? "", createdBy: user.email, createdAt: now, updatedAt: now }))).returning();
  await db.insert(stageEvents).values(created.map((j) => ({ jobId: j.id, fromStage: "new", toStage: "plotter", actorEmail: user.email, actorRole: user.role, note: "ثبت سفارش", createdAt: now })));
  return Response.json({ count: created.length }, { status: 201 });
}
