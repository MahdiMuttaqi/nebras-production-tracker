import { desc, eq, inArray, or } from "drizzle-orm";
import { getDb } from "@/db";
import { jobs, stageEvents } from "@/db/schema";
import { currentUser } from "@/lib/current-user";
export async function GET() {
  const user = await currentUser(); if (!user || user.role === "pending") return Response.json({ jobs: [] }); const db = getDb();
  const rows = user.role === "admin" ? await db.select().from(jobs).orderBy(desc(jobs.updatedAt)).limit(500)
    : await db.select().from(jobs).where(user.role === "sewing1" ? or(eq(jobs.currentStage, "sewing1"), eq(jobs.currentStage, "sewing")) : eq(jobs.currentStage, user.role)).orderBy(desc(jobs.updatedAt)).limit(200);
  const ids = new Set(rows.map((job) => job.id));
  const events = ids.size ? await db.select().from(stageEvents).where(inArray(stageEvents.jobId, [...ids])).orderBy(desc(stageEvents.createdAt)).limit(5000) : [];
  const byJob = new Map<number, typeof events>();
  for (const event of events) byJob.set(event.jobId, [...(byJob.get(event.jobId) ?? []), event]);
  return Response.json({ jobs: rows.map((job) => ({ ...job, events: (byJob.get(job.id) ?? []).slice(0, 30) })) });
}
export async function POST(request: Request) {
  const user = await currentUser(); if (!user || user.role !== "admin") return Response.json({ error: "دسترسی مدیر لازم است" }, { status: 403 });
  const payload = await request.json() as { rows?: Array<{code?: string; quantity?: number; customer?: string; notes?: string}> };
  const rows = (payload.rows ?? []).filter((r) => r.code?.trim()).slice(0, 500); if (!rows.length) return Response.json({ error: "حداقل یک کد وارد کنید" }, { status: 400 });
  const now = new Date().toISOString(), db = getDb();
  const created = await db.insert(jobs).values(rows.map((r) => ({ code: r.code!.trim(), quantity: Math.max(1, Number(r.quantity) || 1), customer: r.customer?.trim() ?? "", occasion: "", notes: r.notes?.trim() ?? "", createdBy: user.email, createdAt: now, updatedAt: now }))).returning();
  await db.insert(stageEvents).values(created.map((j) => ({ jobId: j.id, fromStage: "new", toStage: "plotter", actorEmail: user.email, actorRole: user.role, note: "ثبت سفارش", createdAt: now })));
  return Response.json({ count: created.length }, { status: 201 });
}
