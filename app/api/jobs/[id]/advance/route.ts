import { eq } from "drizzle-orm";
import { getDb } from "@/db";
import { jobs, stageEvents } from "@/db/schema";
import { currentUser } from "@/lib/current-user";
const next: Record<string, string> = { plotter: "calendar", sewing: "warehouse", sewing1: "warehouse", sewing2: "warehouse", sewing3: "warehouse", warehouse: "completed" };
const tailors = ["sewing1", "sewing2", "sewing3"];
export async function POST(request: Request, context: { params: Promise<{ id: string }> }) {
  const user = await currentUser(); if (!user || user.role === "pending") return Response.json({ error: "دسترسی ندارید" }, { status: 403 });
  const { id } = await context.params, db = getDb(); const [job] = await db.select().from(jobs).where(eq(jobs.id, Number(id))).limit(1);
  if (!job || job.status !== "active") return Response.json({ error: "سفارش فعال پیدا نشد" }, { status: 404 });
  if (user.role !== "admin" && user.role !== job.currentStage && !(user.role === "sewing1" && job.currentStage === "sewing")) return Response.json({ error: "این سفارش در واحد شما نیست" }, { status: 403 });
  const body = await request.json().catch(() => ({})) as { route?: "laser" | "direct"; tailor?: string; note?: string }; let to = next[job.currentStage], route = job.route;
  if (job.currentStage === "calendar") {
    if (!body.route) return Response.json({ error: "مسیر بعدی را انتخاب کنید" }, { status: 400 });
    route = body.route;
    if (body.route === "laser") to = "laser";
    else if (body.tailor && tailors.includes(body.tailor)) to = body.tailor;
    else return Response.json({ error: "خیاط مقصد را انتخاب کنید" }, { status: 400 });
  }
  if (job.currentStage === "laser") {
    if (!body.tailor || !tailors.includes(body.tailor)) return Response.json({ error: "خیاط مقصد را انتخاب کنید" }, { status: 400 });
    to = body.tailor;
  }
  if (!to) return Response.json({ error: "مرحله بعد مشخص نیست" }, { status: 400 }); const now = new Date().toISOString();
  await db.update(jobs).set({ currentStage: to, route, status: to === "completed" ? "completed" : "active", updatedAt: now, completedAt: to === "completed" ? now : null }).where(eq(jobs.id, job.id));
  const [event] = await db.insert(stageEvents).values({ jobId: job.id, fromStage: job.currentStage, toStage: to, actorEmail: user.email, actorRole: user.role, note: body.note?.trim() ?? "", createdAt: now }).returning();
  const [updated] = await db.select().from(jobs).where(eq(jobs.id, job.id)).limit(1);
  return Response.json({ ok: true, next: to, job: updated, event });
}
