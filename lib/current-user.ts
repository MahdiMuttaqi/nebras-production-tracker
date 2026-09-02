import { asc, count, eq } from "drizzle-orm";
import { getDb } from "@/db";
import { users } from "@/db/schema";
import { getChatGPTUser } from "@/app/chatgpt-auth";
export const roles = ["pending", "admin", "plotter", "calendar", "laser", "sewing", "warehouse"] as const;
export async function currentUser() {
  const identity = await getChatGPTUser(); if (!identity) return null; const db = getDb();
  const found = await db.select().from(users).where(eq(users.email, identity.email)).limit(1); if (found[0]) return found[0];
  const [{ value }] = await db.select({ value: count() }).from(users); const now = new Date().toISOString();
  await db.insert(users).values({ email: identity.email, displayName: identity.displayName, role: value === 0 ? "admin" : "pending", createdAt: now, updatedAt: now }).onConflictDoNothing();
  return (await db.select().from(users).where(eq(users.email, identity.email)).limit(1))[0] ?? null;
}
export async function allUsers() { return getDb().select().from(users).orderBy(asc(users.createdAt)); }
