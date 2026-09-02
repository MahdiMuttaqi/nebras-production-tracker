import { index, integer, sqliteTable, text, uniqueIndex } from "drizzle-orm/sqlite-core";
export const users = sqliteTable("users", {
  id: integer("id").primaryKey({ autoIncrement: true }), email: text("email").notNull(), displayName: text("display_name").notNull(),
  role: text("role").notNull().default("pending"), createdAt: text("created_at").notNull(), updatedAt: text("updated_at").notNull(),
}, (t) => [uniqueIndex("users_email_uq").on(t.email)]);
export const jobs = sqliteTable("jobs", {
  id: integer("id").primaryKey({ autoIncrement: true }), code: text("code").notNull(), quantity: integer("quantity").notNull().default(1),
  customer: text("customer").notNull().default(""), occasion: text("occasion").notNull().default(""), route: text("route").notNull().default("undecided"),
  currentStage: text("current_stage").notNull().default("plotter"), status: text("status").notNull().default("active"), notes: text("notes").notNull().default(""),
  createdBy: text("created_by").notNull(), createdAt: text("created_at").notNull(), updatedAt: text("updated_at").notNull(), completedAt: text("completed_at"),
}, (t) => [index("jobs_stage_status_idx").on(t.currentStage, t.status), index("jobs_updated_idx").on(t.updatedAt)]);
export const stageEvents = sqliteTable("stage_events", {
  id: integer("id").primaryKey({ autoIncrement: true }), jobId: integer("job_id").notNull().references(() => jobs.id, { onDelete: "cascade" }),
  fromStage: text("from_stage").notNull(), toStage: text("to_stage").notNull(), actorEmail: text("actor_email").notNull(),
  actorRole: text("actor_role").notNull(), note: text("note").notNull().default(""), createdAt: text("created_at").notNull(),
}, (t) => [index("events_job_idx").on(t.jobId, t.createdAt)]);
