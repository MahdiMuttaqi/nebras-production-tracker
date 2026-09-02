import { allUsers, currentUser } from "@/lib/current-user";
export async function GET() { const user = await currentUser(); if (!user) return Response.json({ error: "ورود لازم است" }, { status: 401 }); return Response.json({ user, users: user.role === "admin" ? await allUsers() : [] }); }
