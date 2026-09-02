import { requireChatGPTUser } from "./chatgpt-auth";
import TrackerClient from "./tracker-client";
export default async function Page() { await requireChatGPTUser("/"); return <TrackerClient />; }
