// app/(authenticated)/api/auth/[...nextauth]/route.ts
import NextAuth from "next-auth";
import { options } from "@/features/auth-page/auth-api";

const handler = NextAuth(options);

export { handler as GET, handler as POST };