// src/features/auth-page/auth-api.ts
import NextAuth, { NextAuthOptions } from "next-auth";
import AzureADProvider from "next-auth/providers/azure-ad";
import CredentialsProvider from "next-auth/providers/credentials";
import GitHubProvider from "next-auth/providers/github";
import type { Provider } from "next-auth/providers/index";  // ← 修正
import { hashValue } from "@/features/auth-page/helpers";
import { resolveSlAccess } from "@/lib/sl-dept";

const AAD_SCOPE = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "User.Read",
  "Files.ReadWrite",
].join(" ");

const configureIdentityProvider = () => {
  const providers: Array<Provider> = [];

  const adminEmails = process.env.SL_ADMIN_EMAILS
    ?.split(",")
    .map((email) => email.toLowerCase().trim());

  if (process.env.AUTH_GITHUB_ID && process.env.AUTH_GITHUB_SECRET) {
    providers.push(
      GitHubProvider({
        clientId: process.env.AUTH_GITHUB_ID!,
        clientSecret: process.env.AUTH_GITHUB_SECRET!,
        async profile(profile) {
          return {
            ...profile,
            isAdmin: adminEmails?.includes(profile.email?.toLowerCase()),
          } as any;
        },
      })
    );
  }

  if (
    process.env.AZURE_AD_CLIENT_ID &&
    process.env.AZURE_AD_CLIENT_SECRET &&
    process.env.AZURE_AD_TENANT_ID
  ) {
    providers.push(
      AzureADProvider({
        clientId: process.env.AZURE_AD_CLIENT_ID!,
        clientSecret: process.env.AZURE_AD_CLIENT_SECRET!,
        tenantId: process.env.AZURE_AD_TENANT_ID!,
        authorization: {
          params: { scope: AAD_SCOPE },
        },
        async profile(profile) {
          const p: any = profile ?? {};
          return {
            ...p,
            id: p.sub,
            isAdmin:
              adminEmails?.includes(String(p.email || "").toLowerCase()) ||
              adminEmails?.includes(
                String(p.preferred_username || "").toLowerCase()
              ),
          };
        },
      })
    );
  }

  if (process.env.NODE_ENV === "development") {
    providers.push(
      CredentialsProvider({
        name: "localdev",
        credentials: {
          username: { label: "Username", type: "text", placeholder: "dev" },
          password: { label: "Password", type: "password" },
        },
        async authorize(credentials) {
          const username = credentials?.username || "dev";
          // NEXT_PUBLIC_DEV_USER_EMAIL が設定されていればそのメールを使う（SL ロール判定のため）
          const email =
            process.env.NEXT_PUBLIC_DEV_USER_EMAIL ?? `${username}@localhost`;
          const adminEmails = process.env.SL_ADMIN_EMAILS
            ?.split(",")
            .map((e) => e.toLowerCase().trim()) ?? [];
          const user = {
            id: hashValue(email),
            name: username,
            email,
            isAdmin: adminEmails.includes(email.toLowerCase()),
            image: "",
          };
          console.log("=== DEV USER LOGGED IN ===\n", JSON.stringify(user, null, 2));
          return user as any;
        },
      })
    );
  }

  return providers;
};

async function refreshAzureADAccessToken(token: any) {
  const tenantId = process.env.AZURE_AD_TENANT_ID;
  const clientId = process.env.AZURE_AD_CLIENT_ID;
  const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) return token;
  if (!token?.refreshToken) return token;

  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          grant_type: "refresh_token",
          refresh_token: token.refreshToken,
          scope: AAD_SCOPE,
        }),
      }
    );

    const data: any = await res.json().catch(() => ({}));

    if (!res.ok) {
      return { ...token, refreshError: data };
    }

    const nowSec = Math.floor(Date.now() / 1000);
    const expiresIn = Number(data.expires_in ?? 3600);

    return {
      ...token,
      accessToken: data.access_token ?? token.accessToken,
      accessTokenExpiresAt: nowSec + expiresIn,
      refreshToken: data.refresh_token ?? token.refreshToken,
      refreshError: undefined,
    };
  } catch (e) {
    return { ...token, refreshError: String(e) };
  }
}

export const options: NextAuthOptions = {
  secret: process.env.NEXTAUTH_SECRET,
  providers: [...configureIdentityProvider()],

  callbacks: {
    async jwt({ token, user, profile, account }) {
      if (user) {
        (token as any).isAdmin = (user as any).isAdmin ?? (token as any).isAdmin;
        token.email = (user as any).email ?? token.email;
        token.name = (user as any).name ?? token.name;
        (token as any).picture = (user as any).image ?? (token as any).picture;

        // ★ SlRole / SlDept をトークンに保存（サインイン時に1回計算）
        try {
          const emailForRole = String((user as any).email ?? token.email ?? "");
          const access = resolveSlAccess(emailForRole);
          (token as any).slRole = access.role;
          (token as any).slDept = access.dept;
        } catch {
          // SL_DEPTS / SL_DEPT_DEFAULT 未設定時はスキップ（バッジ非表示）
        }
      }

      const p: any = profile ?? {};
      const profileEmail = p.email || p.preferred_username || p.upn || null;
      if (!token.email && profileEmail) token.email = String(profileEmail);

      if (account?.provider === "azure-ad") {
        if (account.access_token) (token as any).accessToken = account.access_token;
        if (account.expires_at)
          (token as any).accessTokenExpiresAt = Number(account.expires_at);
        const rt = (account as any).refresh_token;
        if (rt) (token as any).refreshToken = rt;
      }

      if (!process.env.AZURE_AD_TENANT_ID) return token;

      const accessToken = (token as any).accessToken as string | undefined;
      const expiresAt = (token as any).accessTokenExpiresAt as number | undefined;
      if (!accessToken || !expiresAt) return token;

      const nowSec = Math.floor(Date.now() / 1000);
      if (nowSec >= expiresAt - 60) {
        return (await refreshAzureADAccessToken(token as any)) as any;
      }

      return token;
    },

    async session({ session, token }) {
      if (session.user) {
        (session.user as any).isAdmin = (token as any).isAdmin as boolean;
        session.user.email = token.email as string;
        session.user.name = token.name as string;
        (session.user as any).image = (token as any).picture as string;
        session.user.slRole = (token as any).slRole;
        session.user.slDept = (token as any).slDept;
      }
      (session as any).accessToken = (token as any).accessToken;
      (session as any).accessTokenExpiresAt = (token as any).accessTokenExpiresAt;
      (session as any).refreshError = (token as any).refreshError;
      return session;
    },
  },
  session: {
    strategy: "jwt",
  },
};
