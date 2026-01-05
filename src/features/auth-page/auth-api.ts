import NextAuth, { NextAuthOptions } from "next-auth";
import AzureADProvider from "next-auth/providers/azure-ad";
import CredentialsProvider from "next-auth/providers/credentials";
import GitHubProvider from "next-auth/providers/github";
import { Provider } from "next-auth/providers/index";
import { hashValue } from "./helpers";

const configureIdentityProvider = () => {
  const providers: Array<Provider> = [];

  const adminEmails = process.env.ADMIN_EMAIL_ADDRESS
    ?.split(",")
    .map((email) => email.toLowerCase().trim());

  // ------------------------
  // GitHub Provider
  // ------------------------
  if (process.env.AUTH_GITHUB_ID && process.env.AUTH_GITHUB_SECRET) {
    providers.push(
      GitHubProvider({
        clientId: process.env.AUTH_GITHUB_ID!,
        clientSecret: process.env.AUTH_GITHUB_SECRET!,
        async profile(profile) {
          return {
            ...profile,
            isAdmin: adminEmails?.includes(profile.email?.toLowerCase()),
          };
        },
      })
    );
  }

  // ------------------------
  // Azure AD Provider
  // ------------------------
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
        async profile(profile) {
          return {
            ...profile,
            // NextAuth requires id
            id: profile.sub,
            isAdmin:
              adminEmails?.includes(profile.email?.toLowerCase()) ||
              adminEmails?.includes(
                profile.preferred_username?.toLowerCase()
              ),
          };
        },
      })
    );
  }

  // ------------------------
  // Local dev Credentials Provider
  // ------------------------
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
          const email = `${username}@localhost`;

          const user = {
            id: hashValue(email),
            name: username,
            email,
            isAdmin: false,
            image: "",
          };

          console.log(
            "=== DEV USER LOGGED IN ===\n",
            JSON.stringify(user, null, 2)
          );

          return user;
        },
      })
    );
  }

  return providers;
};

export const options: NextAuthOptions = {
  secret: process.env.NEXTAUTH_SECRET,
  providers: [...configureIdentityProvider()],

  callbacks: {
    /**
     * JWT に必要なユーザー情報を永続化
     */
    async jwt({ token, user, profile }) {
      // 初回ログイン時（user が来る）
      if (user) {
        token.isAdmin = (user as any).isAdmin ?? token.isAdmin;
        token.email = (user as any).email ?? token.email;
        token.name = (user as any).name ?? token.name;
        token.picture = (user as any).image ?? token.picture;
      }

      // Azure AD / OIDC 対策（profile からも拾う）
      const p: any = profile ?? {};
      const profileEmail =
        p.email ||
        p.preferred_username ||
        p.upn ||
        null;

      if (!token.email && profileEmail) {
        token.email = String(profileEmail);
      }

      return token;
    },

    /**
     * JWT → session.user へ復元
     */
    async session({ session, token }) {
      if (session.user) {
        (session.user as any).isAdmin = token.isAdmin as boolean;
        session.user.email = token.email as string;
        session.user.name = token.name as string;
        (session.user as any).image = token.picture as string;
      }
      return session;
    },
  },

  session: {
    strategy: "jwt",
  },
};

export const handlers = NextAuth(options);
