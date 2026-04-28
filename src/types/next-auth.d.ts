import { DefaultSession } from "next-auth";

// https://next-auth.js.org/getting-started/typescript#module-augmentation

declare module "next-auth" {
  interface Session {
    user: {
      isAdmin: boolean;
      slRole?: "global_admin" | "dept_admin" | "dept_member";
      slDept?: string;
    } & DefaultSession["user"];
  }

  interface Token {
    isAdmin: boolean;
    slRole?: "global_admin" | "dept_admin" | "dept_member";
    slDept?: string;
  }

  interface User {
    isAdmin: boolean;
  }
}
