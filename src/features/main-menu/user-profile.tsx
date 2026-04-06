"use client";

import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/ui/dropdown-menu";
import { menuIconProps } from "@/ui/menu";
import { CircleUserRound, LogOut } from "lucide-react";
import { signOut, useSession } from "next-auth/react";
import { Avatar, AvatarImage } from "../ui/avatar";
import { ThemeToggle } from "./theme-toggle";

const NON_SP_DEPTS = (process.env.NEXT_PUBLIC_SL_DEPT_NON_SP ?? "others")
  .split(",")
  .map((s) => s.trim().toLowerCase())
  .filter(Boolean);

function getRoleLabel(
  slRole?: "global_admin" | "dept_admin" | "dept_member" | null,
  slDept?: string | null
): string {
  const deptLower = (slDept ?? "").trim().toLowerCase();
  if (NON_SP_DEPTS.includes(deptLower)) return "";
  const dept = deptLower.toUpperCase();
  if (slRole === "global_admin") return "全社管理者";
  if (slRole === "dept_admin") return `${dept}管理者`;
  if (slRole === "dept_member") return `${dept}員`;
  return "";
}

export const UserProfile = () => {
  const { data: session } = useSession();
  const roleLabel = getRoleLabel(session?.user?.slRole, session?.user?.slDept);

  return (
    <div className="flex flex-col items-center gap-1 w-full">
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          {session?.user?.image ? (
            <Avatar className="">
              <AvatarImage
                src={session?.user?.image!}
                alt={session?.user?.name!}
              />
            </Avatar>
          ) : (
            <CircleUserRound {...menuIconProps} role="button" />
          )}
        </DropdownMenuTrigger>
        <DropdownMenuContent side="right" className="w-56" align="end">
          <DropdownMenuLabel className="font-normal">
            <div className="flex flex-col gap-2">
              <p className="text-sm font-medium leading-none">
                {session?.user?.name}
              </p>
              <p className="text-xs leading-none text-muted-foreground">
                {session?.user?.email}
              </p>
              {roleLabel && (
                <span className="inline-block self-start bg-primary text-primary-foreground text-xs font-medium px-2 py-0.5 rounded-full leading-tight">
                  {roleLabel}
                </span>
              )}
            </div>
          </DropdownMenuLabel>
          <DropdownMenuSeparator />
          <DropdownMenuLabel className="font-normal">
            <div className="flex flex-col gap-1">
              <p className="text-sm font-medium leading-none">Switch themes</p>
              <ThemeToggle />
            </div>
          </DropdownMenuLabel>
          <DropdownMenuSeparator />
          <DropdownMenuItem
            className="flex gap-2"
            onClick={() => signOut({ callbackUrl: "/" })}
          >
            <LogOut {...menuIconProps} size={18} />
            <span>Log out</span>
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
      {roleLabel && (
        <span className="bg-primary text-primary-foreground text-[8px] font-medium leading-tight text-center rounded px-0.5 py-0.5 w-full block">
          {roleLabel}
        </span>
      )}
    </div>
  );
};
