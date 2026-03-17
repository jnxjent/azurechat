// src/features/ui/markdown/paragraph.tsx
import { cn } from "@/ui/lib";

export const Paragraph = ({
  children,
  className,
}: {
  children: React.ReactNode;
  className?: string;
}) => {
  return <div className={cn(className, "leading-relaxed")}>{children}</div>;
  //                                    ^^^^^^^^^^^^^^^^
  // py-3 を削除、leading-relaxed のみ残す
};

export const paragraph = {
  render: "Paragraph",
};