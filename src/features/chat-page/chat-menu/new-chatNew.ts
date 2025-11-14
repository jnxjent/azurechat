"use client";

import { Button } from "@/features/ui/button";
import { LoadingIndicator } from "@/features/ui/loading";
import { Plus } from "lucide-react";
// React 18.2 では正式APIではないため experimental 名で import します
import { experimental_useFormStatus as useFormStatus } from "react-dom";

export const NewChat = () => {
  const { pending } = useFormStatus();

  return (
    <Button
      aria-disabled={pending}
      size="default"
      className="flex gap-2"
      variant="outline"
    >
      {pending ? <LoadingIndicator isLoading={pending} /> : <Plus size={18} />}
      New Chat
    </Button>
  );
};
