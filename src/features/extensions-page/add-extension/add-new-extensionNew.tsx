"use client";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { LoadingIndicator } from "@/features/ui/loading";
import { Textarea } from "@/features/ui/textarea";
import { useSession } from "next-auth/react";
import { FC } from "react";
// React 18.2 では正式APIではないため experimental 名で import します
import { experimental_useFormState as useFormState } from "react-dom";
import { Button } from "../../ui/button";
import { Input } from "../../ui/input";
import { Label } from "../../ui/label";
import { ScrollArea } from "../../ui/scroll-area";
import {
  Sheet,
  SheetContent,
  SheetFooter,
  SheetHeader,
  SheetTitle,
} from "../../ui/sheet";
import { Switch } from "../../ui/switch";
import {
  AddOrUpdateExtension,
  extensionStore,
  useExtensionState,
} from "../extension-store";
import { AddFunction } from "./add-function";
import { EndpointHeader } from "./endpoint-header";
import { ErrorMessages } from "./error-messages";

interface Props {}

export const AddExtension: FC<Props> = () => {
  const { isOpened, extension } = useExtensionState();

  const { data } = useSession();
  const initialState: ServerActionResponse | undefined = undefined;

  const [formState, formAction] = useFormState(
    AddOrUpdateExtension,
    initialState
  );

  const PublicSwitch = () => {
    if (!data) return null;
    if (data?.user?.isAdmin) {
