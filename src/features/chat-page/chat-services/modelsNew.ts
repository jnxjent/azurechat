 export interface ChatMessageModel {
   id: string;
   role: "user" | "assistant" | "tool" | "function" | "system";
-  content: string;
+  content: string;
+  message?: string;        // ★ 追加：BEの .message を保持
+  message_format?: string; // ★ 追加："markdown" 等（描画分岐で使う）
   name: string;
   multiModalImage?: string;
   createdAt: Date;
   isDeleted: boolean;
   threadId: string;
   type: string;
   userId: string;
 }
