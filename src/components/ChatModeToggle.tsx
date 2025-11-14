
"use client";
// =============================================
// src/components/chatModeToggle.tsx (v11)
// Purpose:
//  - 3段階サイクル（標準 → 熟考 → 即答 → 標準）※auto 廃止
//  - reasoning_effort の定義は別ファイル（reasoning-utils.ts）で管理
//  - Cookie 保存（thinkingMode=standard|thinking|fast）でサーバへ伝達
//  - ボタンは中央に日本語ラベルのみ表示
//  - デザイン調整：高さを低く（h-8）、即答は熟考よりもさらに薄い緑で白文字
// =============================================

import React, { useCallback, useState } from "react";

/** UI表示モード（auto は廃止） */
export type ThinkingMode = "standard" | "thinking" | "fast"; // 標準／熟考／即答

export interface ModeCycleButtonProps {
  value: ThinkingMode;                    // 現在のモード
  onChange: (mode: ThinkingMode) => void; // 変更通知
  className?: string;
}

const labels: Record<ThinkingMode, string> = {
  standard: "標準",
  thinking: "熟考",
  fast: "即答",
};

// 次モード（標準 → 熟考 → 即答 → 標準）
function next(mode: ThinkingMode): ThinkingMode {
  if (mode === "standard") return "thinking";
  if (mode === "thinking") return "fast";
  return "standard";
}

// 前モード（←キー用）
function prev(mode: ThinkingMode): ThinkingMode {
  if (mode === "standard") return "fast";
  if (mode === "thinking") return "standard";
  return "thinking";
}

// Cookie 書き込み（365日）
function setModeCookie(mode: ThinkingMode) {
  try {
    document.cookie = `thinkingMode=${mode}; Path=/; Max-Age=${60 * 60 * 24 * 365}`;
  } catch (e) {
    console.warn("Failed to set thinkingMode cookie:", e);
  }
}

// ★ named export：ボタン
export function ModeCycleButton({ value, onChange, className }: ModeCycleButtonProps) {
  const apply = (m: ThinkingMode) => {
    onChange(m);
    setModeCookie(m);
  };

  const handleClick = () => apply(next(value));

  const handleKeyDown: React.KeyboardEventHandler<HTMLButtonElement> = (e) => {
    if (e.key === "Enter" || e.key === " " || e.key === "ArrowRight") {
      e.preventDefault();
      apply(next(value));
    }
    if (e.key === "ArrowLeft") {
      e.preventDefault();
      apply(prev(value));
    }
  };

  // デザイン設定
  const base =
    "inline-flex h-8 min-w-20 items-center justify-center rounded-md px-3 text-sm " +
    "font-medium shadow-sm transition active:scale-[.98] focus:outline-none " +
    "focus:ring-2 focus:ring-offset-2 ring-1 text-center select-none";

  // 色設計：
  // - 熟考：濃い緑（bg-green-600）
  // - 即答：もっと薄い緑（bg-green-300）＋ 白文字
  // - 標準：既存どおりスレート系
  const themeByMode: Record<ThinkingMode, string> = {
    standard: "bg-slate-200 text-slate-900 ring-slate-300 hover:bg-slate-300",
    thinking: "bg-green-600 text-white ring-green-300 hover:bg-green-700",
    fast: "bg-green-300 text-black ring-green-200 hover:bg-green-400", // ← 薄緑＋白文字
  };

  return (
    <button
      type="button"
      aria-label={`思考モード: ${labels[value]}`}
      title={`クリックで切替（現在: ${labels[value]}）`}
      onClick={handleClick}
      onKeyDown={handleKeyDown}
      className={[base, themeByMode[value], className].filter(Boolean).join(" ")}
    >
      <span className="leading-none">{labels[value]}</span>
    </button>
  );
}

// ローカル状態フック（初期値: 標準）
function useModeState(initial: ThinkingMode = "standard") {
  const [mode, setMode] = useState<ThinkingMode>(initial);
  const set = useCallback((m: ThinkingMode) => setMode(m), []);
  return { mode, setMode: set };
}

// ★ default export：配置用の簡易コンポーネント
export default function ChatModeToggle() {
  const { mode, setMode } = useModeState("standard");
  return (
    <div className="flex items-center gap-2">
      <ModeCycleButton value={mode} onChange={setMode} />
    </div>
  );
}
