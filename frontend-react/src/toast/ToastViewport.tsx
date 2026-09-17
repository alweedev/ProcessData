import { useEffect, useState } from "react";
import { dismissToast, subscribe, type ToastItem } from "./toastStore";

const STYLE: Record<ToastItem["type"], { wrap: string; icon: string }> = {
  success: { wrap: "border-success/30 bg-success-soft text-success", icon: "✓" },
  info: { wrap: "border-info/30 bg-info-soft text-info", icon: "i" },
  danger: { wrap: "border-danger/30 bg-danger-soft text-danger", icon: "!" },
};

export function ToastViewport() {
  const [items, setItems] = useState<ToastItem[]>([]);

  useEffect(() => subscribe(setItems), []);

  return (
    <div className="flex flex-col items-end gap-2">
      {items.map((item) => (
        <div
          key={item.id}
          role="status"
          aria-live="polite"
          className={`flex items-center gap-2.5 rounded-control border px-3.5 py-2 text-sm shadow-card animate-[pd-toast-in_160ms_ease-out] ${STYLE[item.type].wrap}`}
        >
          <span aria-hidden="true" className="shrink-0 text-base font-bold leading-none">
            {STYLE[item.type].icon}
          </span>
          <span className="toast-body">{item.message}</span>
          <button
            type="button"
            aria-label="Fechar"
            onClick={() => dismissToast(item.id)}
            className="ml-1 opacity-60 transition-opacity hover:opacity-100"
          >
            ✕
          </button>
        </div>
      ))}
    </div>
  );
}
