import { useEffect, useState } from "react";
import { dismissToast, subscribe, type ToastItem } from "./toastStore";

const BG: Record<ToastItem["type"], string> = {
  success: "bg-success",
  info: "bg-info",
  danger: "bg-danger",
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
          className={`flex items-center gap-3 rounded-lg px-4 py-2 text-sm text-white shadow-lg ${BG[item.type]}`}
        >
          <span className="toast-body">{item.message}</span>
          <button
            type="button"
            aria-label="Fechar"
            onClick={() => dismissToast(item.id)}
            className="text-white/80 hover:text-white"
          >
            ✕
          </button>
        </div>
      ))}
    </div>
  );
}
