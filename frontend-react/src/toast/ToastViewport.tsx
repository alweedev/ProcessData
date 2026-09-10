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
    <>
      {items.map((item) => (
        <div
          key={item.id}
          role="status"
          aria-live="polite"
          className={`toast show align-items-center text-white ${BG[item.type]} border-0 mb-2`}
        >
          <div className="d-flex">
            <div className="toast-body">{item.message}</div>
            <button
              type="button"
              className="btn-close btn-close-white me-2 m-auto"
              aria-label="Fechar"
              onClick={() => dismissToast(item.id)}
            />
          </div>
        </div>
      ))}
    </>
  );
}
