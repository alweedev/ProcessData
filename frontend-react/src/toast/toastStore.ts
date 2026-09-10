export type ToastType = "success" | "info" | "danger";

export interface ToastItem {
  id: number;
  message: string;
  type: ToastType;
}

type Listener = (items: ToastItem[]) => void;

let items: ToastItem[] = [];
let nextId = 1;
const listeners = new Set<Listener>();

function emit() {
  for (const listener of listeners) listener(items);
}

export function pushToast(message: string, type: ToastType = "success") {
  const id = nextId++;
  items = [...items, { id, message, type }];
  emit();
  window.setTimeout(() => dismissToast(id), 3600);
}

export function dismissToast(id: number) {
  items = items.filter((item) => item.id !== id);
  emit();
}

export function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  listener(items);
  return () => {
    listeners.delete(listener);
  };
}
