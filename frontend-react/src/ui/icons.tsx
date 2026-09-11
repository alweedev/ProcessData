import type { ReactNode } from "react";

interface IconProps {
  className?: string;
}

function svg(path: ReactNode, className = "h-4 w-4") {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.8}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {path}
    </svg>
  );
}

export function IconHome({ className }: IconProps) {
  return svg(
    <>
      <path d="M3 10.5 12 3l9 7.5" />
      <path d="M5 9.5V21h14V9.5" />
      <path d="M9.5 21v-6h5v6" />
    </>,
    className,
  );
}

export function IconUpload({ className }: IconProps) {
  return svg(
    <>
      <path d="M12 15V4" />
      <path d="m7 9 5-5 5 5" />
      <path d="M5 15v4a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-4" />
    </>,
    className,
  );
}

export function IconUserMinus({ className }: IconProps) {
  return svg(
    <>
      <circle cx="9" cy="8" r="3.5" />
      <path d="M3.5 20c0-3.5 2.7-5.5 5.5-5.5s5.5 2 5.5 5.5" />
      <path d="M16.5 12h5" />
    </>,
    className,
  );
}

export function IconSitemap({ className }: IconProps) {
  return svg(
    <>
      <rect x="9" y="3" width="6" height="5" rx="1" />
      <rect x="3" y="16" width="6" height="5" rx="1" />
      <rect x="15" y="16" width="6" height="5" rx="1" />
      <path d="M12 8v4M6 16v-2h12v2" />
    </>,
    className,
  );
}

export function IconClock({ className }: IconProps) {
  return svg(
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M12 7v5l3 2" />
    </>,
    className,
  );
}

export function IconSearch({ className }: IconProps) {
  return svg(
    <>
      <circle cx="11" cy="11" r="7" />
      <path d="m20 20-3.5-3.5" />
    </>,
    className,
  );
}

export function IconInbox({ className }: IconProps) {
  return svg(
    <>
      <path d="M4 13h4l1.5 3h5L16 13h4" />
      <path d="M5 6h14l2 7v5a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-5Z" />
    </>,
    className,
  );
}

export function IconKeyboard({ className }: IconProps) {
  return svg(
    <>
      <rect x="2.5" y="6" width="19" height="12" rx="2" />
      <path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M8 14h8" />
    </>,
    className,
  );
}
