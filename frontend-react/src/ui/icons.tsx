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

export function IconSun({ className }: IconProps) {
  return svg(
    <>
      <circle cx="12" cy="12" r="4" />
      <path d="M12 4V2m0 20v-2m8-8h2M2 12h2m13.657-6.343l1.414-1.414M4.929 19.071l1.414-1.414m0-11.314L4.93 4.93m13.657 13.657l1.414 1.414" />
    </>,
    className,
  );
}

export function IconMoon({ className }: IconProps) {
  return svg(<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79Z" />, className);
}

export function IconCheck({ className }: IconProps) {
  return svg(<path d="m5 12.5 4.5 4.5L19 7.5" />, className);
}

export function IconX({ className }: IconProps) {
  return svg(<path d="M6 6l12 12M18 6 6 18" />, className);
}

export function IconFile({ className }: IconProps) {
  return svg(
    <>
      <path d="M14 3H7a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7z" />
      <path d="M14 3v4h4" />
      <path d="m9.5 12.5 5 5m0-5-5 5" />
    </>,
    className,
  );
}

export function IconAlert({ className }: IconProps) {
  return svg(
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M12 8v5" />
      <path d="M12 16.5h.01" />
    </>,
    className,
  );
}
