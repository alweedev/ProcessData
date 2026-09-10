import { ApiStatusBadge } from "./ApiStatusBadge";
import { MotionToggle } from "./MotionToggle";
import { ThemeToggle } from "./ThemeToggle";

export function AppChrome() {
  return (
    <>
      <ApiStatusBadge />
      <ThemeToggle />
      <MotionToggle />
    </>
  );
}
