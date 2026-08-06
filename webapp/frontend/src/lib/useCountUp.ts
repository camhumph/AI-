import { useEffect, useRef } from "react";

/**
 * Animate a number from its previous value to a new one, writing straight to
 * the DOM node.
 *
 * Ported from the Datum console's tick loop. Two things make it feel right:
 *
 *  - It writes via textContent on a ref instead of setState, so a 620 ms
 *    animation is one paint per frame rather than ~37 React re-renders.
 *  - Quartic ease-out (1 - (1-k)^4) decelerates hard at the end, which reads as
 *    a value *settling* rather than sliding. Linear interpolation on a money
 *    figure looks like a slot machine.
 *
 * Honours prefers-reduced-motion by snapping to the final value.
 */
export function useCountUp(
  value: number,
  format: (v: number) => string,
  durationMs = 620,
) {
  const ref = useRef<HTMLSpanElement | null>(null);
  const from = useRef(0);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;

    const reduced =
      typeof window !== "undefined" &&
      window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;

    const start = from.current;
    const end = Number.isFinite(value) ? value : 0;
    from.current = end;

    if (reduced || start === end) {
      el.textContent = format(end);
      return;
    }

    const t0 = performance.now();
    let raf = 0;

    const tick = () => {
      const k = Math.min(1, (performance.now() - t0) / durationMs);
      const eased = 1 - Math.pow(1 - k, 4);
      el.textContent = format(start + (end - start) * eased);
      if (k < 1) raf = requestAnimationFrame(tick);
    };

    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [value, format, durationMs]);

  return ref;
}
