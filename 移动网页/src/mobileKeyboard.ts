import { useEffect, useRef, useState } from "react";

const MOBILE_BREAKPOINT = 720;
const KEYBOARD_HEIGHT_THRESHOLD = 140;

interface KeyboardViewport {
  viewportWidth: number;
  layoutHeight: number;
  visualHeight: number;
  hasFocusedField: boolean;
}

export function isMobileKeyboardVisible(viewport: KeyboardViewport) {
  return viewport.viewportWidth <= MOBILE_BREAKPOINT
    && viewport.hasFocusedField
    && viewport.layoutHeight - viewport.visualHeight >= KEYBOARD_HEIGHT_THRESHOLD;
}

function isEditableField(element: Element | null) {
  if (!(element instanceof HTMLElement)) return false;
  return element.matches("input, textarea, select, [contenteditable='true']");
}

export function useMobileKeyboardVisible() {
  const [visible, setVisible] = useState(false);
  const layoutHeight = useRef(window.innerHeight);
  const focusTimer = useRef<number | null>(null);

  useEffect(() => {
    const viewport = window.visualViewport;

    const update = () => {
      const hasFocusedField = isEditableField(document.activeElement);
      if (!hasFocusedField) {
        layoutHeight.current = Math.max(window.innerHeight, viewport?.height ?? 0);
      }
      setVisible(isMobileKeyboardVisible({
        viewportWidth: window.innerWidth,
        layoutHeight: layoutHeight.current,
        visualHeight: viewport?.height ?? window.innerHeight,
        hasFocusedField
      }));
    };

    const updateAfterFocus = () => {
      if (focusTimer.current) window.clearTimeout(focusTimer.current);
      focusTimer.current = window.setTimeout(update, 60);
    };
    update();
    window.addEventListener("resize", update, { passive: true });
    document.addEventListener("focusin", updateAfterFocus);
    document.addEventListener("focusout", updateAfterFocus);
    viewport?.addEventListener("resize", update, { passive: true });
    viewport?.addEventListener("scroll", update, { passive: true });

    return () => {
      if (focusTimer.current) window.clearTimeout(focusTimer.current);
      window.removeEventListener("resize", update);
      document.removeEventListener("focusin", updateAfterFocus);
      document.removeEventListener("focusout", updateAfterFocus);
      viewport?.removeEventListener("resize", update);
      viewport?.removeEventListener("scroll", update);
    };
  }, []);

  return visible;
}
