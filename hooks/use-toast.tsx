"use client";

import * as React from "react";

type ToastOptions = {
  title: string;
  description?: string;
};

type ToastState = ToastOptions & {
  id: string;
  open: boolean;
};

type ToastContextValue = {
  toasts: ToastState[];
  toast: (options: ToastOptions) => void;
  dismiss: (id: string) => void;
};

const ToastContext = React.createContext<ToastContextValue | null>(null);

export function ToastStateProvider({ children }: { children: React.ReactNode }) {
  const [toasts, setToasts] = React.useState<ToastState[]>([]);

  const dismiss = React.useCallback((id: string) => {
    setToasts((current) => current.map((toast) => (toast.id === id ? { ...toast, open: false } : toast)));
    window.setTimeout(() => {
      setToasts((current) => current.filter((toast) => toast.id !== id));
    }, 250);
  }, []);

  const toast = React.useCallback(
    (options: ToastOptions) => {
      const id = crypto.randomUUID();
      setToasts((current) => [{ ...options, id, open: true }, ...current].slice(0, 4));
      window.setTimeout(() => dismiss(id), 4500);
    },
    [dismiss]
  );

  return <ToastContext.Provider value={{ toasts, toast, dismiss }}>{children}</ToastContext.Provider>;
}

export function useToast() {
  const context = React.useContext(ToastContext);
  if (!context) {
    throw new Error("useToast must be used within ToastStateProvider");
  }
  return context;
}
