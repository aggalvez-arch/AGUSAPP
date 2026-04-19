"use client";

import { useState } from "react";

type CopyBlockProps = {
  label: string;
  value: string;
};

export default function CopyBlock({ label, value }: CopyBlockProps) {
  const [copied, setCopied] = useState(false);

  const copy = async () => {
    await navigator.clipboard.writeText(value);
    setCopied(true);
    setTimeout(() => setCopied(false), 1200);
  };

  return (
    <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
      <div className="mb-2 flex items-center justify-between">
        <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</p>
        <button
          type="button"
          onClick={copy}
          className="rounded-lg border border-slate-300 bg-white px-2 py-1 text-xs text-slate-700"
        >
          {copied ? "Copiado" : "Copiar"}
        </button>
      </div>
      <pre className="whitespace-pre-wrap text-sm text-slate-800">{value}</pre>
    </div>
  );
}
