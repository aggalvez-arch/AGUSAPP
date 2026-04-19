"use client";

import { useState } from "react";

type GlobalSearchBarProps = {
  defaultQuery?: string;
  onSearch?: (query: string) => void;
};

export default function GlobalSearchBar({
  defaultQuery = "",
  onSearch,
}: GlobalSearchBarProps) {
  const [query, setQuery] = useState(defaultQuery);

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    onSearch?.(query.trim());
  };

  return (
    <form onSubmit={submit} className="w-full">
      <div className="flex w-full items-center gap-2 rounded-2xl border border-slate-200 bg-white p-2 shadow-sm">
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Buscar aerolínea, proceso o regla operativa..."
          className="h-12 w-full rounded-xl px-4 text-sm text-slate-900 outline-none"
          aria-label="Buscar"
        />
        <button
          type="submit"
          className="h-12 rounded-xl bg-slate-900 px-5 text-sm font-medium text-white hover:bg-slate-700"
        >
          Buscar
        </button>
      </div>
    </form>
  );
}
