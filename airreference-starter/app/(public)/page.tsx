import GlobalSearchBar from "../../components/GlobalSearchBar";
import { homeUpdates } from "../../lib/mockData";

export default function HomePage() {
  return (
    <main className="min-h-screen bg-slate-50">
      <section className="mx-auto max-w-5xl px-4 pb-12 pt-16">
        <p className="text-xs font-semibold uppercase tracking-[0.16em] text-slate-500">
          AirReference
        </p>
        <h1 className="mt-3 text-3xl font-semibold tracking-tight text-slate-900">
          La referencia operativa para agencias.
        </h1>
        <p className="mt-3 max-w-2xl text-sm text-slate-600">
          Fuente única de verdad para procedimientos de aerolíneas, con historial,
          verificación y ejecución clara.
        </p>

        <div className="mt-8">
          <GlobalSearchBar />
        </div>
      </section>

      <section className="mx-auto grid max-w-5xl gap-4 px-4 pb-16 md:grid-cols-2">
        {homeUpdates.map((item) => (
          <article key={item.id} className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <div className="flex items-center justify-between">
              <p className="text-sm font-semibold text-slate-900">{item.airline}</p>
              <p className="text-xs text-slate-500">{item.date}</p>
            </div>
            <p className="mt-1 text-xs font-medium uppercase tracking-wide text-slate-500">
              {item.topic}
            </p>
            <p className="mt-2 text-sm text-slate-700">{item.note}</p>
          </article>
        ))}
      </section>
    </main>
  );
}
