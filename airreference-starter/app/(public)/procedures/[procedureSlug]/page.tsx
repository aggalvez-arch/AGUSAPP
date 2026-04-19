import CopyBlock from "../../../../components/CopyBlock";
import OperationalSummaryCard from "../../../../components/OperationalSummaryCard";
import ProcedureChecklist from "../../../../components/ProcedureChecklist";
import { sampleProcedure } from "../../../../lib/mockData";

export default function ProcedurePage() {
  return (
    <main className="min-h-screen bg-slate-50">
      <section className="mx-auto max-w-4xl px-4 py-10">
        <header className="mb-6 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <p className="text-xs uppercase tracking-wide text-slate-500">Procedimiento</p>
          <h1 className="mt-1 text-2xl font-semibold text-slate-900">{sampleProcedure.title}</h1>
          <div className="mt-3 flex flex-wrap gap-2 text-xs">
            <span className="rounded-full bg-slate-100 px-3 py-1 text-slate-700">
              {sampleProcedure.version}
            </span>
            <span className="rounded-full bg-emerald-100 px-3 py-1 text-emerald-800">
              Verificado: {sampleProcedure.verifiedAt}
            </span>
          </div>
        </header>

        <div className="space-y-5">
          <OperationalSummaryCard bullets={sampleProcedure.summary} />
          <ProcedureChecklist steps={sampleProcedure.steps} />
          <CopyBlock label="Plantilla de correo" value={sampleProcedure.template} />

          <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="text-base font-semibold text-slate-900">Fuentes y control</h2>
            <ul className="mt-3 list-disc space-y-1 pl-5 text-sm text-slate-700">
              <li>Fuente oficial aerolínea (portal público)</li>
              <li>Snapshot guardado: 2026-04-18 09:30 UTC</li>
              <li>Historial: v1.1 → v1.2 (ajuste de validación)</li>
            </ul>
          </section>
        </div>
      </section>
    </main>
  );
}
