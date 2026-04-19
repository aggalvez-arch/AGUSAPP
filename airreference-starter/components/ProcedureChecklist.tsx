type ProcedureStep = {
  id: string;
  title: string;
  body: string;
};

type ProcedureChecklistProps = {
  steps: ProcedureStep[];
};

export default function ProcedureChecklist({ steps }: ProcedureChecklistProps) {
  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
      <h2 className="text-base font-semibold text-slate-900">Paso a paso</h2>
      <ol className="mt-4 space-y-4">
        {steps.map((step, index) => (
          <li key={step.id} className="rounded-xl border border-slate-200 p-4">
            <p className="text-xs font-semibold uppercase text-slate-500">Paso {index + 1}</p>
            <h3 className="mt-1 text-sm font-semibold text-slate-900">{step.title}</h3>
            <p className="mt-2 text-sm text-slate-700">{step.body}</p>
          </li>
        ))}
      </ol>
    </section>
  );
}
