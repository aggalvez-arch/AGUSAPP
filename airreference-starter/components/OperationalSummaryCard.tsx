type OperationalSummaryCardProps = {
  title?: string;
  bullets: string[];
};

export default function OperationalSummaryCard({
  title = "Resumen operativo (30s)",
  bullets,
}: OperationalSummaryCardProps) {
  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
      <h2 className="text-base font-semibold text-slate-900">{title}</h2>
      <ul className="mt-3 space-y-2 text-sm text-slate-700">
        {bullets.map((item, i) => (
          <li key={i} className="flex gap-2">
            <span className="mt-[2px] inline-block h-2 w-2 rounded-full bg-slate-500" />
            <span>{item}</span>
          </li>
        ))}
      </ul>
    </section>
  );
}
