# AirReference Starter (base UI)

Este módulo contiene una base de interfaz para empezar AirReference en Next.js:

- Home con buscador
- Procedimiento con resumen operativo + checklist + bloque copiable
- Datos mock para reemplazar luego por Postgres/API

## Archivos clave

- `app/(public)/page.tsx`
- `app/(public)/procedures/[procedureSlug]/page.tsx`
- `components/GlobalSearchBar.tsx`
- `components/OperationalSummaryCard.tsx`
- `components/ProcedureChecklist.tsx`
- `components/CopyBlock.tsx`
- `lib/mockData.ts`

## Integración rápida

1. Copia estas carpetas dentro de tu proyecto Next.js.
2. Ajusta imports si tu estructura cambia.
3. Reemplaza `lib/mockData.ts` con llamadas reales (`/api/...`).
4. Conecta roles/estado/versionado desde backend.
