export const homeUpdates = [
  {
    id: "1",
    airline: "Aerolínea Demo",
    topic: "Reemisión",
    date: "2026-04-18",
    note: "Actualizada validación de penalidad por tarifa base.",
  },
  {
    id: "2",
    airline: "Aerolínea Demo",
    topic: "Void/Refund",
    date: "2026-04-17",
    note: "Se añadió canal de contacto para casos internacionales.",
  },
];

export const sampleProcedure = {
  title: "Reemisión por cambio voluntario (GDS)",
  version: "v1.2",
  verifiedAt: "2026-04-18",
  summary: [
    "Validar regla de tarifa antes de calcular diferencia.",
    "Aplicar penalidad según tipo de fare y ruta.",
    "Registrar SSR/OSI requerido en PNR antes de emitir.",
    "Cerrar con validación final de impuestos y endosos.",
  ],
  steps: [
    {
      id: "s1",
      title: "Validar elegibilidad",
      body: "Confirma que el boleto permita cambio voluntario y revisa ventana temporal aplicable.",
    },
    {
      id: "s2",
      title: "Calcular diferencia + penalidad",
      body: "Calcula ADC/penalidad según reglas vigentes y confirma impuestos recalculados.",
    },
    {
      id: "s3",
      title: "Actualizar PNR",
      body: "Inserta observaciones operativas, SSR/OSI y referencias internas requeridas.",
    },
    {
      id: "s4",
      title: "Reemitir y validar",
      body: "Emite, verifica endosos y documenta evidencia de validación en el caso interno.",
    },
  ],
  template:
    "Asunto: Reemisión aplicada\nPNR: ______\nTicket: ______\nAcción: Cambio voluntario reemitido\nObservaciones: ______",
};
