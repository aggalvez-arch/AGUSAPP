# Diagnóstico y rediseño operativo

## 1) Estado actual observado
El repositorio no contiene implementación funcional previa de Google Apps Script (solo `.gitkeep`).
Por lo tanto, no existe flujo operativo ejecutable para analizar línea a línea.

## 2) Riesgos deducidos por ausencia de implementación
- No hay estructura de hojas estandarizada.
- No hay criterios de decisión explícitos ni trazabilidad.
- No hay separación entre decisiones automáticas y revisión humana.
- No hay controles de concurrencia ni proceso batch.

## 3) Diseño propuesto
Se implementa un motor de procesamiento con:
- Botón principal único: `procesarTodo`.
- Hojas estándar: `RAW`, `ANALISIS`, `REVISION`, `CONFIG`.
- Puntaje de confianza 0-100 y motivo de decisión.
- Estados normalizados de punta a punta.
- Revisión humana solo por excepción.
- Configuración parametrizable sin tocar código.

## 4) Política de decisión
### Decisión automática segura
- `AUTO_APROBAR` cuando el puntaje supera umbral de auto aprobación,
  no hay errores críticos y no hay categoría sensible.

### Decisión para revisión humana
- Registros con datos críticos inválidos.
- Registros de puntaje intermedio.
- Registros de bajo puntaje o con señales de riesgo.

## 5) Rendimiento y mantenibilidad
- Lectura en bloque de datos por hoja.
- Escritura en bloque para altas de análisis/revisión.
- Actualizaciones concentradas por columna para estados.
- Configuración centralizada en `CONFIG`.
- Funciones pequeñas con responsabilidades separadas.
