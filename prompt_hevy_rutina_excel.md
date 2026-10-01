# Prompt para analizar historial de Hevy y crear una rutina en Excel

Voy a adjuntar mi historial de entrenamiento exportado desde **Hevy en formato CSV**.

Quiero que trabajes en dos etapas.

## ETAPA 1 — Analizar mi historial real

Antes de crear una rutina, analiza completamente el archivo y determina:

- cantidad de sesiones y período cubierto;
- frecuencia real de entrenamiento por semana;
- ejercicios más utilizados;
- volumen aproximado por grupo muscular;
- frecuencia por grupo muscular;
- evolución de cargas y repeticiones;
- ejercicios donde estoy progresando;
- ejercicios donde estoy estancado;
- posibles desequilibrios entre tren superior/inferior y patrones de movimiento;
- adherencia real a las sesiones A/B;
- duración de los entrenamientos si existe esa información;
- cargas recientes que puedan utilizarse como referencia para una nueva rutina.

No mezcles indiscriminadamente datos muy antiguos con mi nivel actual. Para definir las cargas iniciales da mayor peso a las **últimas 6–10 semanas**.

Si existe RPE en el archivo, úsalo. Si no existe, indícalo explícitamente y no inventes valores históricos.

Antes de diseñar la rutina, entrégame un breve diagnóstico con los principales hallazgos.

---

## ETAPA 2 — Crear un nuevo plan

Después del análisis, crea una rutina orientada principalmente a **hipertrofia y progresión de fuerza**, estructurada en:

- **Superior A**
- **Inferior A**
- Descanso
- **Superior B**
- **Inferior B**
- **Brazos en casa**
- Descanso

El objetivo es entrenar:

- 4 días en gimnasio;
- 1 día en casa dedicado exclusivamente a bíceps y tríceps.

En casa dispongo de:

- mancuernas;
- barra Z;
- barra recta;
- polea simple.

Evita duplicar volumen innecesariamente. El día de brazos debe **redistribuir el volumen semanal**, no simplemente sumar muchas series adicionales.

Para cada ejercicio define:

1. nombre del ejercicio;
2. número de series efectivas;
3. rango de repeticiones;
4. **carga inicial sugerida basada en mi historial real**;
5. RIR objetivo;
6. RPE equivalente;
7. descanso;
8. regla exacta de progresión.

Usa como referencia general:

- compuestos principales: RIR 1–2;
- ejercicios secundarios: RIR 1–2;
- aislados: RIR 1–2;
- evitar fallo sistemático en ejercicios compuestos.

Usa **doble progresión**:

- mantener el peso mientras no se complete el máximo del rango en todas las series;
- aumentar carga cuando todas las series lleguen al máximo del rango con técnica limpia y al menos RIR 1;
- si durante dos sesiones consecutivas no alcanzo el mínimo del rango, recomendar una reducción de aproximadamente 5–10 %.

Si un ejercicio no tiene suficiente historial para proponer una carga fiable, escribe **“Calibrar”** en vez de inventar un peso.

---

## VOLUMEN

Busca un volumen semanal razonable y sostenible. Como referencia inicial:

- Pecho: ~8–12 series directas;
- Espalda: ~10–14;
- Bíceps: ~8–10 directas;
- Tríceps: ~8–10 directas;
- Cuádriceps: ~8–12;
- Femoral: ~8–10;
- Pantorrilla: ~6–8;
- Deltoides lateral/posterior: volumen suficiente sin duplicar innecesariamente el trabajo de presses.

Ajusta estos valores si mi historial demuestra que conviene hacerlo.

No agregues ejercicios solo para aumentar el número total. Prioriza una rutina que pueda completar consistentemente.

---

## EXCEL FINAL

Una vez que la rutina esté definida, crea un archivo **Excel `.xlsx`** profesional y ordenado.

Debe incluir estas hojas:

1. `Resumen`
2. `Superior A`
3. `Inferior A`
4. `Superior B`
5. `Inferior B`
6. `Brazos Casa`
7. `RIR-RPE`

En cada hoja de entrenamiento utiliza estas columnas:

| Ejercicio | Series | Reps | Carga inicial | RIR objetivo | RPE | Descanso | Regla de progresión |
|---|---:|---|---|---|---|---|---|

La hoja `Resumen` debe contener:

- distribución semanal;
- objetivo de cada sesión;
- volumen semanal aproximado por grupo muscular;
- reglas generales de progresión.

La hoja `RIR-RPE` debe incluir al menos:

| RIR | RPE | Interpretación |
|---:|---:|---|
| 3 | 7 | Quedaban ~3 repeticiones |
| 2 | 8 | Quedaban ~2 repeticiones |
| 1 | 9 | Quedaba ~1 repetición |
| 0 | 10 | No quedaban repeticiones limpias |

Dale al Excel formato visual profesional:

- encabezados claros;
- colores suaves;
- columnas ajustadas;
- filtros;
- congelar encabezados;
- diferenciar visualmente carga, RIR/RPE y progresión.

**No me entregues solo una tabla en el chat. Quiero que generes el archivo `.xlsx` para descargar.**

Finalmente, resume en 5–10 puntos por qué la nueva rutina es mejor que la estructura que venía realizando.

---

## CONTEXTO OPCIONAL PARA MEJORAR EL RESULTADO

Puedes agregar al inicio del prompt algo como:

> Tengo acceso normal a máquinas, poleas, barras y mancuernas. Prefiero evitar [ejercicios que no te gustan]. No tengo lesiones actualmente. Mi objetivo principal es hipertrofia, manteniendo o aumentando fuerza. Quiero que las sesiones duren idealmente entre 75 y 100 minutos.

Si después quieres importar el Excel a Hevy con un script, agrega además:

> **Importante:** conserva exactamente los nombres de las hojas y las columnas indicadas, porque posteriormente este Excel será leído por un script Python que lo importará mediante la API de Hevy.

Flujo esperado:

**Hevy CSV → ChatGPT → análisis → Excel → Python → Hevy API**
