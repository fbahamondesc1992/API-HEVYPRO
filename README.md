# API-HEVYPRO

Importador seguro de una rutina de entrenamiento desde Excel hacia Hevy Pro mediante la API pública de Hevy.

## Qué incluye

- `import_hevy_routine.py`: importador principal.
- `Rutina_Entrenamiento_5_Dias_Hevy.xlsx`: rutina de 5 días preparada para importar.
- `hevy_mapping.example.json`: ejemplo de mapeo entre nombres del Excel y ejercicios de Hevy.
- `.gitignore`: protege archivos locales, API keys y mapeos generados.

## Funcionamiento

El script trabaja en dos etapas:

1. **Preview seguro**: lee el Excel, valida la API key, descarga el catálogo de ejercicios de Hevy y muestra cómo resolverá cada ejercicio. No escribe nada en Hevy.
2. **Aplicación**: con `--apply`, crea las rutinas que no existan.

## Requisitos

- Hevy Pro.
- Python 3.
- API key de Hevy obtenida desde `https://hevy.com/settings?developer`.

La API key **no debe guardarse en el repositorio**.

## Uso

```bash
python3 import_hevy_routine.py
```

Revisa el preview y, cuando todo esté correcto:

```bash
python3 import_hevy_routine.py --apply
```

También puedes usar la API key solo durante la sesión actual de Terminal:

```bash
export HEVY_API_KEY="TU_CLAVE"
python3 import_hevy_routine.py
```

## Notas de compatibilidad

- La conexión principal se valida contra `/v1/exercise_templates`.
- `/v1/user/info` es informativo y no bloquea el proceso.
- Las rutinas se crean inicialmente en **My Routines** con `folder_id: null`.
- El rango objetivo de repeticiones, RIR, RPE y regla de progresión quedan en las notas de cada ejercicio.
- Si el Excel define un descanso como rango, por ejemplo `2:30–3:00`, se usa el extremo superior.
- El script evita crear una rutina si ya existe otra con el mismo título.

## Seguridad

Nunca subas al repositorio:

- API keys.
- archivos `.env`.
- `hevy_mapping.json` generado localmente si contiene preferencias personales que no quieras versionar.

El repositorio está pensado para mantenerse privado.
