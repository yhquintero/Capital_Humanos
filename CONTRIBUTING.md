# Contribuir a Capital Humano

Gracias por ayudar a mejorar esta herramienta. Las contribuciones deben conservar la trazabilidad de las fuentes, la precisión de los cálculos salariales y la privacidad de las personas.

## Antes de empezar

- Para errores o propuestas pequeñas, abre un issue sin incluir información personal ni salarios reales.
- Para cambios grandes, abre primero una propuesta para acordar el alcance y los criterios de aceptación.
- Consulta [`docs/ROADMAP.md`](docs/ROADMAP.md), [`docs/MEJORAS_PROPUESTAS.md`](docs/MEJORAS_PROPUESTAS.md) y el [registro de cambios](CHANGELOG.md) para evitar trabajo duplicado.
- No cambies escalas, archivos fuente ni referencias normativas sin documentar su procedencia y revisión. No inventes importes faltantes ni extrapoles un costo total de nómina a partir de datos por cargo.

## Preparar el entorno y ejecutar pruebas

Requiere Python 3.10 o posterior. Desde la raíz del repositorio:

```bash
python -m pip install -r "TES/Salariales/Nueva Escala 3210/requirements.txt" -r requirements-dev.txt
python -m pytest -q
```

Si el cambio afecta una sección de Streamlit, comprueba también el flujo correspondiente en `tests/test_streamlit_app.py`. Para cambios de accesibilidad, actualiza la documentación y ejecuta las verificaciones manuales pendientes indicadas en [`docs/ACCESSIBILITY.md`](docs/ACCESSIBILITY.md); las pruebas automáticas por sí solas no justifican una declaración de conformidad WCAG.

## Requisitos para una contribución

- Mantén los datos de origen intactos y deja los cálculos reproducibles.
- Usa el formato monetario de interfaz y exportaciones **`$ 3,210.00`**.
- Añade pruebas para cambios funcionales y actualiza README/documentación cuando cambie el comportamiento.
- No añadas datos de empleados, plantillas reales no públicas, credenciales, certificados privados ni exportaciones con información sensible.
- Usa exclusivamente datos ficticios o anonimizados en ejemplos y pruebas.
- Describe las pruebas realizadas, las fuentes modificadas y cualquier limitación que quede abierta.

## Pull requests

1. Crea una rama de trabajo desde `main`.
2. Mantén cada cambio enfocado y explica su motivación y efecto.
3. Ejecuta las pruebas relevantes y confirma que no se alteraron fuentes por accidente.
4. Solicita revisión; no presentes cálculos o propuestas como aprobaciones oficiales.

La licencia del proyecto está pendiente de definición; este documento no sustituye esa decisión. La revisión y aceptación de un pull request no implica aprobación institucional de los cálculos ni de las fuentes.

---

# Contributing to Capital Humano

Thanks for helping improve this tool. Contributions must preserve source traceability, salary-calculation accuracy, and personal privacy.

## Before you start

- Open an issue for small bugs or suggestions, without personal information or real salary data.
- For larger changes, first propose the scope and acceptance criteria.
- Review [`docs/ROADMAP.md`](docs/ROADMAP.md), [`docs/MEJORAS_PROPUESTAS.md`](docs/MEJORAS_PROPUESTAS.md), and the [changelog](CHANGELOG.md) to avoid duplicate work.
- Do not change source files, salary scales, or legal references without documenting their provenance and review. Do not invent missing amounts or extrapolate total payroll costs from position-level data.

## Set up and test

Python 3.10 or newer is required. From the repository root:

```bash
python -m pip install -r "TES/Salariales/Nueva Escala 3210/requirements.txt" -r requirements-dev.txt
python -m pytest -q
```

For Streamlit changes, also test the relevant flow in `tests/test_streamlit_app.py`. For accessibility changes, update the documentation and perform the outstanding manual checks in [`docs/ACCESSIBILITY.md`](docs/ACCESSIBILITY.md); automated tests alone do not support a WCAG conformance claim.

## Contribution requirements

- Keep source data unchanged and calculations reproducible.
- Use the **`$ 3,210.00`** monetary format in the interface and exports.
- Add tests for functional changes and update documentation when behavior changes.
- Do not add employee data, non-public real templates, credentials, private certificates, or exports containing sensitive information.
- Use fictional or anonymized data in examples and tests.
- Describe tests performed, sources modified, and any remaining limitations.

## Pull requests

1. Create a working branch from `main`.
2. Keep each change focused and explain its motivation and effect.
3. Run relevant tests and confirm that source files were not changed accidentally.
4. Request review; do not present calculations or proposals as official approvals.

The project's license is pending a decision; this document does not replace that decision. Pull request review and acceptance do not constitute institutional approval of calculations or sources.
