# Accesibilidad de la interfaz

## Alcance de la revisión

CH-06 se trabaja con objetivo WCAG 2.2 AA, no como certificación ni como afirmación de conformidad. Los flujos representativos son:

1. navegar entre las siete secciones desde el panel;
2. cambiar el factor salarial y consultar la escala;
3. comparar dos escenarios por grupo o cargo;
4. filtrar y revisar incidencias de la plantilla;
5. revisar el estado de fuentes y descargar informes.

## Cambios en curso

La interfaz incluye un indicador de foco visible de alto contraste, un color de texto secundario que alcanza contraste AA sobre blanco, ajustes de maquetación para anchos móviles y respeto por la preferencia de movimiento reducido. `test_accessibility_contract.py` comprueba algunos pares de color seleccionados y la presencia de estos estilos; no sustituye las verificaciones manuales descritas abajo.

## Verificaciones que faltan

Los tests automáticos del repositorio no inspeccionan el DOM renderizado por navegador ni sus componentes internos. Antes de cerrar CH-06 se deben registrar:

- ejecución de un analizador automático sobre los flujos y páginas definidos, sin hallazgos graves/críticos sin resolver;
- comprobación manual con teclado de orden de tabulación, controles, pestañas, foco visible y ausencia de trampas;
- comprobación de reflujo a 320 CSS px y zoom, incluidos gráficos, tablas y descargas;
- prueba manual de etiquetas, nombres, estados, mensajes de error y cambios dinámicos con lector de pantalla (al menos una combinación de navegador y lector documentada);
- incidencias, excepciones justificadas, herramientas y versiones utilizadas.

No se declara que la aplicación cumpla WCAG ni se cierra CH-06 hasta realizar y documentar esas comprobaciones. La revisión independiente de CH-14 sigue siendo una mejora separada.
