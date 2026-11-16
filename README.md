# 🛡️ ATIC Word Sanitizer

Herramienta automatizada en **PowerShell** diseñada para la limpieza y reconstrucción masiva de documentos de Microsoft Word (`.doc`, `.docx`, `.docm`) afectados por corrupción estructural, macros maliciosas o código inyectado.

El script soluciona el problema clásico de los bloqueos en la automatización COM de Word utilizando una arquitectura de **proceso aislado (*worker*) con control de tiempo de espera (*timeout*)**, garantizando que un archivo severamente dañado no detenga ni cuelgue el proceso completo.

---

## 🚀 ¿Cómo funciona?

A diferencia de los métodos tradicionales de guardado directo que pueden arrastrar partes del paquete XML corrupto o macros ocultas, este script realiza una **reconstrucción limpia**:
1. Abre el documento original en **modo seguro** (solo lectura, sin actualizar vínculos y con macros completamente deshabilitadas).
2. Crea un documento nuevo y limpio en blanco.
3. Transfiere el contenido estructurado mediante rangos internos de Word.
4. Exporta el resultado como un archivo `.docx` estándar totalmente sanitizado.
5. Controla cada archivo mediante un subproceso independiente (`_worker.ps1`); si un archivo se congela, el sistema aplica un *timeout*, fuerza el cierre de la instancia y continúa con el siguiente documento de forma desatendida.

---

## 📂 Estructura del Proyecto

```text
📦 atic-word-sanitizer/
├── 📜 ATIC.ps1          # Script principal interactivo (menú y bucle de control)
├── 📜 _worker.ps1           # Subproceso encargado de interactuar de forma segura con Word COM
└── 🚀 sanitizer.bat          # Lanzador rápido para Windows (opcional)
