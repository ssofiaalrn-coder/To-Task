To-Task es una aplicación web académica ligera desarrollada con Google Apps Script  y una interfaz web moderna y responsiva integrada con HTML, CSS y JavaScript. Está diseñada específicamente para que los estudiantes universitarios puedan organizar su carga académica, gestionar fechas de entrega, consultar su calendario de pendientes y apoyarse en un panel de control rápido.

Para el Backend (Código.gs):
Utiliza Google Sheets como una base de datos en la nube. Administra la lógica de negocio para registrar, consultar, validar y actualizar tareas y entregas académicas, integración con Gemini.

Para el Frontend (index.html):
(HTML, CSS y JavaScript). Presenta una interfaz de usuario organizada en pestañas verticales:
Por Hacer: Panel visual (tareas totales, pendientes, entregas próximas).
Añadir Entrega: Formulario con botón para registrar asignaciones o trabajos.
Calendario de Entregas: Formato de calendario con el listado general de deberes y sus estados de urgencia.
Chat IA / Búsqueda: Motor de búsqueda para apoyo.
Configuración: Herramientas de mantenimiento del sistema y reseteo de datos.

Para las hojas de calculo (googel sheets):
Materias: Código, Título / Materia, Modalidad, Categoría, Prioridad (1-5), Fecha Creación
Tareas: Código, Fecha Entrega, Estado, Avance (%), Usuario, Timestamp, Descripción / Notas, Prioridad
Categorias:	Lista de categorías
Estados:	pendiente, en proceso, completada, atrasada. 


