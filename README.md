# CRM INTEGRAL | GROUP EJ MORALES S.U.R.L. — Aplicación de escritorio

Aplicación de escritorio para que cada **especialista (comercial)** trabaje su libro
individual de CRM y envíe su información actualizada a **jefatura**, quien la
**consolida diariamente** en el libro maestro y revisa las ofertas.

No requiere internet y trabaja directamente sobre los libros Excel existentes,
manteniendo los nombres de hojas y encabezados (sigue siendo compatible con
Power Query si algún día quieren usarlo).

---

## 1. Requisitos e instalación

1. Tener instalado **Python 3.10 o superior** (se descarga una sola vez desde
   https://www.python.org/downloads/ — en Windows marcar la casilla
   **"Add Python to PATH"**).
2. Copiar la carpeta completa de la aplicación a cualquier ubicación del PC
   (por ejemplo `C:\CRM_GroupEJ`).
3. Doble clic en **`Iniciar_CRM.bat`** (Windows) o ejecutar
   `./Iniciar_CRM.sh` (Linux/Mac).

La carpeta `libs` ya incluye todo lo necesario (openpyxl). **No hay que instalar
nada más ni tener internet.**

---

## 2. Flujo de trabajo diario

```
   ESPECIALISTA (cada comercial)                    JEFATURA
   ─────────────────────────────                    ─────────────────────────
   1. Abre su libro individual               1. Abre el libro maestro
   2. Registra clientes, oportunidades,      2. "Importar paquete..."
      ofertas, actividades, etc.                recibe los archivos que
   3. Guarda                                    le enviaron los comerciales
   4. "Enviar a jefatura"                     3. Resuelve conflictos si los hay
      genera un archivo Paquete_...json      4. El maestro queda consolidado
      y se lo envía a la jefa                  y se registra en LOG_CONSOLIDACION
      (correo, USB, etc.)                    5. Revisa y aprueba ofertas
                                            6. "Exportar retroalimentación"
                                               (opcional) para devolver a cada
                                               especialista el estado de sus
                                               ofertas
```

### Rol ESPECIALISTA
- **Abrir libro / Nuevo libro**: al crear uno nuevo se copia la plantilla
  oficial `Plantilla_Libro_Individual.xlsx` (se sugiere nombrarlo
  `CRM_Comercial_02_GroupEJ_Morales.xlsx`, etc.).
- En el menú lateral están todas las hojas del CRM (Clientes, Proveedores,
  Prospectos, Ofertas, Importaciones, Contenedores, Ventas, Logística,
  Facturas_Cobros, Incidencias, etc.).
- **Nueva fila**: abre un formulario con los campos exactos de la hoja. El ID
  se genera automáticamente (CLI-0001, OPO-0001, OFE-0001, IMP-0001, ...).
  Los campos con catálogo (Etapa, Estado_revision, Estado de importación /
  embarque / contenedor, Prioridad) muestran lista desplegable.
- **Editar** (o doble clic en una fila) y **Eliminar**.
- **Buscar** filtra la tabla por cualquier columna.
- **Guardar** escribe los cambios en el Excel.
- **Enviar a jefatura**: genera en la subcarpeta `Paquetes_a_Jefatura` un
  archivo `Paquete_<nombre>_<fecha>.json` con toda su información y un resumen
  de los cambios desde su último envío. Ese archivo es el que se envía a la
  jefa. Guardar el libro antes de enviar es lo recomendado.
- **Importar retroalimentación**: cuando la jefa le devuelve un archivo
  `Retro_...json`, esta opción aplica a su libro el estado de revisión de sus
  ofertas y observaciones de jefatura.

### Rol JEFATURA
- Abre el libro maestro (o crea uno nuevo desde `Plantilla_Maestro_Jefatura.xlsx`).
- **Importar paquete...**: selecciona uno o varios archivos `Paquete_...json`.
  La aplicación muestra un resumen (filas nuevas, filas a completar y
  conflictos) y aplica la consolidación:
  - **Filas nuevas** del especialista → se agregan al maestro.
  - **Campos vacíos en el maestro** que el especialista sí llenó → se completan
    automáticamente.
  - **Conflictos** (mismo ID con datos distintos) → se abre un panel donde la
    jefa decide caso por caso: conservar el maestro o tomar la versión del
    especialista (también puede decidir todo de una vez).
  - Cada consolidación queda registrada en la hoja **LOG_CONSOLIDACION** y
    actualiza "Ultima_consolidacion" en Equipo.
- **Panel**: indicadores (clientes, oportunidades abiertas, ofertas pendientes
  de revisión, ventas, seguimientos vencidos), cola de ofertas que requieren su
  revisión y registro de consolidaciones.
- **Exportar retroalimentación**: genera un `Retro_<especialista>_<fecha>.json`
  para devolverle a cada comercial el estado de sus ofertas (aprobada,
  modificaciones solicitadas, observaciones, etc.).

---

## 3. Reglas de consolidación (coinciden con la guía interna del Excel)

1. Clientes se deduplican por `ID_Cliente` y Proveedores por `ID_Proveedor`.
   Si hay diferencias para el mismo ID, la aplicación lo marca como
   **conflicto** y no sobrescribe sin decisión de la jefa.
2. En las tablas de movimientos se conservan todos los IDs únicos.
3. Las hojas y encabezados se mantienen idénticos: el maestro sigue
   funcionando con Power Query si se desea.
4. Cada cambio marca automáticamente `Fecha_modificacion`.
5. Las ofertas requieren revisión de jefatura antes de presentarse al cliente
   (campo `Estado_revision` con el catálogo oficial).

---

## 4. Preguntas frecuentes

**¿Cómo viaja la información de un PC a otro?**
Con el archivo `Paquete_...json` (es pequeño) por correo, mensajería o USB.
La jefa lo importa con un clic. No hace falta enviar el Excel completo.

**¿Puedo seguir usando Excel a la vez?**
Sí, pero evite tener el mismo libro abierto en Excel y en la aplicación al
mismo tiempo (Excel bloquea el archivo). La aplicación lee y escribe el mismo
formato.

**¿Qué pasa si un comercial ya envió su paquete y luego hace más cambios?**
Simplemente genera otro paquete: la aplicación solo importará las novedades.

**¿Y si dos especialistas crean el mismo ID?**
La aplicación lo detecta al consolidar y lo presenta como conflicto para que
la jefa decida.

**¿Dónde están los archivos?**
- Libros Excel: donde usted los guarde.
- Paquetes del especialista: subcarpeta `Paquetes_a_Jefatura` junto a su libro.
- Retroalimentaciones de jefatura: subcarpeta `Retroalimentaciones` junto al
  maestro.
- Un archivo auxiliar `<libro>.xlsx.snapshot.json` se crea junto a cada libro
  para calcular los cambios entre envíos (no borrarlo).

**¿Seguridad?**
Como indica la guía del propio CRM, Excel no ofrece seguridad robusta por
fila; use permisos de carpeta/archivo del sistema operativo si necesita
limitar el acceso.

---

## 5. Contenido de la carpeta

| Archivo | Descripción |
|---|---|
| `CRM_GroupEJ.py` | Aplicación completa |
| `Iniciar_CRM.bat` | Lanzador para Windows (doble clic) |
| `Iniciar_CRM.sh` | Lanzador para Linux/Mac |
| `Plantilla_Libro_Individual.xlsx` | Plantilla para nuevos especialistas |
| `Plantilla_Maestro_Jefatura.xlsx` | Plantilla del libro maestro |
| `libs/` | Librerías incluidas (openpyxl), no requiere internet |
| `README.md` | Este manual |

*CRM INTEGRAL | GROUP EJ MORALES S.U.R.L.*
