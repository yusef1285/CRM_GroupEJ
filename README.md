# CRM INTEGRAL | GROUP EJ MORALES S.U.R.L. — v2.0

CRM de escritorio para Windows. **Los datos se introducen directamente en la aplicación**
(no hace falta abrir Excel). Cada especialista trabaja de forma individual y envía su
información actualizada a jefatura, que la consolida cada día.

## Descarga (Windows)

[CRM_GroupEJ.exe — última versión](https://github.com/yusef1285/CRM_GroupEJ/releases/latest/download/CRM_GroupEJ.exe)

No requiere instalar nada. Doble clic para abrir. Si Windows muestra «Windows protegió
tu equipo», pulse «Más información» y «Ejecutar de todas formas» (el .exe no está firmado).

## Cómo funciona

1. **Primer uso: crear un perfil.** Elija *Especialista* o *Jefatura*, escriba su nombre y un
   código corto (por ejemplo `C01`). El código forma parte de los identificadores
   (`CLI-C01-0001`, `OFE-C01-0004`) para que nunca se repitan entre especialistas.
2. **Trabajar como en un CRM real.** Menú lateral con los módulos: Comercial, Proveedores y
   compras, Importaciones, Operaciones, Gestión. En cada uno: nuevo, editar, eliminar, buscar,
   filtrar y ordenar. Los formularios traen listas desplegables, fechas con «Hoy/Ahora»,
   selección de clientes/proveedores existentes y cálculos automáticos (por ejemplo
   Importe = Cantidad × Precio). Todo se guarda al instante.
3. **Panel de control:** indicadores, agenda de seguimientos vencidos y próximos, ofertas
   pendientes y actividad reciente. En el cliente, la **ficha 360°** reúne todo su historial.
4. **Enviar a jefatura (especialista):** *Enviar a jefatura* genera un paquete con lo nuevo o
   modificado desde el último envío. Se manda por correo o carpeta compartida.
5. **Consolidar (jefatura):** *Importar paquetes* (uno o una carpeta completa). Se evitan
   duplicados por ID, los conflictos se resuelven campo por campo y las ofertas pasan por
   revisión de jefatura (y autorización de Dirección cuando aplique).
6. **Retroalimentación:** jefatura devuelve a cada especialista el estado de sus ofertas con
   *Devolver retroalimentación*; el especialista la carga con *Recibir retroalimentación*.

## Compartir paquetes por carpeta común

Todo el equipo puede usar una misma carpeta en lugar de mandar archivos por correo.
Sirve con: carpeta de red (Windows), OneDrive, Google Drive (con escritorio remoto),
Dropbox o un USB.

**Preparación (una sola vez):**

1. Elija la carpeta común. En una red de Windows, por ejemplo
   `\SERVIDOR\CRM_GroupEJ`; con OneDrive/Google Drive, la carpeta sincronizada
   (por ejemplo `C:\Users\su_usuario\OneDrive - Empresa\CRM_GroupEJ`).
2. Comparta esa ruta con todos los especialistas y con jefatura.
3. Cada persona, en su propio PC, abre la aplicación y configura una sola vez:
   **Sincronización > Carpeta compartida del equipo...** y selecciona esa carpeta.

**Uso diario:**

1. **Especialista:** al usar *Enviar a jefatura*, la aplicación ofrece guardar el paquete
   directamente en la subcarpeta `paquetes` de la carpeta compartida (evita el paso del
   correo). El archivo incluye su nombre y la fecha/hora, así que no se pisa con el de
   otro especialista.
2. **Jefatura:** *Sincronización > (Jefatura) Importar desde la carpeta compartida...*
   importa **todos** los paquetes pendientes de una sola vez, en orden.
3. **Retroalimentación:** cuando jefatura genera retroalimentación, se deja además una
   copia en la subcarpeta `retroalimentaciones`; el especialista la carga con
   *Recibir retroalimentación* (el diálogo ya abre en esa carpeta).

Estructura resultante de la carpeta compartida:

```
CRM_GroupEJ/
├── paquetes/                <- aquí dejan los especialistas
│   ├── Paquete_Comercial_01_20261109_1720.json
│   └── Paquete_Comercial_02_20261109_1805.json
└── retroalimentaciones/      <- aquí deja jefatura
    └── Retro_Comercial_01_20261109.json
```

Consejos: importen los paquetes con frecuencia (idealmente diario); después de importar,
jefatura puede mover los archivos ya procesados a una subcarpeta `paquetes/procesados`
para que no se mezclen con los nuevos; no edite los archivos .json a mano.

## Excel

Excel ya no es necesario para operar. Sigue disponible para:
- **Exportar** cualquier módulo o **todo el CRM** a un libro .xlsx (menú Archivo).
- **Importar** datos existentes desde sus libros Excel actuales (Archivo → Importar).

## Datos y respaldos

Los datos viven en una base local SQLite en `%APPDATA%\CRM_GroupEJ` (Archivo → Abrir carpeta de
datos). La aplicación crea respaldos automáticos al abrir un perfil y bajo demanda.

## Desarrollo

```
pip install openpyxl
python CRM_GroupEJ.py                 # ejecutar
python CRM_GroupEJ.py --selftest out.txt   # pruebas internas
```

El ejecutable se compila con GitHub Actions (`.github/workflows/build-exe.yml`).
