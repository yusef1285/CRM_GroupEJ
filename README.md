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
