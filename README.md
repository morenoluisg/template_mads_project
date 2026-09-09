# template_mads_project

Template para gestionar planes de trabajo y proyectos del dominio de datos de OTIC MADS usando un archivo Excel como fuente operativa y GitHub Projects como tablero de seguimiento.

El objetivo es convertir el plan de trabajo en issues administrables con enfoque Scrum: cada fila del Excel se transforma en una tarjeta, los productos se gestionan como milestones, los meses se convierten en sprints mensuales y los responsables se homologan por rol y, cuando exista usuario de GitHub, por assignee.

## Alcance

Este repositorio sirve como base para:

- Consolidar planes de trabajo del dominio de datos desde Excel.
- Crear o actualizar issues en GitHub a partir de las actividades del plan.
- Crear milestones asociados a productos o entregables.
- Configurar un GitHub Project con vistas de seguimiento.
- Administrar sprints mensuales, estados, roles y fechas de ejecucion.
- Repetir la sincronizacion cuando el Excel cambie.

Project actual de referencia:

- GitHub Project: <https://github.com/users/morenoluisg/projects/4>
- Repositorio: <https://github.com/morenoluisg/template_mads_project>

## Estructura

```text
.
+-- data/
|   +-- raw/
|       +-- Plan de Trabajo.xlsx
|       +-- acta_formato.docx
+-- scripts/
|   +-- sync_github_project.py
+-- .env.example
+-- requirements.txt
+-- README.md
```

## Fuente De Datos

El archivo principal es:

```text
data/raw/Plan de Trabajo.xlsx
```

La hoja del Excel debe conservar las columnas esperadas:

- `Lineamiento MGGTI`
- `PLAN OTIC 2026`
- `Entregable (Soporte)`
- `Descripcion`
- `Actividad`
- `ID`
- `Responsable`
- `Tareas relacionadas`
- `Mes`
- `Enero` a `Diciembre`
- `Relacion de producto`
- `ESTADO`
- `LINK`

Cada fila util del Excel se convierte en un issue. El script usa el campo `ID` como llave de sincronizacion. Si hay IDs repetidos, conserva el primero con su ID original y crea llaves derivadas, por ejemplo `I-29`, `I-29-2`, `I-29-3`.

## Modelo En GitHub

### Issues

Cada issue contiene:

- Codigo del plan en el titulo, por ejemplo `[D-01]`.
- Actividad como titulo principal.
- Contexto del lineamiento MGGTI.
- Entregable esperado.
- Producto o milestone asociado.
- Rol responsable.
- Mes original del Excel.
- Sprint mensual.
- Ventana de ejecucion completa.
- Descripcion y tareas relacionadas.
- Estado inicial tomado del Excel.
- Link o soporte, si existe.

### Milestones

Los milestones representan productos o entregables asociados. Se crean desde la columna `Relacion de producto`.

Cuando el producto tiene meses asociados, el milestone recibe fecha de vencimiento al cierre del ultimo mes identificado.

### Sprints

Los sprints se modelan como un campo `Sprint` del GitHub Project:

- `Sprint 01 - Enero 2026`
- `Sprint 02 - Febrero 2026`
- ...
- `Sprint 12 - Diciembre 2026`

El script toma el mes o rango de meses del Excel y asigna el sprint inicial correspondiente.

### Fechas

Los campos `Inicio` y `Fin` se calculan usando meses completos:

- Inicio: primer dia del primer mes asignado.
- Fin: ultimo dia del ultimo mes asignado.

Ejemplo: `Abr/Dic` se convierte en `2026-04-01` a `2026-12-31`.

### Estados Scrum

El campo `Status` usa los estados nativos del Project:

- `Todo`
- `In Progress`
- `Done`

El corte se controla con `WORKPLAN_TODAY` en `.env`.

Con `WORKPLAN_TODAY=2026-09-09`:

- Meses anteriores a septiembre quedan en `Done`.
- Actividades que incluyen septiembre quedan en `In Progress`.
- Meses posteriores quedan en `Todo`.

## Vistas Del Project

El script crea y mantiene estas vistas:

- `Backlog`: tablero filtrado para excluir `Done`.
- `Priority board`: tablero filtrado para excluir `Done`.
- `Team items`: tabla general para seguimiento del equipo.
- `Roadmap`: vista tipo Roadmap para revisar tiempos.
- `My items`: tabla filtrada por `assignee:@me`.
- `Milestones`: tabla orientada a productos y entregables.

GitHub limita por API algunos ajustes visuales finos del Roadmap, pero los items quedan con `Inicio`, `Fin`, `Sprint`, `Milestone`, `Rol` y `Status`, que son los campos necesarios para administrarlo desde la interfaz.

## Configuracion

Copie la plantilla de configuracion:

```powershell
Copy-Item .env.example .env
```

Edite `.env`:

```env
GITHUB_TOKEN=
GITHUB_REPOSITORY=morenoluisg/template_mads_project
GITHUB_PROJECT_OWNER=morenoluisg
GITHUB_PROJECT_TITLE=template_mads_project
WORKPLAN_EXCEL_PATH=data/raw/Plan de Trabajo.xlsx
WORKPLAN_YEAR=2026
WORKPLAN_TODAY=2026-09-09

ROLE_ARQUITECTO_DE_DATOS=morenoluisg
ROLE_CIENTIFICO_DE_DATOS_JR=leonard-max
```

Importante: `.env` esta ignorado por Git. No suba tokens ni secretos al repositorio.

## Roles Y Usuarios GitHub

Los responsables del Excel se convierten en:

- Campo `Rol` dentro del Project.
- Etiquetas `rol: ...` en cada issue.
- Assignees cuando exista mapeo a usuario de GitHub y el usuario tenga acceso al repo.

Ejemplos:

```env
ROLE_ARQUITECTO_DE_DATOS=morenoluisg
ROLE_CIENTIFICO_DE_DATOS_JR=leonard-max
```

Para agregar otro rol, use el nombre normalizado en mayusculas, sin tildes y con guiones bajos:

```env
ROLE_DESARROLLADOR=usuario-github
ROLE_DEVOPS=usuario-github
ROLE_ARQUITECTO_SIG=usuario-github
ROLE_PROFESIONAL_SIG=usuario-github
ROLE_ING_INDUSTRIAL=usuario-github
```

Si un usuario no puede ser asignado, revise que tenga acceso al repositorio.

## Instalacion

Este proyecto usa Python 3.12.

Crear o activar el entorno virtual:

```powershell
.\.venv\Scripts\activate
```

Instalar dependencias:

```powershell
pip install -r requirements.txt
```

## Sincronizacion

Ejecutar:

```powershell
.\.venv\Scripts\python.exe scripts\sync_github_project.py
```

El script realiza estas acciones:

1. Lee el Excel configurado en `WORKPLAN_EXCEL_PATH`.
2. Crea o reutiliza el GitHub Project con nombre `GITHUB_PROJECT_TITLE`.
3. Vincula el Project al repositorio.
4. Crea labels por rol y sprint.
5. Crea milestones por producto.
6. Crea o actualiza issues por cada fila del plan.
7. Asigna `Rol`, `Sprint`, `Inicio`, `Fin`, `Milestone` y `Status`.
8. Crea o actualiza las vistas del Project.

La sincronizacion es idempotente: se puede ejecutar de nuevo despues de modificar el Excel.

## Actualizar El Plan

Para actualizar el tablero:

1. Modifique `data/raw/Plan de Trabajo.xlsx`.
2. Mantenga las columnas esperadas.
3. Verifique que cada fila tenga `ID`, `Actividad`, `Responsable`, `Mes` y `Relacion de producto` cuando aplique.
4. Ejecute nuevamente el script de sincronizacion.
5. Revise el Project en GitHub.

## Buenas Practicas Operativas

- Mantener el Excel como fuente inicial de planeacion.
- Gestionar el dia a dia desde GitHub Projects.
- Usar `Status` para Scrum: `Todo`, `In Progress`, `Done`.
- Usar `Milestone` para seguimiento por producto.
- Usar `Sprint` para seguimiento mensual.
- Usar `Roadmap` para revisar ventanas de ejecucion.
- Mantener actualizado el mapeo de roles a usuarios GitHub.
- Evitar cambios manuales en titulos con prefijo `[ID]`, porque el script los usa para reconocer issues existentes.

## Seguridad

- No almacenar tokens reales en el repositorio.
- Usar `.env` local para credenciales.
- Revocar tokens que hayan sido compartidos por chat o correo.
- Preferir tokens con permisos minimos necesarios: repositorio, issues y projects.

## Estado Actual De Referencia

Con corte `2026-09-09`, el Project quedo con:

- 70 items sincronizados.
- 32 milestones/productos.
- 32 items en `Done`.
- 13 items en `In Progress`.
- 25 items en `Todo`.
- 67 items con fechas `Inicio` y `Fin`.

Los 3 items sin fechas corresponden a filas del Excel sin mes asignado.
