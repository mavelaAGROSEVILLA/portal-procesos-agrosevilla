# Portal de procesos Agrosevilla

## Estado actual

El web part `ProcessView` ya no muestra la pantalla de bienvenida de SPFx. Actualmente funciona como un catálogo navegable de procesos de SharePoint.

La navegación implementada es:

```text
Departamento -> Grupo de procesos -> Proceso -> Ficha del proceso
```

El sitio utilizado para las pruebas es:

```text
https://agrosevilla.sharepoint.com/sites/Gestindocumentaldeprocesos
```

## Funcionalidades implementadas

### Catálogo

- Carga inicial de departamentos reales de SharePoint.
- Selección de un departamento.
- Carga diferida de sus grupos de procesos.
- Selección de un grupo.
- Carga diferida de los procesos relacionados.
- Caché en memoria durante la sesión para evitar consultas repetidas.
- Una única rama principal seleccionada.
- Acción `Cerrar rama`.
- Estados de carga, error y ausencia de datos en cada nivel.
- Ordenación consistente por código y nombre.
- Contador de grupos de procesos en cada departamento.
- Contador de procesos en cada grupo.

Al abrir la ficha de un proceso, el árbol se compacta y conserva únicamente la ruta seleccionada:

```text
Departamento seleccionado -> Grupo seleccionado -> Proceso seleccionado
```

La ficha se muestra debajo de esta ruta.

### Búsqueda global

El encabezado incluye un buscador global con autocompletado mientras se escribe.

- Busca departamentos, grupos y procesos.
- Empieza a consultar a partir de dos caracteres.
- Usa un debounce de 300 ms para evitar una petición por cada tecla.
- Muestra el tipo de resultado y su código.
- Al seleccionar un departamento o grupo, lleva a su rama.
- Al seleccionar un proceso, carga su rama y abre su ficha.

### Ficha de proceso

Al pulsar `Ver ficha` se muestra:

- Código.
- Nombre.
- Responsable.
- Estado.
- Tipo de proceso.
- Descripción.
- Documentos asociados.

El responsable se obtiene desde el campo persona `PropietarioProceso` mediante `$expand`.

### Documentos asociados

La aplicación consulta la biblioteca `Documentos de procesos` mediante el lookup `ProcesoRelated`.

Cada documento utiliza `FileRef` y se abre en una pestaña nueva mediante:

```text
 target="_blank"
 rel="noreferrer"
```

Cuando todavía no hay documentos, se muestra un estado vacío en lugar de datos simulados.

La ficha incluye previsualizaciones específicas cuando existen estos documentos:

- `TipoDocumental = Ficha Maestra` y archivo PDF: visor PDF embebido con enlace para abrirlo en una pestaña nueva.
- `TipoDocumental = BPMN` y archivo PNG: visor de imagen con zoom mediante rueda o botones, desplazamiento arrastrando y opción para abrir la imagen en una pestaña nueva.
- El zoom BPMN permite ampliar hasta el 1000%.

- `TipoDocumental = Carpeta documental`: se consulta el contenido de la carpeta mediante PnPjs y se muestran sus archivos y subcarpetas debajo del visor BPMN.
- El nombre de la carpeta y cada elemento de su contenido abren la ubicación de SharePoint en una pestaña nueva.

Si el documento no existe o su extensión no coincide con el tipo esperado, se muestra un mensaje informativo y se conserva el enlace en la lista de documentos.

### Alta de información

El encabezado dispone de la acción `Añadir`, con dos opciones:

- `Grupo de procesos`.
- `Proceso`.

#### Crear un grupo

El formulario solicita:

- Código.
- Nombre del grupo.
- Departamento.

El grupo se crea en la lista `Grupo de procesos`.

#### Crear un proceso

El formulario solicita:

- Código.
- Nombre del proceso.
- Departamento.
- Grupo de procesos.
- Responsable mediante correo corporativo.
- Descripción.

El selector funciona de forma dependiente:

1. Se selecciona un departamento.
2. Solo aparecen los grupos de ese departamento.
3. Al seleccionar un grupo, se ajusta el departamento propietario.

El responsable se resuelve con `ensureUser` y se guarda su ID de SharePoint.

Los procesos nuevos se guardan en la lista:

```text
Cartera de procesos
```

La navegación existente continúa leyendo procesos desde:

```text
Procesos
```

Después de guardar, se invalida la caché correspondiente y la rama se vuelve a cargar automáticamente.

## Arquitectura

### Web part

- [ProcessViewWebPart.ts](src/webparts/processView/ProcessViewWebPart.ts)
  - Configura PnPjs una sola vez.
  - Usa el contexto real de SPFx.
  - Pasa la instancia `SPFI` al componente raíz.

- [ProcessView.tsx](src/webparts/processView/components/ProcessView.tsx)
  - Actúa como adaptador funcional hacia `Home`.

### Modelos

- [ProcessModels.ts](src/models/ProcessModels.ts)
  - `IDepartment`
  - `IProcessGroup`
  - `IProcess`
  - `IDocumentoProceso`

Las respuestas de SharePoint se transforman en estos modelos antes de llegar a React.

### Servicios

- [SharePointService.ts](src/services/SharePointService.ts)
  - Acceso común a PnPjs.
  - Resolución dinámica de nombres internos.
  - Diagnóstico de campos SharePoint en consola durante desarrollo.
  - Conversión segura de textos, IDs y personas.

- [DepartmentService.ts](src/services/DepartmentService.ts)
  - Lectura y ordenación de departamentos.

- [ProcessGroupService.ts](src/services/ProcessGroupService.ts)
  - Lectura por departamento.
  - Lectura global para los formularios.
  - Creación de grupos.

- [ProcessService.ts](src/services/ProcessService.ts)
  - Lectura de procesos por grupo.
  - Conteo de procesos agrupados por grupo.
  - Lectura desde `Procesos`.
  - Creación desde `Cartera de procesos`.

- [DocumentService.ts](src/services/DocumentService.ts)
  - Lectura de documentos asociados a un proceso.

- [DocumentPreview.tsx](src/pages/Process/DocumentPreview.tsx)
  - Previsualización de PDF y diagrama BPMN.
  - Zoom y desplazamiento del PNG BPMN.

- [SearchService.ts](src/services/SearchService.ts)
  - Búsqueda global de departamentos, grupos y procesos.

### Componentes de interfaz

- [GlobalSearch.tsx](src/pages/Home/GlobalSearch.tsx)
  - Entrada de búsqueda y autocompletado.

- [CreateItemPanel.tsx](src/pages/Home/CreateItemPanel.tsx)
  - Formularios de alta de grupos y procesos.

### Páginas

- [Home.tsx](src/pages/Home/Home.tsx)
  - Estado local de la navegación.
  - Cachés de grupos y procesos.
  - Apertura de formularios y ficha de proceso.

- [Department.tsx](src/pages/Department/Department.tsx)
  - Nivel de departamentos.

- [ProcessGroup.tsx](src/pages/ProcessGroup/ProcessGroup.tsx)
  - Nivel de grupos.

- [Process.tsx](src/pages/Process/Process.tsx)
  - Nivel de procesos.

- [ProcessDetail.tsx](src/pages/Process/ProcessDetail.tsx)
  - Ficha del proceso y documentos asociados.

- [CreateItemPanel.tsx](src/pages/Home/CreateItemPanel.tsx)
  - Formularios de alta de grupos y procesos.

## Configuración de SharePoint

Los servicios no codifican el dominio ni una URL absoluta. PnPjs se configura usando el contexto del web part:

```text
this.context.pageContext.web
```

La página inicial de Heft está configurada en:

[serve.json](config/serve.json)

con el sitio:

```text
/sites/Gestindocumentaldeprocesos
```

Los nombres internos se resuelven consultando los metadatos de cada lista. Esto permite trabajar con columnas que visualmente aparecen como `Código`, pero internamente pueden ser `Title`, `Codigo` u otro nombre.

## Prueba local

Configurar el tenant y arrancar Heft:

```powershell
$env:SPFX_SERVE_TENANT_DOMAIN="agrosevilla.sharepoint.com"
npm start
```

Abrir el workbench hospedado:

```text
https://agrosevilla.sharepoint.com/sites/Gestindocumentaldeprocesos/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true
```

Para probar la aplicación:

1. Mantener `npm start` ejecutándose.
2. Abrir el workbench del sitio de procesos.
3. Añadir el web part `ProcessView`.
4. Seleccionar un departamento.
5. Seleccionar un grupo.
6. Seleccionar un proceso y pulsar `Ver ficha`.
7. Probar `Añadir` desde el encabezado.
8. Escribir al menos dos caracteres en `Buscar en el catálogo` y seleccionar una sugerencia.

No ejecutar `npm run build` mientras `npm start` está ejecutándose, porque ambos procesos limpian o regeneran artefactos de desarrollo.

## Validación realizada

- TypeScript sin errores en los archivos modificados.
- Compilación Heft realizada correctamente durante el desarrollo.
- Webpack de desarrollo iniciado correctamente en `https://localhost:4321`.
- Jest no contiene suites de pruebas actualmente.
- No se han cambiado versiones de SPFx, React, TypeScript o PnPjs.
- No se han añadido dependencias externas.
- No se ha utilizado Gulp.

## Limitaciones conocidas

- La biblioteca de documentos necesita contener elementos con el lookup `ProcesoRelated` para poder probar los enlaces.
- El responsable del proceso debe introducirse como correo corporativo válido para que `ensureUser` pueda resolverlo.
- Las columnas Choice como `Status` y `TipoProceso` se muestran en la ficha, pero no se editan mediante texto libre para evitar valores no válidos.
- Los avisos de `spserviceworker.js`, iconos duplicados, foto de usuario 404 y políticas del workbench pertenecen al entorno de SharePoint y no al catálogo.

## Próximos incrementos posibles

- Detalle ampliado del proceso.
- Consulta y filtros de documentos por área y tipo documental.
- Sistemas relacionados.
- Procesos relacionados.
- Edición controlada de procesos.
- Valores Choice cargados dinámicamente desde los metadatos de SharePoint.
- KPIs y métricas del proceso.
- Pruebas automatizadas de servicios y componentes.
