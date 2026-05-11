# Prompt agnostico para crear un organizador de mesas para eventos

Quiero crear una aplicacion web para organizar invitados y mesas de eventos, especialmente casamientos, bar/bat mitzva, fiestas o reuniones grandes. La aplicacion debe ser multiusuario y cada usuario debe poder administrar varios eventos independientes.

## Objetivo principal

Construir una herramienta donde una persona pueda crear un evento, cargar invitados, crear mesas con distintas capacidades y asignar invitados a esas mesas de forma visual, simple y persistente.

La aplicacion debe priorizar el flujo operativo real: entrar, elegir el evento, ver los invitados sin asignar, ver el plano/listado de mesas, mover invitados entre mesas y guardar automaticamente los cambios.

## Requisitos funcionales

### Usuarios y sesiones

- Permitir registro de usuario con nombre, email y password.
- Permitir inicio de sesion con email y password.
- Mantener la sesion activa entre visitas.
- Permitir cerrar sesion.
- Cada usuario solo puede ver y modificar sus propios eventos.

### Eventos

- Cada usuario puede crear multiples eventos.
- Cada evento debe tener al menos:
  - id unico
  - nombre
  - fecha del evento
  - fecha de creacion
  - usuario propietario
- Mostrar un panel lateral o seccion llamada "Tus eventos" con todos los eventos del usuario.
- Mostrar la cantidad total de eventos.
- Permitir crear un evento nuevo desde un modal o formulario dedicado.
- El alta de evento debe pedir nombre y fecha del evento.
- El campo de fecha del evento debe mantener el mismo ancho, tipografia y estilo visual que el resto de los campos del formulario.
- Permitir renombrar el evento activo mientras la fecha del evento no haya pasado.
- Permitir cambiar la fecha del evento activo mientras la fecha del evento no haya pasado.
- Permitir borrar un evento mientras la fecha del evento no haya pasado.
- Al ingresar a la aplicacion, si el usuario ya tiene eventos, debe abrirse automaticamente el ultimo evento usado o, si no existe, el primer evento disponible.
- El evento activo debe quedar visualmente abierto/destacado en el panel "Tus eventos".
- El titulo principal debe indicar claramente el evento activo, por ejemplo: "Organizador de Mesas - Nombre del evento".
- Si se crea un evento nuevo, ese evento debe pasar a ser el evento activo inmediatamente y debe quedar abierto en "Tus eventos".
- Cuando la fecha del evento ya paso, el evento debe quedar en modo solo lectura.
- En modo solo lectura se puede ver el evento y exportar las asignaciones a Excel.
- En modo solo lectura ya no se puede:
  - cambiar nombre ni fecha
  - borrar el evento
  - importar archivos
  - agregar, editar o borrar invitados
  - agregar, borrar, reordenar o editar mesas
  - mover invitados de mesa o asignarlos

### Mesas

- Cada evento tiene sus propias mesas.
- Permitir crear una mesa individual.
- Permitir crear un bloque de varias mesas al mismo tiempo.
- Cada mesa debe tener:
  - numero
  - nombre editable o derivado, por ejemplo "Mesa 1"
  - capacidad
  - tipo opcional, por ejemplo hombres/mujeres/sin tipo
  - posicion u orden visual
- Permitir editar una mesa:
  - numero
  - nombre
  - capacidad
  - tipo
- Permitir borrar una mesa.
- Si se borra una mesa, los invitados asignados a ella deben volver a "Sin asignar".
- Mostrar claramente si una mesa esta llena, tiene lugares disponibles o excede su capacidad.

### Invitados

- Cada evento tiene sus propios invitados.
- Cada invitado debe tener:
  - id unico
  - nombre completo
  - genero o grupo, usando al menos H/M o equivalente configurable
  - estado de confirmacion
  - fila de origen del archivo importado, si aplica
  - mesa asignada, si aplica
- Permitir agregar invitados manualmente.
- Permitir editar invitados.
- Permitir borrar invitados.
- Mostrar una lista de invitados sin asignar.
- Permitir filtrar invitados por:
  - texto de busqueda
  - genero/grupo
- Mostrar estadisticas del evento:
  - invitados visibles
  - invitados asignados
  - invitados sin asignar
  - cantidad de mesas
  - hombres/mujeres o grupos equivalentes
  - confirmados

### Asignacion visual

- La pantalla principal debe mostrar:
  - panel de acciones
  - panel de eventos
  - resumen del evento
  - filtros
  - lista de invitados sin asignar
  - area de mesas
- Debe ser posible mover invitados entre:
  - lista "Sin asignar"
  - cualquier mesa
  - una mesa y otra mesa
- La asignacion puede implementarse con drag and drop, botones, menues o cualquier interaccion clara, pero debe ser comoda para uso repetido.
- Los cambios deben guardarse automaticamente o de forma muy clara para el usuario.

### Importacion y exportacion

- Permitir importar invitados desde Excel.
- Permitir importar asignaciones desde Excel.
- Antes de importar asignaciones, pedir confirmacion indicando a que evento activo se van a aplicar.
- Permitir exportar las asignaciones actuales a Excel.
- El archivo exportado debe incluir informacion util para reconstruir o revisar:
  - tipo de registro
  - nombre
  - genero/grupo
  - confirmado
  - mesa
  - fila original
  - capacidad de mesa
- La importacion debe afectar solamente al evento activo.

## Persistencia de datos

La aplicacion debe guardar en una base de datos persistente:

- usuarios
- sesiones
- eventos
- mesas
- invitados
- revision o metadata del estado del evento, si se necesita sincronizacion

Modelo sugerido, independiente de tecnologia:

```text
users
- id
- name
- email
- password_hash
- created_at

sessions
- token
- user_id
- created_at

events
- id
- user_id
- name
- event_date
- created_at

event_meta
- event_id
- key
- value

tables
- event_id
- number
- name
- type
- capacity
- position

guests
- id
- event_id
- name
- gender
- confirmed
- source_row
- table_number
```

El sistema debe validar siempre que el evento pertenezca al usuario autenticado antes de leer o modificar datos.

## API o capa de backend esperada

No importa la tecnologia, pero la aplicacion debe exponer operaciones equivalentes a:

### Sesion

- registrar usuario
- iniciar sesion
- obtener usuario actual
- cerrar sesion

### Eventos

- listar eventos del usuario
- crear evento
- renombrar evento
- borrar evento

### Estado del evento

- cargar mesas e invitados de un evento
- guardar mesas e invitados de un evento
- opcional: manejar revision para evitar sobrescribir cambios recientes

## UX y reglas de interfaz

- No usar `alert`, `prompt` ni `confirm`.
- Las acciones sensibles deben usar modales o dialogos propios:
  - login
  - registro
  - crear evento
  - editar evento
  - agregar/editar invitado
  - confirmar importacion
  - borrar datos importantes
- La interfaz debe ser clara, operativa y densa, mas parecida a una herramienta de gestion que a una landing page.
- Evitar pantallas de marketing; la primera pantalla despues de login debe ser la herramienta.
- El usuario siempre debe saber en que evento esta trabajando.
- El evento activo debe coincidir en:
  - titulo principal
  - resumen superior
  - panel "Tus eventos"
  - datos cargados en mesas e invitados
- Si no hay eventos, mostrar un estado vacio claro e invitar a crear el primer evento.
- Si no hay mesas, mostrar un estado vacio en el area de mesas.
- Si no hay invitados sin asignar, mostrar estado vacio en la lista correspondiente.

## Comportamiento importante

- Al abrir la app con una sesion valida:
  1. cargar la sesion
  2. cargar los eventos del usuario
  3. elegir el evento activo
  4. cargar mesas e invitados de ese evento
  5. mostrar ese mismo evento como abierto/destacado en "Tus eventos"
- Al crear un evento:
  1. guardarlo
  2. actualizar la lista de eventos
  3. seleccionarlo como activo
  4. abrirlo/destacarlo en "Tus eventos"
  5. mostrar su plano vacio
- Al cambiar de evento:
  1. guardar o descartar de forma segura los cambios pendientes
  2. limpiar el estado local anterior
  3. cargar el nuevo evento
  4. actualizar titulo, resumen, panel de eventos, mesas e invitados
- Al borrar el evento activo:
  1. borrarlo
  2. elegir otro evento disponible como activo
  3. si no queda ninguno, mostrar estado vacio

## Seguridad basica

- Nunca guardar passwords en texto plano.
- Hashear passwords con un algoritmo adecuado.
- Las sesiones deben usar tokens dificiles de adivinar.
- Las rutas o acciones privadas deben requerir sesion valida.
- Todas las operaciones sobre eventos, mesas e invitados deben verificar propiedad del usuario.
- Validar inputs del lado servidor/backend, no solo en frontend.

## Resultado esperado

Entregar una aplicacion funcional donde un usuario pueda:

1. Registrarse o iniciar sesion.
2. Crear un evento.
3. Agregar mesas.
4. Importar o crear invitados.
5. Asignar invitados a mesas.
6. Ver estadisticas del evento.
7. Exportar asignaciones.
8. Volver mas tarde y encontrar el mismo estado guardado.

La implementacion puede hacerse con cualquier stack tecnologico, siempre que cumpla los comportamientos anteriores y mantenga separada la informacion de cada usuario y cada evento.
