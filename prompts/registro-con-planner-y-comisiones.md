# Prompt para registro con planner, validacion por codigo y comisiones

Quiero agregar una logica de registro donde cada usuario nuevo quede asociado a un planner, y esa asociacion quede validada mediante un codigo enviado al planner.

## Objetivo

Cuando una persona se registra en la aplicacion, no debe quedar habilitada automaticamente para continuar el flujo normal hasta que seleccione un planner y valide un codigo de 6 digitos enviado a ese planner.

Esto sirve para que el sistema pueda saber que usuarios fueron dados de alta bajo cada planner, que eventos creo cada usuario, y asi poder calcular o revisar comisiones por planner.

## Flujo de registro

### 1. Seleccion de planner durante el registro

En el formulario de registro de usuario, ademas de los datos actuales:

- nombre
- email
- password

debe mostrarse un desplegable con los planners disponibles.

El usuario debe seleccionar obligatoriamente un planner antes de completar el alta.

El desplegable debe mostrar informacion clara del planner, por ejemplo:

- nombre del planner
- empresa o marca, si aplica
- email o identificador interno, si hace falta distinguirlos

No debe permitirse registrar una cuenta sin planner asociado.

### 2. Alta pendiente de validacion

Cuando el usuario completa el formulario y selecciona un planner:

1. El sistema crea el usuario en estado pendiente de validacion.
2. Guarda la relacion entre el usuario y el planner seleccionado.
3. Genera un codigo numerico de 6 digitos.
4. Asocia ese codigo al intento de validacion del usuario.
5. Envia el codigo al planner, no al usuario.
6. Muestra al usuario una pantalla o modal indicando que necesita ingresar el codigo enviado al planner.

El usuario no debe poder continuar al sistema principal hasta validar ese codigo.

### 3. Envio del codigo al planner

El codigo de 6 digitos debe enviarse al planner asociado al registro.

El mensaje debe indicar claramente:

- que una persona se esta registrando en la plataforma
- nombre del usuario
- email del usuario
- codigo de validacion de 6 digitos
- que el planner debe compartir ese codigo con el usuario si corresponde aprobarlo

Ejemplo conceptual:

```text
Nuevo registro asociado a tu cuenta de planner.

Usuario: Juan Perez
Email: juan@email.com

Codigo de validacion: 123456

Comparte este codigo con el usuario solo si corresponde aprobar su alta.
```

El canal de envio puede ser email, WhatsApp, SMS u otro mecanismo, segun la tecnologia disponible. La logica debe quedar desacoplada para poder cambiar el proveedor de envio mas adelante.

### 4. Pantalla de espera del usuario

Despues de registrarse, el usuario debe quedar en una pantalla de validacion.

Esa pantalla debe decir algo como:

```text
Tu cuenta esta pendiente de validacion.

Enviamos un codigo de 6 digitos al planner que seleccionaste.
Pedile ese codigo para continuar.
```

Debe mostrar un campo para ingresar el codigo.

Acciones esperadas:

- ingresar codigo
- reenviar codigo, si corresponde
- cerrar sesion o volver al login

No debe mostrar la aplicacion principal ni permitir crear eventos hasta que la validacion sea correcta.

### 5. Validacion del codigo

Cuando el usuario ingresa el codigo:

1. El backend verifica que el codigo corresponda a ese usuario.
2. Verifica que el codigo no este vencido.
3. Verifica que no haya sido usado antes.
4. Si es correcto, marca al usuario como validado.
5. Marca el codigo como usado.
6. Permite al usuario continuar el flujo normal de la aplicacion.

Una vez validado, el usuario debe comportarse como cualquier usuario normal:

- puede iniciar sesion
- puede crear eventos
- puede administrar mesas e invitados
- sus eventos quedan asociados a su usuario
- su usuario queda asociado al planner seleccionado

Si el codigo es incorrecto, debe mostrarse un error claro sin usar `alert`.

Ejemplo:

```text
El codigo ingresado no es correcto. Verifica el codigo con tu planner.
```

### 6. Expiracion y reenvio

El codigo deberia tener vencimiento, por ejemplo 15 o 30 minutos.

Si el codigo vence:

- el usuario debe poder pedir un nuevo codigo
- el sistema genera un nuevo codigo
- invalida el codigo anterior
- envia el nuevo codigo al planner

Debe evitarse el abuso del reenvio con algun limite razonable, por ejemplo:

- no permitir reenviar mas de una vez por minuto
- limitar cantidad de reenvios por hora

## Modelo de datos sugerido

Agregar entidades o campos equivalentes a estos:

```text
planners
- id
- name
- company_name
- email
- phone
- active
- created_at

users
- id
- name
- email
- password_hash
- planner_id
- planner_verified_at
- status
- created_at

planner_verification_codes
- id
- user_id
- planner_id
- code_hash
- expires_at
- used_at
- created_at
```

El campo `status` del usuario puede tener valores como:

```text
pending_planner_verification
active
disabled
```

El codigo no deberia guardarse en texto plano si se puede evitar. Idealmente se guarda hasheado, igual que una contrasena o token sensible.

## Reglas de seguridad

- El usuario no puede elegir o cambiar planner despues de validarse, salvo que lo haga un administrador.
- El codigo debe pertenecer al usuario y planner correctos.
- Un codigo usado no puede volver a usarse.
- Un codigo vencido no puede validarse.
- El backend debe validar siempre el estado del usuario antes de permitir acceso al sistema principal.
- Un usuario pendiente no debe poder crear eventos, mesas ni invitados.
- No confiar solamente en el frontend para bloquear el acceso.

## Backend administrativo para planners y comisiones

Debe existir un backend o panel administrativo que permita revisar los usuarios dados de alta por planner, y los eventos creados por esos usuarios.

El objetivo es poder calcular comisiones o controlar actividad comercial.

### Vista por planner

El administrador debe poder ver una lista de planners con metricas como:

- nombre del planner
- cantidad de usuarios asociados
- cantidad de usuarios validados
- cantidad de eventos creados por esos usuarios
- fecha del ultimo registro
- estado del planner

Al entrar a un planner, debe verse el detalle.

### Detalle de usuarios por planner

Para cada planner, mostrar los usuarios asociados:

- nombre del usuario
- email
- estado de validacion
- fecha de registro
- fecha de validacion
- cantidad de eventos creados
- ultimo evento creado

Debe poder filtrarse por:

- planner
- estado del usuario
- rango de fechas
- usuario/email

### Detalle de eventos por planner

El sistema debe poder mostrar eventos agrupados por planner, aunque tecnicamente los eventos pertenezcan a usuarios.

La relacion seria:

```text
planner -> users -> events
```

Para cada evento, mostrar:

- nombre del evento
- usuario propietario
- planner asociado al usuario
- fecha de creacion
- cantidad de invitados
- cantidad de mesas
- estado del evento, si existe
- metricas utiles para comision, si aplica

### Reporte para comisiones

Debe existir una vista o exportacion que permita calcular comisiones por planner.

El reporte deberia poder incluir:

- planner
- usuario
- email del usuario
- evento
- fecha de creacion del evento
- cantidad de invitados
- cantidad de mesas
- plan contratado o monto, si existe en el futuro
- comision calculada o campo manual, si aplica

Aunque inicialmente no haya pagos integrados, la estructura debe dejar preparada la relacion para cobrar comision por actividad generada por cada planner.

## API o endpoints esperados

La implementacion deberia exponer operaciones equivalentes a:

### Planners

- listar planners activos para el desplegable de registro
- crear planner desde backend/admin
- editar planner
- activar/desactivar planner

### Registro y validacion

- registrar usuario con planner seleccionado
- generar codigo de validacion
- enviar codigo al planner
- validar codigo ingresado por usuario
- reenviar codigo
- consultar estado de validacion del usuario actual

### Administracion

- listar usuarios por planner
- listar eventos por planner
- obtener metricas por planner
- exportar reporte de comisiones

## UX esperada

- El registro debe ser claro y no parecer un paso opcional.
- El selector de planner debe estar dentro del flujo normal de alta.
- Despues del registro, el usuario debe entender que esta esperando un codigo que recibio el planner.
- No usar `alert`, `prompt` ni `confirm`.
- Los errores y estados deben mostrarse con mensajes propios de la interfaz.
- Si el planner esta inactivo, no debe aparecer en el desplegable.
- Si no hay planners disponibles, no debe permitirse el registro y debe mostrarse un mensaje claro.

## Comportamiento completo esperado

Flujo final:

1. El usuario abre la app.
2. Elige registrarse.
3. Completa nombre, email y password.
4. Selecciona un planner desde un desplegable obligatorio.
5. Envia el registro.
6. El sistema crea el usuario como pendiente.
7. El sistema genera un codigo de 6 digitos.
8. El sistema envia ese codigo al planner.
9. El usuario queda en una pantalla de validacion.
10. El usuario pide el codigo al planner.
11. El usuario ingresa el codigo.
12. Si el codigo es correcto, el usuario queda activo.
13. El usuario continua el flujo normal de la app.
14. Todos los eventos que cree ese usuario quedan indirectamente asociados al planner.
15. El backend administrativo permite ver usuarios y eventos agrupados por planner para calcular comisiones.
