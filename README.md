# MesasCasamiento

App web para organizar invitados y mesas de un evento.

## Estado actual

La app ahora trabaja con:

- registro e inicio de sesion por usuario
- multiples eventos por usuario
- persistencia por evento en Cloudflare Pages Functions + D1
- gestion de mesas con distintas capacidades
- importacion de invitados desde Excel
- importacion/exportacion de asignaciones por CSV

## Reglas de UX

- no usar `alert`, `prompt` ni `confirm`
- toda interaccion sensible debe pasar por modal
- aplica a login, registro, alta/edicion de invitados y gestion del evento

## Modelo funcional

Cada usuario puede:

- registrarse
- crear varios eventos
- ver su lista de eventos
- entrar a un evento y configurar mesas
- agregar mesas individuales o en bloque
- definir capacidades distintas por mesa
- cargar invitados y asignarlos

## Estructura principal

- frontend: [app.js](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/app.js), [index.html](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/index.html), [styles.css](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/styles.css)
- API Pages Functions:
  - [session.js](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/functions/api/session.js)
  - [events.js](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/functions/api/events.js)
  - [state.js](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/functions/api/state.js)
  - [\_db.js](/Users/rubenbenayon/Documents/Desarrollos/MesasCasamiento/functions/api/_db.js)

## Correr local

1. Instala dependencias:

```bash
npm install
```

2. Levanta la app:

```bash
npm run dev
```

3. Abre la URL local que te muestre Wrangler. Normalmente:

```text
http://127.0.0.1:8788
```

4. Registra una cuenta desde el modal de acceso y crea tu primer evento.

## Scripts

```bash
npm run dev
npm run dev:3000
```

`dev:3000` sirve si `8788` esta ocupado.

## D1 / Pages

La app espera un binding `DB` en Pages Functions.

Referencia de la base usada en Pages:

```json
{
  "d1_databases": [
    {
      "binding": "mesascasamiento_db",
      "database_name": "mesascasamiento-db",
      "database_id": "aa915d6b-ea69-492d-8d3d-6d8ebc615090"
    }
  ]
}
```

En produccion, el binding activo para Functions debe llamarse `DB`.

## Notas de desarrollo

- Wrangler crea el binding local de D1 al correr `pages dev`
- el schema se crea automaticamente al primer acceso
- si vienes de una base local vieja, la app recrea las tablas legacy necesarias para adaptarlas al esquema multievento
- si Wrangler da problemas de logs en macOS:

```bash
HOME=/tmp npm run dev
```

## Archivos locales no versionados

El repo ya ignora:

- `node_modules/`
- `.wrangler/`
- `Invitados Sharon y Ariel.xlsx`
- `mesas_asignacion.csv`

Esos archivos pueden existir localmente, pero no deben subirse al repo.
