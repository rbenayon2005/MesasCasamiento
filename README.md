# MesasCasamiento

App web para organizar invitados y mesas de un casamiento.

## Regla del proyecto

- No usar `alert`, `prompt` ni `confirm`.
- Todas las interacciones con usuario deben ser via modal.
- Aplica especialmente a:
  - Login (usuario/password)
  - Alta de invitado (nombre, apellido y genero por selector `Hombre`/`Mujer`)

## Produccion

- URL: `https://mesascasamiento.pages.dev`
- Persistencia: Cloudflare Pages Functions + D1

## Credenciales de acceso

- Usuario: `adminmesas`
- Password: `mesas2026`

La app pide estas credenciales al abrir y las guarda en `localStorage`.

## D1 configurada

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

En Pages el binding activo para Functions debe llamarse `DB`.
