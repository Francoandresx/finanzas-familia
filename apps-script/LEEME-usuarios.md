# Usuarios y registro de quién agrega cada gasto

La app funciona igual que siempre hasta que completes estos pasos. Es un script **aparte**: no modifica tu hoja ni tu Apps Script actual.

## 1. Crear el script (una sola vez)
1. Entra a https://script.google.com y pulsa **Nuevo proyecto**.
2. Borra lo que trae y pega todo el contenido de `apps-script/usuarios-bitacora.gs`.
3. En ⚙ **Configuración del proyecto → Propiedades del script**, agrega:
   - `HOJA_ID`: el ID de tu hoja de Google (en su URL, entre `/d/` y `/edit`).
   - `CLAVE`: la misma clave que ya usa la app (el valor de `?k=...` en tu enlace).
4. **Implementar → Nueva implementación → Aplicación web**. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona**. Autoriza los permisos que pida.
5. Copia la URL que termina en `/exec`.

## 2. Conectar la app
Pega esa URL en `index.html`, en la línea `var EXEC2_URL = "";`, o pásasela a Claude para que la ponga.

## 3. Agregar a la familia
La primera vez que se use, el script crea la pestaña **USUARIOS** en tu hoja con **Franco** y **Anny**. Puedes editarla a mano:

| nombre | color | activo |
|---|---|---|
| Franco | #38D591 | TRUE |
| Anny | #F2BE22 | TRUE |

Colores sugeridos de la paleta: `#38D591`, `#F2BE22`, `#F29727`, `#26A699`, `#A3365A`, `#3B6E8C`. Para quitar a alguien, pon `FALSE` en «activo».

## Cómo funciona
- Al abrir la app por primera vez en un teléfono, cada persona elige su perfil («¿Quién eres?»). Se puede cambiar en **Ajustes → Cambiar de usuario**.
- Cada gasto, ingreso, ajuste o conteo se anota en la pestaña **BITACORA** con el nombre de quien lo hizo.
- Franco tiene como foto un T-rex animado y Anny una oveja animada (circulares, ya incluidas en la app). Otros nombres usan un círculo con su inicial.
- En **Historial** aparece el autor en cada movimiento y la pestaña **Actividad** con todo lo que hizo cada persona.

## Límites (importante)
- Elegir perfil **no es una clave**: cualquiera con el enlace de la app puede elegir cualquier perfil. Sirve para llevar el registro entre personas de confianza, no como seguridad.
- Si no hay internet al guardar, el registro se reintenta después (hasta 50 pendientes por teléfono).
- Los movimientos que ya existían antes no tienen autor.
