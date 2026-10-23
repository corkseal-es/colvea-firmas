# Firmas de correo Colvea

Firmas de correo de Colvea / CorkSeal para Outlook (Exchange Online / Microsoft 365).

## Qué hay en esta carpeta

| Archivo | Para qué sirve |
|---|---|
| `generador.html` | Formulario para crear una firma a mano: rellenas los datos, ves la vista previa y la copias o descargas. |
| `scripts/comerciales.csv` | Lista de personas con firma (email, nombre, cargo, teléfono, dirección, mapa). |
| `scripts/plantilla-firma.html` | Diseño de la firma con huecos `{{NOMBRE}}`, `{{CARGO}}`, etc. |
| `scripts/aplicar-firmas.ps1` | Script que genera las firmas desde el CSV y las pone en el Outlook de cada persona. |
| `scripts/salida/` | Firmas generadas por el script para revisarlas (no se sube a Git). |
| `firma_*-colvea.html` | Firmas individuales ya creadas. |
| `img/` | Logo e iconos de redes. Se cargan desde GitHub, así que **cualquier cambio aquí hay que subirlo** (commit + push). |

---

## Poner o actualizar las firmas en Outlook (lo habitual)

### 1. Editar la lista de personas

Abre `scripts/comerciales.csv` con el **Bloc de notas**. Una línea por persona, datos separados por `;`:

```
Email;Nombre;Cargo;Telefono;Direccion;Mapa
maria@corkseal.es;María López;Comercial;+34 600 11 22 33;;
```

- Si `Direccion` y `Mapa` se dejan vacíos, se usan los de la oficina de Terrassa.
- Si lo editas con Excel, guarda como **"CSV UTF-8"** o los acentos saldrán mal.

### 2. Abrir PowerShell en la carpeta

```
cd C:\Users\Edgar\Documents\firmas-colvea
Set-ExecutionPolicy -Scope Process Bypass -Force
```

La segunda línea solo permite ejecutar el script en esa ventana.

### 3. Generar y revisar (no toca ningún buzón)

```
.\scripts\aplicar-firmas.ps1
```

Abre los archivos de `scripts\salida\` en el navegador y comprueba que todo está bien.

### 4. Aplicar

Una sola persona:

```
.\scripts\aplicar-firmas.ps1 -Aplicar -Solo maria@corkseal.es
```

Todos los del CSV:

```
.\scripts\aplicar-firmas.ps1 -Aplicar
```

- Pide iniciar sesión con una cuenta **administradora** de Microsoft 365. No hace falta darse permisos sobre cada buzón.
- Cada persona debe salir con `OK` en verde. Si sale `ERROR` en rojo, lee el mensaje.
- Sustituye la firma que tuviera puesta antes.
- La firma se añade **en correos nuevos y en respuestas/reenvíos**.
- Puede tardar desde unos minutos hasta unas horas en verse en Outlook web.

### 5. Comprobar

```
Get-MailboxMessageConfiguration maria@corkseal.es | Select AutoAddSignature, AutoAddSignatureOnReply
```

Las dos columnas deben salir en `True`. Después abre Outlook web de esa persona y crea un correo nuevo.

---

## Firmas en la nube (importante)

Outlook web puede guardar las firmas "en la nube" (*roaming signatures*). Mientras eso esté activo, **ignora la firma puesta por PowerShell**. Por eso, al principio, en algunos buzones funcionaba y en otros no.

El script lo comprueba y avisa. Para desactivarlo (solo hace falta una vez):

```
.\scripts\aplicar-firmas.ps1 -Aplicar -DesactivarFirmasEnLaNube -Solo edgar@corkseal.es
```

A tener en cuenta:

- Afecta a **toda la organización**.
- Las firmas creadas a mano en Outlook web no se borran: simplemente dejan de usarse.
- Si alguien sigue viendo la firma antigua, bórrala a mano en Outlook web (⚙️ Configuración → Cuentas → Firmas) y vuelve a lanzar el script para esa persona.
- Si algún día se reactivan (`Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $false`), las firmas antiguas volverían a aparecer.
- Comprobar el estado: `Get-OrganizationConfig | Select PostponeRoamingSignaturesUntilLater` (`True` = desactivadas, que es lo que queremos).

**Outlook de escritorio clásico** guarda las firmas en cada ordenador; este script no las cambia. Ahí hay que pegarla a mano (usa `generador.html` → *Copiar firma*).

---

## Crear una firma a mano

Abre `generador.html` en el navegador, rellena los datos y pulsa:

- **Copiar firma** → pégala en la configuración de firma del correo.
- **Descargar HTML** → guarda el archivo `firma_nombre-apellido-colvea.html`.

---

## Cambiar el diseño, logo o iconos

- **Diseño de la firma:** editar `scripts/plantilla-firma.html` (para el script) **y** la función `firmaHTML` de `generador.html` (para el formulario). Son dos copias del mismo diseño: si cambias una, cambia la otra.
- **Logo o iconos:** sustituir el archivo en `img/` (mismo nombre, mejor a doble tamaño del que se muestra: iconos de 64×64 o más, logo de 230 px de ancho o más) y subirlo:

  ```
  git add img
  git commit -m "Actualizar iconos"
  git push
  ```

  Las imágenes se sirven a través de `wsrv.nl`, que guarda copia. Para forzar que cargue la versión nueva, cambia `v=2` por `v=3` en las URLs de las imágenes (en la plantilla y en el generador) y vuelve a aplicar las firmas.

Si los iconos se ven borrosos, casi siempre es porque la imagen subida a GitHub es más pequeña de lo que se muestra o porque no se ha hecho push.
