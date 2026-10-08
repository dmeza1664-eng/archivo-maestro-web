# Evaluación técnica de CRM — Pastelería Pepe

**Estado:** borrador interno. No fusionar hasta que la comparativa comercial (precios y funciones) se cruce con este informe.

**Alcance:** encaje técnico con los datos que hoy existen en este repo (`archivo-maestro-web`), con la app de galletas (referencia en GitHub) y con WhatsApp Business API. No evalúa precio, UX comercial ni “quién vende más pastel”.

**Fecha de las verificaciones de documentación oficial:** 8 de octubre de 2026.

---

## Hallazgo principal

**Hoy no hay un padrón de cliente de mostrador identificado.** El sistema de ventas (`pepes_devBI`) registra tickets por sucursal, bodega y producto. El campo que en algunos extractos se llama “cliente” no es una persona con teléfono o correo: en Planta León son sucursales (`SUC. ...`) o “público general”. No hay tablas de lealtad, pedidos especiales con cliente, ni teléfonos en las 7 tablas `dbo` que este repo puede leer.

Eso cambia el CRM que conviene. No se trata de “migrar una base de clientes”. Se trata de **elegir la capa que va a *crear* esa base** (WhatsApp, tiendas de galletas, y más adelante el mostrador), y de no intentar sincronizar cada ticket de piso.

La única fuente *identificada* cercana a un CRM que se pudo documentar está **fuera de este repo**: las tiendas de consignación de galletas en Firebase (`stores`: nombre, dirección, teléfono, encargado). El repositorio `dmeza1664-eng/app-galleta` **no existe o no es accesible** (GitHub API 404). Se usó como referencia `dmeza1664-eng/dashboard-pepes` (“Dashboard galletas”).

**Recomendación técnica:** **Kommo**, porque el problema real es capturar conversaciones de WhatsApp (pedidos de pastel, mayoreo, sucursal) y convertirlas en contactos sin padrón previo. HubSpot es la alternativa si se prioriza API, importación masiva y SDK Android + Firebase para la app. Odoo CRM/POS no conviene como primer paso: duplicaría el POS que ya vive en SQL Server.

**Esfuerzo de la arquitectura mínima (Kommo o HubSpot):** **14–22 días-persona**. No incluye sustituir el pronóstico ni el POS.

---

## 1. Qué datos de clientes existen hoy

### 1.1 Qué sí hay (este repo)

Este repo es el **Archivo Maestro**: pronóstico de demanda, producción y reparto por sucursal. Las fuentes de datos documentadas son:

| Fuente | Qué contiene | Dónde está en el repo |
| --- | --- | --- |
| SQL Server `pepes_devBI` (copia espejo, solo lectura) | `dbo.Ventas`, `dbo.VentaDet`, `dbo.Productos`, `dbo.Bodegas`, `dbo.Sucursales`, `dbo.Produccion`, `dbo.ProduccionDet` | `scripts/lib/pepes-fuente-base.cjs`, `README.md` |
| Extractos CSV de esas mismas consultas | Venta mensual producto × bodega × sucursal; producción mensual | `scripts/lib/pepes-fuente-archivos.cjs` |
| App web / backend MySQL-TiDB | `productos`, `ventas_diarias` (fecha, producto, canal, **cliente** texto), stock, producción, pronóstico, usuarios internos | `backend/schema.sql` |
| Catálogo | 113 SKUs (`stock_ideal`); 111 se pronostican | `datos/catalogo/mapeo-productos.csv` |

Columnas reales de la venta que el pronóstico lee (`VENTAS_COLUMNAS` en `scripts/lib/pepes-datos.cjs`):

`mes, nProductoPK, cCodigo, cDescripcion, nBodegaPK, bodegaNombre, nSucursalPK, sucursalNombre, qty, subtotal, total, lineas` (+ `qtyPromocion` en modo base).

**No hay** `telefono`, `email`, `nClientePK`, RFC de cliente final, ni lista de lealtad.

El backend sí tiene una columna `ventas_diarias.cliente VARCHAR(180)` (`backend/schema.sql`). Es una **etiqueta de agrupación** (fecha + producto + canal + cliente). En la documentación del modelo, el “cliente” de Planta León es la sucursal destino:

> tickets de sep-2025: 89.6% de las piezas a clientes "SUC. ...", 0% a público general, 100% ligado a un pedido de sucursal  
> — `CONTROL_MODELO_PRONOSTICO.md`, `App.jsx`

Es decir: cuando el ticket trae “cliente”, en el canal de planta **es otra sucursal**, no un comprador con WhatsApp.

`dbo.Produccion` / `dbo.ProduccionDet` son **órdenes de planta** (estados `Peticion` → `Finalizado`), no pedidos especiales de un particular.

Los `usuarios` del backend son operadores del Archivo Maestro (`admin` / `operador` / `consulta`), no clientes.

### 1.2 Qué no hay

En las tablas permitidas y en los extractos de este repo **no aparece**:

- Tabla `dbo.Clientes` (ni equivalente) en el perímetro que el código puede consultar.
- Teléfono, celular, WhatsApp o correo de comprador de mostrador.
- Ticket con persona identificada (nombre + medio de contacto).
- Programa de lealtad / puntos / cupón nominativo.
- Pedido especial (pastel personalizado) ligado a un contacto.
- Consentimiento / opt-in de marketing.

`pepes_espeBI` **no se menciona en este repo**. El modo base y el ejemplo de entorno usan solo `pepes_devBI` (`config/pronostico-automatico.env.example`). No se pudo verificar desde aquí si `pepes_espeBI` tiene un padrón de clientes: hay que preguntar a TI, con un `SELECT` de solo lectura sobre `INFORMATION_SCHEMA.TABLES`. Hasta que eso exista, el diseño debe asumir que **no hay clientes identificados en SQL Server**.

El conector de pronóstico **rechaza** cualquier tabla fuera de las 7 `dbo` listadas (`TABLAS_PERMITIDAS` en `pepes-fuente-base.cjs`). Aunque TI tenga más tablas, este repo hoy no las lee.

### 1.3 App de galletas (referencia externa)

| Intento | Resultado |
| --- | --- |
| `gh repo view dmeza1664-eng/app-galleta` | **404** — el repo no existe o no es público/accesible con las credenciales de este entorno |
| `dmeza1664-eng/dashboard-pepes` | Accesible. Descripción: “Dashboard galletas”. Firebase proyecto `cookie-1d48c` |
| `dmeza1664-eng/dashboard-galletas` | Accesible. Un HTML de supervisor / corte de caja |

De `dashboard-pepes` (reglas Firestore + `dashboard.js`):

- App Android en campo (versiones 1.0.3 / 1.0.4 citadas en `firestore.rules`); el dashboard es web sobre Firebase Auth + Firestore.
- Colecciones: `users`, `products`, `stores`, `visits`, `visitDetails`, `routeInventories`, `cashCuts`, `dailyStoreAssignments`, `storeConsignments`, `consignmentSwaps`.
- Roles: Administrador, Supervisor, Contadora, Repartidor.
- Una **tienda** (`stores`) sí es un contacto B2B: `nombre`, `municipio`, `colonia`, `direccion`, `zona`, **`telefono`**, **`encargado`**.
- No hay colección de consumidor final de pastelería. Son tiendas de consignación de galletas (Chocolate, Ate, Nuez, Alfajor) visitadas por repartidores.

Eso **sí** es importable a un CRM (empresas + encargado). No está en este repo; hay que leerlo desde Firestore con una cuenta de servicio, no con la `apiKey` pública del dashboard.

El usuario mencionó “backend en Azure API + Firebase functions”. En los repos accesibles de la org **no apareció** el código de esas Functions ni de una API Azure. Queda como supuesto operativo, no verificado en código.

### 1.4 Implicación para el CRM

| Pregunta | Respuesta |
| --- | --- |
| ¿Se puede “cargar el CRM con los clientes de Pepe”? | **No**, no hay archivo de clientes de mostrador. |
| ¿Qué se puede cargar el día 1? | Sucursales (empresas), productos del catálogo, y —si TI autoriza Firestore— tiendas de galletas con teléfono. |
| ¿De dónde saldrán los clientes? | WhatsApp (pedidos), captura en sucursal, y la app de galletas. |
| ¿Hay que sincronizar cada ticket de `dbo.Ventas`? | **No.** Son ventas anónimas de piso. Como mucho, KPIs por sucursal/SKU. |

---

## 2. Evaluación técnica de los seis CRM

Criterios: API REST y límites, webhooks, importación masiva, camino desde SQL Server, WhatsApp oficial (Cloud API / Business Platform de Meta), SDK o Firebase para la app, multi-sucursal. Cifras y URLs de documentación oficial consultada el 2026-10-08. Donde un proveedor no publica un conector nativo a SQL Server, se dice explícitamente.

Leyenda de encaje: **alto** / **medio** / **bajo** para *este* contexto (cadena de pastelerías en México, sin padrón, WhatsApp primero, POS ya en SQL Server, app Android + Firebase).

### 2.1 HubSpot

| Tema | Hecho verificado | Fuente |
| --- | --- | --- |
| API | REST CRM Objects (`/crm/objects/{version}/contacts`, `companies`, `products`, …). Auth OAuth o token de app privada. | [Contacts API](https://developers.hubspot.com/docs/api-reference/latest/crm/objects/contacts/guide), [Using Object APIs](https://developers.hubspot.com/docs/api-reference/latest/crm/using-object-apis) |
| Límites | Apps privadas: Free/Starter **100 req / 10 s** y **250 000 / día** por cuenta; Professional **190 / 10 s** y **625 000 / día**; Enterprise **190 / 10 s** y **1 000 000 / día**. Apps OAuth de marketplace: **110 req / 10 s** por cuenta instalada. HTTP 429 al exceder. | [API usage guidelines](https://developers.hubspot.com/docs/developer-tooling/platform/usage-guidelines) |
| Webhooks | Hasta **1 000** suscripciones por app. Hasta 100 eventos por POST. Concurrencia por defecto 10. Los POST del webhook **no cuentan** contra el rate limit. HTTPS obligatorio. | [Webhooks guide](https://developers.hubspot.com/docs/api-reference/latest/webhooks/guide) |
| Import masivo | Imports API: **80 000 000** filas/día; por archivo **1 048 576** filas o **512 MB**. También importador guiado en UI. | [Imports API](https://developers.hubspot.com/docs/api-reference/latest/crm/imports/guide) |
| SQL Server | **No hay conector oficial** HubSpot ↔ SQL Server. La vía soportada es API / Imports + un job propio (el mismo patrón que ya usa el pronóstico: `mssql` + solo lectura). Los “conectores SQL Server” que aparecen en búsqueda son de terceros (p. ej. ZappySys), no de HubSpot. | Guía de uso + ausencia en el catálogo de APIs oficiales |
| WhatsApp | Integración nativa con WhatsApp Business (bandeja / help desk). Requiere **Marketing Hub o Service Hub Professional o Enterprise**, cuenta Meta Business y número verificado. Coexistencia con la app WhatsApp Business. | [WhatsApp coexistence](https://knowledge.hubspot.com/inbox/connect-a-whatsapp-number-to-hubspot-using-coexistence), [producto](https://www.hubspot.com/products/whatsapp-integration), [help desk](https://knowledge.hubspot.com/help-desk/connect-a-whatsapp-channel-to-help-desk) |
| App / Firebase | **SDK oficial Android** de chat (`mobile-chat-sdk-android`) con `HubspotFirebaseMessagingService` (FCM). Encaja con una app que ya usa Firebase. | [Mobile chat SDK Android](https://developers.hubspot.com/docs/api-reference/latest/conversations/chat-configuration/mobile-chat-sdk/android), [repo HubSpot](https://github.com/HubSpot/mobile-chat-sdk-android/) |
| Multi-sucursal | Teams jerárquicos (API `/settings/teams`). No hay objeto “sucursal”: se modelan como Companies + propiedad `sucursal` / teams. | [Teams API](https://developers.hubspot.com/docs/api-reference/latest/account/settings/teams/guide) |

**Encaje:** alto en ingeniería (mejor API e import del grupo; único con SDK Android + Firebase oficial). Medio en el problema de negocio inmediato: WhatsApp nativo existe pero en planes Professional/Enterprise, y HubSpot está pensado para marketing sobre una base de contactos que **aún no existe**.

### 2.2 Zoho CRM

| Tema | Hecho verificado | Fuente |
| --- | --- | --- |
| API | REST v8 (`/crm/v8/...`). Créditos en ventana móvil de 24 h. Insert/Update/Upsert: 1 crédito / 10 registros, máx. 100 registros/llamada. | [API limits v8](https://www.zoho.com/crm/developer/docs/api/v8/api-limits.html) |
| Límites | Free 5 000 créditos; Standard/Starter 50 000 + 250/licencia (tope 100 000); Professional 50 000 + 500/licencia (tope 3 M); Enterprise/Zoho One 50 000 + 1 000/licencia (tope 5 M); Ultimate ilimitado. Concurrencia org/app: 5 / 10 / 15 / 20 / 25 según edición. Sub-concurrencia 10 en APIs pesadas. | misma página |
| Webhooks | Se crean en `/crm/v8/settings/automation/webhooks` y se ligan a workflow rules. | [Create Webhook API](https://www.zoho.com/crm/developer/docs/api/v8/create-webhook.html), [Workflow rules](https://www.zoho.com/crm/developer/docs/api/v8/config-workflow.html) |
| Import masivo | UI: CSV / XLS / XLSX / VCF. Contacts exige **Last Name**. Bulk Write API: ZIP con **un** CSV, **25 000** registros, **25 MB**, 200 columnas; **500 créditos** por job. | [Import data](https://help.zoho.com/portal/en/kb/crm/data-administration/import-data/articles/import-data), [Bulk Write](https://www.zoho.com/crm/developer/docs/api/v8/bulk-write/overview.html), [limitaciones Bulk Write](https://www.zoho.com/crm/developer/docs/api/v8/bulk-write/limitations.html) |
| SQL Server | **Zoho Databridge** conecta SQL Server on-prem, pero la documentación oficial es para **Zoho Analytics** y **Zoho Creator**, no un conector nativo “SQL Server → módulo Contacts de CRM”. Hacia CRM: CSV, Bulk Write o un job propio con la API. | [Databridge / Analytics](https://help.zoho.com/portal/en/kb/analytics/user-guide/import-connect-to-data/databases-and-datalakes/articles/sql-server-3-3-2021), [Creator SQL Server](https://help.zoho.com/portal/en/kb/creator/developer-guide/microservices/connector-references/articles/sql-server-connector) |
| WhatsApp | Integración nativa WhatsApp Business (Setup → Channels → Business Messaging). Exige Meta Business verificado, WABA y número **no usado en otro producto**. “Migration of existing phone numbers is not supported yet.” | [WhatsApp Business Integration](https://help.zoho.com/portal/en/kb/crm/connect-with-customers/business-messaging/articles/business-messaging-using-whatsapp-for-business-integration-with-zoho-crm), [página de producto](https://www.zoho.com/crm/whatsapp.html) |
| App / Firebase | API REST + SDK móviles de Zoho. **No** hay SDK oficial documentado equivalente al de HubSpot para incrustar chat en una app Firebase propia. | Docs de API; no figura un Mobile Chat SDK + FCM oficial comparable |
| Multi-sucursal | Territories (API de territorios; operaciones masivas descuentan 50 créditos/territorio). | [API limits](https://www.zoho.com/crm/developer/docs/api/v8/api-limits.html) |

**Encaje:** medio-alto. WhatsApp oficial y territories cubren sucursales. Databridge **no** resuelve solo el CRM. Encaja si Pepe ya se inclina por el ecosistema Zoho (Books, Campaigns); no es el mejor capturador de chats.

### 2.3 Kommo (ex amoCRM)

| Tema | Hecho verificado | Fuente |
| --- | --- | --- |
| API | REST HTTPS en el subdominio (`https://{cuenta}.kommo.com`), no en `www.kommo.com`. TLS 1.2. | [Limitations](https://developers.kommo.com/docs/limitations) |
| Límites | **≤ 7 req/s**. Máx. **250** entidades por GET y por POST (recomienda ≤ 50). HTTP 429 al exceder; 403 si se bloquea la IP. 50 pipelines; 100 etapas/pipeline; 100 webhooks/cuenta; 10 listas. | misma página |
| Webhooks | Planes Advanced / Pro / Enterprise para webhooks por API. Respuesta en **2 s**. Reintentos 5 min / 15 min / 15 min / 1 h. Chats API: 200 en 5 s; un solo envío (sin reintento de mensaje). | [Webhooks](https://developers.kommo.com/docs/webhooks-general), [Chats webhooks](https://developers.kommo.com/reference/receiving-chat-webhooks) |
| Import masivo | UI: XLS / XLSX / ODS / CSV y Google Sheets; un archivo puede crear leads + contactos + compañías. API: alta/actualización por lotes de contactos y compañías (hasta 250). | [Import or export data](https://support.kommo.com/docs/import-data-into-kommo), [Add contacts](https://developers.kommo.com/reference/add-contacts) |
| SQL Server | **Sin conector oficial.** Sync = job propio contra API v4 o CSV. Los “conectores SQL Server” de iPaaS son de terceros. | Docs de API / import; no hay conector SQL en developers.kommo.com |
| WhatsApp | Integración nativa **WhatsApp Business sobre Cloud API de Meta** (la API antigua se dejó de mantener el 2024-10-01). Varios números según plan: Base 1/asiento; Advanced 3 + 1 por asiento extra; Pro/Enterprise **ilimitados**. Import de contactos de la app WhatsApp Business (hasta 24 h). | [Connect WhatsApp](https://support.kommo.com/docs/connect-whatsapp-business-to-kommo), [sunsetting API antigua](https://support.kommo.com/docs/whatsapp-business-api-sunsetting), [Cloud API](https://www.kommo.com/blog/whatsapp-cloud-api/), [límites de plan](https://support.kommo.com/docs/plans-limits) |
| App / Firebase | Chats API para canales propios. **No** hay SDK Android + FCM oficial tipo HubSpot. | [Chats API](https://developers.kommo.com/reference/receiving-chat-webhooks) |
| Multi-sucursal | Hasta 50 pipelines (una por región / grupo de sucursales) y 100 fuentes por integración. Varios números WhatsApp en planes altos. | [Limitations](https://developers.kommo.com/docs/limitations), [planes](https://support.kommo.com/docs/plans-limits) |

**Encaje:** **el más alto para el problema real.** CRM conversacional: el chat *es* el alta de cliente. WhatsApp oficial nativo. Multi-pipeline para sucursales. API suficiente para el volumen de Pepe (no van a empujar tickets de piso). Debilidad: 7 req/s y sin SDK móvil/Firebase.

### 2.4 Bitrix24

| Tema | Hecho verificado | Fuente |
| --- | --- | --- |
| API | REST (cloud y on-prem). Incoming webhooks y aplicaciones OAuth. `batch` hasta 50 métodos / HTTP. Cloud: timeout **60 s** por request. | [REST limits](https://apidocs.bitrix24.com/limits.html) |
| Límites | Leaky bucket. Cloud no-Enterprise: **Y = 2 req/s**, **X = 50**. Enterprise: **5 req/s**, **X = 250**. Sostenible ≈ 172 800 HTTP/día en tarifa normal. `QUERY_LIMIT_EXCEEDED` (503). Además tope de tiempo acumulado por método (`OPERATION_TIME_LIMIT`, 429). On-prem: los topes los pone el servidor. REST en cloud solo en planes comerciales (Vibe+; Essentials no). | [limits](https://apidocs.bitrix24.com/limits.html), [error codes](https://apidocs.bitrix24.com/error-codes.html) |
| Webhooks | Cola de eventos. Si el handler no responde o falla, **Bitrix24 no reenvía**. Recomiendan cola propia. | [Performance / events](https://apidocs.bitrix24.com/settings/performance/index.html) |
| Import masivo | UI CSV (crea, no actualiza deals existentes). API `crm.item.batchImport`: **hasta 20** ítems del mismo tipo. `batch` 50 llamadas. | [Import CRM](https://helpdesk.bitrix24.com/open/25766211/), [batchImport](https://apidocs.bitrix24.com/api-reference/crm/universal/import/crm-item-batch-import.html) |
| SQL Server | Cloud: sin conector oficial a un SQL Server ajeno. On-prem el producto usa su propia base; **no** es un sync hacia `pepes_devBI`. Misma vía: job + REST. | Docs REST; no hay connector MSSQL → CRM |
| WhatsApp | Contact Center / Open Channels. Vías oficiales documentadas: **Twilio**, **Edna.io**, **Instant WhatsApp**. Marketing de producto habla de Cloud API; el help desk lista esas tres. Hay conectores de marketplace (p. ej. Gupshup). WhatsApp comercial; respuesta en ventana de 24 h (Twilio). | [Communication channels](https://helpdesk.bitrix24.com/open/25935795/), [Connect WhatsApp / Twilio](https://helpdesk.bitrix24.com/open/10222132/), [Open Channels](https://helpdesk.bitrix24.com/open/25385203/), [página WhatsApp](https://www.bitrix24.com/tools/crm/integrations/whatsapp.php) |
| App / Firebase | REST + apps de marketplace. **No** SDK Android + FCM oficial comparable al de HubSpot. | Catálogo de API |
| Multi-sucursal | Open Channels por departamento / cola; un canal de mensajería por Open Channel. Departamentos y colas (uniforme / primero disponible / todos). Encaja bien “un WhatsApp por sucursal o por zona”. | [Open Channels](https://helpdesk.bitrix24.com/open/25385203/) |

**Encaje:** medio-alto en multi-sucursal (Open Channels). WhatsApp oficial **no es un conector único y limpio**: depende de Twilio, Edna o un BSP. Plataforma pesada (intranet + CRM). Import API chico (20 registros). On-prem solo justifica si TI quiere hospedarlo; no acerca `pepes_devBI`.

### 2.5 Odoo (CRM / POS)

| Tema | Hecho verificado | Fuente |
| --- | --- | --- |
| API | **JSON-2** actual: `POST /json/2/{modelo}/{metodo}`, `Authorization: bearer {API_KEY}`. XML-RPC / JSON-RPC **deprecados**: retiro previsto Odoo 22 (otoño 2028) y Online 21.1 (invierno 2027). | [External JSON-2 API](https://www.odoo.com/documentation/19.0/developer/reference/external_api.html), [External RPC API](https://www.odoo.com/documentation/19.0/developer/reference/external_rpc_api.html) |
| Límites | Odoo no publica un rate limit tipo HubSpot. El cupo de mensajes WhatsApp es el de **Meta** (p. ej. 200 req/h por app/WABA por defecto; 5 000/h si el WABA está activo). | [WhatsApp Cloud API overview](https://developers.facebook.com/docs/whatsapp/cloud-api/overview/) |
| Webhooks | Studio: URL POST para que un sistema externo cree/actualice registros. También se puede emitir webhook hacia fuera. | [Webhooks](https://www.odoo.com/documentation/saas-18.3/applications/studio/automated_actions/webhooks.html) |
| Import masivo | UI CSV / XLSX en cualquier modelo (`res.partner`, productos, pedidos). External IDs para no duplicar. Método ORM `load()` para lotes. | [Export and import](https://www.odoo.com/documentation/19.0/applications/essentials/export_import_data.html) |
| SQL Server | Importador documenta “export/import from an SQL application” vía CSV + External ID. **No** hay réplica nativa SQL Server → Odoo. Un job `mssql` → JSON-2 es viable y alineado con este repo. | misma página de import |
| WhatsApp | App nativa **solo Enterprise** (no Community). Cloud API de Meta: callback URL + verify token + campos `messages`, plantillas, etc. | [WhatsApp Odoo 19](https://www.odoo.com/documentation/19.0/applications/productivity/whatsapp.html) |
| App / Firebase | JSON-2 / RPC. Módulos de terceros para POS+WhatsApp. **Sin** SDK Android + FCM oficial. | Docs externas |
| Multi-sucursal | Multi-company; contacto restringible a una compañía. POS por tienda. Contactos persona/empresa. | [Contacts](https://www.odoo.com/documentation/19.0/applications/essentials/contacts.html) |

**Encaje:** bajo **como CRM añadido** al stack actual. Alto **solo si** se decidiera reemplazar el POS/ERP de `pepes_devBI` (fuera de alcance y en conflicto con “no tocar pronóstico/distribución”). WhatsApp nativo = Enterprise. Esfuerzo de implantación el más alto del grupo.

### 2.6 Pipedrive

| Tema | Hecho verificado | Fuente |
| --- | --- | --- |
| API | REST v1 y v2 (`https://{dominio}.pipedrive.com/api/v2/...`). Token o OAuth. SDK oficiales **Node.js y PHP**. | [API reference](https://developers.pipedrive.com/docs/api/v1) |
| Límites | Presupuesto diario de **tokens**: `30 000 × multiplicador de plan × asientos` (Lite 1, Growth 2, Premium 5, Ultimate 7). Costes típicos: GET 1 entidad = 2 tokens; listados = 20; update = 10; search = 40. Burst / 2 s: Lite 20, Growth 40, Premium 100, Ultimate 120 (token); Search 10 / 2 s. 429 al agotar; 403 si se insiste. | [Rate limiting](https://pipedrive.readme.io/docs/core-api-concepts-rate-limiting) |
| Webhooks | v2 por defecto desde 2025-03-17. No consumen el rate limit. Reintentos 3 / 30 / 150 s. Ban 30 min a las 10 fallas; se borra a los 3 días sin éxito. | [Guide for webhooks](https://pipedrive.readme.io/docs/guide-for-webhooks), [API webhooks](https://developers.pipedrive.com/docs/api/v1/Webhooks) |
| Import masivo | UI: XLS / XLSX / CSV, **≤ 50 000 filas**, **≤ 50 MB**, una pestaña. Persons, orgs, leads, deals, products. | [Import spreadsheets](https://support.pipedrive.com/en/article/importing-data-into-pipedrive-with-spreadsheets) |
| SQL Server | **Sin conector oficial.** Job + API v2 o CSV. | Docs de import / API |
| WhatsApp | Integración nativa en **Growth+**, documentada como **beta** y no visible en todas las cuentas. Cloud API / coexistencia; importa historial (180 días texto, 14 días media) en coexistencia. Alternativa: Channel API + proveedor (tutorial oficial usa Cloud API). | [WhatsApp setup](https://support.pipedrive.com/en/article/whatsapp-integration-setup), [messaging tutorial](https://developers.pipedrive.com/tutorials/building-messaging-app-integration-with-pipedrive) |
| App / Firebase | SDK Node/PHP. **No** SDK Android + FCM oficial. | [API reference / client libraries](https://developers.pipedrive.com/docs/api/v1) |
| Multi-sucursal | Pipelines y stages; equipos. Pensado para fuerza de ventas B2B, no para 20+ mostradores. | [Import a stage/pipeline](https://support.pipedrive.com/en/article/importing-deals-into-a-specific-stage-or-pipeline) |

**Encaje:** bajo-medio. API clara (y hay SDK Node, el lenguaje de este repo), pero el producto es pipeline B2B. WhatsApp aún beta. No resuelve captura masiva de mostrador ni consignación de galletas mejor que Kommo o HubSpot.

### 2.7 Comparativa resumida

| Criterio | HubSpot | Zoho CRM | **Kommo** | Bitrix24 | Odoo | Pipedrive |
| --- | --- | --- | --- | --- | --- | --- |
| API REST documentada | Sí, excelente | Sí (créditos) | Sí (7 req/s) | Sí (leaky bucket) | JSON-2 | Sí (tokens/día) |
| Import masivo fuerte | **Sí (millones)** | Sí (25 k Bulk) | Sí (CSV + 250/API) | Débil (20/API) | Sí (CSV/XLSX) | Sí (50 k) |
| Conector oficial SQL Server → CRM | No | No (Databridge ≠ CRM) | No | No | No (CSV/API) | No |
| WhatsApp Cloud API nativo | Sí (Prof/Ent) | Sí (nativo; no migra número) | **Sí, núcleo del producto** | Vía Twilio / Edna / BSP | Sí, **solo Enterprise** | Beta Growth+ |
| SDK Android + Firebase | **Sí, oficial** | No comparable | No | No | No | No |
| Multi-sucursal | Teams | Territories | Pipelines + varios WA | **Open Channels** | Multi-company + POS | Pipelines |
| Encaje Pepe (sin padrón) | Alto (plan B) | Medio | **Alto (plan A)** | Medio | Bajo (salvo reemplazo POS) | Bajo |

Ninguno ofrece un conector oficial “lee `pepes_devBI` y llena contactos”. El patrón correcto es el que el pronóstico ya usa: **usuario de solo lectura + job + CSV/API**.

---

## 3. Arquitectura mínima recomendada

Premisas:

1. El POS / `pepes_devBI` **sigue siendo la fuente de verdad de venta y producción**.
2. El CRM **es la fuente de verdad de personas y conversaciones**.
3. No se sincronizan tickets anónimos.
4. WhatsApp Business API entra **por el conector nativo del CRM** (no se monta un BSP paralelo el día 1).
5. Credenciales solo en variables de entorno, igual que `PEPES_DB_*`.

### 3.1 Qué se sincroniza

| Objeto | Origen | Destino CRM | Frecuencia | Notas |
| --- | --- | --- | --- | --- |
| Sucursales | `dbo.Sucursales` (espejo) o CSV de extracto | Company / Account, `external_id = pepes-suc-{nSucursalPK}` | Semanal o al alta de sucursal | Incluir Planta León marcada como no-tienda |
| Productos de catálogo | `datos/catalogo/mapeo-productos.csv` + `dbo.Productos` | Product, `external_id = pepes-prod-{nProductoPK}` | Semanal | Solo `inCatalog = true` |
| KPI venta sucursal × mes | Agregado ya permitido (`ventasMes` de `pepes-fuente-base.cjs`) | Propiedades de la Company (piezas, importe) | Diario de madrugada, después de la recopia | **No** líneas de ticket |
| Tiendas galletas | Firestore `stores` | Company + Contact (encargado, teléfono) | Diario | Proyecto `cookie-1d48c`; cuenta de servicio |
| Contactos de mostrador / pastel | **WhatsApp** (y captura manual) | Contact + Lead / Deal | Tiempo real (nativo) | Aquí nace el padrón |
| Pedidos especiales | Pipeline del CRM (no `dbo.Produccion`) | Deal / Lead | Tiempo real | Si un día se liga a planta, será un proyecto aparte |
| Usuarios internos Archivo Maestro | No sincronizar | — | — | No son clientes |

Fuera de alcance mínimo: lealtad, email marketing masivo, POS Odoo, sync bidireccional CRM → SQL Server, chat embebido en la app Android.

### 3.2 Dibujo

```text
pepes_devBI (solo lectura)
        │  job nocturno (mismo patrón que pronostico-automatico)
        ▼
   CSV / API  ──►  CRM (Kommo recomendado; HubSpot alternativa)
                        ▲
Firebase stores ────────┘  Cloud Function o job diario
(app galletas)

WhatsApp Cloud API ──► conector NATIVO del CRM ──► Contactos + chats
                              │
                         sucursal / zona (pipeline o Open Channel)

App Android galletas ── Firebase ── operativa (visitas, corte)
        └── (fase 2, solo si HubSpot) Mobile Chat SDK + FCM
```

El job SQL Server **reutiliza** `configDesdeEntorno` + `asegurarSoloLectura`. No se abren tablas nuevas sin revisión de TI. Si hace falta `dbo.Clientes` en `pepes_espeBI`, primero un inventario de esquema; no se adivina.

### 3.3 Estimación de esfuerzo (días-persona)

| Bloque | Días | Qué sale |
| --- | --- | --- |
| A. Alta CRM + WABA Meta + pipelines/zonas + usuarios | 3–5 | Número oficial, plantillas, colas por zona |
| B. Import inicial CSV (sucursales fixture/reales + productos) | 1 | Los CSV de `scripts/crm/` |
| C. Job espejo → CRM (sucursales, productos, KPI mensual) | 5–8 | Script Node, secretos en env, log sin contraseña |
| D. Job Firestore `stores` → CRM | 4–6 | Requiere acceso de lectura a `cookie-1d48c` |
| E. Pruebas, mapeo de campos, handoff a sucursales | 1–2 | Guía de 1 página para encargados |
| **Mínimo (A+B+C+E, sin galletas)** | **10–16** | CRM usable por WhatsApp + catálogo |
| **Mínimo + galletas (A–E)** | **14–22** | También mayoreo de galletas |
| F. (Opcional) HubSpot Mobile Chat SDK en Android | +5–8 | Solo si se elige HubSpot y se abre `app-galleta` |
| G. (No recomendado ahora) Odoo CRM+POS+WA Enterprise | 40+ | Reemplazo de stack; fuera de este PR |

Riesgos que mueven la cifra: TI no entrega lectura a Firestore; el número de WhatsApp ya está en otro BSP (Zoho no migra; Kommo sí documenta migración); el plan comercial no incluye WhatsApp nativo.

---

## 4. Prototipo de importación (este PR)

Script: `scripts/crm/generar-csv-importacion-crm.cjs`

- Sin credenciales, sin red, sin `pepes_devBI`.
- Productos: `datos/catalogo/mapeo-productos.csv` (`inCatalog = true`).
- Sucursales: nombres que **ya están en los fixtures de prueba** del repo (no un dump real).
- Tiendas de galletas: `scripts/crm/fixtures/tiendas-galletas-sinteticas.json` (teléfonos `+52550000000x`, inventados).

```bash
node scripts/crm/generar-csv-importacion-crm.cjs
node scripts/crm/generar-csv-importacion-crm-test.cjs
```

Salida en `scripts/crm/ejemplos/`:

| Archivo | Uso en el CRM |
| --- | --- |
| `empresas-sucursales.csv` | Companies / Accounts / Organizations |
| `productos.csv` | Products |
| `contactos-tiendas-galletas.csv` | Contacts (Last Name = “Ejemplo”) |
| `manifiesto.json` | Trazabilidad de fuentes |

Mapeo rápido: HubSpot Companies/Contacts/Products; Zoho Accounts/Contacts/Products; Kommo Companies/Contacts; Bitrix24 Companies/Contacts; Odoo `res.partner` + `product.template` (External ID = `external_id`); Pipedrive Organizations/Persons/Products.

**No importar** `contactos-tiendas-galletas.csv` a producción: son personas inventadas.

---

## 5. Cómo cruzar esto con la comparativa comercial

La ficha comercial debería penalizar o bonificar, además del precio:

1. **WhatsApp Cloud API nativo** (sin BSP extra) y cuántos números entran en el plan.
2. Si el número actual de Pepe **se puede migrar** (Zoho documenta que no; Kommo sí).
3. Si WhatsApp exige Professional/Enterprise (HubSpot, Odoo).
4. Costo de asientos × sucursales vs. un equipo de planta + un equipo de galletas.
5. Que **no** se compre “módulo POS” ni “lealtad” el día 1: no hay datos que alimentarlos.

Decisión sugerida al comité:

1. **Kommo** si el caso de uso #1 es WhatsApp (pastel, mayoreo, sucursal).
2. **HubSpot** si además quieren marketing y, más adelante, chat en la app Android (único SDK + Firebase oficial de la lista).
3. Aplazar Odoo hasta que exista un proyecto de reemplazo de ERP/POS.
4. Pipedrive solo si el criterio comercial es “pipeline B2B barato”; técnicamente es el que menos cubre mostrador + WhatsApp.
5. Antes de firmar: que TI confirme si `pepes_espeBI` tiene `Clientes` / teléfonos. Si los tiene, se reabre la sección 1 y baja el peso de “captura desde cero”.

---

## 6. Límites de esta evaluación

- No se abrió `pepes_devBI` ni `pepes_espeBI` (no hay credenciales en el repo; está bien).
- No se leyó `dmeza1664-eng/app-galleta` (404).
- No se listó el esquema completo de SQL Server fuera de las 7 tablas permitidas.
- Los precios y empaquetados comerciales quedan para la otra comparativa; aquí solo se anota cuando un requisito técnico (WhatsApp nativo) está atado a un tier.
- La documentación de Bitrix24 describe **varias** vías de WhatsApp; no se afirma un único conector Cloud API de primer partido tan cerrado como el de Kommo o HubSpot.

---

## 7. URLs oficiales citadas

- HubSpot API limits: https://developers.hubspot.com/docs/developer-tooling/platform/usage-guidelines
- HubSpot webhooks: https://developers.hubspot.com/docs/api-reference/latest/webhooks/guide
- HubSpot imports: https://developers.hubspot.com/docs/api-reference/latest/crm/imports/guide
- HubSpot WhatsApp: https://knowledge.hubspot.com/inbox/connect-a-whatsapp-number-to-hubspot-using-coexistence
- HubSpot Android SDK: https://developers.hubspot.com/docs/api-reference/latest/conversations/chat-configuration/mobile-chat-sdk/android
- Zoho CRM API limits v8: https://www.zoho.com/crm/developer/docs/api/v8/api-limits.html
- Zoho Bulk Write: https://www.zoho.com/crm/developer/docs/api/v8/bulk-write/overview.html
- Zoho WhatsApp: https://help.zoho.com/portal/en/kb/crm/connect-with-customers/business-messaging/articles/business-messaging-using-whatsapp-for-business-integration-with-zoho-crm
- Zoho Databridge SQL Server (Analytics): https://help.zoho.com/portal/en/kb/analytics/user-guide/import-connect-to-data/databases-and-datalakes/articles/sql-server-3-3-2021
- Kommo limits: https://developers.kommo.com/docs/limitations
- Kommo webhooks: https://developers.kommo.com/docs/webhooks-general
- Kommo WhatsApp: https://support.kommo.com/docs/connect-whatsapp-business-to-kommo
- Bitrix24 limits: https://apidocs.bitrix24.com/limits.html
- Bitrix24 WhatsApp canales: https://helpdesk.bitrix24.com/open/25935795/
- Odoo JSON-2: https://www.odoo.com/documentation/19.0/developer/reference/external_api.html
- Odoo WhatsApp (Enterprise): https://www.odoo.com/documentation/19.0/applications/productivity/whatsapp.html
- Odoo import: https://www.odoo.com/documentation/19.0/applications/essentials/export_import_data.html
- Pipedrive rate limits: https://pipedrive.readme.io/docs/core-api-concepts-rate-limiting
- Pipedrive import: https://support.pipedrive.com/en/article/importing-data-into-pipedrive-with-spreadsheets
- Pipedrive WhatsApp: https://support.pipedrive.com/en/article/whatsapp-integration-setup
- Meta Cloud API (límites de mensajes): https://developers.facebook.com/docs/whatsapp/cloud-api/overview/
