# Sistema Valoración Actividad Productiva (VAP)
**ADR – Architecture Decision Record**
**Versión:** 1.3
**Fecha última actualización:** 24 Marzo 2026


## 1.Antecedentes

Hasta este momento el sistema consta de varias fuentes de datos provenientes de diferentes personas o departamentos, con diferentes formas de transmisión hacia la persona que los centraliza: formularios, emails, Google Sheets diversos. 

Todos estos datos deben ser "volcados" en un Google Sheet final para calcular los importes de los salarios mensuales de parte del personal.

Todo este proceso es complejo, actualmente tedioso de gestionar y requiere una inversión de tiempo elevada, la cual se puede reducir haciendo los procesos más eficaces y automatizados.


### Notas Importantes

Este sistema implica principalmente el cálculo de salarios mensuales de dos grupos de personal: **secretarias** y **profesores de autoescuelas**.

Para poder avanzar de forma ordenada, el proyecto se estructura en fases, mejorando primero el sistema para un grupo y después para el otro, dado que tanto los datos como sus fuentes no siempre coinciden entre ambos grupos.

---

## 2.Objetivos

1º Mejorar los procesos para actualizar los listados de personal a los cuales se les debe calcular los salarios mensualmente. 
2º Automatizar y mejorar las cargas de datos de las diferentes fuentes al proceso para el calculo final. 

---

## 3. Decisiones Técnicas

| # | Decisión | Justificación | Alternativas descartadas |
|---|----------|---------------|--------------------------|
| DT-01 | Usar Google Sheets como BBDD central de personal | Ya es la herramienta corporativa adoptada; no requiere infraestructura nueva | Airtable, Notion DB, base de datos SQL |
| DT-02 | Crear una Webapp propia para reporte de complementos | Los formularios nativos de Google clasificaban envíos masivos como SPAM y obligaban a resolver reCaptcha constantemente | Google Forms estándar |
| DT-03 | Asignar cada responsable a un complemento concreto en BBDD | Permite cambiar responsables o complementos sin tocar el código de la webapp | Hardcodear responsables en el script |
| DT-04 | Vincular el formulario de Horas Extras directamente a `VAP_Secretarias` | Evita tener que actualizar el formulario manualmente cada vez que hay altas o bajas de personal | Lista manual en el formulario |
| DT-05 | Gestionar envíos por mes + año en el timestamp | Hace el sistema duradero ante el paso de los años sin necesidad de reestructurar las hojas | Solo mes sin año |
| DT-06 | Separar la BBDD de personal (`VAP_BBDD_Personal`) del resto de hojas operativas | Centraliza las altas/bajas en un único punto, reduciendo errores de sincronización | Mantener listas dispersas por cada hoja operativa |


---

## 4. Estructura y Recursos

### 4.1 Estructura General
```
VAP
├── Secretarias
│   ├── Complementos Salariales
│   │   ├── Webapp Reporte Complementos (VAP_index_Reportar_Complementos_Secretarias)
│   │   ├── Webapp Consulta Envíos (VAP_index_Consulta_Complementos_Secretarias)
│   │   └── Google Sheet Carga Datos (VAP_Carga_Datos_Mensuales)
│   └── Horas Extras
│       ├── Formulario Horas Extra Secretarias (EXTRA SECRETARIAS 2026)
│       └── Script Volcado → VINCULACIÓN GASTOS PERSONAL
└── Profesores
    └── [Pendiente – Ver Fases]
```

### 4.2 Recursos

**Carpeta raíz del proyecto:** `D. Data Sistema VAP` (Google Drive)

| Recurso | Nombre | Descripción |
|---------|--------|-------------|
| BBDD Personal | `VAP_BBDD_Personas` | Google Sheet que centraliza todo el personal del sistema |
| Datos Mensual Complementos Salariales Secretarias | `VAP_Carga_Datos_Mensuales_Complementos_Secretarias` | Recibe los volcados de la webapp de complementos salariales secretarias |
| Webapp Reporte | `VAP_index_Reportar_Complementos_Secretarias` | Permite reporte masivo de complementos por responsable |
| Webapp Consulta | `VAP_index_Consulta_Complementos_Secretarias` | Permite a responsables consultar sus envíos realizados |
| Formulario HH.EE. Secretarias | `EXTRA SECRETARIAS 2026` | Formulario de horas extras vinculado a BBDD |


### 4.3 Scripts

#### Horas Extras Secretarias

| Script | Ubicación GAS | Descripción |
|--------|--------------|-------------|
| `Scripts Formulario Horas Extras Secretarias` | Spreadsheet `VAP_BBDD_Personas` | Sincroniza la pregunta lista "SECRETARIA" del formulario `EXTRA SECRETARIAS 2026` con las secretarias activas de `VAP_Secretarias`. Al recibir un envío (`onFormSubmit`), busca el `ID_Secre` por nombre y lo escribe automáticamente en la hoja de respuestas. Incluye control de frecuencia (cache 30 s) y trigger de respaldo horario. |
| `Script Volcado Horas Extras Secretarias` | Hoja de respuestas de `EXTRA SECRETARIAS 2026` | Vuelca automáticamente (trigger `onFormSubmit`) los datos de horas extras al spreadsheet `VINCULACIÓN GASTOS PERSONAL`. Toma el **último** registro por empleado en caso de duplicados. Usa `LockService` para evitar escrituras concurrentes. |

**Columnas volcadas por `Script Volcado Horas Extras Secretarias`:**

| Concepto | Col. origen (`EXTRA SECRETARIAS 2026`) | Col. destino (`VINCULACIÓN GASTOS PERSONAL`) |
|----------|----------------------------------------|----------------------------------------------|
| Nº HORAS EXTRA | G | AE |
| € OTROS EXTRAS | I | AH |
| Nombre empleado | E | C (clave de cruce) |

**Triggers configurados en `Scripts Formulario Horas Extras Secretarias`:**

| Función | Tipo de trigger | Evento |
|---------|----------------|--------|
| `onFormSubmit` | `forForm` | Al enviar el formulario |
| `onSpreadsheetEditTrigger` | `forSpreadsheet` | `onEdit` en `VAP_BBDD_Personas` |
| `onSpreadsheetChangeTrigger` | `forSpreadsheet` | `onChange` en `VAP_BBDD_Personas` (altas/bajas de filas) |
| `actualizarPreguntaSecretarias` | `timeBased` | Cada hora (respaldo) |

---

#### Complementos Salariales Secretarias

| Script | Ubicación GAS | Descripción |
|--------|--------------|-------------|
| `VAP_Export_Complementos_Mensuales_Secretarias` | `VAP_Carga_Datos_Mensuales_Complementos_Secretarias` | Script principal (`Code.gs`) que concentra la lógica de servidor de ambas webapps (reporte y consulta), la exportación mensual a XLSX y el envío de emails de confirmación/aviso. |

**Funciones expuestas por `VAP_Export_Complementos_Mensuales_Secretarias`:**

| Función GAS | Webapp / Acción | Descripción |
|-------------|----------------|-------------|
| `vap_bootstrapForEmail(email)` | Webapp Reporte | Valida el responsable contra `VAP_Responsables` e inicializa la sesión: devuelve sus complementos asignados, lista de secretarias activas y valores por defecto de mes/año |
| `vap_submitBatch(payload)` | Webapp Reporte | Guarda el envío o corrección de importes en la hoja `Data`; genera `Batch_ID` único, escribe en `Logs` y envía email de confirmación al responsable y aviso al gestor |
| `vap_listCorrectionTargets(params)` | Webapp Reporte | Devuelve los `Batch_ID` previos disponibles para corregir, filtrados por responsable, mes, año y concepto |
| `vap_consultaBootstrap(email)` | Webapp Consulta | Valida el responsable e inicializa la sesión de consulta: devuelve los períodos (año/mes) con envíos existentes |
| `vap_consultaBatches(params)` | Webapp Consulta | Devuelve los últimos 5 envíos del responsable para el período indicado, con detalle por secretaria |
| `vap_generarExcelMensual()` | Menú Google Sheet | Genera el Excel mensual fusionado (ID + Nombre de secretaria + columna por concepto) y lo guarda en la carpeta de Drive configurada |
| `onOpen()` | Google Sheet | Añade el menú "Sistema VAP → Generar Excel mensual" al abrir el spreadsheet |

**Orden de conceptos en el Excel exportado:** `APTOS`, `FINANCIACIONES`, `RESEÑAS`, `SABADOS_50`, `SABADOS_60`, `HORAS_PUNTOS`


### 4.4 Hojas dentro de `VAP_BBDD_Personas`

| Hoja | Propósito |
|------|-----------|
| `VAP_Secretarias` | Listado de secretarias con estado activo/inactivo |
| `VAP_Complementos_Secretarias` | Catálogo de complementos salariales disponibles |
| `VAP_Responsables` | Listado de responsables (managers) del sistema |
| `VAP_Responsables_Complementos` | Relación entre responsables y complementos asignados |
| `VAP_Profesores` | Listado de profesores (uso futuro) |
| `Config` | Listado de valores necesarios para script varios |

---

## 5. Fases del Proyecto

### FASE 1 – Secretarias

#### Fase 1.A – Complementos Salariales

- [x] Creación de `VAP_BBDD_Personas` como fuente única de verdad para el personal activo
- [x] Webapp de reporte masivo de complementos con lista dinámica de secretarias activas
- [x] Validación de que el responsable corresponde al complemento que reporta
- [x] Gestión de envíos y correcciones por mes/año
- [x] Email de confirmación al responsable y aviso al gestor al recibir nueva remesa
- [x] Webapp de consulta de envíos realizados por los responsables
- [x] Exportación mensual a demanda de datos fusionados por secretaria y complemento

#### Fase 1.B – Horas Extras de Secretarias

- [x] Vinculación del formulario `EXTRA SECRETARIAS 2026` a `VAP_Secretarias` (lista dinámica, sin mantenimiento manual)
- [x] Actualización del script de volcado para aceptar las nuevas columnas del formulario
- [x] Volcado de datos hacia `VINCULACIÓN GASTOS PERSONAL`

### FASE 2 – Profesores

- [ ] Revisar y completar la hoja `VAP_Profesores` en `VAP_BBDD_Personas`
- [ ] Identificar fuentes de datos actuales y su estructura
- [ ] Definir complementos salariales específicos del colectivo
- [ ] Evaluar reutilización de la webapp de secretarias o necesidad de versión propia
- [ ] Diseñar el flujo de volcado hacia el cálculo final de nóminas

---

## 6. Registro de Trabajo

### Sesión 1 - 25/02/26
- Análisis del sistema existente y definición de objetivos
- Creación de `VAP_BBDD_Personas` y sus hojas internas
- Diseño inicial de la webapp de reporte de complementos
- Desarrollo y despliegue de `VAP_index_Reportar_Complementos_Secretarias`
- Implementación de la lógica de validación responsable ↔ complemento
- Implementación de gestión de envíos y correcciones por mes/año
- Desarrollo y despliegue script `VAP_Export_Complementos_Mensuales_Secretarias`


### Sesión 2 - 24/03/26
- Implementación de emails de confirmación y aviso al gestor
- Desarrollo y despliegue de `VAP_index_Consulta_Complementos_Secretarias`
- Vinculación del formulario `EXTRA SECRETARIAS 2026` con `VAP_Secretarias`
- Actualización del script de volcado de horas extras con las nuevas columnas con script `Script Volcado Horas Extras Secretarias`
- Actualización de validacion de respeustas en formulario `EXTRA SECRETARIAS 2026` con script `Scripts Formulario Horas Extras Secretarias`

---

## 7. Riesgos y Limitaciones Conocidas

| ID | Riesgo / Limitación | Impacto | Medida actual o propuesta |
|----|---------------------|---------|---------------------------|
| R-01 | Dependencia total de Google Workspace | Alto | Asumida como decisión corporativa. Documentar exportaciones periódicas |
| R-02 | Permisos de acceso a la webapp no gestionados por rol de Google | Medio | La webapp valida internamente que el responsable corresponde al complemento |
| R-03 | Si una secretaria cambia de nombre en BBDD, los registros históricos quedan desvinculados | Medio | Usar siempre `ID_Secre` como clave primaria en lugar del nombre |
| R-04 | ~~El script de volcado de horas extras es manual (no automático)~~ **RESUELTO** | ~~Medio~~ | El script `Script Volcado Horas Extras Secretarias` tiene trigger `onFormSubmit` activo: el volcado se ejecuta automáticamente al recibir cada envío del formulario |

---

## 8. Glosario

| Término | Definición |
|---------|------------|
| VAP | Sistema de Valoración de Actividad Productiva |
| BBDD | Base de Datos |
| Complemento salarial | Concepto variable que se suma al salario base mensual |
| Responsable / Manager | Persona autorizada a reportar un complemento concreto |
| `ID_Secre` | Identificador único de cada secretaria en la BBDD |
| Webapp | Aplicación web generada con Google Apps Script |
| HH.EE. | Horas Extras |