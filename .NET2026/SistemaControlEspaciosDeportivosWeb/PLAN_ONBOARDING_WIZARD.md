# Plan de Trabajo - Onboarding Wizard (Administrador de Complejo)

## Fase 1 - Diseno funcional cerrado

### Objetivo
Implementar un wizard de primer ingreso para administradores de complejo deportivo, con avance persistente y validacion final obligatoria antes de habilitar la operacion normal.

### Flujo general del wizard
1. Paso 1: Configuracion
2. Paso 2: Maestros
3. Paso 3: Sedes
4. Paso 4: Espacios deportivos
5. Paso 5: Resumen y finalizar

### Regla de primer ingreso
- Si onboarding esta incompleto: redirigir automaticamente al wizard.
- Si onboarding esta completo: acceso normal al panel.

### Regla de reingreso
- Si queda incompleto, en el siguiente login retoma el paso pendiente.
- Mostrar progreso en formato `X/5`.

### Regla de bloqueo operativo mientras este incompleto
- Permitir acceso a: `Configuracion`, `Maestros`, `Sedes`, `Espacios` y `Onboarding`.
- Bloquear acceso a modulos operativos (reservas, pagos, comprobantes, reportes) hasta finalizar.

## Datos obligatorios por paso

### Paso 1 - Configuracion (minimo)
1. `NombreComercial`
2. `TipoDocumento`
3. `MonedaId`

Condicional para comprobante electronico:
1. `TipoDocumento = 6` (RUC)
2. `PorcentajeIgv > 0`
3. `DireccionFiscal` no vacia
4. `CodigoUbigeo` valido (6 digitos)

### Paso 2 - Maestros (minimo)
1. Al menos 1 `TipoDeporte` activo
2. Al menos 1 `TipoSuelo` activo
3. Al menos 1 `FormaPago` activa
4. Al menos 1 `Moneda` activa
5. Al menos 1 `TipoDocumentoComprobante` activo
6. Al menos 1 `SerieDocumentoComprobante` activa

### Paso 3 - Sedes (minimo)
Al menos 1 sede activa con:
1. `Nombre`
2. `Direccion`
3. `CodigoUbigeo` valido
4. `HoraApertura`
5. `HoraCierre`
6. Al menos 1 servicio seleccionado en `ServiciosSeleccionados`

Recomendados (no bloqueantes):
1. `Telefono` o `WhatsappContacto`
2. Dias de atencion configurados

### Paso 4 - Espacios (minimo)
Al menos 1 espacio activo con:
1. `SedeId`
2. `TipoDeporteId`
3. `TipoSueloId`
4. `Codigo`
5. `Nombre`
6. Tarifa valida: al menos un rango con `Precio > 0`, `HoraFin > HoraInicio` y sin cruces por dia

### Paso 5 - Resumen
1. Mostrar checklist con estado por requisito
2. Solo permitir finalizar si todo esta completo

## UX de ayuda contextual

### Regla
- Cada campo del wizard tendra microtexto: "Para que sirve este dato".
- Longitud maxima recomendada: 1 linea (120 caracteres aprox).

### Ejemplo
- `PoliticaConfirmacionPago`: "Define si la reserva se confirma sin pago, con adelanto o con pago total."

## Persistencia de avance (para fase SQL/SP)

### Tabla de estado propuesta
`NegocioOnboardingEstado`

Campos base:
1. `NegocioId` (PK)
2. `PasoActual`
3. `Completado` (bit)
4. `FechaUltimaActualizacionUtc`
5. `UsuarioUltimaActualizacion`
6. `FechaCompletadoUtc` (nullable)
7. `UsuarioCompletado` (nullable)

### SP requeridos
1. `sp_OnboardingEstado_Obtener`
2. `sp_OnboardingEstado_GuardarAvance`
3. `sp_OnboardingEstado_MarcarCompletado`
4. `sp_OnboardingChecklist_Validar`

## Criterios de aceptacion
1. Usuario nuevo admin entra y cae al wizard automaticamente.
2. Si sale sin completar, al volver retoma avance.
3. No puede finalizar sin cumplir todos los minimos definidos.
4. Al completar, no vuelve a mostrarse el wizard en ingresos siguientes.
5. Todos los campos del wizard incluyen microresumen de utilidad.
