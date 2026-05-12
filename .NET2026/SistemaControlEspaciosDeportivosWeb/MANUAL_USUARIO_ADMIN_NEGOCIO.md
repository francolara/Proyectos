# Manual de Usuario - Administrador de Complejo Deportivo

Versión: 1.0  
Fecha: 04/05/2026  
Sistema: La Zona Deportiva - Panel Administrador de Negocio

## Nota técnica de codificación
Este manual debe mantenerse en UTF-8.  
Si se regenera DOCX/PDF automáticamente, la lectura de este archivo debe hacerse en UTF-8 y evitar literales con acentos hardcodeados fuera del contenido del `.md`.

## 1. Objetivo
Este manual explica el uso operativo del panel de administración del complejo deportivo para gestionar reservas, clientes, pagos, usuarios, promociones, cupones, reportes y configuración del negocio.

## 2. Perfil de usuario
Rol objetivo: usuarios con acceso al panel del negocio (Administrador, Operador u otro rol con permisos habilitados por módulo).

## 3. Recomendaciones antes de iniciar
1. Verificar que el negocio tenga suscripción activa.
2. Definir sedes y espacios antes de registrar reservas.
3. Configurar políticas de confirmación y datos fiscales en Configuración.
4. Revisar maestros (tipos de suelo, deporte, documentos, formas de pago).

---

## 4. Estructura del panel
Menú lateral principal del Administrador de Negocio:
1. Dashboard
2. Maestros
3. Sedes
4. Espacio deportivo
5. Clientes
6. Reservas
7. Pagos
8. Comprobantes
9. Usuarios
10. Promociones
11. Cupones
12. Reportes
13. Configuración
14. Mi suscripción

Adicionalmente:
- Notificaciones de reservas web en tiempo real en el lateral.
- Tarjeta de contexto de Complejo Deportivo (nombre e identificador).

---

## 5. Dashboard

### Evidencia visual
![Dashboard](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Dashboard.png)
### 5.1 Para qué sirve
Vista resumen del estado operativo del negocio.

### 5.2 Qué muestra
- Indicadores de ocupación/reservas.
- Resumen por estados (pendiente, confirmada, pagada, cancelada, no asistió).
- Accesos rápidos por módulo.

### 5.3 Buenas prácticas
1. Revisar Dashboard al iniciar el turno.
2. Validar no-show/cancelaciones antes de cerrar caja.
3. Usar notificaciones para atender reservas web rápido.

---
## 6. Maestros

### Evidencia visual
![Maestros](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Maestros.png)
### 6.1 Para qué sirve
Mantiene catálogos base del negocio.

### 6.2 Catálogos
1. Tipos de deporte
2. Tipos de suelo
3. Tipos de documento
4. Monedas
5. Formas de pago
6. Series por documento

### 6.3 Reglas
- No eliminar catálogos en uso activo si impactan reservas/comprobantes.
- Mantener formas de pago coherentes con operación real.
- Configurar series por documento con formato y correlativo válidos antes de emitir comprobantes.
- Mantener una serie activa por tipo de documento según sede/negocio para evitar conflictos de numeración.

---
## 7. Sedes

### Evidencias visuales
![Sedes listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Sedes_listado.png)
![Sedes registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Sedes_registro.png)
### 7.1 Para qué sirve
Gestión de sedes del complejo.

### 7.2 Datos clave
- Nombre y dirección
- Teléfono y canales de contacto
- Servicios del complejo deportivo
- Configuración de notificaciones
- Horarios de atención y fechas inhabilitadas

### 7.3 Reglas
1. Una sede inactiva no debería recibir nuevas reservas operativas.
2. Mantener teléfono/correo actualizados para comunicación con clientes.

---
## 8. Espacio deportivo

### Evidencias visuales
![Espacio deportivo listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\eSPACIODEPORTIVO_LISTADO.png)
![Espacio deportivo registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\eSPACIODEPORTIVO_REGISTRO.png)
### 8.1 Para qué sirve
Gestión de canchas/espacios reservables por sede.

### 8.2 Datos clave
- Nombre del espacio
- Tipo de deporte y tipo de suelo
- Precio base
- Estado activo
- Configuración operativa del espacio

### 8.3 Reglas
1. Validar precio base correcto antes de abrir agenda.
2. Desactivar espacios fuera de servicio para evitar errores de reserva.

---
## 9. Clientes

### Evidencias visuales
![Clientes listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\clientes_listado.png)
![Clientes registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\clientes_Registro.png)
### 9.1 Para qué sirve
Padrón de clientes del negocio.

### 9.2 Funciones
- Alta/edición de cliente
- Datos de contacto
- Tipo y número de documento

### 9.3 Reglas
- Evitar duplicados de cliente con mismo documento.
- Completar teléfono/correo para recordatorios y seguimiento.

---
## 10. Reservas

### Evidencias visuales
![Reservas listado 1](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Reservas_litado1.png)
![Reservas listado 2](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Reserva_Listado2.png)
![Reserva registro 1](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Reserva_Registro.png)
![Reserva registro 2](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Reserva_Registro2.png)
### 10.1 Para qué sirve
Gestión completa de agenda y registro de reservas.

### 10.2 Funciones
1. Calendario por sede y espacio.
2. Registro de reserva desde modal.
3. Edición de reserva existente.
4. Cambio de estado.
5. Registro de pagos/adelantos.
6. Validación de cupones en nueva reserva.

### 10.3 Reglas funcionales importantes
1. El cupón solo se aplica en la creación de reserva (no en reservas ya existentes).
2. Al validar cupón, el total debe reflejar el descuento aplicado.
3. El precio guardado no debe volver a descontarse en la grabación.
4. El contador de uso del cupón incrementa al uso exitoso.
5. Si se alcanza el máximo de usos, el cupón deja de ser válido.
6. Si el pago acumulado cubre el total final, el estado debe pasar a Pagada.
7. Si no cubre el total final, el estado se mantiene según reglas de confirmación.

### 10.4 Estados de reserva (referencia operativa)
- Pendiente
- Confirmada
- Pagada
- Cancelada
- No asistió
- Bloqueada / no atención

### 10.5 Buenas prácticas
1. Registrar pagos el mismo día para mantener saldos correctos.
2. Verificar horario disponible antes de confirmar.
3. Auditar cambios de estado al cierre del turno.

---
## 11. Pagos

### Evidencias visuales
![Pagos listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Pagos_Listado.png)
![Pago registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Pago_Registro.png)
### 11.1 Para qué sirve
Control de movimientos de pago por reserva.

### 11.2 Funciones
- Alta de pago
- Edición/anulación según permisos
- Conciliación de adelantos y saldos

### 11.3 Reglas
1. El estado de la reserva depende del total pagado vs total de reserva.
2. Registrar forma de pago y referencia operativa cuando aplique.

---
## 12. Comprobantes

### Evidencias visuales
![Comprobantes listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\comprobante_Listado.png)
![Comprobante registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Comprobnate_Registro.png)
### 12.1 Para qué sirve
Gestión de comprobantes asociados a operaciones.

### 12.2 Funciones
- Emisión y consulta de comprobantes
- Vinculación con cliente y reserva

### 12.3 Reglas
- Verificar datos fiscales en Configuración antes de emitir.

---
## 13. Usuarios

### Evidencia visual
![Usuario registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\usuario_registro.png)
### 13.1 Para qué sirve
Gestión de accesos al panel de negocio.

### 13.2 Funciones
- Alta de usuario
- Asignación de rol
- Configuración de permisos por módulo

### 13.3 Reglas
1. Aplicar principio de mínimo privilegio.
2. Revisar permisos de cajas/pagos/comprobantes con especial cuidado.

---
## 14. Promociones

### Evidencias visuales
![Promociones listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Promocion_Listado.png)
![Promociones registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Promocion_Registro.png)
### 14.1 Para qué sirve
Descuentos/promociones por rango de fechas y reglas de horario.

### 14.2 Funciones
- Crear, editar, activar/inactivar promociones
- Visualizar impacto en cotización

### 14.3 Reglas
1. Validar vigencia (fecha inicio/fin) y alcance (sede/espacio).
2. Evitar solapes no deseados de descuentos.

---
## 15. Cupones

### Evidencias visuales
![Cupones listado](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\cupones_listado.png)
![Cupones registro](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Cupones_Registro.png)
### 15.1 Para qué sirve
Gestión de cupones con control de uso.

### 15.2 Funciones
1. Crear cupón con código y nombre.
2. Definir vigencia (inicio/fin).
3. Definir cantidad máxima de usos.
4. Restringir por sede y/o espacio deportivo.
5. Consultar listado paginado (20 en 20).
6. Ver KPIs del módulo.

### 15.3 Reglas
1. Cupón válido solo dentro de vigencia.
2. Cupón inactivo no aplica.
3. Cupón sin stock de usos no aplica.
4. El uso se acumula por cada aplicación exitosa.
5. Debe verse en cotización y respetarse en grabación final.

---
## 16. Reportes

### Evidencia visual
![Reportes](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\Reportes.png)
### 16.1 Para qué sirve
Análisis operativo/comercial del negocio.

### 16.2 Uso recomendado
1. Corte diario de ingresos y reservas.
2. Revisión semanal de ocupación por espacio.
3. Seguimiento mensual de clientes y conversión.

---
## 17. Configuración

### Evidencia visual
![Configuración](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\configuracion.png)
### 17.1 Para qué sirve
Parámetros estructurales del complejo.

### 17.2 Secciones clave
1. Datos del complejo deportivo
2. Logo del complejo deportivo
3. Política de confirmación de reserva
4. Parámetros de pago y emisión

### 17.3 Reglas
1. Cualquier cambio impacta módulos operativos.
2. Validar con prueba corta luego de guardar.

---

## 18. Mi suscripción

### Evidencia visual
![Mi suscripción](C:\Users\Franco Lara\Desktop\GitHub Proyectos\Proyectos\Pantallasos\mususcripcion.png)
### 18.1 Para qué sirve
Consulta del estado comercial del negocio.

### 18.2 Estados comunes
- Prueba
- Contrato activo
- Vencido / suspendido

### 18.3 Recomendación
- Monitorear vencimiento para evitar bloqueo de módulos.

---

## 19. Reglas de reserva pública (impacto operativo)
1. Se permite aplicar cupón en reserva pública.
2. Reserva pública actualmente restringida a bloques de 1 hora.
3. Hora fin se ajusta automáticamente en +1 hora respecto de hora inicio.
4. Si no cumple 60 minutos exactos, la grabación se rechaza por validaciones de interfaz, backend y SP.
5. El calendario público diferencia visualmente reservas en estado "Reservada" y "Confirmada" con estilos de color más claros para lectura rápida.

---

## 20. Errores frecuentes y solución
### 20.1 "No se pudo obtener el detalle"
Causa probable: error en endpoint o consulta de soporte.  
Acción: revisar logs y verificar datos base (correo/teléfono de contacto).

### 20.2 Cupón válido en pantalla pero no coincide en total final
Causa probable: descuento aplicado dos veces o recálculo inconsistente.  
Acción: revisar flujo de validación y total final guardado.

### 20.3 Pago total no cambia a Pagada
Causa probable: validación de estado contra total desactualizado.  
Acción: verificar monto final de reserva y acumulado de pagos.

### 20.4 Correo de recordatorio no sale
Causa probable: remitente no permitido en proveedor de correo.  
Acción: validar remitente configurado/autorizado en Brevo.

---

## 21. Checklist operativo diario
1. Revisar notificaciones web.
2. Confirmar reservas del día.
3. Registrar pagos y validar estados.
4. Verificar reservas con cupón.
5. Cerrar pendientes/cancelaciones.
6. Revisar reporte rápido de caja/ingresos.

---

## 22. Anexo A - Evidencias sugeridas (capturas)
Agregar capturas de estas pantallas para versión ilustrada del manual:
1. Dashboard principal.
2. Alta de sede.
3. Alta de espacio deportivo.
4. Nueva reserva (modal admin).
5. Validación de cupón en reserva.
6. Módulo de cupones (KPI + listado + formulario).
7. Registro de pago.
8. Emisión de comprobante.
9. Configuración del complejo.
10. Mi suscripción.

---

## 23. Ajustes recientes de experiencia visual (portal público)
1. Barra superior pública y acciones de sesión con estilo más claro y consistente (incluye botón de cierre con variante visual de alerta).
2. Buscador principal de espacios con mayor jerarquía visual y botón de búsqueda con icono.
3. Sección de beneficios y tarjetas de espacios mejoradas para lectura móvil/escritorio (contraste, profundidad, jerarquía tipográfica y estados hover).
4. Tarjetas de disponibilidad en reserva pública con mejor distinción visual entre estado Reservada y Confirmada.
5. En registro, el texto de relación fue ajustado a: "Relación con el complejo deportivo".

---

## 24. Control de cambios del manual
- 04/05/2026: Versión inicial operativa para Administrador de Negocio.
- 11/05/2026: Actualización por mejoras funcionales y visuales en portal público, refinamiento de etiquetas de negocio/complejo y ajustes de experiencia en reserva pública.
- 12/05/2026: Las pantallas quedan integradas dentro de cada sección del manual (se retira el anexo final consolidado).
