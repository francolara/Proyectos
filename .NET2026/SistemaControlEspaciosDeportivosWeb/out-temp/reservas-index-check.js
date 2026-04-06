(() => {
            const negocioId = 1;
            const calendarioEl = document.getElementById('calendarioReservas');
            const calendarioActivo = !!calendarioEl;
            const formFiltros = document.getElementById('filtrosForm');
            const disponibilidadMensajeEl = document.getElementById('disponibilidadMensaje');
            const disponibilidadSlotsEl = document.getElementById('disponibilidadSlots');
            const disponibilidadFechaFiltroEl = document.getElementById('disponibilidadFechaFiltro');
            const pendientesMensajeEl = document.getElementById('pendientesMensaje');
            const pendientesSlotsEl = document.getElementById('pendientesSlots');
            const pendientesFechaFiltroEl = document.getElementById('pendientesFechaFiltro');
            const reservaModalEl = document.getElementById('reservaModal');
            const reservaModal = new bootstrap.Modal(reservaModalEl);
            const reservaModalForm = document.getElementById('reservaModalForm');
            const reservaModalTitulo = document.getElementById('reservaModalTitulo');
            const reservaModalError = document.getElementById('reservaModalError');
            const reservaModalDisponibilidad = document.getElementById('reservaModalDisponibilidad');
            const btnGuardarReservaModal = document.getElementById('btnGuardarReservaModal');
            const reservaModalAccionConflicto = document.getElementById('reservaModalAccionConflicto');
            const btnVerReservaConflicto = document.getElementById('btnVerReservaConflicto');
            const getReservaModalUrl = '@Url.Action("ObtenerReservaModal", "Reservas")';
            const guardarReservaModalUrl = '@Url.Action("GuardarReservaModal", "Reservas")';
            const validarDisponibilidadModalUrl = '@Url.Action("ValidarDisponibilidadModal", "Reservas")';
            const cambiarEstadoRapidoUrl = '@Url.Action("CambiarEstadoRapido", "Reservas")';
            const comboEspaciosUrl = '@Url.Action("ObtenerEspaciosFiltro", "Reservas")';
            const resumenDiaOperativoUrl = '@Url.Action("ResumenDiaOperativo", "Reservas")';
            const historialReservaUrl = '@Url.Action("HistorialReserva", "Reservas")';
            const sedeSelect = document.getElementById('sedeId');
            const espacioSelect = document.getElementById('espacioDeportivoId');
            const estadoSelect = document.getElementById('estado');
            const legendEstadoButtons = Array.from(document.querySelectorAll('[data-legend-estado]'));
            const btnEditarReservaListadoEls = Array.from(document.querySelectorAll('[data-accion-editar-reserva]'));
            const kpiOcupacionDiaEl = document.getElementById('kpiOcupacionDia');
            const kpiSlotsOcupadosDiaEl = document.getElementById('kpiSlotsOcupadosDia');
            const kpiSlotsLibresDiaEl = document.getElementById('kpiSlotsLibresDia');
            const kpiReservasActivasDiaEl = document.getElementById('kpiReservasActivasDia');
            const kpiPendientesDiaEl = document.getElementById('kpiPendientesDia');
            const hoyDate = new Date();
            const hoyIso = `${hoyDate.getFullYear()}-${String(hoyDate.getMonth() + 1).padStart(2, '0')}-${String(hoyDate.getDate()).padStart(2, '0')}`;
            const puedeCrearReserva = true;
            const puedeEditarReserva = true;
            const estadosOcultosCalendario = new Set();
            let timerValidacion = null;
            let conflictoReservaIdActual = null;
            let calendar = null;
            let recargandoEspacios = false;
            let disponibilidadRenderSeq = 0;
            let pendientesRenderSeq = 0;
            const calendarAxisWidth = '122px';
            function pad2(n) {
                return String(n).padStart(2, '0');
            }

            function formatDateIso(date) {
                return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
            }

            function esFechaPasadaIso(fechaIso) {
                if (!fechaIso) return false;
                return fechaIso < hoyIso;
            }

            function formatTime(h, m) {
                return `${pad2(h)}:${pad2(m)}`;
            }

            function hexToRgba(hexColor, alpha) {
                const hex = String(hexColor || '').trim().replace('#', '');
                if (hex.length !== 6) return `rgba(100,116,139,${alpha})`;
                const r = parseInt(hex.slice(0, 2), 16);
                const g = parseInt(hex.slice(2, 4), 16);
                const b = parseInt(hex.slice(4, 6), 16);
                if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return `rgba(100,116,139,${alpha})`;
                return `rgba(${r},${g},${b},${alpha})`;
            }

            function toLocalDateTimeString(date) {
                return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}:${pad2(date.getSeconds())}`;
            }

            function formatearRangoCabecera(fechaInicio, fechaFinExclusiva) {
                if (!fechaInicio || !fechaFinExclusiva) return '';
                const fin = new Date(fechaFinExclusiva.getTime() - 1);
                const diaInicio = fechaInicio.getDate();
                const diaFin = fin.getDate();
                const mesInicio = capitalizar(fechaInicio.toLocaleDateString('es-PE', { month: 'long' }));
                const mesFin = capitalizar(fin.toLocaleDateString('es-PE', { month: 'long' }));
                const anio = fin.getFullYear();
                return `${diaInicio} de ${mesInicio} al ${diaFin} de ${mesFin} del ${anio}`;
            }

            function renderToolbarSubtitle(fechaInicio, fechaFinExclusiva) {
                if (!calendarioEl) return;
                const toolbarTitle = calendarioEl.querySelector('.fc-toolbar-title');
                if (!toolbarTitle) return;

                const chunk = toolbarTitle.closest('.fc-toolbar-chunk');
                if (chunk) {
                    const subtitle = chunk.querySelector('.sc-fc-toolbar-subtitle');
                    if (subtitle) subtitle.remove();
                }

                const titulo = formatearRangoCabecera(fechaInicio, fechaFinExclusiva);
                if (titulo) {
                    toolbarTitle.textContent = titulo;
                }
            }

            function capitalizar(texto) {
                if (!texto) return '';
                return texto.charAt(0).toUpperCase() + texto.slice(1);
            }

            function pintarCabeceraHorarios() {
                if (!calendarioEl) return;
                const axisHeader = calendarioEl.querySelector('.fc-col-header .fc-timegrid-axis');
                if (!axisHeader) return;
                axisHeader.classList.add('sc-time-axis-header');
                const cushion = axisHeader.querySelector('.fc-col-header-cell-cushion');
                if (cushion) {
                    cushion.innerHTML = '<span class="sc-axis-header-label">Horarios</span>';
                    return;
                }
                axisHeader.innerHTML = '<span class="sc-axis-header-label">Horarios</span>';
            }

            function aplicarEstiloCalendarioFase1() {
                if (!calendarioEl) return;
                calendarioEl.classList.add('sc-calendar-ready');
            }

            function aplicarFiltroLeyendaCalendario() {
                if (!calendarioEl) return;
                ['pendiente', 'confirmada', 'finalizada', 'cancelada', 'noshow', 'bloqueada']
                    .forEach((estado) => {
                        calendarioEl.classList.toggle(`sc-hide-${estado}`, estadosOcultosCalendario.has(estado));
                    });
            }

            function sincronizarLeyendaCalendario() {
                legendEstadoButtons.forEach((btn) => {
                    const estado = String(btn.dataset.legendEstado || '').trim().toLowerCase();
                    const oculto = estadosOcultosCalendario.has(estado);
                    btn.classList.toggle('is-off', oculto);
                    btn.setAttribute('aria-pressed', oculto ? 'false' : 'true');
                });
            }

            function estadoLeyendaDesdeEvento(evento) {
                const tipoEvento = String(evento?.extendedProps?.tipoEvento || '').trim().toUpperCase();
                if (tipoEvento === 'BLOQUEO' || tipoEvento === 'NO_ATENCION') return 'bloqueada';
                const codigoRaw = String(evento?.extendedProps?.estadoCodigo ?? '').trim().toUpperCase();
                switch (codigoRaw) {
                    case '1':
                    case 'PENDIENTE': return 'pendiente';
                    case '2':
                    case 'CONFIRMADA': return 'confirmada';
                    case '3':
                    case 'EN_USO': return 'finalizada';
                    case '4':
                    case 'PAGADA':
                    case 'FINALIZADA': return 'finalizada';
                    case '5':
                    case 'CANCELADA': return 'cancelada';
                    case '6':
                    case 'NO_ASISTIO':
                    case 'NO_SHOW': return 'noshow';
                    default: return '';
                }
            }

            function actualizarContadoresLeyendaCalendario() {
                if (!calendar || legendEstadoButtons.length === 0) return;
                const contadores = {
                    pendiente: 0,
                    confirmada: 0,
                    finalizada: 0,
                    cancelada: 0,
                    noshow: 0,
                    bloqueada: 0
                };
                const rangoInicio = calendar.view?.activeStart;
                const rangoFin = calendar.view?.activeEnd;
                const eventos = calendar.getEvents();

                eventos.forEach((evento) => {
                    if (!evento || !evento.start) return;
                    if (rangoInicio && evento.start < rangoInicio) return;
                    if (rangoFin && evento.start >= rangoFin) return;

                    const idEvento = String(evento.id || '');
                    const tipoEvento = String(evento.extendedProps?.tipoEvento || '').trim().toUpperCase();
                    if (idEvento.startsWith('BG-')) return;
                    if (evento.display === 'background' && tipoEvento !== 'BLOQUEO' && tipoEvento !== 'NO_ATENCION') return;

                    const estadoLeyenda = estadoLeyendaDesdeEvento(evento);
                    if (!estadoLeyenda || !Object.prototype.hasOwnProperty.call(contadores, estadoLeyenda)) return;
                    contadores[estadoLeyenda] += 1;
                });

                legendEstadoButtons.forEach((btn) => {
                    const estado = String(btn.dataset.legendEstado || '').trim().toLowerCase();
                    if (estado === 'bloqueada') return;
                    const total = contadores[estado] ?? 0;
                    const countEl = btn.querySelector('[data-legend-count]');
                    if (countEl) {
                        countEl.textContent = `(${total})`;
                    }
                });
            }

            function activarHoverColumnasCalendario() {
                if (!calendarioEl) return;
                const cols = calendarioEl.querySelectorAll('.fc-timegrid-col');
                cols.forEach((col) => {
                    if (col.dataset.scHoverBound === '1') return;
                    col.dataset.scHoverBound = '1';
                    col.addEventListener('mouseenter', () => col.classList.add('sc-col-hover'));
                    col.addEventListener('mouseleave', () => col.classList.remove('sc-col-hover'));
                });
            }

            function classByEstado(estadoCodigo, tipoEvento) {
                if (tipoEvento === 'BLOQUEO' || tipoEvento === 'NO_ATENCION') return 'is-bloqueada';
                const codigoRaw = String(estadoCodigo ?? '').trim();
                const codigo = codigoRaw.toUpperCase();
                switch (codigo) {
                    case '1':
                    case 'PENDIENTE': return 'is-pendiente';
                    case '2':
                    case 'CONFIRMADA': return 'is-confirmada';
                    case '3':
                    case 'EN_USO': return 'is-finalizada';
                    case '4':
                    case 'PAGADA':
                    case 'FINALIZADA': return 'is-finalizada';
                    case '5':
                    case 'CANCELADA': return 'is-cancelada';
                    case '6':
                    case 'NO_ASISTIO':
                    case 'NO_SHOW': return 'is-noshow';
                    default: return 'is-reservada';
                }
            }

            function colorTextoByEstado(estadoCodigo, tipoEvento) {
                const clase = classByEstado(estadoCodigo, tipoEvento);
                switch (clase) {
                    case 'is-pendiente': return '#7c2d12';
                    case 'is-confirmada': return '#166534';
                    case 'is-finalizada': return '#0f4e63';
                    case 'is-cancelada': return '#991b1b';
                    case 'is-noshow': return '#5b21b6';
                    default: return '#1e3a8a';
                }
            }

            function textoEstadoPorCodigo(estado) {
                switch (Number(estado)) {
                    case 1: return 'Pendiente';
                    case 2: return 'Confirmada';
                    case 3: return 'Pagada';
                    case 4: return 'Pagada';
                    case 5: return 'Cancelada';
                    case 6: return 'No Asistio';
                    default: return 'Sin estado';
                }
            }

            function obtenerTextoComboPorValor(selectEl, value) {
                if (!selectEl || value === null || value === undefined) return '-';
                const val = String(value);
                const opt = Array.from(selectEl.options || []).find(o => String(o.value) === val);
                return opt ? String(opt.text || '-').trim() : '-';
            }

            function limpiarErrorModal() {
                reservaModalError.classList.add('d-none');
                reservaModalError.textContent = '';
            }

            function limpiarDisponibilidadModal() {
                reservaModalDisponibilidad.classList.add('d-none');
                reservaModalDisponibilidad.classList.remove('alert-success', 'alert-warning');
                reservaModalDisponibilidad.textContent = '';
                ['modalEspacioDeportivoId', 'modalFecha', 'modalHoraInicio', 'modalHoraFin']
                    .forEach((id) => {
                        const el = document.getElementById(id);
                        if (!el) return;
                        el.classList.remove('is-invalid', 'is-valid');
                    });
                conflictoReservaIdActual = null;
                reservaModalAccionConflicto.classList.add('d-none');
            }

            function mostrarErrorModal(msg) {
                reservaModalError.textContent = msg || 'No se pudo guardar la reserva.';
                reservaModalError.classList.remove('d-none');
            }

            function mostrarDisponibilidadModal(disponible, mensaje, conflictoTipo, conflictoId) {
                reservaModalDisponibilidad.classList.remove('d-none');
                reservaModalDisponibilidad.classList.toggle('alert-success', disponible);
                reservaModalDisponibilidad.classList.toggle('alert-warning', !disponible);
                reservaModalDisponibilidad.textContent = mensaje;
                btnGuardarReservaModal.disabled = !disponible;
                ['modalEspacioDeportivoId', 'modalFecha', 'modalHoraInicio', 'modalHoraFin']
                    .forEach((id) => {
                        const el = document.getElementById(id);
                        if (!el) return;
                        el.classList.toggle('is-invalid', !disponible);
                        el.classList.toggle('is-valid', disponible);
                    });

                conflictoReservaIdActual = null;
                reservaModalAccionConflicto.classList.add('d-none');
                if (!disponible && String(conflictoTipo || '').toUpperCase() === 'RESERVA' && conflictoId) {
                    conflictoReservaIdActual = conflictoId;
                    btnVerReservaConflicto.textContent = `Ver reserva #${conflictoId}`;
                    reservaModalAccionConflicto.classList.remove('d-none');
                }
            }

            function setValue(id, value) {
                const el = document.getElementById(id);
                el.value = value ?? '';
            }

            function getFiltroEspacioSeleccionado() {
                return espacioSelect?.value || '';
            }

            function getFiltroSedeSeleccionada() {
                return sedeSelect?.value || '';
            }

            function limpiarCacheResumenDia() {
                // Consulta backend en cada render; no usamos cache local para evitar desfases.
            }

            async function obtenerResumenDiaBackend(fechaIso) {
                const sedeId = getFiltroSedeSeleccionada();
                const espacioId = getFiltroEspacioSeleccionado();
                if (!sedeId || !espacioId || !fechaIso) {
                    return null;
                }

                const params = new URLSearchParams();
                params.append('negocioId', String(negocioId));
                params.append('fecha', fechaIso);
                params.append('sedeId', String(sedeId));
                params.append('espacioDeportivoId', String(espacioId));
                params.append('_ts', Date.now().toString());

                const response = await fetch(`${resumenDiaOperativoUrl}?${params.toString()}`, { cache: 'no-store' });
                const payload = await response.json();
                if (!response.ok || !payload.ok) {
                    throw new Error(payload.mensaje || 'No se pudo consultar la disponibilidad diaria.');
                }

                return payload;
            }

            async function recargarEspaciosPorSede(valorPreferido = '', conservarActual = true) {
                if (!sedeSelect || !espacioSelect) return;
                if (!sedeSelect.value) {
                    espacioSelect.innerHTML = '<option value="">Selecciona sede primero</option>';
                    espacioSelect.value = '';
                    espacioSelect.disabled = true;
                    limpiarCacheResumenDia();
                    return;
                }

                const valorActual = espacioSelect.value || '';
                const valorObjetivo = valorPreferido || (conservarActual ? valorActual : '');
                const params = new URLSearchParams();
                params.append('negocioId', String(negocioId));
                if (sedeSelect.value) {
                    params.append('sedeId', sedeSelect.value);
                }

                try {
                    recargandoEspacios = true;
                    const r = await fetch(`${comboEspaciosUrl}?${params.toString()}`, { cache: 'no-store' });
                    const payload = await r.json();
                    if (!r.ok || !payload.ok) {
                        throw new Error(payload.mensaje || 'No se pudo cargar espacios por sede.');
                    }

                    espacioSelect.innerHTML = '';
                    const optSelecciona = document.createElement('option');
                    optSelecciona.value = '';
                    optSelecciona.textContent = 'Selecciona espacio';
                    espacioSelect.appendChild(optSelecciona);

                    let encontrado = false;
                    (payload.items || []).forEach((item) => {
                        const opt = document.createElement('option');
                        opt.value = String(item.value ?? '');
                        opt.textContent = String(item.text ?? '');
                        if (valorObjetivo && opt.value === String(valorObjetivo)) {
                            encontrado = true;
                        }
                        espacioSelect.appendChild(opt);
                    });

                    espacioSelect.value = encontrado ? String(valorObjetivo) : '';
                    espacioSelect.disabled = false;
                } catch (err) {
                    console.error(err);
                    espacioSelect.innerHTML = '<option value="">Selecciona espacio</option>';
                    espacioSelect.value = '';
                    espacioSelect.disabled = true;
                } finally {
                    recargandoEspacios = false;
                }
            }

            function buscarBloqueoBackendEnRango(inicio, fin, espacioId) {
                if (!inicio || !fin || !espacioId || !calendar) return null;
                const espacioTexto = String(espacioId);
                const eventos = calendar.getEvents();
                for (const ev of eventos) {
                    const tipo = String(ev.extendedProps.tipoEvento || '').toUpperCase();
                    if (tipo !== 'NO_ATENCION' && tipo !== 'BLOQUEO') continue;
                    if (String(ev.extendedProps.espacioDeportivoId ?? '') !== espacioTexto) continue;
                    const evInicio = ev.start;
                    const evFin = ev.end ?? ev.start;
                    if (!evInicio || !evFin) continue;
                    if (inicio < evFin && fin > evInicio) {
                        return {
                            motivo: String(ev.extendedProps.motivo || '').trim(),
                            estadoTexto: String(ev.extendedProps.estadoTexto || '').trim(),
                            titulo: String(ev.title || '').trim()
                        };
                    }
                }

                return null;
            }

            function validarAperturaReserva(fechaHora) {
                if (!getFiltroSedeSeleccionada()) {
                    return { ok: false, mensaje: 'Selecciona una sede para continuar.' };
                }
                const espacioId = getFiltroEspacioSeleccionado();
                if (!espacioId) {
                    return { ok: false, mensaje: 'Selecciona un espacio deportivo para validar disponibilidad.' };
                }
                if (formatDateIso(fechaHora) < hoyIso) {
                    return { ok: false, mensaje: 'No se permiten reservas en fechas pasadas.' };
                }
                const fin = new Date(fechaHora.getTime() + (60 * 60000));
                const bloqueo = buscarBloqueoBackendEnRango(fechaHora, fin, espacioId);
                if (bloqueo) {
                    return { ok: false, mensaje: bloqueo.motivo || bloqueo.titulo || bloqueo.estadoTexto || 'Horario bloqueado/no atencion.' };
                }
                return { ok: true, mensaje: '' };
            }

            function abrirModalNuevaReserva(fechaIso, horaInicio, horaFin, espacioId) {
                if (!getFiltroSedeSeleccionada()) {
                    alert('Selecciona una sede antes de crear una reserva.');
                    return;
                }
                limpiarErrorModal();
                limpiarDisponibilidadModal();
                reservaModalTitulo.textContent = 'Nueva reserva';
                setValue('modalReservaId', 0);
                const fechaInicial = fechaIso && !esFechaPasadaIso(fechaIso) ? fechaIso : hoyIso;
                setValue('modalFecha', fechaInicial);
                setValue('modalHoraInicio', horaInicio || '18:00');
                setValue('modalHoraFin', horaFin || '19:00');
                setValue('modalEstado', 1);
                setValue('modalTotal', '0.00');
                setValue('modalAdelanto', '0.00');
                setValue('modalEspacioDeportivoId', espacioId || getFiltroEspacioSeleccionado());
                setValue('modalClienteId', '');
                reservaModal.show();
                validarDisponibilidadModal();
            }

            function abrirModalEditarReserva(reservaId) {
                limpiarErrorModal();
                limpiarDisponibilidadModal();
                fetch(`${getReservaModalUrl}?negocioId=${negocioId}&id=${reservaId}`)
                    .then(async r => {
                        const payload = await r.json();
                        if (!r.ok || !payload.ok) {
                            throw new Error(payload.mensaje || 'No se pudo obtener la reserva.');
                        }
                        return payload;
                    })
                    .then(data => {
                        reservaModalTitulo.textContent = `Editar reserva #${data.id}`;
                        setValue('modalReservaId', data.id);
                        setValue('modalEspacioDeportivoId', data.espacioDeportivoId);
                        setValue('modalClienteId', data.clienteId);
                        setValue('modalFecha', data.fecha);
                        setValue('modalHoraInicio', data.horaInicio);
                        setValue('modalHoraFin', data.horaFin);
                        setValue('modalEstado', data.estado);
                        setValue('modalTotal', data.total ?? '0.00');
                        setValue('modalAdelanto', data.adelanto ?? '0.00');
                        reservaModal.show();
                        validarDisponibilidadModal();
                    })
                    .catch(err => alert(err.message));
            }

            function validarDisponibilidadModal() {
                limpiarErrorModal();
                const reservaId = document.getElementById('modalReservaId').value || '';
                const espacioId = document.getElementById('modalEspacioDeportivoId').value || '';
                const fecha = document.getElementById('modalFecha').value || '';
                const horaInicio = document.getElementById('modalHoraInicio').value || '';
                const horaFin = document.getElementById('modalHoraFin').value || '';

                if (!espacioId || !fecha || !horaInicio || !horaFin) {
                    limpiarDisponibilidadModal();
                    btnGuardarReservaModal.disabled = true;
                    return;
                }
                if (esFechaPasadaIso(fecha)) {
                    mostrarDisponibilidadModal(false, 'No se permite registrar reservas en fechas pasadas.', null, null);
                    return;
                }

                const params = new URLSearchParams();
                params.append('negocioId', negocioId);
                if (reservaId && Number(reservaId) > 0) params.append('reservaId', reservaId);
                params.append('espacioDeportivoId', espacioId);
                params.append('fecha', fecha);
                params.append('horaInicio', horaInicio);
                params.append('horaFin', horaFin);

                fetch(`${validarDisponibilidadModalUrl}?${params.toString()}`)
                    .then(async r => {
                        const payload = await r.json();
                        if (!r.ok || !payload.ok) {
                            throw new Error(payload.mensaje || 'No se pudo validar disponibilidad.');
                        }
                        mostrarDisponibilidadModal(payload.disponible, payload.mensaje, payload.conflictoTipo, payload.conflictoId);
                    })
                    .catch(err => {
                        mostrarDisponibilidadModal(false, err.message, null, null);
                    });
            }

            function validarDisponibilidadModalDebounce() {
                if (timerValidacion) {
                    clearTimeout(timerValidacion);
                }
                timerValidacion = setTimeout(validarDisponibilidadModal, 220);
            }

            function obtenerTooltipBloqueo(arg) {
                const motivoApi = arg.event.extendedProps.motivo || arg.event.extendedProps.detalle || arg.event.title;
                if (motivoApi) return String(motivoApi);
                const estadoTexto = arg.event.extendedProps.estadoTexto;
                if (estadoTexto) return String(estadoTexto);
                return 'Horario bloqueado/no atencion';
            }

            async function renderDisponibilidadDia(dateClick) {
                const requestSeq = ++disponibilidadRenderSeq;
                if (!calendar || !disponibilidadSlotsEl || !disponibilidadMensajeEl) return;
                let fechaTrabajo = dateClick;
                if (!fechaTrabajo && disponibilidadFechaFiltroEl?.value) {
                    fechaTrabajo = new Date(`${disponibilidadFechaFiltroEl.value}T00:00:00`);
                }
                if (!fechaTrabajo || Number.isNaN(fechaTrabajo.getTime())) {
                    disponibilidadMensajeEl.textContent = 'Selecciona una fecha valida para ver disponibilidad.';
                    return;
                }
                const espacioId = document.getElementById('espacioDeportivoId').value;
                disponibilidadSlotsEl.innerHTML = '';
                if (disponibilidadFechaFiltroEl) {
                    disponibilidadFechaFiltroEl.value = formatDateIso(fechaTrabajo);
                }

                if (!getFiltroSedeSeleccionada()) {
                    disponibilidadMensajeEl.textContent = 'Selecciona una sede para ver disponibilidad segun su horario de atencion.';
                    return;
                }

                if (!espacioId) {
                    disponibilidadMensajeEl.textContent = 'Selecciona un espacio en el filtro para ver disponibilidad exacta por dia.';
                    return;
                }

                const fechaIso = formatDateIso(fechaTrabajo);
                if (esFechaPasadaIso(fechaIso)) {
                    disponibilidadMensajeEl.textContent = `La fecha ${fechaIso} ya paso. Solo se pueden registrar reservas desde hoy.`;
                    return;
                }
                try {
                    const resumen = await obtenerResumenDiaBackend(fechaIso);
                    if (requestSeq !== disponibilidadRenderSeq) return;
                    pintarKpiOperativoDia(resumen?.kpi, resumen?.totalPendientes ?? 0);
                    const slots = resumen?.slotsDisponibles || [];
                    const ahora = new Date();
                    const esHoySeleccionado = fechaIso === hoyIso;
                    let slotsVisibles = 0;
                    disponibilidadSlotsEl.innerHTML = '';
                    for (const slot of slots) {
                        const horaInicio = String(slot.horaInicio || '').trim();
                        const horaFin = String(slot.horaFin || '').trim();
                        if (!horaInicio || !horaFin) continue;
                        if (esHoySeleccionado) {
                            const partesInicio = horaInicio.split(':');
                            const hh = Number(partesInicio[0]);
                            const mm = Number(partesInicio[1] || '0');
                            if (Number.isNaN(hh) || Number.isNaN(mm)) continue;
                            const inicioSlotHoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate(), hh, mm, 0, 0);
                            if (inicioSlotHoy <= ahora) continue;
                        }
                        const btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = 'btn btn-sm sc-slot-btn';
                        btn.textContent = `${horaInicio} - ${horaFin}`;
                        btn.addEventListener('click', () => {
                            abrirModalNuevaReserva(fechaIso, horaInicio, horaFin, espacioId);
                        });
                        disponibilidadSlotsEl.appendChild(btn);
                        slotsVisibles += 1;
                    }

                    if (slotsVisibles === 0) {
                        disponibilidadMensajeEl.textContent = `No hay horas libres para ${fechaIso} en el espacio seleccionado.`;
                        return;
                    }
                    disponibilidadMensajeEl.textContent = `Horas disponibles para ${fechaIso}: ${slotsVisibles} bloque(s) libre(s).`;
                } catch (err) {
                    if (requestSeq !== disponibilidadRenderSeq) return;
                    disponibilidadMensajeEl.textContent = err?.message || 'No se pudo consultar disponibilidad en este momento.';
                }
            }

            async function renderPendientesConfirmarDia(dateClick) {
                const requestSeq = ++pendientesRenderSeq;
                if (!calendar || !pendientesSlotsEl || !pendientesMensajeEl) return;
                let fechaTrabajo = dateClick;
                if (!fechaTrabajo && pendientesFechaFiltroEl?.value) {
                    fechaTrabajo = new Date(`${pendientesFechaFiltroEl.value}T00:00:00`);
                }
                if (!fechaTrabajo || Number.isNaN(fechaTrabajo.getTime())) {
                    pendientesMensajeEl.textContent = 'Selecciona una fecha valida para revisar reservas pendientes.';
                    return;
                }
                const espacioId = document.getElementById('espacioDeportivoId').value;
                pendientesSlotsEl.innerHTML = '';
                if (pendientesFechaFiltroEl) {
                    pendientesFechaFiltroEl.value = formatDateIso(fechaTrabajo);
                }

                if (!getFiltroSedeSeleccionada()) {
                    pendientesMensajeEl.textContent = 'Selecciona una sede para revisar reservas pendientes.';
                    return;
                }

                if (!espacioId) {
                    pendientesMensajeEl.textContent = 'Selecciona un espacio en el filtro para mostrar pendientes.';
                    return;
                }

                const fechaIso = formatDateIso(fechaTrabajo);
                try {
                    const resumen = await obtenerResumenDiaBackend(fechaIso);
                    if (requestSeq !== pendientesRenderSeq) return;
                    pintarKpiOperativoDia(resumen?.kpi, resumen?.totalPendientes ?? 0);
                    const pendientes = resumen?.pendientes || [];
                    pendientesSlotsEl.innerHTML = '';

                    if (pendientes.length === 0) {
                        pendientesMensajeEl.textContent = `No hay reservas pendientes para ${fechaIso}.`;
                        return;
                    }

                    const pendientesUnicos = new Set();
                    for (const item of pendientes) {
                        const reservaId = Number(item.reservaId || 0);
                        if (!reservaId) continue;
                        if (pendientesUnicos.has(reservaId)) continue;
                        pendientesUnicos.add(reservaId);
                        const hi = String(item.horaInicio || '').trim();
                        const hf = String(item.horaFin || '').trim();
                        const titulo = String(item.titulo || `Reserva #${reservaId}`).trim();
                        const btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = 'btn btn-sm sc-pendiente-slot-btn';
                        btn.innerHTML = `<span class="sc-pendiente-slot-hora">${hi} - ${hf}</span><span class="sc-pendiente-slot-titulo">${titulo}</span>`;
                        btn.addEventListener('click', () => abrirModalEditarReserva(reservaId));
                        pendientesSlotsEl.appendChild(btn);
                    }

                    pendientesMensajeEl.textContent = `Reservas pendientes para ${fechaIso}: ${pendientesUnicos.size} registro(s).`;
                } catch (err) {
                    if (requestSeq !== pendientesRenderSeq) return;
                    pendientesMensajeEl.textContent = err?.message || 'No se pudo consultar pendientes en este momento.';
                }
            }

            function pintarKpiOperativoDia(kpi, totalPendientes) {
                if (!kpiOcupacionDiaEl) return;
                const data = kpi || {};
                kpiOcupacionDiaEl.textContent = `${Number(data.ocupacionPct || 0).toFixed(2)}%`;
                if (kpiSlotsOcupadosDiaEl) kpiSlotsOcupadosDiaEl.textContent = String(data.slotsOcupados ?? 0);
                if (kpiSlotsLibresDiaEl) kpiSlotsLibresDiaEl.textContent = String(data.slotsLibres ?? 0);
                if (kpiReservasActivasDiaEl) kpiReservasActivasDiaEl.textContent = String(data.reservasActivas ?? 0);
                if (kpiPendientesDiaEl) kpiPendientesDiaEl.textContent = String(totalPendientes ?? 0);
            }

            function armarUrlEventos(info) {
                const sedeId = sedeSelect?.value || '';
                const espacioDeportivoId = espacioSelect?.value || '';
                const estado = estadoSelect?.value || '';
                const params = new URLSearchParams();
                params.append('negocioId', negocioId);
                params.append('start', info.startStr);
                params.append('end', info.endStr);
                params.append('_ts', Date.now().toString());
                if (sedeId) params.append('sedeId', sedeId);
                if (espacioDeportivoId) params.append('espacioDeportivoId', espacioDeportivoId);
                if (estado) params.append('estado', estado);
                return '@Url.Action("CalendarioEventos", "Reservas")' + '?' + params.toString();
            }

            function notificarCalendario(mensaje, tipo = 'warning') {
                const texto = String(mensaje || '').trim();
                if (!texto) return;
                const div = document.createElement('div');
                div.className = `alert alert-${tipo} shadow-sm`;
                div.style.position = 'fixed';
                div.style.top = '1rem';
                div.style.right = '1rem';
                div.style.zIndex = '2000';
                div.style.maxWidth = '420px';
                div.style.margin = '0';
                div.textContent = texto;
                document.body.appendChild(div);
                window.setTimeout(() => {
                    if (div.parentElement) div.parentElement.removeChild(div);
                }, 3200);
            }

            async function procesarMovimientoReserva(info, mensajeErrorPorDefecto) {
                const reservaId = info?.event?.extendedProps?.reservaId;
                if (!reservaId) {
                    info?.revert?.();
                    return;
                }

                const controller = new AbortController();
                const timeoutId = window.setTimeout(() => controller.abort(), 12000);
                try {
                    const response = await fetch('@Url.Action("MoverEvento", "Reservas")', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        signal: controller.signal,
                        body: JSON.stringify({
                            negocioId: negocioId,
                            reservaId: reservaId,
                            inicio: toLocalDateTimeString(info.event.start),
                            fin: toLocalDateTimeString(info.event.end ? info.event.end : info.event.start)
                        })
                    });

                    const payload = await response.json().catch(() => ({}));
                    if (!response.ok || !payload?.ok) {
                        throw new Error(payload?.mensaje || mensajeErrorPorDefecto);
                    }

                    calendar?.refetchEvents();
                    limpiarCacheResumenDia();
                    void renderDisponibilidadDia(info.event.start ?? new Date());
                    void renderPendientesConfirmarDia(info.event.start ?? new Date());
                } catch (err) {
                    info?.revert?.();
                    calendar?.unselect();
                    calendar?.refetchEvents();
                    const mensaje = err?.name === 'AbortError'
                        ? 'La validacion tardÃ³ demasiado. Se restauro la reserva a su posicion original.'
                        : (err?.message || mensajeErrorPorDefecto);
                    notificarCalendario(mensaje, 'warning');
                } finally {
                    window.clearTimeout(timeoutId);
                }
            }

            if (calendarioActivo) {
                calendar = new FullCalendar.Calendar(calendarioEl, {
                    initialView: 'timeGridWeek',
                    locale: 'es',
                    firstDay: 1,
                    navLinks: false,
                    allDaySlot: false,
                    dayHeaderContent: function(arg) {
                        const nombreDia = capitalizar(arg.date.toLocaleDateString('es-PE', { weekday: 'long' }));
                        const fechaCorta = arg.date.toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit' });
                        return {
                            html: `<span class="sc-day-header-wrap"><span class="sc-day-header-title">${nombreDia}</span><span class="sc-day-header-date">${fechaCorta}</span></span>`
                        };
                    },
                    dayHeaderClassNames: function() { return ['sc-header-day-cell']; },
                    slotLabelFormat: {
                        hour: '2-digit',
                        minute: '2-digit',
                        hour12: false
                    },
                    slotLabelContent: function(arg) {
                        const hh = pad2(arg.date.getHours());
                        const mm = pad2(arg.date.getMinutes());
                        return { html: hh + ':' + mm };
                    },
                    slotLabelInterval: '01:00:00',
                    slotDuration: '00:30:00',
                    slotMinTime: '06:00:00',
                    slotMaxTime: '24:00:00',
                    nowIndicator: true,
                    editable: true,
                    selectable: true,
                    height: 'auto',
                    expandRows: true,
                    buttonText: {
                        today: 'Hoy',
                        week: 'Semana',
                        day: 'Dia'
                    },
                    customButtons: {
                        nuevaReserva: {
                            text: 'Nueva reserva',
                            click: function() {
                                abrirModalNuevaReserva();
                            }
                        }
                    },
                    headerToolbar: {
                        left: 'prev,next today',
                        center: 'title',
                        right: puedeCrearReserva ? 'timeGridWeek,timeGridDay nuevaReserva' : 'timeGridWeek,timeGridDay'
                    },
                    eventClassNames: function(arg) {
                        if (arg.event.display === 'background') {
                            return [];
                        }
                        return ['sc-fc-event', classByEstado(arg.event.extendedProps.estadoCodigo, arg.event.extendedProps.tipoEvento)];
                    },
                    eventContent: function(arg) {
                        if (arg.event.display === 'background') {
                            return { html: '' };
                        }
                        const timeText = arg.timeText ? `<div class="sc-fc-event-time">${arg.timeText}</div>` : '';
                        return { html: `<div class="sc-fc-event-inner"><div class="sc-fc-event-title">${arg.event.title || ''}</div>${timeText}</div>` };
                    },
                    eventDidMount: function(arg) {
                        if (arg.event.display === 'background') {
                            const tipo = String(arg.event.extendedProps.tipoEvento || '').toUpperCase();
                            if (tipo === 'NO_ATENCION' || tipo === 'BLOQUEO') {
                                const colorBloqueo = String(arg.event.backgroundColor || arg.event.borderColor || '#64748b').trim();
                                arg.el.style.setProperty('--sc-block-bg-a', hexToRgba(colorBloqueo, 0.18));
                                arg.el.style.setProperty('--sc-block-bg-b', hexToRgba(colorBloqueo, 0.1));
                                arg.el.style.setProperty('--sc-block-border', hexToRgba(colorBloqueo, 0.32));
                                arg.el.textContent = '';
                                const motivoTexto = String(arg.event.title || '').trim();
                                const estadoTexto = String(arg.event.extendedProps.estadoTexto || '').trim();
                                const etiquetaTexto = motivoTexto || estadoTexto;
                                if (etiquetaTexto) {
                                    const motivo = document.createElement('span');
                                    motivo.className = 'sc-bg-motivo-inline';
                                    motivo.textContent = etiquetaTexto;
                                    arg.el.appendChild(motivo);
                                }
                                const textoTooltip = obtenerTooltipBloqueo(arg);
                                if (textoTooltip) {
                                    arg.el.setAttribute('title', textoTooltip);
                                    arg.el.setAttribute('data-bs-toggle', 'tooltip');
                                    arg.el.setAttribute('data-bs-placement', 'top');
                                    arg.el.setAttribute('data-bs-custom-class', 'sc-tooltip');
                                    if (window.bootstrap && bootstrap.Tooltip) {
                                        const tip = new bootstrap.Tooltip(arg.el, {
                                            container: 'body',
                                            trigger: 'hover'
                                        });
                                        arg.el._scTooltip = tip;
                                    }
                                }
                            }
                            return;
                        }
                        const estadoCodigo = arg.event.extendedProps.estadoCodigo;
                        const tipo = arg.event.extendedProps.tipoEvento;
                        const colorTexto = colorTextoByEstado(estadoCodigo, tipo);
                        arg.el.style.setProperty('--fc-event-text-color', colorTexto);
                        arg.el.style.setProperty('--sc-event-text-color', colorTexto);
                        arg.el.querySelectorAll('.sc-fc-event-title, .sc-fc-event-time, .fc-event-main, .fc-event-title, .fc-event-time').forEach(el => {
                            el.style.color = colorTexto;
                        });
                    },
                    eventWillUnmount: function(arg) {
                        if (arg.el && arg.el._scTooltip && typeof arg.el._scTooltip.dispose === 'function') {
                            arg.el._scTooltip.dispose();
                        }
                    },
                    datesSet: function(info) {
                        renderToolbarSubtitle(info.start, info.end);
                        pintarCabeceraHorarios();
                        aplicarEstiloCalendarioFase1();
                        activarHoverColumnasCalendario();
                        aplicarFiltroLeyendaCalendario();
                        actualizarContadoresLeyendaCalendario();
                    },
                    eventsSet: function() {
                        actualizarContadoresLeyendaCalendario();
                    },
                    events: function(info, success, failure) {
                        fetch(armarUrlEventos(info), { cache: 'no-store' })
                            .then(r => r.json())
                            .then(data => {
                                const eventosApi = data.flatMap(ev => {
                                    const tipo = String(ev.tipoEvento || '');
                                    const colorEvento = ev.backgroundColor || ev.borderColor || ev.color || '#64748b';
                                    const base = { ...ev };
                                    delete base.backgroundColor;
                                    delete base.borderColor;
                                    delete base.textColor;
                                    delete base.color;
                                    if (tipo === 'NO_ATENCION') {
                                        return [{ ...base, display: 'background', classNames: ['sc-no-atencion-bg'], editable: false, overlap: false, backgroundColor: colorEvento, borderColor: colorEvento }];
                                    }
                                    if (tipo === 'BLOQUEO') {
                                        return [{ ...base, display: 'background', classNames: ['sc-bloqueo-bg'], editable: false, overlap: false, backgroundColor: colorEvento, borderColor: colorEvento }];
                                    }
                                    const estadoClase = classByEstado(base.estadoCodigo, base.tipoEvento);
                                const fondoEstado = {
                                    ...base,
                                    id: `BG-${base.id}`,
                                        title: '',
                                        display: 'background',
                                        editable: false,
                                        overlap: true,
                                        classNames: ['sc-reserva-slot-bg', estadoClase]
                                    };
                                base.textColor = colorTextoByEstado(base.estadoCodigo, base.tipoEvento);
                                return [fondoEstado, base];
                            });
                            success(eventosApi);
                            setTimeout(() => {
                                const fechaRef = disponibilidadFechaFiltroEl?.value || hoyIso;
                                void renderDisponibilidadDia(new Date(`${fechaRef}T00:00:00`));
                                const fechaPendienteRef = pendientesFechaFiltroEl?.value || hoyIso;
                                void renderPendientesConfirmarDia(new Date(`${fechaPendienteRef}T00:00:00`));
                            }, 0);
                        })
                        .catch(err => failure(err));
                },
                    dateClick: function(info) {
                        if (!info.allDay) {
                            const inicio = info.date;
                            const validacion = validarAperturaReserva(inicio);
                            if (!validacion.ok) {
                                alert(validacion.mensaje);
                                return;
                            }
                            const fin = new Date(inicio.getTime() + (60 * 60000));
                            abrirModalNuevaReserva(
                                formatDateIso(inicio),
                                formatTime(inicio.getHours(), inicio.getMinutes()),
                                formatTime(fin.getHours(), fin.getMinutes()),
                                getFiltroEspacioSeleccionado()
                            );
                            return;
                        }
                        renderDisponibilidadDia(info.date);
                    },
                    select: function(info) {
                        renderDisponibilidadDia(info.start);
                        renderPendientesConfirmarDia(info.start);
                    },
                    eventDrop: function(info) {
                        void procesarMovimientoReserva(info, 'No se pudo mover la reserva.');
                    },
                    eventResize: function(info) {
                        void procesarMovimientoReserva(info, 'No se pudo redimensionar la reserva.');
                    },
                    eventClick: function(info) {
                        const reservaId = info.event.extendedProps.reservaId;
                        if (!reservaId) {
                            return;
                        }
                        abrirModalEditarReserva(reservaId);
                    }
                });

                calendar.render();
                renderToolbarSubtitle(calendar.view.currentStart, calendar.view.currentEnd);
                pintarCabeceraHorarios();
                aplicarEstiloCalendarioFase1();
                activarHoverColumnasCalendario();
                aplicarFiltroLeyendaCalendario();
                actualizarContadoresLeyendaCalendario();
                void renderDisponibilidadDia(new Date(`${disponibilidadFechaFiltroEl?.value || hoyIso}T00:00:00`));
                void renderPendientesConfirmarDia(new Date(`${pendientesFechaFiltroEl?.value || hoyIso}T00:00:00`));
            }
            document.getElementById('modalFecha').setAttribute('min', hoyIso);
            if (disponibilidadFechaFiltroEl) {
                if (!disponibilidadFechaFiltroEl.value) {
                    disponibilidadFechaFiltroEl.value = hoyIso;
                }
                disponibilidadFechaFiltroEl.setAttribute('min', hoyIso);
                disponibilidadFechaFiltroEl.addEventListener('change', () => {
                    const fechaRef = disponibilidadFechaFiltroEl.value;
                    if (!fechaRef || !calendar) return;
                    calendar.gotoDate(fechaRef);
                    setTimeout(() => {
                        limpiarCacheResumenDia();
                        void renderDisponibilidadDia(new Date(`${fechaRef}T00:00:00`));
                    }, 80);
                });
            }
            if (pendientesFechaFiltroEl) {
                if (!pendientesFechaFiltroEl.value) {
                    pendientesFechaFiltroEl.value = hoyIso;
                }
                pendientesFechaFiltroEl.setAttribute('min', hoyIso);
                pendientesFechaFiltroEl.addEventListener('change', () => {
                    const fechaRef = pendientesFechaFiltroEl.value;
                    if (!fechaRef || !calendar) return;
                    calendar.gotoDate(fechaRef);
                    setTimeout(() => {
                        limpiarCacheResumenDia();
                        void renderPendientesConfirmarDia(new Date(`${fechaRef}T00:00:00`));
                    }, 80);
                });
            }
            btnEditarReservaListadoEls.forEach((btn) => {
                btn.addEventListener('click', () => {
                    const reservaId = Number(btn.dataset.reservaId || 0);
                    if (!reservaId) return;
                    abrirModalEditarReserva(reservaId);
                });
            });

            legendEstadoButtons.forEach((btn) => {
                btn.addEventListener('click', () => {
                    const estado = String(btn.dataset.legendEstado || '').trim().toLowerCase();
                    if (!estado) return;
                    if (estadosOcultosCalendario.has(estado)) {
                        estadosOcultosCalendario.delete(estado);
                    } else {
                        estadosOcultosCalendario.add(estado);
                    }
                    sincronizarLeyendaCalendario();
                    aplicarFiltroLeyendaCalendario();
                    actualizarContadoresLeyendaCalendario();
                });
            });
            sincronizarLeyendaCalendario();
            actualizarContadoresLeyendaCalendario();

            if (sedeSelect && espacioSelect) {
                const espacioInicial = espacioSelect.value || '';
                recargarEspaciosPorSede(espacioInicial, true);
            }

            btnVerReservaConflicto.addEventListener('click', () => {
                if (!conflictoReservaIdActual) return;
                abrirModalEditarReserva(conflictoReservaIdActual);
            });

            ['modalEspacioDeportivoId', 'modalFecha', 'modalHoraInicio', 'modalHoraFin']
                .forEach(id => {
                    const el = document.getElementById(id);
                    if (!el) return;
                    el.addEventListener('change', validarDisponibilidadModalDebounce);
                    el.addEventListener('keyup', validarDisponibilidadModalDebounce);
                });

            reservaModalForm.addEventListener('submit', function (e) {
                e.preventDefault();
                limpiarErrorModal();
                if (btnGuardarReservaModal.disabled) {
                    mostrarErrorModal('Debes corregir la disponibilidad antes de guardar.');
                    return;
                }
                const data = new FormData(reservaModalForm);
                fetch(guardarReservaModalUrl, {
                    method: 'POST',
                    body: data
                })
                .then(async r => {
                    const payload = await r.json();
                    if (!r.ok || !payload.ok) {
                        throw new Error(payload.mensaje || 'No se pudo guardar la reserva.');
                    }
                    reservaModal.hide();
                    if (calendar) {
                        calendar.refetchEvents();
                        limpiarCacheResumenDia();
                        void renderDisponibilidadDia(new Date(document.getElementById('modalFecha').value + 'T00:00:00'));
                        void renderPendientesConfirmarDia(new Date(document.getElementById('modalFecha').value + 'T00:00:00'));
                        setTimeout(() => calendar.refetchEvents(), 250);
                    }
                })
                .catch(err => mostrarErrorModal(err.message));
            });

            formFiltros.addEventListener('submit', () => {
                const inicio = document.getElementById('fechaDesde').value;
                if (calendar && inicio) {
                    calendar.gotoDate(inicio);
                }
            });
            const enviarFiltros = () => {
                if (!formFiltros) return;
                if (typeof formFiltros.requestSubmit === 'function') {
                    formFiltros.requestSubmit();
                    return;
                }
                formFiltros.submit();
            };

            if (sedeSelect) {
                sedeSelect.addEventListener('change', async () => {
                    await recargarEspaciosPorSede('', false);
                    limpiarCacheResumenDia();
                    if (disponibilidadMensajeEl) disponibilidadMensajeEl.textContent = 'Selecciona un espacio en el filtro para ver disponibilidad exacta por dia.';
                    if (disponibilidadSlotsEl) disponibilidadSlotsEl.innerHTML = '';
                    if (pendientesMensajeEl) pendientesMensajeEl.textContent = 'Selecciona un espacio en el filtro para mostrar pendientes.';
                    if (pendientesSlotsEl) pendientesSlotsEl.innerHTML = '';
                    if (calendar) {
                        renderToolbarSubtitle(calendar.view.currentStart, calendar.view.currentEnd);
                        calendar.refetchEvents();
                    }
                });
            }

            if (espacioSelect) {
                espacioSelect.addEventListener('change', () => {
                    if (recargandoEspacios) return;
                    limpiarCacheResumenDia();
                    if (!calendarioActivo) {
                        enviarFiltros();
                        return;
                    }
                    if (calendar) {
                        renderToolbarSubtitle(calendar.view.currentStart, calendar.view.currentEnd);
                        calendar.refetchEvents();
                        const fechaDisponibilidad = disponibilidadFechaFiltroEl?.value || hoyIso;
                        const fechaPendiente = pendientesFechaFiltroEl?.value || hoyIso;
                        void renderDisponibilidadDia(new Date(`${fechaDisponibilidad}T00:00:00`));
                        void renderPendientesConfirmarDia(new Date(`${fechaPendiente}T00:00:00`));
                    }
                });
            }

            if (estadoSelect) {
                estadoSelect.addEventListener('change', () => {
                    if (calendar) {
                        renderToolbarSubtitle(calendar.view.currentStart, calendar.view.currentEnd);
                        calendar.refetchEvents();
                        return;
                    }
                    enviarFiltros();
                });
            }

        })();
