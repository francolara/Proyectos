// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

document.addEventListener("click", function (event) {
    const row = event.target.closest(".workspace-clickable-row");
    if (!row) {
        return;
    }

    if (event.target.closest("a, button, input, select, textarea, label, form")) {
        return;
    }

    const editUrl = row.dataset.editUrl;
    if (!editUrl) {
        return;
    }

    window.location.href = editUrl;
});

document.addEventListener("DOMContentLoaded", function () {
    const storageKey = "sisadm-theme";
    const root = document.documentElement;
    const toggles = Array.from(document.querySelectorAll("[data-theme-toggle]"));
    const labels = Array.from(document.querySelectorAll("[data-theme-label]"));
    if (toggles.length === 0 && labels.length === 0) {
        return;
    }

    const applyTheme = function (theme) {
        const resolvedTheme = theme === "dark" ? "dark" : "light";
        root.setAttribute("data-theme", resolvedTheme);
        root.setAttribute("data-bs-theme", resolvedTheme);

        try {
            localStorage.setItem(storageKey, resolvedTheme);
        } catch {
            // Ignora errores de almacenamiento y mantiene el tema aplicado.
        }

        const isDark = resolvedTheme === "dark";
        toggles.forEach(function (toggle) {
            toggle.setAttribute("aria-pressed", isDark ? "true" : "false");
        });

        labels.forEach(function (label) {
            label.textContent = isDark ? "Modo oscuro" : "Modo claro";
        });
    };

    toggles.forEach(function (toggle) {
        toggle.addEventListener("click", function () {
            const nextTheme = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
            applyTheme(nextTheme);
        });
    });

    applyTheme(root.getAttribute("data-theme"));
});

// Firma: FRANCO LARA - 04/08/2026 | Controla de forma reutilizable el despliegue accesible de filtros avanzados en reportes contables.
document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll("[data-report-filter-toggle]").forEach(function (toggle) {
        const targetId = toggle.getAttribute("data-target");
        const filters = targetId ? document.getElementById(targetId) : null;
        if (!filters) {
            return;
        }

        const syncAdvancedFilters = function () {
            filters.classList.toggle("d-none", !toggle.checked);
            toggle.setAttribute("aria-expanded", toggle.checked ? "true" : "false");
        };

        toggle.addEventListener("change", syncAdvancedFilters);
        syncAdvancedFilters();
    });
});

document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll("[data-password-toggle]").forEach(function (toggle) {
        const targetId = toggle.getAttribute("data-password-target");
        const input = targetId
            ? document.getElementById(targetId)
            : toggle.closest(".auth-password-field")?.querySelector("input");

        if (!(input instanceof HTMLInputElement)) {
            return;
        }

        const icon = toggle.querySelector("i");

        const syncState = function () {
            const visible = input.type === "text";
            toggle.setAttribute("aria-pressed", visible ? "true" : "false");
            toggle.setAttribute("aria-label", visible ? "Ocultar contrasena" : "Mostrar contrasena");

            if (icon) {
                icon.classList.toggle("bi-eye", !visible);
                icon.classList.toggle("bi-eye-slash", visible);
            }
        };

        toggle.addEventListener("click", function () {
            input.type = input.type === "password" ? "text" : "password";
            syncState();
            input.focus({ preventScroll: true });

            if (typeof input.setSelectionRange === "function") {
                const cursor = input.value.length;
                input.setSelectionRange(cursor, cursor);
            }
        });

        syncState();
    });
});

document.addEventListener("DOMContentLoaded", function () {
    const personaForm = document.querySelector("[data-persona-form='true']");
    if (!personaForm) {
        return;
    }

    const tipoPersona = personaForm.querySelector("[data-persona-tipo]");
    const seccionNatural = personaForm.querySelector("[data-persona-natural-section]");
    const seccionJuridica = personaForm.querySelector("[data-persona-juridica-section]");
    const departamento = personaForm.querySelector("[data-ubigeo-departamento]");
    const provincia = personaForm.querySelector("[data-ubigeo-provincia]");
    const distrito = personaForm.querySelector("[data-ubigeo-distrito]");
    const provinciasUrl = personaForm.dataset.provinciasUrl;
    const distritosUrl = personaForm.dataset.distritosUrl;
    const consultaDocumentoUrl = personaForm.dataset.consultaDocumentoUrl;
    const tipoDocumento = personaForm.querySelector("[name='Formulario.TipoDocumento']");
    const numeroDocumento = personaForm.querySelector("[data-persona-numero-documento]");
    const botonConsultarDocumento = personaForm.querySelector("[data-persona-consultar-documento]");
    const consultaMensaje = personaForm.querySelector("[data-persona-consulta-mensaje]");
    const apellidoPaterno = personaForm.querySelector("[name='Formulario.ApellidoPaterno']");
    const apellidoMaterno = personaForm.querySelector("[name='Formulario.ApellidoMaterno']");
    const nombres = personaForm.querySelector("[name='Formulario.Nombres']");
    const razonSocial = personaForm.querySelector("[name='Formulario.RazonSocial']");
    const direccion = personaForm.querySelector("[name='Formulario.Direccion']");

    const alternarTipoPersona = function () {
        const esNatural = tipoPersona && tipoPersona.value === "N";
        if (seccionNatural) {
            seccionNatural.classList.toggle("is-hidden", !esNatural);
        }
        if (seccionJuridica) {
            seccionJuridica.classList.toggle("is-hidden", esNatural);
        }
    };

    const poblarSelect = function (select, items, selectedValue) {
        if (!select) {
            return;
        }

        const placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "Seleccione";
        select.innerHTML = "";
        select.appendChild(placeholder);

        items.forEach(function (item) {
            const option = document.createElement("option");
            option.value = item.value;
            option.textContent = item.text;
            if (selectedValue && selectedValue === item.value) {
                option.selected = true;
            }
            select.appendChild(option);
        });
    };

    const cargarSelect = async function (url, sourceValue, targetSelect, selectedValue) {
        if (!url || !sourceValue || !targetSelect) {
            poblarSelect(targetSelect, [], null);
            return;
        }

        const response = await fetch(url + "?" + new URLSearchParams({ [targetSelect.dataset.sourceName]: sourceValue }));
        if (!response.ok) {
            poblarSelect(targetSelect, [], null);
            return;
        }

        const items = await response.json();
        poblarSelect(targetSelect, items, selectedValue);
    };

    const mostrarMensajeConsulta = function (tipo, mensaje) {
        if (!consultaMensaje) {
            return;
        }

        consultaMensaje.className = "alert mb-3";
        consultaMensaje.classList.add(tipo === "ok" ? "alert-success" : "alert-warning");
        consultaMensaje.textContent = mensaje;
        consultaMensaje.classList.remove("d-none");
    };

    const limpiarMensajeConsulta = function () {
        if (!consultaMensaje) {
            return;
        }

        consultaMensaje.className = "alert d-none mb-3";
        consultaMensaje.textContent = "";
    };

    const aplicarUbigeo = async function (codigoUbigeo) {
        if (!departamento || !provincia || !distrito) {
            return;
        }

        if (!codigoUbigeo || codigoUbigeo.length < 6) {
            departamento.value = "";
            poblarSelect(provincia, [], null);
            poblarSelect(distrito, [], null);
            return;
        }

        const codigoDepartamento = codigoUbigeo.slice(0, 2);
        const codigoProvincia = codigoUbigeo.slice(0, 4);

        departamento.value = codigoDepartamento;
        await cargarSelect(provinciasUrl, codigoDepartamento, provincia, codigoProvincia);
        await cargarSelect(distritosUrl, codigoProvincia, distrito, codigoUbigeo);
    };

    const aplicarConsultaDocumento = async function () {
        limpiarMensajeConsulta();

        const tipoDocumentoValue = tipoDocumento ? String(tipoDocumento.value || "").trim() : "";
        const numeroDocumentoValue = numeroDocumento ? String(numeroDocumento.value || "").trim() : "";
        if (!consultaDocumentoUrl || !tipoDocumentoValue || !numeroDocumentoValue) {
            mostrarMensajeConsulta("error", "Seleccione tipo de documento e ingrese el numero antes de consultar.");
            return;
        }

        if (botonConsultarDocumento) {
            botonConsultarDocumento.disabled = true;
        }

        try {
            const response = await fetch(`${consultaDocumentoUrl}?${new URLSearchParams({ tipoDocumento: tipoDocumentoValue, numeroDocumento: numeroDocumentoValue }).toString()}`, { cache: "no-store" });
            const payload = await response.json();
            if (!response.ok || !payload.ok) {
                throw new Error(payload.mensaje || "No se pudo consultar el documento.");
            }

            if (payload.tipoPersona && tipoPersona) {
                tipoPersona.value = payload.tipoPersona;
                alternarTipoPersona();
            }
            if (payload.tipoDocumento && tipoDocumento) {
                tipoDocumento.value = payload.tipoDocumento;
            }
            if (payload.numeroDocumento && numeroDocumento) {
                numeroDocumento.value = payload.numeroDocumento;
            }
            if (typeof payload.razonSocial === "string" && razonSocial) {
                razonSocial.value = payload.razonSocial;
            }
            if (typeof payload.apellidoPaterno === "string" && apellidoPaterno) {
                apellidoPaterno.value = payload.apellidoPaterno;
            }
            if (typeof payload.apellidoMaterno === "string" && apellidoMaterno) {
                apellidoMaterno.value = payload.apellidoMaterno;
            }
            if (typeof payload.nombres === "string" && nombres) {
                nombres.value = payload.nombres;
            }
            if (typeof payload.direccion === "string" && direccion) {
                direccion.value = payload.direccion;
            }
            if (typeof payload.codigoUbigeo === "string") {
                await aplicarUbigeo(payload.codigoUbigeo);
            }

            mostrarMensajeConsulta("ok", "Documento consultado correctamente desde Migo.");
        } catch (error) {
            mostrarMensajeConsulta("error", error?.message || "No se pudo consultar el documento.");
        } finally {
            if (botonConsultarDocumento) {
                botonConsultarDocumento.disabled = false;
            }
        }
    };

    if (provincia) {
        provincia.dataset.sourceName = "codigoDepartamento";
    }
    if (distrito) {
        distrito.dataset.sourceName = "codigoProvincia";
    }

    if (tipoPersona) {
        tipoPersona.addEventListener("change", alternarTipoPersona);
        alternarTipoPersona();
    }

    botonConsultarDocumento?.addEventListener("click", function () {
        aplicarConsultaDocumento().catch(function () {
            mostrarMensajeConsulta("error", "No se pudo consultar el documento.");
        });
    });

    if (departamento && provincia) {
        departamento.addEventListener("change", async function () {
            await cargarSelect(provinciasUrl, departamento.value, provincia, null);
            poblarSelect(distrito, [], null);
        });
    }

    if (provincia && distrito) {
        provincia.addEventListener("change", async function () {
            await cargarSelect(distritosUrl, provincia.value, distrito, null);
        });
    }
});

document.addEventListener("DOMContentLoaded", function () {
    const overlayId = "global-loading-overlay";

    const ensureLoadingOverlay = function () {
        let overlay = document.getElementById(overlayId);
        if (overlay) {
            return overlay;
        }

        overlay = document.createElement("div");
        overlay.id = overlayId;
        overlay.className = "workspace-loading-overlay";
        overlay.setAttribute("aria-hidden", "true");
        overlay.innerHTML = `
            <div class="workspace-loading-box" role="status" aria-live="polite" aria-busy="true">
                <span class="workspace-loading-spinner"></span>
                <strong>Cargando...</strong>
                <small>Espere mientras se procesa la solicitud.</small>
            </div>`;

        document.body.appendChild(overlay);
        return overlay;
    };

    const showLoadingOverlay = function () {
        const overlay = ensureLoadingOverlay();
        overlay.classList.add("is-visible");
        overlay.setAttribute("aria-hidden", "false");
        document.body.classList.add("workspace-loading-active");
    };

    const hideLoadingOverlay = function () {
        const overlay = document.getElementById(overlayId);
        if (!overlay) {
            return;
        }

        overlay.classList.remove("is-visible");
        overlay.setAttribute("aria-hidden", "true");
        document.body.classList.remove("workspace-loading-active");
    };

    const disableSubmitControls = function (form) {
        form.querySelectorAll("button[type='submit'], input[type='submit']").forEach(function (element) {
            if (!(element instanceof HTMLButtonElement || element instanceof HTMLInputElement)) {
                return;
            }

            if (!element.dataset.originalDisabled) {
                element.dataset.originalDisabled = element.disabled ? "true" : "false";
            }

            element.disabled = true;
        });
    };

    document.querySelectorAll("form").forEach(function (form) {
        form.addEventListener("submit", function () {
            if (form.dataset.skipLoading === "true") {
                return;
            }

            if (typeof form.checkValidity === "function" && !form.checkValidity()) {
                hideLoadingOverlay();
                return;
            }

            window.setTimeout(function () {
                const jqueryForm = window.jQuery ? window.jQuery(form) : null;
                if (jqueryForm && typeof jqueryForm.valid === "function" && !jqueryForm.valid()) {
                    hideLoadingOverlay();
                    return;
                }

                showLoadingOverlay();
                disableSubmitControls(form);
            }, 0);
        });

        form.addEventListener("invalid", function () {
            hideLoadingOverlay();
        }, true);
    });

    document.querySelectorAll("a[data-loading-link='true']").forEach(function (link) {
        link.addEventListener("click", function () {
            showLoadingOverlay();
        });
    });

    document.querySelectorAll("form.workspace-filter-card[method='get']").forEach(function (form) {
        const anio = form.querySelector("select[name='anio']");
        const mes = form.querySelector("select[name='mes']");

        const submitFilter = function () {
            if (typeof form.requestSubmit === "function") {
                form.requestSubmit();
                return;
            }

            showLoadingOverlay();
            form.submit();
        };

        anio?.addEventListener("change", submitFilter);
        mes?.addEventListener("change", submitFilter);
    });
});

document.addEventListener("DOMContentLoaded", function () {
    const culture = "es-PE";
    const dateInputs = Array.from(document.querySelectorAll("input[type='date']"));
    if (dateInputs.length === 0) {
        return;
    }

    const monthFormatter = new Intl.DateTimeFormat(culture, { month: "long" });
    const yearFormatter = new Intl.DateTimeFormat(culture, { year: "numeric" });
    const weekdayFormatter = new Intl.DateTimeFormat(culture, { weekday: "short" });
    const weekdayReference = new Date(Date.UTC(2026, 6, 6));
    const weekdayLabels = Array.from({ length: 7 }, function (_, index) {
        const date = new Date(weekdayReference);
        date.setUTCDate(weekdayReference.getUTCDate() + index);
        return weekdayFormatter.format(date).replace(".", "").slice(0, 2).toUpperCase();
    });

    let activePicker = null;

    const parseValue = function (value) {
        if (!value || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
            return null;
        }

        const parts = value.split("-").map(Number);
        return new Date(parts[0], parts[1] - 1, parts[2]);
    };

    const formatDisplay = function (value) {
        const parsed = parseValue(value);
        if (!parsed) {
            return "";
        }

        const day = String(parsed.getDate()).padStart(2, "0");
        const month = String(parsed.getMonth() + 1).padStart(2, "0");
        const year = parsed.getFullYear();
        return `${day}/${month}/${year}`;
    };

    const formatValue = function (date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, "0");
        const day = String(date.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    };

    const compareDateOnly = function (left, right) {
        return left.getFullYear() === right.getFullYear()
            && left.getMonth() === right.getMonth()
            && left.getDate() === right.getDate();
    };

    const closeActivePicker = function () {
        if (!activePicker) {
            return;
        }

        activePicker.root.classList.remove("is-open");
        activePicker.panel.hidden = true;
        activePicker.trigger.setAttribute("aria-expanded", "false");
        activePicker = null;
    };

    const positionPickerPanel = function (picker) {
        if (!picker || picker.panel.hidden) {
            return;
        }

        const rootRect = picker.root.getBoundingClientRect();
        const margin = 12;
        const preferredWidth = Math.max(rootRect.width, 320);

        picker.panel.style.width = `${preferredWidth}px`;
        picker.panel.style.minWidth = `${Math.min(preferredWidth, 320)}px`;
        picker.panel.style.maxWidth = `${Math.max(preferredWidth, 320)}px`;

        const panelRect = picker.panel.getBoundingClientRect();
        const availableBelow = window.innerHeight - rootRect.bottom - margin;
        const availableAbove = rootRect.top - margin;
        const openUpwards = panelRect.height > availableBelow && availableAbove > availableBelow;

        let top = openUpwards
            ? rootRect.top - panelRect.height - 8
            : rootRect.bottom + 8;

        let left = rootRect.left;
        const maxLeft = window.innerWidth - panelRect.width - margin;
        left = Math.min(Math.max(left, margin), Math.max(margin, maxLeft));

        if (top < margin) {
            top = margin;
        }

        picker.panel.style.top = `${top}px`;
        picker.panel.style.left = `${left}px`;
    };

    const renderCalendar = function (picker) {
        const titleStrong = picker.panel.querySelector("[data-calendar-title]");
        const titleSpan = picker.panel.querySelector("[data-calendar-year]");
        const grid = picker.panel.querySelector("[data-calendar-grid]");
        const selectedDate = parseValue(picker.original.value);
        const minDate = parseValue(picker.original.min);
        const maxDate = parseValue(picker.original.max);
        const today = new Date();

        titleStrong.textContent = monthFormatter.format(picker.viewDate).replace(/^\w/, function (char) {
            return char.toUpperCase();
        });
        titleSpan.textContent = yearFormatter.format(picker.viewDate);
        grid.innerHTML = "";

        const start = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth(), 1);
        const end = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth() + 1, 0);
        const leadingDays = (start.getDay() + 6) % 7;
        const totalCells = 42;

        for (let index = 0; index < totalCells; index += 1) {
            const dayDate = new Date(start);
            dayDate.setDate(start.getDate() - leadingDays + index);

            const button = document.createElement("button");
            button.type = "button";
            button.className = "app-date-picker-day";
            button.textContent = String(dayDate.getDate());
            button.dataset.value = formatValue(dayDate);

            if (dayDate.getMonth() !== picker.viewDate.getMonth()) {
                button.classList.add("is-other-month");
            }

            if (compareDateOnly(dayDate, today)) {
                button.classList.add("is-today");
            }

            if (selectedDate && compareDateOnly(dayDate, selectedDate)) {
                button.classList.add("is-selected");
            }

            if ((minDate && dayDate < minDate) || (maxDate && dayDate > maxDate)) {
                button.disabled = true;
            }

            button.addEventListener("click", function () {
                picker.original.value = button.dataset.value;
                picker.display.value = formatDisplay(button.dataset.value);
                picker.original.dispatchEvent(new Event("input", { bubbles: true }));
                picker.original.dispatchEvent(new Event("change", { bubbles: true }));
                closeActivePicker();
            });

            grid.appendChild(button);
        }
    };

    const openPicker = function (picker) {
        if (picker.disabled) {
            return;
        }

        if (activePicker && activePicker !== picker) {
            closeActivePicker();
        }

        const selectedDate = parseValue(picker.original.value);
        picker.viewDate = selectedDate
            ? new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1)
            : new Date(new Date().getFullYear(), new Date().getMonth(), 1);

        renderCalendar(picker);
        picker.root.classList.add("is-open");
        picker.panel.hidden = false;
        picker.trigger.setAttribute("aria-expanded", "true");
        activePicker = picker;
        positionPickerPanel(picker);
    };

    const createPicker = function (original) {
        if (original.dataset.calendarEnhanced === "true") {
            return;
        }

        original.dataset.calendarEnhanced = "true";

        const originalId = original.id || `date-${Math.random().toString(36).slice(2, 8)}`;
        const displayId = `${originalId}__display`;
        const wrapper = document.createElement("div");
        wrapper.className = "app-date-picker";

        const control = document.createElement("div");
        control.className = `app-date-picker-control${original.disabled ? " is-disabled" : ""}`;

        const display = document.createElement("input");
        display.type = "text";
        display.className = "form-control app-date-picker-display";
        display.id = displayId;
        display.readOnly = true;
        display.placeholder = "Seleccione fecha";
        display.value = formatDisplay(original.value);
        display.disabled = original.disabled;
        display.autocomplete = "off";

        const trigger = document.createElement("button");
        trigger.type = "button";
        trigger.className = "app-date-picker-trigger";
        trigger.innerHTML = "<i class='bi bi-calendar3'></i>";
        trigger.setAttribute("aria-label", "Abrir calendario");
        trigger.setAttribute("aria-expanded", "false");
        trigger.disabled = original.disabled;

        const panel = document.createElement("div");
        panel.className = "app-date-picker-panel";
        panel.hidden = true;
        panel.innerHTML = `
            <div class="app-date-picker-header">
                <button type="button" class="app-date-picker-nav" data-calendar-prev aria-label="Mes anterior">
                    <i class="bi bi-chevron-left"></i>
                </button>
                <div class="app-date-picker-title">
                    <strong data-calendar-title></strong>
                    <span data-calendar-year></span>
                </div>
                <button type="button" class="app-date-picker-nav" data-calendar-next aria-label="Mes siguiente">
                    <i class="bi bi-chevron-right"></i>
                </button>
            </div>
            <div class="app-date-picker-weekdays">${weekdayLabels.map(label => `<span>${label}</span>`).join("")}</div>
            <div class="app-date-picker-grid" data-calendar-grid></div>
            <div class="app-date-picker-footer">
                <button type="button" data-calendar-clear>Limpiar</button>
                <button type="button" data-calendar-today>Hoy</button>
            </div>`;

        original.parentNode.insertBefore(wrapper, original);
        wrapper.appendChild(control);
        control.appendChild(display);
        control.appendChild(trigger);
        wrapper.appendChild(original);
        document.body.appendChild(panel);

        original.classList.add("app-date-picker-native");
        original.tabIndex = -1;

        if (original.id) {
            document.querySelectorAll(`label[for='${originalId}']`).forEach(function (label) {
                label.setAttribute("for", display.id);
            });
        }

        const picker = {
            root: wrapper,
            original: original,
            display: display,
            trigger: trigger,
            panel: panel,
            viewDate: parseValue(original.value) || new Date(),
            disabled: original.disabled
        };

        display.addEventListener("click", function () {
            if (activePicker === picker) {
                closeActivePicker();
                return;
            }

            openPicker(picker);
        });

        trigger.addEventListener("click", function () {
            if (activePicker === picker) {
                closeActivePicker();
                return;
            }

            openPicker(picker);
        });

        panel.querySelector("[data-calendar-prev]").addEventListener("click", function () {
            picker.viewDate = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth() - 1, 1);
            renderCalendar(picker);
        });

        panel.querySelector("[data-calendar-next]").addEventListener("click", function () {
            picker.viewDate = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth() + 1, 1);
            renderCalendar(picker);
        });

        panel.querySelector("[data-calendar-clear]").addEventListener("click", function () {
            if (picker.original.required) {
                return;
            }

            picker.original.value = "";
            picker.display.value = "";
            picker.original.dispatchEvent(new Event("input", { bubbles: true }));
            picker.original.dispatchEvent(new Event("change", { bubbles: true }));
            closeActivePicker();
        });

        panel.querySelector("[data-calendar-today]").addEventListener("click", function () {
            const today = new Date();
            picker.original.value = formatValue(today);
            picker.display.value = formatDisplay(picker.original.value);
            picker.original.dispatchEvent(new Event("input", { bubbles: true }));
            picker.original.dispatchEvent(new Event("change", { bubbles: true }));
            closeActivePicker();
        });

        original.addEventListener("change", function () {
            display.value = formatDisplay(original.value);
        });
    };

    dateInputs.forEach(createPicker);

    document.addEventListener("click", function (event) {
        if (!activePicker) {
            return;
        }

        if (activePicker.root.contains(event.target) || activePicker.panel.contains(event.target)) {
            return;
        }

        closeActivePicker();
    });

    window.addEventListener("resize", function () {
        if (activePicker) {
            positionPickerPanel(activePicker);
        }
    });

    window.addEventListener("scroll", function () {
        if (activePicker) {
            positionPickerPanel(activePicker);
        }
    }, true);

    document.addEventListener("keydown", function (event) {
        if (event.key === "Escape") {
            closeActivePicker();
        }
    });
});
