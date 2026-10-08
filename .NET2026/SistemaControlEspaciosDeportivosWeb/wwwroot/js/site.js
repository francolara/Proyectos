// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
(function () {
    const configurarMensajesValidacion = function () {
        if (!window.jQuery || !jQuery.validator) return;

        jQuery.extend(jQuery.validator.messages, {
            required: "Este campo es obligatorio.",
            remote: "Corrige este campo.",
            email: "Ingresa un correo electronico valido.",
            url: "Ingresa una URL valida.",
            date: "Ingresa una fecha valida.",
            dateISO: "Ingresa una fecha valida (ISO).",
            number: "Ingresa un numero valido.",
            digits: "Ingresa solo digitos.",
            creditcard: "Ingresa una tarjeta valida.",
            equalTo: "Los valores no coinciden.",
            extension: "Ingresa un valor con una extension valida.",
            maxlength: jQuery.validator.format("Ingresa como maximo {0} caracteres."),
            minlength: jQuery.validator.format("Ingresa al menos {0} caracteres."),
            rangelength: jQuery.validator.format("Ingresa un valor entre {0} y {1} caracteres."),
            range: jQuery.validator.format("Ingresa un valor entre {0} y {1}."),
            max: jQuery.validator.format("Ingresa un valor menor o igual a {0}."),
            min: jQuery.validator.format("Ingresa un valor mayor o igual a {0}.")
        });
    };

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", configurarMensajesValidacion, { once: true });
    } else {
        configurarMensajesValidacion();
    }
})();

/* Firma: FRANCO LARA - 02/10/2026 | Centraliza en español los mensajes de validación nativos y de jQuery para todos los formularios del sistema. */
document.addEventListener("DOMContentLoaded", function () {
    const marcaMensajePropio = "mensajeValidacionEspanol";
    const mensajes = {
        valueMissing: "Este campo es obligatorio.",
        typeMismatch: "Ingresa un valor con el formato correcto.",
        patternMismatch: "El formato ingresado no es valido.",
        tooLong: "Reduce la cantidad de caracteres.",
        tooShort: "Ingresa mas caracteres.",
        rangeUnderflow: "Ingresa un valor mayor o igual al permitido.",
        rangeOverflow: "Ingresa un valor menor o igual al permitido.",
        stepMismatch: "Ingresa un valor permitido.",
        badInput: "Ingresa un valor valido."
    };

    const obtenerMensaje = function (control) {
        const validity = control.validity;
        if (!validity) return null;

        return Object.keys(mensajes).find(function (regla) {
            return validity[regla];
        }) || null;
    };

    document.addEventListener("invalid", function (event) {
        const control = event.target;
        if (!(control instanceof HTMLInputElement || control instanceof HTMLSelectElement || control instanceof HTMLTextAreaElement)) return;

        const regla = obtenerMensaje(control);
        if (!regla) return;

        control.setCustomValidity(mensajes[regla]);
        control.dataset[marcaMensajePropio] = "true";
    }, true);

    const limpiarMensaje = function (event) {
        const control = event.target;
        if (!(control instanceof HTMLInputElement || control instanceof HTMLSelectElement || control instanceof HTMLTextAreaElement)) return;
        if (control.dataset[marcaMensajePropio] !== "true") return;

        control.setCustomValidity("");
        delete control.dataset[marcaMensajePropio];
    };

    document.addEventListener("input", limpiarMensaje, true);
    document.addEventListener("change", limpiarMensaje, true);
});

(function () {
    function normalizar(texto) {
        return (texto || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim();
    }

    function resolverMetaKpi(etiqueta) {
        const t = normalizar(etiqueta);

        if (/(ingreso|cobranza|cobrado|monto|pago|ticket|saldo|recaud)/.test(t)) {
            return "kpi-tone-green";
        }
        if (/(pendiente|alerta|vencer|vencid|no show|cancelad|inactivo|mantenimiento|anulad)/.test(t)) {
            return "kpi-tone-amber";
        }
        if (/(critico|rechazad|error|caid|bloque|sin ingreso)/.test(t)) {
            return "kpi-tone-red";
        }
        if (/(cliente|usuario|equipo)/.test(t)) {
            return "kpi-tone-blue";
        }
        if (/(reserva|dia|fecha|periodo|ocupacion|sede|espacio)/.test(t)) {
            return "kpi-tone-blue";
        }
        return "kpi-tone-blue";
    }

    function estilizarKpis() {
        const cards = document.querySelectorAll(".kpi-card, .sc-kpi-card, .sc-dash-kpi-card, .sc-kpi-item, .sc-reservas-metric");
        cards.forEach((card) => {
            const labelNode = card.querySelector(".kpi-card-label, .sc-dash-kpi-label, p, span");
            if (!labelNode) return;

            const tone = resolverMetaKpi(labelNode.textContent || "");
            card.classList.remove("kpi-tone-blue", "kpi-tone-green", "kpi-tone-amber", "kpi-tone-red");
            card.classList.add(tone);
        });
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", estilizarKpis);
    } else {
        estilizarKpis();
    }
})();

// Firma: FRANCO LARA - 02/09/2026 | Conserva el color semantico de los KPI sin insertar iconos decorativos en las tarjetas administrativas.

// Firma: FRANCO LARA - 08/07/2026 | Replica modo oscuro administrativo y ventana de carga del sistema administrativo para navegacion interna, formularios y acciones principales del panel.
// Firma: FRANCO LARA - 15/07/2026 | Ventana de carga global omite submits interceptados por AJAX para evitar overlays bloqueados en modales administrativos.
document.addEventListener("DOMContentLoaded", function () {
    const body = document.body;
    const root = document.documentElement;
    if (!body || !body.classList.contains("sc-admin-theme-shell")) {
        return;
    }

    const storageKey = "sc-admin-theme";
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

/* Selector visual comun para controles de fecha. Mantiene el input date original para el model binding. */
document.addEventListener("DOMContentLoaded", function () {
    const selector = "input[type='date']:not([data-date-picker='native'])";
    const culture = "es-PE";
    const pickers = new Set();
    let activePicker = null;

    const monthFormatter = new Intl.DateTimeFormat(culture, { month: "long" });
    const yearFormatter = new Intl.DateTimeFormat(culture, { year: "numeric" });
    const weekdayFormatter = new Intl.DateTimeFormat(culture, { weekday: "short" });
    const fullDateFormatter = new Intl.DateTimeFormat(culture, {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric"
    });
    const weekdayReference = new Date(Date.UTC(2026, 6, 6));
    const weekdayLabels = Array.from({ length: 7 }, function (_, index) {
        const date = new Date(weekdayReference);
        date.setUTCDate(weekdayReference.getUTCDate() + index);
        return weekdayFormatter.format(date).replace(".", "").slice(0, 2).toUpperCase();
    });

    const parseValue = function (value) {
        if (!value || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
            return null;
        }

        const parts = value.split("-").map(Number);
        const parsed = new Date(parts[0], parts[1] - 1, parts[2]);
        if (parsed.getFullYear() !== parts[0]
            || parsed.getMonth() !== parts[1] - 1
            || parsed.getDate() !== parts[2]) {
            return null;
        }

        return parsed;
    };

    const formatValue = function (date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, "0");
        const day = String(date.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    };

    const formatDisplay = function (value) {
        const parsed = parseValue(value);
        if (!parsed) {
            return "";
        }

        return `${String(parsed.getDate()).padStart(2, "0")}/${String(parsed.getMonth() + 1).padStart(2, "0")}/${parsed.getFullYear()}`;
    };

    const compareDateOnly = function (left, right) {
        return left.getFullYear() === right.getFullYear()
            && left.getMonth() === right.getMonth()
            && left.getDate() === right.getDate();
    };

    const isWithinBounds = function (date, original) {
        const minDate = parseValue(original.min);
        const maxDate = parseValue(original.max);
        const candidate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        return (!minDate || candidate >= minDate) && (!maxDate || candidate <= maxDate);
    };

    const isBlocked = function (picker) {
        return picker.original.disabled || picker.original.readOnly;
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
        let top = openUpwards ? rootRect.top - panelRect.height - 8 : rootRect.bottom + 8;
        let left = rootRect.left;

        left = Math.min(Math.max(left, margin), Math.max(margin, window.innerWidth - panelRect.width - margin));
        top = Math.max(top, margin);
        picker.panel.style.top = `${top}px`;
        picker.panel.style.left = `${left}px`;
    };

    const closeActivePicker = function (restoreFocus) {
        if (!activePicker) {
            return;
        }

        const picker = activePicker;
        picker.root.classList.remove("is-open");
        picker.panel.hidden = true;
        picker.trigger.setAttribute("aria-expanded", "false");
        activePicker = null;

        if (restoreFocus) {
            picker.display.focus({ preventScroll: true });
        }
    };

    const renderCalendar = function (picker) {
        const titleStrong = picker.panel.querySelector("[data-calendar-title]");
        const titleSpan = picker.panel.querySelector("[data-calendar-year]");
        const grid = picker.panel.querySelector("[data-calendar-grid]");
        const selectedDate = parseValue(picker.original.value);
        const minDate = parseValue(picker.original.min);
        const maxDate = parseValue(picker.original.max);
        const today = new Date();

        titleStrong.textContent = monthFormatter.format(picker.viewDate).replace(/^\w/, function (character) {
            return character.toUpperCase();
        });
        titleSpan.textContent = yearFormatter.format(picker.viewDate);
        grid.innerHTML = "";

        const start = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth(), 1);
        const leadingDays = (start.getDay() + 6) % 7;

        for (let index = 0; index < 42; index += 1) {
            const dayDate = new Date(start);
            dayDate.setDate(start.getDate() - leadingDays + index);

            const button = document.createElement("button");
            button.type = "button";
            button.className = "app-date-picker-day";
            button.textContent = String(dayDate.getDate());
            button.dataset.value = formatValue(dayDate);
            button.setAttribute("aria-label", fullDateFormatter.format(dayDate));

            if (dayDate.getMonth() !== picker.viewDate.getMonth()) {
                button.classList.add("is-other-month");
            }
            if (compareDateOnly(dayDate, today)) {
                button.classList.add("is-today");
            }
            if (selectedDate && compareDateOnly(dayDate, selectedDate)) {
                button.classList.add("is-selected");
                button.setAttribute("aria-current", "date");
            }
            if (isBlocked(picker) || (minDate && dayDate < minDate) || (maxDate && dayDate > maxDate)) {
                button.disabled = true;
            }

            button.addEventListener("click", function () {
                picker.original.value = button.dataset.value;
                picker.original.dispatchEvent(new Event("input", { bubbles: true }));
                picker.original.dispatchEvent(new Event("change", { bubbles: true }));
                closeActivePicker(true);
            });

            grid.appendChild(button);
        }
    };

    const syncPicker = function (picker, rerender) {
        if (!picker.original.isConnected) {
            picker.panel.remove();
            pickers.delete(picker);
            return;
        }

        const blocked = isBlocked(picker);
        picker.display.value = formatDisplay(picker.original.value);
        picker.display.disabled = picker.original.disabled;
        picker.display.setAttribute("aria-disabled", blocked ? "true" : "false");
        picker.trigger.disabled = blocked;
        picker.control.classList.toggle("is-disabled", picker.original.disabled);
        picker.control.classList.toggle("is-readonly", picker.original.readOnly);
        picker.display.classList.toggle("is-invalid", picker.original.classList.contains("is-invalid") || picker.original.classList.contains("input-validation-error"));

        const clearButton = picker.panel.querySelector("[data-calendar-clear]");
        clearButton.hidden = picker.original.required;
        clearButton.disabled = blocked;

        const todayButton = picker.panel.querySelector("[data-calendar-today]");
        todayButton.disabled = blocked || !isWithinBounds(new Date(), picker.original);

        if (blocked && activePicker === picker) {
            closeActivePicker(false);
        } else if (rerender && activePicker === picker) {
            renderCalendar(picker);
            positionPickerPanel(picker);
        }
    };

    const syncAllPickers = function () {
        pickers.forEach(function (picker) {
            syncPicker(picker, false);
        });
    };

    const openPicker = function (picker) {
        syncPicker(picker, false);
        if (isBlocked(picker)) {
            return;
        }

        if (activePicker && activePicker !== picker) {
            closeActivePicker(false);
        }

        const selectedDate = parseValue(picker.original.value);
        const baseDate = selectedDate || new Date();
        picker.viewDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
        renderCalendar(picker);
        picker.root.classList.add("is-open");
        picker.panel.hidden = false;
        picker.trigger.setAttribute("aria-expanded", "true");
        activePicker = picker;
        positionPickerPanel(picker);
    };

    const createPicker = function (original) {
        if (!(original instanceof HTMLInputElement) || original.dataset.calendarEnhanced === "true") {
            return;
        }

        original.dataset.calendarEnhanced = "true";
        const originalId = original.id || `date-${Math.random().toString(36).slice(2, 10)}`;
        const displayId = `${originalId}__display`;
        const panelId = `${originalId}__calendar`;
        const wrapper = document.createElement("div");
        wrapper.className = `app-date-picker${original.classList.contains("form-control-sm") ? " is-compact" : ""}`;

        const control = document.createElement("div");
        control.className = "app-date-picker-control";

        const display = document.createElement("input");
        display.type = "text";
        display.className = `form-control app-date-picker-display${original.classList.contains("form-control-sm") ? " form-control-sm" : ""}`;
        display.id = displayId;
        display.readOnly = true;
        display.placeholder = original.placeholder || "Seleccione fecha";
        display.autocomplete = "off";
        display.setAttribute("inputmode", "none");
        display.setAttribute("aria-haspopup", "dialog");
        display.setAttribute("aria-controls", panelId);
        if (original.getAttribute("aria-describedby")) {
            display.setAttribute("aria-describedby", original.getAttribute("aria-describedby"));
        }

        const trigger = document.createElement("button");
        trigger.type = "button";
        trigger.className = "app-date-picker-trigger";
        trigger.innerHTML = "<i class='bi bi-calendar3' aria-hidden='true'></i>";
        trigger.setAttribute("aria-label", "Abrir calendario");
        trigger.setAttribute("aria-expanded", "false");
        trigger.setAttribute("aria-controls", panelId);

        const panel = document.createElement("div");
        panel.className = "app-date-picker-panel";
        panel.id = panelId;
        panel.hidden = true;
        panel.setAttribute("role", "dialog");
        panel.setAttribute("aria-label", "Seleccionar fecha");
        panel.innerHTML = `
            <div class="app-date-picker-header">
                <button type="button" class="app-date-picker-nav" data-calendar-prev aria-label="Mes anterior">
                    <i class="bi bi-chevron-left" aria-hidden="true"></i>
                </button>
                <div class="app-date-picker-title">
                    <strong data-calendar-title></strong>
                    <span data-calendar-year></span>
                </div>
                <button type="button" class="app-date-picker-nav" data-calendar-next aria-label="Mes siguiente">
                    <i class="bi bi-chevron-right" aria-hidden="true"></i>
                </button>
            </div>
            <div class="app-date-picker-weekdays" aria-hidden="true">${weekdayLabels.map(function (label) { return `<span>${label}</span>`; }).join("")}</div>
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
        original.setAttribute("aria-hidden", "true");

        if (original.id) {
            document.querySelectorAll(`label[for='${CSS.escape(originalId)}']`).forEach(function (label) {
                label.setAttribute("for", displayId);
            });
        }

        const picker = {
            root: wrapper,
            control: control,
            original: original,
            display: display,
            trigger: trigger,
            panel: panel,
            viewDate: parseValue(original.value) || new Date()
        };
        pickers.add(picker);

        const togglePicker = function () {
            if (activePicker === picker) {
                closeActivePicker(false);
            } else {
                openPicker(picker);
            }
        };

        display.addEventListener("click", togglePicker);
        display.addEventListener("keydown", function (event) {
            if (["Enter", " ", "ArrowDown"].includes(event.key)) {
                event.preventDefault();
                openPicker(picker);
            }
        });
        trigger.addEventListener("click", togglePicker);

        panel.querySelector("[data-calendar-prev]").addEventListener("click", function () {
            picker.viewDate = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth() - 1, 1);
            renderCalendar(picker);
        });
        panel.querySelector("[data-calendar-next]").addEventListener("click", function () {
            picker.viewDate = new Date(picker.viewDate.getFullYear(), picker.viewDate.getMonth() + 1, 1);
            renderCalendar(picker);
        });
        panel.querySelector("[data-calendar-clear]").addEventListener("click", function () {
            if (picker.original.required || isBlocked(picker)) {
                return;
            }

            picker.original.value = "";
            picker.original.dispatchEvent(new Event("input", { bubbles: true }));
            picker.original.dispatchEvent(new Event("change", { bubbles: true }));
            closeActivePicker(true);
        });
        panel.querySelector("[data-calendar-today]").addEventListener("click", function () {
            const today = new Date();
            if (isBlocked(picker) || !isWithinBounds(today, picker.original)) {
                return;
            }

            picker.original.value = formatValue(today);
            picker.original.dispatchEvent(new Event("input", { bubbles: true }));
            picker.original.dispatchEvent(new Event("change", { bubbles: true }));
            closeActivePicker(true);
        });

        original.addEventListener("input", function () { syncPicker(picker, true); });
        original.addEventListener("change", function () { syncPicker(picker, true); });
        original.addEventListener("invalid", function () {
            syncPicker(picker, false);
            display.focus({ preventScroll: false });
        });

        const stateObserver = new MutationObserver(function () {
            syncPicker(picker, true);
        });
        stateObserver.observe(original, {
            attributes: true,
            attributeFilter: ["class", "disabled", "readonly", "required", "min", "max", "value"]
        });

        original.form?.addEventListener("reset", function () {
            window.setTimeout(function () { syncPicker(picker, true); }, 0);
        });

        syncPicker(picker, false);
    };

    document.querySelectorAll(selector).forEach(createPicker);

    const documentObserver = new MutationObserver(function (mutations) {
        mutations.forEach(function (mutation) {
            mutation.addedNodes.forEach(function (node) {
                if (!(node instanceof Element)) {
                    return;
                }

                if (node.matches(selector)) {
                    createPicker(node);
                }
                node.querySelectorAll(selector).forEach(createPicker);
            });
        });
    });
    documentObserver.observe(document.body, { childList: true, subtree: true });

    document.addEventListener("click", function (event) {
        if (activePicker && !activePicker.root.contains(event.target) && !activePicker.panel.contains(event.target)) {
            closeActivePicker(false);
        }

        queueMicrotask(syncAllPickers);
    });
    document.addEventListener("change", function () { queueMicrotask(syncAllPickers); });
    document.addEventListener("shown.bs.modal", syncAllPickers);

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
    window.addEventListener("pageshow", syncAllPickers);
    document.addEventListener("keydown", function (event) {
        if (event.key === "Escape") {
            closeActivePicker(true);
        }
    });

    window.AppDatePicker = {
        refresh: syncAllPickers
    };
});

document.addEventListener("DOMContentLoaded", function () {
    const body = document.body;
    if (!body || !body.classList.contains("sc-admin-theme-shell")) {
        return;
    }

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
        overlay.innerHTML = [
            "<div class=\"workspace-loading-box\" role=\"status\" aria-live=\"polite\" aria-busy=\"true\">",
            "  <span class=\"workspace-loading-spinner\"></span>",
            "  <strong>Cargando...</strong>",
            "  <small>Espere mientras se procesa la solicitud.</small>",
            "</div>"
        ].join("");

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

            element.disabled = true;
        });
    };

    const sameOriginNavigation = function (href) {
        if (!href || href.startsWith("#") || href.startsWith("javascript:")) {
            return false;
        }

        try {
            const targetUrl = new URL(href, window.location.href);
            if (targetUrl.origin !== window.location.origin) {
                return false;
            }

            const currentWithoutHash = `${window.location.pathname}${window.location.search}`;
            const targetWithoutHash = `${targetUrl.pathname}${targetUrl.search}`;
            return currentWithoutHash !== targetWithoutHash || !targetUrl.hash;
        } catch {
            return false;
        }
    };

    const shouldHandleLink = function (link) {
        if (!(link instanceof HTMLAnchorElement)) {
            return false;
        }

        if (link.target === "_blank" || link.hasAttribute("download") || link.dataset.skipLoading === "true") {
            return false;
        }

        if (link.hasAttribute("data-bs-toggle") || link.hasAttribute("data-bs-dismiss")) {
            return false;
        }

        return sameOriginNavigation(link.getAttribute("href"));
    };

    const shouldHandleActionButton = function (button) {
        if (!(button instanceof HTMLButtonElement || button instanceof HTMLInputElement)) {
            return false;
        }

        if (button.disabled || button.dataset.skipLoading === "true") {
            return false;
        }

        if (button.type === "submit" || button.form) {
            return false;
        }

        if (button.hasAttribute("data-bs-toggle") || button.hasAttribute("data-bs-dismiss")) {
            return false;
        }

        const sourceText = button instanceof HTMLInputElement
            ? (button.value || "")
            : (button.textContent || "");
        const normalized = sourceText
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim();

        return /(grabar|guardar|listar|consultar|buscar|filtrar|procesar|generar)/.test(normalized);
    };

    document.querySelectorAll("form").forEach(function (form) {
        form.addEventListener("submit", function (event) {
            if (form.dataset.skipLoading === "true") {
                return;
            }

            if (typeof form.checkValidity === "function" && !form.checkValidity()) {
                hideLoadingOverlay();
                return;
            }

            window.setTimeout(function () {
                if (event.defaultPrevented) {
                    hideLoadingOverlay();
                    return;
                }

                if (window.jQuery) {
                    const jqueryForm = window.jQuery(form);
                    if (typeof jqueryForm.valid === "function" && !jqueryForm.valid()) {
                        hideLoadingOverlay();
                        return;
                    }
                }

                showLoadingOverlay();
                disableSubmitControls(form);
            }, 0);
        });

        form.addEventListener("invalid", function () {
            hideLoadingOverlay();
        }, true);
    });

    document.querySelectorAll(".sc-admin-sidebar a[href], .sc-admin-main a[href]").forEach(function (link) {
        link.addEventListener("click", function () {
            if (!shouldHandleLink(link)) {
                return;
            }

            showLoadingOverlay();
        });
    });

    document.querySelectorAll(".sc-admin-main .btn, .sc-admin-sidebar .btn").forEach(function (button) {
        button.addEventListener("click", function () {
            if (!shouldHandleActionButton(button)) {
                return;
            }

            showLoadingOverlay();
        });
    });

    window.addEventListener("pageshow", function () {
        hideLoadingOverlay();
    });
});
