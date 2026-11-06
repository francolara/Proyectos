// Firma: FRANCO LARA - 05/11/2026 | Cabecera HYT y posicionamiento limitado a la ventana en todos los calendarios.
// Firma: FRANCO LARA - 03/11/2026 | Calendario FralseCont con limpieza de paneles y eventos para formularios emergentes.
const hytDateScopes=new WeakMap();
window.hytDisposeDates=function(root){const scope=hytDateScopes.get(root);if(scope){scope.controller.abort();scope.panels.forEach(panel=>panel.remove());hytDateScopes.delete(root);}};
window.hytInitDates = function (root = document) {
    window.hytDisposeDates(root);const scope={controller:new AbortController(),panels:[]};hytDateScopes.set(root,scope);
    const culture = "es-PE";
    const todayValue = () => {
        const parts=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:'America/Lima',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date()).map(p=>[p.type,p.value]));
        return new Date(+parts.year,+parts.month-1,+parts.day);
    };
    const dateInputs = Array.from(root.querySelectorAll("input[type='date']"));
    if (dateInputs.length === 0) {
        return;
    }

    const monthFormatter = new Intl.DateTimeFormat(culture, { month: "long" });
    const yearFormatter = new Intl.DateTimeFormat(culture, { year: "numeric" });
    const weekdayFormatter = new Intl.DateTimeFormat(culture, { weekday: "short", timeZone: "UTC" });
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

    const closeActivePicker = function (restoreFocus = false) {
        if (!activePicker) {
            return;
        }

        activePicker.root.classList.remove("is-open");
        activePicker.panel.hidden = true;
        activePicker.trigger.setAttribute("aria-expanded", "false");
        if (restoreFocus) activePicker.trigger.focus();
        activePicker = null;
    };

    const positionPickerPanel = function (picker) {
        if (!picker || picker.panel.hidden) {
            return;
        }

        const rootRect = picker.root.getBoundingClientRect();
        const margin = 12;
        const preferredWidth = Math.min(Math.max(rootRect.width, 360), window.innerWidth - margin * 2);

        picker.panel.style.width = `${preferredWidth}px`;
        picker.panel.style.minWidth = `${preferredWidth}px`;
        picker.panel.style.maxWidth = `${preferredWidth}px`;

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

        top = Math.min(Math.max(top, margin), Math.max(margin, window.innerHeight - panelRect.height - margin));
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
        const today = todayValue();

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
            button.setAttribute("aria-label", dayDate.toLocaleDateString(culture, {dateStyle:"full"}));

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
                closeActivePicker(true);
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
            : new Date(todayValue().getFullYear(), todayValue().getMonth(), 1);

        renderCalendar(picker);
        picker.root.classList.add("is-open");
        picker.panel.hidden = false;
        picker.trigger.setAttribute("aria-expanded", "true");
        activePicker = picker;
        positionPickerPanel(picker);
        picker.panel.querySelector(".is-selected:not(:disabled), .app-date-picker-day:not(:disabled)")?.focus();
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
        control.className = `app-date-picker-control${(original.disabled || original.readOnly) ? " is-disabled" : ""}`;

        const display = document.createElement("input");
        display.type = "text";
        display.className = "form-control app-date-picker-display";
        display.id = displayId;
        display.readOnly = true;
        display.placeholder = "Seleccione fecha";
        display.value = formatDisplay(original.value);
        display.disabled = (original.disabled || original.readOnly);
        display.autocomplete = "off";

        const trigger = document.createElement("button");
        trigger.type = "button";
        trigger.className = "app-date-picker-trigger";
        trigger.innerHTML = "<svg width='20' height='20' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.6' aria-hidden='true'><rect x='3' y='5' width='18' height='16' rx='2'/><path d='M7 3v4m10-4v4M3 11h18'/></svg>";
        trigger.setAttribute("aria-label", "Abrir calendario");
        trigger.setAttribute("aria-expanded", "false");
        trigger.disabled = (original.disabled || original.readOnly);

        const panel = document.createElement("div");
        panel.className = "app-date-picker-panel";
        panel.hidden = true;
        panel.setAttribute("role", "dialog");
        panel.setAttribute("aria-label", "Calendario");
        panel.innerHTML = `
            <div class="app-date-picker-banner">Seleccionar fecha</div>
            <div class="app-date-picker-header">
                <button type="button" class="app-date-picker-nav" data-calendar-prev aria-label="Mes anterior">
                    <span aria-hidden="true">‹</span>
                </button>
                <div class="app-date-picker-title">
                    <strong data-calendar-title></strong>
                    <span data-calendar-year></span>
                </div>
                <button type="button" class="app-date-picker-nav" data-calendar-next aria-label="Mes siguiente">
                    <span aria-hidden="true">›</span>
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
        (original.closest("dialog") || document.body).appendChild(panel);scope.panels.push(panel);

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
            viewDate: parseValue(original.value) || todayValue(),
            disabled: (original.disabled || original.readOnly)
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
            closeActivePicker(true);
        });

        panel.querySelector("[data-calendar-today]").addEventListener("click", function () {
            const today = todayValue();
            picker.original.value = formatValue(today);
            picker.display.value = formatDisplay(picker.original.value);
            picker.original.dispatchEvent(new Event("input", { bubbles: true }));
            picker.original.dispatchEvent(new Event("change", { bubbles: true }));
            closeActivePicker(true);
        });

        original.addEventListener('invalid', event=>{event.preventDefault();display.setAttribute('aria-invalid','true');display.focus();});
        original.addEventListener("change", function () {
            display.removeAttribute('aria-invalid');
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
    }, {signal:scope.controller.signal});

    window.addEventListener("resize", function () {
        if (activePicker) {
            positionPickerPanel(activePicker);
        }
    }, {signal:scope.controller.signal});

    window.addEventListener("scroll", function () {
        if (activePicker) {
            positionPickerPanel(activePicker);
        }
    }, {capture:true,signal:scope.controller.signal});

    document.addEventListener("keydown", function (event) {
        if (event.key === "Escape") {
            closeActivePicker(true);
        }
    }, {signal:scope.controller.signal});
};
document.addEventListener("DOMContentLoaded", () => window.hytInitDates());
