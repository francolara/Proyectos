// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
(function () {
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
})();

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
            return { tone: "kpi-tone-green", icon: "bi-cash-stack" };
        }
        if (/(pendiente|alerta|vencer|vencid|no show|cancelad|inactivo|mantenimiento|anulad)/.test(t)) {
            return { tone: "kpi-tone-amber", icon: "bi-exclamation-triangle" };
        }
        if (/(critico|rechazad|error|caid|bloque|sin ingreso)/.test(t)) {
            return { tone: "kpi-tone-red", icon: "bi-shield-exclamation" };
        }
        if (/(cliente|usuario|equipo)/.test(t)) {
            return { tone: "kpi-tone-blue", icon: "bi-people" };
        }
        if (/(reserva|dia|fecha|periodo|ocupacion|sede|espacio)/.test(t)) {
            return { tone: "kpi-tone-blue", icon: "bi-calendar3" };
        }
        return { tone: "kpi-tone-blue", icon: "bi-bar-chart-line" };
    }

    function estilizarKpis() {
        const cards = document.querySelectorAll(".kpi-card, .sc-kpi-card, .sc-dash-kpi-card, .sc-kpi-item, .sc-reservas-metric");
        cards.forEach((card) => {
            const labelNode = card.querySelector(".kpi-card-label, .sc-dash-kpi-label, p, span");
            if (!labelNode) return;

            const meta = resolverMetaKpi(labelNode.textContent || "");
            card.classList.remove("kpi-tone-blue", "kpi-tone-green", "kpi-tone-amber", "kpi-tone-red");
            card.classList.add(meta.tone);

            let iconWrap = card.querySelector(".kpi-context-icon");
            if (!iconWrap) {
                iconWrap = document.createElement("span");
                iconWrap.className = "kpi-context-icon";
                card.appendChild(iconWrap);
            }
            iconWrap.innerHTML = `<i class="bi ${meta.icon}" aria-hidden="true"></i>`;
        });
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", estilizarKpis);
    } else {
        estilizarKpis();
    }
})();

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
