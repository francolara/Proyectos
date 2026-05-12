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
