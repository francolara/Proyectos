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
