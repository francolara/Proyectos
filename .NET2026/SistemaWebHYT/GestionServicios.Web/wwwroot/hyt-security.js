document.addEventListener("click", async event => {
    const button = event.target.closest("[data-security-copy]");
    if (!button) return;
    const group = button.closest("[data-security-password]");
    const input = group.querySelector("input");
    const status = group.querySelector("[role=status]");
    try {
        await navigator.clipboard.writeText(input.value);
        status.textContent = "Contraseña copiada.";
    } catch {
        status.textContent = "Usa el icono del ojo para ver y copiar la contraseña.";
    }
});

const profile = document.querySelector("[data-security-profile]");
const count = profile?.querySelector("[data-security-count]");
if (count) {
    const rows = [...profile.querySelectorAll("[data-security-module]")];
    const permissions = row => [...row.querySelectorAll("input[name=Permissions]")];
    const refresh = () => {
        for (const row of rows) {
            const boxes = permissions(row);
            const all = row.querySelector("[data-security-all]");
            all.checked = boxes.every(box => box.checked);
            all.indeterminate = !all.checked && boxes.some(box => box.checked);
        }
        count.textContent = profile.querySelectorAll("input[name=Permissions]:checked").length + " permisos seleccionados";
    };
    profile.addEventListener("change", event => {
        const box = event.target;
        const row = box.closest("[data-security-module]");
        if (!row) return;
        const boxes = permissions(row);
        if (box.hasAttribute("data-security-all")) {
            for (const permission of boxes) permission.checked = box.checked;
        } else {
            const consult = row.querySelector("[data-security-consult]");
            if (consult && box !== consult && box.checked) consult.checked = true;
            if (consult === box && !box.checked)
                for (const permission of boxes) permission.checked = false;
        }
        refresh();
    });
    refresh();
}

const auditModule = document.getElementById("audit-entity");
const auditAction = document.getElementById("audit-action");
if (auditModule && auditAction) {
    const options = [...auditAction.options].filter(option => option.value);
    const updateActions = (preserveSelection = false) => {
        const selected = auditAction.value;
        for (const option of options) {
            const available = !auditModule.value || option.dataset.auditModules.split(",").includes(auditModule.value);
            option.hidden = !available && !(preserveSelection && option.value === selected);
            option.disabled = option.hidden;
        }
        auditAction.value = options.some(option => option.value === selected && !option.disabled) ? selected : "";
    };
    auditModule.addEventListener("change", () => updateActions());
    updateActions(true);
}
