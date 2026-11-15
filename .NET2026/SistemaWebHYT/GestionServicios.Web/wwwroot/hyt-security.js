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

// Firma: FRANCO LARA - 14/11/2026 | Navegación, búsqueda y selección global de permisos del perfil. Enlaza las opciones compartidas con su fila original.
const profile = document.querySelector("[data-security-profile]");
const count = profile?.querySelector("[data-security-count]");
if (count) {
    const rows = [...profile.querySelectorAll("[data-security-module]")].filter(row => row.querySelector("[data-security-all]"));
    const permissions = row => [...row.querySelectorAll("input[name=Permissions]")];
    const allPermissions = rows.flatMap(permissions);
    const selectAll = profile.querySelector('[data-profile-select="all"]');
    const clearAll = profile.querySelector('[data-profile-select="none"]');
    const setAll = checked => {
        for (const box of allPermissions) box.checked = checked;
        refresh();
    };
    const refresh = () => {
        for (const row of rows) {
            const boxes = permissions(row);
            const all = row.querySelector("[data-security-all]");
            all.checked = boxes.every(box => box.checked);
            all.indeterminate = !all.checked && boxes.some(box => box.checked);
        }
        const selected = allPermissions.filter(box => box.checked).length;
        count.textContent = selected + (selected === 1 ? " permiso seleccionado" : " permisos seleccionados");
        if (selectAll) selectAll.disabled = selected === allPermissions.length;
        if (clearAll) clearAll.disabled = selected === 0;
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
    selectAll?.addEventListener("click", () => setAll(true));
    clearAll?.addEventListener("click", () => setAll(false));
    refresh();
}

const navigation = profile?.querySelector("[data-profile-navigation]");
if (navigation) {
    const tabs = [...navigation.querySelectorAll("[data-profile-tab]")];
    const panels = tabs.map(tab => document.getElementById(tab.getAttribute("aria-controls")));
    const search = navigation.querySelector("[data-profile-search]");
    const results = navigation.querySelector("[data-profile-results]");
    const normalize = value => value.toLocaleLowerCase("es").normalize("NFD").replace(/\p{Diacritic}/gu, "");
    const rows = panels.map(panel => [...panel.querySelectorAll("[data-profile-row]")].map(row => ({
        row,
        terms: normalize([row.textContent, row.dataset.profileRow,
            ...[...row.querySelectorAll("[data-profile-action]")].map(action => action.dataset.profileAction)].join(" "))
    })));
    let active = 0;

    const render = () => {
        const query = normalize(search.value.trim());
        let found = 0;
        tabs.forEach((tab, index) => {
            tab.setAttribute("aria-selected", String(index === active));
            tab.tabIndex = index === active ? 0 : -1;
            let matches = 0;
            for (const {row, terms} of rows[index]) {
                row.hidden = query !== "" && !terms.includes(query);
                if (!row.hidden) matches++;
            }
            found += matches;
            const panel = panels[index];
            panel.hidden = query ? matches === 0 : index !== active;
            panel.setAttribute("role", query ? "region" : "tabpanel");
            panel.setAttribute("aria-labelledby", tab.id);
        });
        navigation.querySelector("[role=tablist]").hidden = Boolean(query);
        results.hidden = !query;
        if (query) results.textContent = found === 0 ? "No hay coincidencias." :
            found === 1 ? "1 opción encontrada." : found + " opciones encontradas.";
    };

    const activate = index => {
        active = index;
        search.value = "";
        render();
        tabs[active].focus();
    };
    navigation.hidden = false;
    search.addEventListener("input", render);
    tabs.forEach((tab, index) => {
        tab.addEventListener("click", () => activate(index));
        tab.addEventListener("keydown", event => {
            const next = { ArrowLeft: (index + tabs.length - 1) % tabs.length,
                ArrowRight: (index + 1) % tabs.length, Home: 0, End: tabs.length - 1 }[event.key];
            if (next === undefined) return;
            event.preventDefault();
            activate(next);
        });
    });
    profile.addEventListener("click", event => {
        const link = event.target.closest("[data-profile-owner]");
        if (!link) return;
        const target = document.getElementById(link.hash.slice(1));
        const index = panels.indexOf(target?.closest("[data-profile-panel]"));
        if (index < 0) return;
        event.preventDefault();
        activate(index);
        const control = target.querySelector("input[name=Permissions]");
        if (control) control.focus();
        else {
            target.tabIndex = -1;
            target.focus();
        }
    });
    render();
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
