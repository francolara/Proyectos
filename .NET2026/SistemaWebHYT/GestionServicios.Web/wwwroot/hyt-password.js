document.addEventListener("click", event => {
    const button = event.target.closest("[data-hyt-password-toggle]");
    if (!button) return;

    const input = document.getElementById(button.getAttribute("aria-controls"));
    if (!input) return;

    const show = input.type === "password";
    input.type = show ? "text" : "password";
    button.querySelector("[data-password-show]").toggleAttribute("hidden", show);
    button.querySelector("[data-password-hide]").toggleAttribute("hidden", !show);
    const label = show ? "Ocultar contraseña" : "Mostrar contraseña";
    button.setAttribute("aria-label", label);
    button.title = label;
});
