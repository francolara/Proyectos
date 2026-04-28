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
