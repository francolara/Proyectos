using Microsoft.AspNetCore.Identity;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public static class IdentityErrorTranslator
{
    public static string Translate(IdentityError error)
    {
        return error.Code switch
        {
            "DuplicateUserName" => "Ya existe una cuenta registrada con este correo.",
            "DuplicateEmail" => "Ya existe una cuenta registrada con este correo.",
            "InvalidEmail" => "Ingrese un correo electronico valido.",
            "InvalidUserName" => "Ingrese un correo electronico valido.",
            "PasswordTooShort" => "La contrasena debe tener al menos 6 caracteres.",
            "PasswordRequiresNonAlphanumeric" => "La contrasena debe incluir al menos un caracter especial.",
            "PasswordRequiresDigit" => "La contrasena debe incluir al menos un numero.",
            "PasswordRequiresLower" => "La contrasena debe incluir al menos una letra minuscula.",
            "PasswordRequiresUpper" => "La contrasena debe incluir al menos una letra mayuscula.",
            "PasswordMismatch" => "La contrasena ingresada no es correcta.",
            "LoginAlreadyAssociated" => "Este acceso externo ya esta vinculado a otra cuenta.",
            _ => TranslateFallbackDescription(error.Description)
        };
    }

    private static string TranslateFallbackDescription(string? description)
    {
        if (string.IsNullOrWhiteSpace(description))
        {
            return "No se pudo validar la operacion de usuario. Revise los datos ingresados.";
        }

        if (description.Contains("non alphanumeric", StringComparison.OrdinalIgnoreCase))
        {
            return "La contrasena debe incluir al menos un caracter especial.";
        }

        if (description.Contains("one digit", StringComparison.OrdinalIgnoreCase))
        {
            return "La contrasena debe incluir al menos un numero.";
        }

        if (description.Contains("one lowercase", StringComparison.OrdinalIgnoreCase))
        {
            return "La contrasena debe incluir al menos una letra minuscula.";
        }

        if (description.Contains("one uppercase", StringComparison.OrdinalIgnoreCase))
        {
            return "La contrasena debe incluir al menos una letra mayuscula.";
        }

        if (description.Contains("already taken", StringComparison.OrdinalIgnoreCase)
            || description.Contains("already exists", StringComparison.OrdinalIgnoreCase))
        {
            return "Ya existe una cuenta registrada con ese dato.";
        }

        if (description.Contains("invalid email", StringComparison.OrdinalIgnoreCase))
        {
            return "Ingrese un correo electronico valido.";
        }

        return "No se pudo validar la operacion de usuario. Revise los datos ingresados.";
    }
}
