using Microsoft.AspNetCore.Identity;

namespace SistemaControlEspaciosDeportivosWeb.Data;

public class SpanishIdentityErrorDescriber : IdentityErrorDescriber
{
    public override IdentityError DefaultError() => new() { Code = nameof(DefaultError), Description = "Ocurrio un error inesperado." };
    public override IdentityError ConcurrencyFailure() => new() { Code = nameof(ConcurrencyFailure), Description = "Error de concurrencia, vuelve a intentarlo." };
    public override IdentityError PasswordMismatch() => new() { Code = nameof(PasswordMismatch), Description = "La contrasena actual es incorrecta." };
    public override IdentityError InvalidToken() => new() { Code = nameof(InvalidToken), Description = "El token no es valido." };
    public override IdentityError LoginAlreadyAssociated() => new() { Code = nameof(LoginAlreadyAssociated), Description = "Este inicio de sesion ya esta asociado a otra cuenta." };
    public override IdentityError InvalidUserName(string? userName) => new() { Code = nameof(InvalidUserName), Description = $"El usuario '{userName}' no es valido." };
    public override IdentityError InvalidEmail(string? email) => new() { Code = nameof(InvalidEmail), Description = $"El correo '{email}' no es valido." };
    public override IdentityError DuplicateUserName(string userName) => new() { Code = nameof(DuplicateUserName), Description = $"El usuario '{userName}' ya existe." };
    public override IdentityError DuplicateEmail(string email) => new() { Code = nameof(DuplicateEmail), Description = $"El correo '{email}' ya esta registrado." };
    public override IdentityError InvalidRoleName(string? role) => new() { Code = nameof(InvalidRoleName), Description = $"El rol '{role}' no es valido." };
    public override IdentityError DuplicateRoleName(string role) => new() { Code = nameof(DuplicateRoleName), Description = $"El rol '{role}' ya existe." };
    public override IdentityError UserAlreadyHasPassword() => new() { Code = nameof(UserAlreadyHasPassword), Description = "El usuario ya tiene contrasena asignada." };
    public override IdentityError UserLockoutNotEnabled() => new() { Code = nameof(UserLockoutNotEnabled), Description = "El bloqueo de usuario no esta habilitado." };
    public override IdentityError UserAlreadyInRole(string role) => new() { Code = nameof(UserAlreadyInRole), Description = $"El usuario ya pertenece al rol '{role}'." };
    public override IdentityError UserNotInRole(string role) => new() { Code = nameof(UserNotInRole), Description = $"El usuario no pertenece al rol '{role}'." };
    public override IdentityError PasswordTooShort(int length) => new() { Code = nameof(PasswordTooShort), Description = $"La contrasena debe tener al menos {length} caracteres." };
    public override IdentityError PasswordRequiresNonAlphanumeric() => new() { Code = nameof(PasswordRequiresNonAlphanumeric), Description = "La contrasena debe incluir al menos un simbolo." };
    public override IdentityError PasswordRequiresDigit() => new() { Code = nameof(PasswordRequiresDigit), Description = "La contrasena debe incluir al menos un numero." };
    public override IdentityError PasswordRequiresLower() => new() { Code = nameof(PasswordRequiresLower), Description = "La contrasena debe incluir al menos una letra minuscula." };
    public override IdentityError PasswordRequiresUpper() => new() { Code = nameof(PasswordRequiresUpper), Description = "La contrasena debe incluir al menos una letra mayuscula." };
}
