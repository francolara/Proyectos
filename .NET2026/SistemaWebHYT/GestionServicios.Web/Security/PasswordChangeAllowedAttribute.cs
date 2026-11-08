namespace GestionServicios.Web.Security;

[AttributeUsage(AttributeTargets.Method | AttributeTargets.Class)]
public sealed class PasswordChangeAllowedAttribute : Attribute;
