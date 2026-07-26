namespace SistemaAdministrativoWeb.Infrastructure.Security;

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, Inherited = true)]
public sealed class AllowRestrictedSubscriptionAttribute : Attribute;
