namespace GestionServicios.Domain;

public static class BusinessCodes
{
    public const string Soles = "PEN";
    public const string Dollars = "USD";
    public const string Invoice = "01";
    public const string Receipt = "03";
    public const string IncludesTax = "CON_IGV";
    public const string ExcludesTax = "SIN_IGV";
}

public enum AttentionType { Invitation = 1, Emergency = 2 }
