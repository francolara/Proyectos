namespace GestionServicios.Contracts;

public sealed record SearchFilter(string Search, DateOnly? From, DateOnly? To);
