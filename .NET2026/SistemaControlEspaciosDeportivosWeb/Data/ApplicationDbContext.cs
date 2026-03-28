using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using SistemaControlEspaciosDeportivosWeb.Models;

namespace SistemaControlEspaciosDeportivosWeb.Data;

public class ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : IdentityDbContext<ApplicationUser>(options)
{
    public DbSet<Negocio> Negocios => Set<Negocio>();
    public DbSet<Sede> Sedes => Set<Sede>();
    public DbSet<TipoDeporte> TiposDeporte => Set<TipoDeporte>();
    public DbSet<TipoSuelo> TiposSuelo => Set<TipoSuelo>();
    public DbSet<EspacioDeportivo> EspaciosDeportivos => Set<EspacioDeportivo>();
    public DbSet<Cliente> Clientes => Set<Cliente>();
    public DbSet<Tarifa> Tarifas => Set<Tarifa>();
    public DbSet<Reserva> Reservas => Set<Reserva>();
    public DbSet<Pago> Pagos => Set<Pago>();
    public DbSet<UsuarioNegocio> UsuariosNegocio => Set<UsuarioNegocio>();
    public DbSet<ModuloSistema> ModulosSistema => Set<ModuloSistema>();
    public DbSet<RolNegocioPermiso> RolesNegocioPermiso => Set<RolNegocioPermiso>();
    public DbSet<UsuarioNegocioPermiso> UsuariosNegocioPermiso => Set<UsuarioNegocioPermiso>();
    public DbSet<ComprobanteElectronico> ComprobantesElectronicos => Set<ComprobanteElectronico>();
    public DbSet<ComprobanteDetalle> ComprobantesDetalle => Set<ComprobanteDetalle>();
    public DbSet<BitacoraAuditoria> BitacoraAuditoria => Set<BitacoraAuditoria>();

    protected override void OnModelCreating(ModelBuilder builder)
    {
        base.OnModelCreating(builder);

        builder.Entity<UsuarioNegocio>()
            .HasIndex(un => new { un.UsuarioId, un.NegocioId })
            .IsUnique();

        builder.Entity<EspacioDeportivo>()
            .HasIndex(e => new { e.SedeId, e.Codigo })
            .IsUnique();

        builder.Entity<Tarifa>()
            .HasIndex(t => new { t.EspacioDeportivoId, t.DiaSemana, t.HoraInicio, t.HoraFin });

        builder.Entity<Reserva>()
            .HasIndex(r => new { r.EspacioDeportivoId, r.Fecha, r.HoraInicio, r.HoraFin });

        builder.Entity<Cliente>()
            .HasIndex(c => new { c.TipoDocumento, c.NumeroDocumento })
            .IsUnique();

        builder.Entity<ModuloSistema>()
            .HasIndex(m => m.Codigo)
            .IsUnique();

        builder.Entity<ComprobanteElectronico>()
            .HasIndex(c => new { c.NegocioId, c.TipoComprobante, c.Serie, c.Numero })
            .IsUnique();

        builder.Entity<ComprobanteElectronico>()
            .HasIndex(c => c.ReservaId)
            .IsUnique();

        builder.Entity<ComprobanteDetalle>()
            .HasIndex(cd => new { cd.ComprobanteElectronicoId, cd.Item })
            .IsUnique();

        builder.Entity<RolNegocioPermiso>()
            .HasIndex(rp => new { rp.RolNegocio, rp.ModuloSistemaId })
            .IsUnique();

        builder.Entity<UsuarioNegocioPermiso>()
            .HasIndex(up => new { up.UsuarioNegocioId, up.ModuloSistemaId })
            .IsUnique();

        builder.Entity<BitacoraAuditoria>()
            .HasIndex(b => new { b.NegocioId, b.Modulo, b.FechaRegistro });

        builder.Entity<Reserva>()
            .HasOne(r => r.ComprobanteElectronico)
            .WithOne(c => c.Reserva)
            .HasForeignKey<ComprobanteElectronico>(c => c.ReservaId)
            .OnDelete(DeleteBehavior.Restrict);

        builder.Entity<Negocio>()
            .HasMany(n => n.ComprobantesElectronicos)
            .WithOne(c => c.Negocio)
            .HasForeignKey(c => c.NegocioId)
            .OnDelete(DeleteBehavior.Restrict);

        builder.Entity<Cliente>()
            .HasMany(c => c.ComprobantesElectronicos)
            .WithOne(ce => ce.Cliente)
            .HasForeignKey(ce => ce.ClienteId)
            .OnDelete(DeleteBehavior.Restrict);

        builder.Entity<UsuarioNegocioPermiso>()
            .HasOne(up => up.UsuarioNegocio)
            .WithMany(un => un.Permisos)
            .HasForeignKey(up => up.UsuarioNegocioId)
            .OnDelete(DeleteBehavior.Cascade);

        builder.Entity<UsuarioNegocioPermiso>()
            .HasOne(up => up.ModuloSistema)
            .WithMany(m => m.UsuariosPermiso)
            .HasForeignKey(up => up.ModuloSistemaId)
            .OnDelete(DeleteBehavior.Restrict);

        builder.Entity<RolNegocioPermiso>()
            .HasOne(rp => rp.ModuloSistema)
            .WithMany(m => m.RolesPermiso)
            .HasForeignKey(rp => rp.ModuloSistemaId)
            .OnDelete(DeleteBehavior.Restrict);

        builder.Entity<Reserva>()
            .Property(r => r.Total)
            .HasPrecision(10, 2);

        builder.Entity<Reserva>()
            .Property(r => r.Adelanto)
            .HasPrecision(10, 2);

        builder.Entity<Reserva>()
            .Property(r => r.Saldo)
            .HasPrecision(10, 2);

        builder.Entity<Tarifa>()
            .Property(t => t.Precio)
            .HasPrecision(10, 2);

        builder.Entity<Pago>()
            .Property(p => p.Monto)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteElectronico>()
            .Property(c => c.SubTotal)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteElectronico>()
            .Property(c => c.Igv)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteElectronico>()
            .Property(c => c.Total)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteDetalle>()
            .Property(cd => cd.Cantidad)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteDetalle>()
            .Property(cd => cd.ValorUnitario)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteDetalle>()
            .Property(cd => cd.PrecioUnitario)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteDetalle>()
            .Property(cd => cd.BaseIgv)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteDetalle>()
            .Property(cd => cd.Igv)
            .HasPrecision(10, 2);

        builder.Entity<ComprobanteDetalle>()
            .Property(cd => cd.Total)
            .HasPrecision(10, 2);
    }
}
