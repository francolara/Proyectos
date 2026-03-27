using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SistemaControlEspaciosDeportivosWeb.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddAuditoriaBitacora : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "FechaActualizacion",
                table: "Sedes",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaCreacion",
                table: "Sedes",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<string>(
                name: "UsuarioActualizacion",
                table: "Sedes",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCreacion",
                table: "Sedes",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaActualizacion",
                table: "Reservas",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioActualizacion",
                table: "Reservas",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCreacion",
                table: "Reservas",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaActualizacion",
                table: "Pagos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaCreacion",
                table: "Pagos",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<string>(
                name: "UsuarioActualizacion",
                table: "Pagos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCreacion",
                table: "Pagos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaActualizacion",
                table: "EspaciosDeportivos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaCreacion",
                table: "EspaciosDeportivos",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<string>(
                name: "UsuarioActualizacion",
                table: "EspaciosDeportivos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCreacion",
                table: "EspaciosDeportivos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "FechaActualizacion",
                table: "ComprobantesElectronicos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioActualizacion",
                table: "ComprobantesElectronicos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCreacion",
                table: "ComprobantesElectronicos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.CreateTable(
                name: "BitacoraAuditoria",
                columns: table => new
                {
                    Id = table.Column<long>(type: "bigint", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NegocioId = table.Column<int>(type: "int", nullable: true),
                    Modulo = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false),
                    Accion = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: false),
                    Entidad = table.Column<string>(type: "nvarchar(80)", maxLength: 80, nullable: false),
                    EntidadId = table.Column<string>(type: "nvarchar(80)", maxLength: 80, nullable: false),
                    UsuarioId = table.Column<string>(type: "nvarchar(450)", maxLength: 450, nullable: false),
                    UsuarioNombre = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    UsuarioCorreo = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    DetalleJson = table.Column<string>(type: "nvarchar(4000)", maxLength: 4000, nullable: true),
                    FechaRegistro = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BitacoraAuditoria", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_BitacoraAuditoria_NegocioId_Modulo_FechaRegistro",
                table: "BitacoraAuditoria",
                columns: new[] { "NegocioId", "Modulo", "FechaRegistro" });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "BitacoraAuditoria");

            migrationBuilder.DropColumn(
                name: "FechaActualizacion",
                table: "Sedes");

            migrationBuilder.DropColumn(
                name: "FechaCreacion",
                table: "Sedes");

            migrationBuilder.DropColumn(
                name: "UsuarioActualizacion",
                table: "Sedes");

            migrationBuilder.DropColumn(
                name: "UsuarioCreacion",
                table: "Sedes");

            migrationBuilder.DropColumn(
                name: "FechaActualizacion",
                table: "Reservas");

            migrationBuilder.DropColumn(
                name: "UsuarioActualizacion",
                table: "Reservas");

            migrationBuilder.DropColumn(
                name: "UsuarioCreacion",
                table: "Reservas");

            migrationBuilder.DropColumn(
                name: "FechaActualizacion",
                table: "Pagos");

            migrationBuilder.DropColumn(
                name: "FechaCreacion",
                table: "Pagos");

            migrationBuilder.DropColumn(
                name: "UsuarioActualizacion",
                table: "Pagos");

            migrationBuilder.DropColumn(
                name: "UsuarioCreacion",
                table: "Pagos");

            migrationBuilder.DropColumn(
                name: "FechaActualizacion",
                table: "EspaciosDeportivos");

            migrationBuilder.DropColumn(
                name: "FechaCreacion",
                table: "EspaciosDeportivos");

            migrationBuilder.DropColumn(
                name: "UsuarioActualizacion",
                table: "EspaciosDeportivos");

            migrationBuilder.DropColumn(
                name: "UsuarioCreacion",
                table: "EspaciosDeportivos");

            migrationBuilder.DropColumn(
                name: "FechaActualizacion",
                table: "ComprobantesElectronicos");

            migrationBuilder.DropColumn(
                name: "UsuarioActualizacion",
                table: "ComprobantesElectronicos");

            migrationBuilder.DropColumn(
                name: "UsuarioCreacion",
                table: "ComprobantesElectronicos");
        }
    }
}
