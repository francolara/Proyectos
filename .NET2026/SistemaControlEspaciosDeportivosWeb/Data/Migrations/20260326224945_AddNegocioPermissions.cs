using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SistemaControlEspaciosDeportivosWeb.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddNegocioPermissions : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ModulosSistema",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Codigo = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(120)", maxLength: 120, nullable: false),
                    Activo = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ModulosSistema", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "RolesNegocioPermiso",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    RolNegocio = table.Column<int>(type: "int", nullable: false),
                    ModuloSistemaId = table.Column<int>(type: "int", nullable: false),
                    PuedeVer = table.Column<bool>(type: "bit", nullable: false),
                    PuedeCrear = table.Column<bool>(type: "bit", nullable: false),
                    PuedeEditar = table.Column<bool>(type: "bit", nullable: false),
                    PuedeEliminar = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_RolesNegocioPermiso", x => x.Id);
                    table.ForeignKey(
                        name: "FK_RolesNegocioPermiso_ModulosSistema_ModuloSistemaId",
                        column: x => x.ModuloSistemaId,
                        principalTable: "ModulosSistema",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "UsuariosNegocioPermiso",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    UsuarioNegocioId = table.Column<int>(type: "int", nullable: false),
                    ModuloSistemaId = table.Column<int>(type: "int", nullable: false),
                    PuedeVer = table.Column<bool>(type: "bit", nullable: false),
                    PuedeCrear = table.Column<bool>(type: "bit", nullable: false),
                    PuedeEditar = table.Column<bool>(type: "bit", nullable: false),
                    PuedeEliminar = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_UsuariosNegocioPermiso", x => x.Id);
                    table.ForeignKey(
                        name: "FK_UsuariosNegocioPermiso_ModulosSistema_ModuloSistemaId",
                        column: x => x.ModuloSistemaId,
                        principalTable: "ModulosSistema",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_UsuariosNegocioPermiso_UsuariosNegocio_UsuarioNegocioId",
                        column: x => x.UsuarioNegocioId,
                        principalTable: "UsuariosNegocio",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_ModulosSistema_Codigo",
                table: "ModulosSistema",
                column: "Codigo",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_RolesNegocioPermiso_ModuloSistemaId",
                table: "RolesNegocioPermiso",
                column: "ModuloSistemaId");

            migrationBuilder.CreateIndex(
                name: "IX_RolesNegocioPermiso_RolNegocio_ModuloSistemaId",
                table: "RolesNegocioPermiso",
                columns: new[] { "RolNegocio", "ModuloSistemaId" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_UsuariosNegocioPermiso_ModuloSistemaId",
                table: "UsuariosNegocioPermiso",
                column: "ModuloSistemaId");

            migrationBuilder.CreateIndex(
                name: "IX_UsuariosNegocioPermiso_UsuarioNegocioId_ModuloSistemaId",
                table: "UsuariosNegocioPermiso",
                columns: new[] { "UsuarioNegocioId", "ModuloSistemaId" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "RolesNegocioPermiso");

            migrationBuilder.DropTable(
                name: "UsuariosNegocioPermiso");

            migrationBuilder.DropTable(
                name: "ModulosSistema");
        }
    }
}
