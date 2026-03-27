using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SistemaControlEspaciosDeportivosWeb.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddComprobanteElectronico : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "DireccionFiscal",
                table: "Clientes",
                type: "nvarchar(250)",
                maxLength: 250,
                nullable: true);

            migrationBuilder.CreateTable(
                name: "ComprobantesElectronicos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NegocioId = table.Column<int>(type: "int", nullable: false),
                    ReservaId = table.Column<int>(type: "int", nullable: false),
                    ClienteId = table.Column<int>(type: "int", nullable: false),
                    TipoComprobante = table.Column<int>(type: "int", nullable: false),
                    Serie = table.Column<string>(type: "nvarchar(4)", maxLength: 4, nullable: false),
                    Numero = table.Column<int>(type: "int", nullable: false),
                    FechaEmision = table.Column<DateTime>(type: "datetime2", nullable: false),
                    TipoMoneda = table.Column<int>(type: "int", nullable: false),
                    CodigoTipoOperacionSunat = table.Column<string>(type: "nvarchar(4)", maxLength: 4, nullable: false),
                    CodigoTipoDocumentoClienteSunat = table.Column<string>(type: "nvarchar(4)", maxLength: 4, nullable: false),
                    CodigoHashCpe = table.Column<string>(type: "nvarchar(8)", maxLength: 8, nullable: true),
                    NumeroTicketSunat = table.Column<string>(type: "nvarchar(40)", maxLength: 40, nullable: true),
                    CodigoRespuestaSunat = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    MensajeRespuestaSunat = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true),
                    SubTotal = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    Igv = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    Total = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    Estado = table.Column<int>(type: "int", nullable: false),
                    FechaRegistro = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ComprobantesElectronicos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_ComprobantesElectronicos_Clientes_ClienteId",
                        column: x => x.ClienteId,
                        principalTable: "Clientes",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_ComprobantesElectronicos_Negocios_NegocioId",
                        column: x => x.NegocioId,
                        principalTable: "Negocios",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_ComprobantesElectronicos_Reservas_ReservaId",
                        column: x => x.ReservaId,
                        principalTable: "Reservas",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "ComprobantesDetalle",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ComprobanteElectronicoId = table.Column<int>(type: "int", nullable: false),
                    Item = table.Column<int>(type: "int", nullable: false),
                    Descripcion = table.Column<string>(type: "nvarchar(250)", maxLength: 250, nullable: false),
                    Cantidad = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    UnidadMedidaSunat = table.Column<string>(type: "nvarchar(3)", maxLength: 3, nullable: false),
                    ValorUnitario = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    PrecioUnitario = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    BaseIgv = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    Igv = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    Total = table.Column<decimal>(type: "decimal(10,2)", precision: 10, scale: 2, nullable: false),
                    AfectacionIgvSunat = table.Column<string>(type: "nvarchar(2)", maxLength: 2, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ComprobantesDetalle", x => x.Id);
                    table.ForeignKey(
                        name: "FK_ComprobantesDetalle_ComprobantesElectronicos_ComprobanteElectronicoId",
                        column: x => x.ComprobanteElectronicoId,
                        principalTable: "ComprobantesElectronicos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_ComprobantesDetalle_ComprobanteElectronicoId_Item",
                table: "ComprobantesDetalle",
                columns: new[] { "ComprobanteElectronicoId", "Item" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_ComprobantesElectronicos_ClienteId",
                table: "ComprobantesElectronicos",
                column: "ClienteId");

            migrationBuilder.CreateIndex(
                name: "IX_ComprobantesElectronicos_NegocioId_TipoComprobante_Serie_Numero",
                table: "ComprobantesElectronicos",
                columns: new[] { "NegocioId", "TipoComprobante", "Serie", "Numero" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_ComprobantesElectronicos_ReservaId",
                table: "ComprobantesElectronicos",
                column: "ReservaId",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ComprobantesDetalle");

            migrationBuilder.DropTable(
                name: "ComprobantesElectronicos");

            migrationBuilder.DropColumn(
                name: "DireccionFiscal",
                table: "Clientes");
        }
    }
}
