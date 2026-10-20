/****** Object:  StoredProcedure [dbo].[spu_Docventas_Mostrar_Docventas_cab]    Script Date: 13/10/2026 21:34:11 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Carga la cabecera del documento de forma idempotente y sin fijar la base de despliegue.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[spu_Docventas_Mostrar_Docventas_cab]
@idEmpresa     CHAR(2),
@idSucursal    VARCHAR(8),
@idDocumento   CHAR(2),
@idSerie	   CHAR(4),
@iddocventas   CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY

	SELECT idDocumento , idDocVentas, idSerie, idPerCliente, GlsCliente, RucCliente, dirCliente, idAlmacen, idPerVendedor, fecEmision, estDocVentas,TotalValorVenta , TotalIGVVenta, TotalPrecioVenta, 
	idMoneda, TipoCambio, obsdocventas, Partida, llegada, idPerChofer, glsChofer, idPerEmpTrans, GlsEmpTrans, idVehiculo, GlsVehiculo, Placa, Marca, Color, Modelo, CodInsCrip, Bultos, Brevete, 
	idLista, GlsVendedor, idMotivoTraslado, GlsFormasPago, GlsFecVectos, idMotivoNCD, GlsMotivoNCD, totalLetras, FecIniTraslado, rucEmpTrans, FecPago, GlsDocReferencia, simboloMonBruto, 
	simboloMonIGV, simboloMonNeto, indVale, idValesCab, idEmpresa, idSucursal, idCentroCosto, GlsCentroCosto, estDocImportado, idmovcaja, idPerVendedorCampo, GlsVendedorCampo, TotalDsctoVV, 
	TotalDsctoPV, TotalExonerado, TotalBaseImponible, estGuiaImportado, idTipoTicket, idUsuarioAnulacion, FechaAnulacion, FecRegistro, idUsuarioReg, Partida2, Llegada2, numOrdenCompra, idValeCabRec, 
	indAprobado, ComisionVtas, indLetra, idupp, idFormaPago, indTrasladoConta,glsFormapago , vtercerosCliente, vtercerosdireccion, vterceroCodigo, vterceroRuc, IndReq, IdContacto, 
	GlsMoneda, indTrasladoContaFin, idProvCliente,ObsDocVentas2 , idtienda, Tipovale, IdPeriodo, IdComprobante, GlsContacto, totalTallaPeso, IDArea, GlsTrabajo, GlsPlaca, GlsAreaCliente, 
	IdUsuarioAprueba, FechaAprueba, IdUsuarioDesaprueba, FechaDesaprueba, DiasEntrega, FechaEntrega, IndTerminado, indVtaGratuita, ObsRegVentas, SucursalDestino, 
	AlmacenDestino, GlsSucursalDestino, GlsAlmacenDestino, idValesCabIng, TipoValeIng, IndValeIng, GlsObsAnula, FechaTermino, FechaTerminoSistema, idAreaCli, Prioridad, 
	IndCerrado, IdUsuarioCierre, FechaCierre, IdUsuarioAbre, FechaAbre, ObsCierre, ObsAbre, PeriodoAD, IdCtaCorrienteOriAD, IdMovBancosCabOriAD, IdCtaCorrienteDesAD, 
	IdMovBancosCabDesAD, IdSucursalDestino, FechaPT, IdUsuarioPT, IndTransGratuitaMP, IndTransGratuita, PeriodoRec, indPeriodoRecEstado, IndAfectoPerc, 
	TotalIvap, TipoDescuentoGlobal, DescuentoGlobal, TotalPorcentajeDescuentoGlobal, TotalDescuentoGlobalGravado, TotalDescuentoGlobalExonerado, 
	TotalDescuento, IndAceptadoSunat, IndBajaSunat, IndEnviadoSunat, IndAtribucionNC,IndVContado,iddirrecojo,TotalCosto,ubigeollegada,ubigeoPartida 
	FROM docventas d 
	WHERE d.idEmpresa = @idEmpresa 
	AND d.idSucursal = @idSucursal 
	AND d.idDocumento = @idDocumento
	AND d.idDocVentas = @iddocventas
	AND d.idSerie = @idSerie
	;

    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
GO
