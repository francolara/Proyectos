/*
Firma: Codex - 09/04/2026
Descripcion: Crea/actualiza SP para guardar checks de emision de comprobantes en Negocios.
*/
USE [DbSportCenter]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_ActualizarEmision
    @NegocioId INT,
    @EmisionComprobantesElectronicos BIT = 0,
    @EmisionReciboInterno BIT = 0,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        UPDATE n
        SET
            n.EmisionComprobantesElectronicos = @EmisionComprobantesElectronicos,
            n.EmisionReciboInterno = @EmisionReciboInterno
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el club para actualizar emision.', 16, 1);

        DECLARE @EntidadIdAuditoria NVARCHAR(80);
        SET @EntidadIdAuditoria = CONVERT(NVARCHAR(80), @NegocioId);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CONFIGURACION',
            @Accion = N'EDIT',
            @Entidad = N'Negocio',
            @EntidadId = @EntidadIdAuditoria,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
