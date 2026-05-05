USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: FRANCO LARA
-- Create date: 03/05/2026
CREATE OR ALTER PROCEDURE [dbo].[Sp_Cupones_Actualizar]
    @Id INT,
    @NegocioId INT,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @CodigoCupon NVARCHAR(30),
    @Nombre NVARCHAR(150),
    @TipoDescuento NVARCHAR(20),
    @ValorDescuento DECIMAL(10,2),
    @CantidadMaxUsos INT,
    @FechaInicio DATE,
    @FechaFin DATE,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @CodigoCupon = UPPER(LTRIM(RTRIM(@CodigoCupon)));
        IF @CodigoCupon = N'' RAISERROR('El codigo de cupon es obligatorio.', 16, 1);
        IF @FechaFin < @FechaInicio RAISERROR('La fecha fin no puede ser menor a la fecha inicio.', 16, 1);
        IF @CantidadMaxUsos <= 0 RAISERROR('La cantidad maxima de usos debe ser mayor a cero.', 16, 1);
        IF @TipoDescuento NOT IN (N'PORCENTAJE', N'MONTO_FIJO') RAISERROR('Tipo de descuento no valido.', 16, 1);
        IF @TipoDescuento = N'PORCENTAJE' AND (@ValorDescuento <= 0 OR @ValorDescuento > 100) RAISERROR('El porcentaje debe estar entre 0.01 y 100.', 16, 1);
        IF @TipoDescuento = N'MONTO_FIJO' AND @ValorDescuento <= 0 RAISERROR('El monto fijo debe ser mayor a cero.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Cupones WHERE NegocioId = @NegocioId AND CodigoCupon = @CodigoCupon AND Id <> @Id)
            RAISERROR('El codigo de cupon ya existe para este negocio.', 16, 1);

        UPDATE c
           SET c.SedeId = @SedeId,
               c.EspacioDeportivoId = @EspacioDeportivoId,
               c.CodigoCupon = @CodigoCupon,
               c.Nombre = @Nombre,
               c.TipoDescuento = @TipoDescuento,
               c.ValorDescuento = @ValorDescuento,
               c.CantidadMaxUsos = @CantidadMaxUsos,
               c.FechaInicio = @FechaInicio,
               c.FechaFin = @FechaFin,
               c.Activo = @Activo,
               c.FechaActualizacion = SYSUTCDATETIME(),
               c.UsuarioActualizacion = @Usuario
        FROM dbo.Cupones c
        WHERE c.Id = @Id
          AND c.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cupon para actualizar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CUPONES', @Accion = N'EDIT', @Entidad = N'Cupon', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
