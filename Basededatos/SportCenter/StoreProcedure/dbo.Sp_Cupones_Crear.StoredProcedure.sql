
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: FRANCO LARA
-- Create date: 03/05/2026
-- Firma: FRANCO LARA - 01/10/2026 | Asigna el correlativo visible de cupon por negocio, independiente del codigo canjeable.
-- Firma: FRANCO LARA - 06/10/2026 | Persiste el codigo canonico de moneda del negocio.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Cupones_Crear]
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
        DECLARE @CodigoMoneda NVARCHAR(10);
        SELECT @CodigoMoneda = n.CodigoMoneda FROM dbo.Negocios n WHERE n.Id = @NegocioId AND n.Activo = 1;
        IF @CodigoMoneda IS NULL RAISERROR('El negocio debe configurar una moneda valida antes de crear cupones.', 16, 1);

        SET @CodigoCupon = UPPER(LTRIM(RTRIM(@CodigoCupon)));
        IF @CodigoCupon = N'' RAISERROR('El codigo de cupon es obligatorio.', 16, 1);
        IF @FechaFin < @FechaInicio RAISERROR('La fecha fin no puede ser menor a la fecha inicio.', 16, 1);
        IF @CantidadMaxUsos <= 0 RAISERROR('La cantidad maxima de usos debe ser mayor a cero.', 16, 1);
        IF @TipoDescuento NOT IN (N'PORCENTAJE', N'MONTO_FIJO') RAISERROR('Tipo de descuento no valido.', 16, 1);
        IF @TipoDescuento = N'PORCENTAJE' AND (@ValorDescuento <= 0 OR @ValorDescuento > 100) RAISERROR('El porcentaje debe estar entre 0.01 y 100.', 16, 1);
        IF @TipoDescuento = N'MONTO_FIJO' AND @ValorDescuento <= 0 RAISERROR('El monto fijo debe ser mayor a cero.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Cupones WHERE NegocioId = @NegocioId AND CodigoCupon = @CodigoCupon)
            RAISERROR('El codigo de cupon ya existe para este negocio.', 16, 1);

        DECLARE @NumeroPorNegocio INT;
        EXEC dbo.Sp_NegocioCorrelativos_ObtenerSiguiente @NegocioId, N'CUPON', @Usuario, @NumeroPorNegocio OUTPUT;

        INSERT INTO dbo.Cupones
        (
            NumeroPorNegocio, NegocioId, SedeId, EspacioDeportivoId, CodigoCupon, Nombre, TipoDescuento, ValorDescuento, CodigoMoneda,
            CantidadMaxUsos, CantidadUsosActuales, FechaInicio, FechaFin, Activo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NumeroPorNegocio, @NegocioId, @SedeId, @EspacioDeportivoId, @CodigoCupon, @Nombre, @TipoDescuento, @ValorDescuento, @CodigoMoneda,
            @CantidadMaxUsos, 0, @FechaInicio, @FechaFin, @Activo, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CUPONES', @Accion = N'CREATE', @Entidad = N'Cupon', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
