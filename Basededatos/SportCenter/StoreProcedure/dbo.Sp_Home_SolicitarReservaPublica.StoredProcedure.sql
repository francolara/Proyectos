USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Home_SolicitarReservaPublica]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 10_Home_Solicitudes_Publicas.sql (linea 35)
-- Firma: Codex - 14/04/2026 | Convierte flujo publico: ahora crea reserva real (canal CLIENTE_WEB), reutiliza/crea cliente y aplica politica de confirmacion/pago del negocio.
-- Firma: Codex - 16/04/2026 | Registro publico autenticado: recibe UsuarioId opcional y guarda relacion en ReservasUsuariosPublicos para historial del perfil publico.
-- Firma: Codex - 17/04/2026 | Elimina INSERT-EXEC para crear reserva y usa salida @ReservaId de Sp_Reservas_Crear (evita error ROLLBACK dentro de INSERT-EXEC).
-- Firma: Codex - 18/04/2026 | Bloquea reservas publicas sobre espacios con AdministracionPrivada activada.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_SolicitarReservaPublica]
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Nombres NVARCHAR(120),
    @Apellidos NVARCHAR(120),
    @NombreEquipo NVARCHAR(120) = NULL,
    @TipoDocumento NVARCHAR(20) = N'0',
    @NumeroDocumento NVARCHAR(20) = NULL,
    @Telefono NVARCHAR(30) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @Comentario NVARCHAR(300) = NULL,
    @UsuarioId NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NombresNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Nombres)), N'');
        DECLARE @ApellidosNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Apellidos)), N'');
        DECLARE @NombreEquipoNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@NombreEquipo)), N'');
        DECLARE @TipoDocumentoNorm NVARCHAR(20) = COALESCE(NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''), N'0');
        DECLARE @NumeroDocumentoNorm NVARCHAR(20) = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        DECLARE @TelefonoNorm NVARCHAR(30) = NULLIF(LTRIM(RTRIM(@Telefono)), N'');
        DECLARE @CorreoNorm NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@Correo)), N'');
        DECLARE @ComentarioNorm NVARCHAR(300) = NULLIF(LTRIM(RTRIM(@Comentario)), N'');
        DECLARE @NombresORazonSocial NVARCHAR(200);
        DECLARE @NumeroDocumentoInsert NVARCHAR(20);
        DECLARE @NegocioId INT;
        DECLARE @ClienteId INT;
        DECLARE @PrecioFinal DECIMAL(10,2);
        DECLARE @ReservaId INT;

        IF @NombresNorm IS NULL OR @ApellidosNorm IS NULL
            RAISERROR('Nombres y apellidos son obligatorios.', 16, 1);

        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        SELECT
            @NegocioId = s.NegocioId
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE e.Id = @EspacioDeportivoId
          AND e.Estado = 1
          AND COALESCE(e.AdministracionPrivada, 0) = 0
          AND s.Activo = 1
          AND n.Activo = 1;

        IF @NegocioId IS NULL
            RAISERROR('El espacio deportivo no esta disponible.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('El horario seleccionado ya no esta disponible.', 16, 1);

        IF @NumeroDocumentoNorm IS NOT NULL
        BEGIN
            SELECT TOP 1
                @ClienteId = c.Id
            FROM dbo.Clientes c
            WHERE c.NegocioId = @NegocioId
              AND c.Activo = 1
              AND NULLIF(LTRIM(RTRIM(c.TipoDocumento)), N'') = @TipoDocumentoNorm
              AND NULLIF(LTRIM(RTRIM(c.NumeroDocumento)), N'') = @NumeroDocumentoNorm
            ORDER BY c.Id DESC;
        END

        IF @ClienteId IS NULL
        BEGIN
            SET @NombresORazonSocial = LTRIM(RTRIM(CONCAT(@NombresNorm, N' ', @ApellidosNorm)));
            SET @NumeroDocumentoInsert = COALESCE(@NumeroDocumentoNorm, N'');

            DECLARE @ClienteInsert TABLE (Id INT);
            INSERT INTO @ClienteInsert (Id)
            EXEC dbo.Sp_Clientes_Crear
                @NegocioId = @NegocioId,
                @NombresORazonSocial = @NombresORazonSocial,
                @Nombres = @NombresNorm,
                @Apellidos = @ApellidosNorm,
                @NombreEquipo = @NombreEquipoNorm,
                @TipoDocumento = @TipoDocumentoNorm,
                @NumeroDocumento = @NumeroDocumentoInsert,
                @Telefono = @TelefonoNorm,
                @Correo = @CorreoNorm,
                @DireccionFiscal = NULL,
                @CodigoUbigeo = NULL,
                @Activo = 1,
                @Usuario = N'portal-web';

            SELECT TOP 1 @ClienteId = Id FROM @ClienteInsert;
        END

        DECLARE @Cotizacion TABLE
        (
            Mensaje NVARCHAR(200),
            PrecioBase DECIMAL(10,2),
            DescuentoPct DECIMAL(5,2),
            PrecioFinal DECIMAL(10,2),
            MonedaSimbolo NVARCHAR(10),
            MonedaNombre NVARCHAR(80),
            PoliticaConfirmacionPago TINYINT,
            PorcentajeAdelantoMinimo DECIMAL(5,2)
        );

        INSERT INTO @Cotizacion
        EXEC dbo.Sp_Reservas_Cotizar
            @NegocioId = @NegocioId,
            @EspacioDeportivoId = @EspacioDeportivoId,
            @Fecha = @Fecha,
            @HoraInicio = @HoraInicio,
            @HoraFin = @HoraFin;

        SELECT TOP 1
            @PrecioFinal = PrecioFinal
        FROM @Cotizacion;

        IF @PrecioFinal IS NULL OR @PrecioFinal <= 0
            RAISERROR('No se pudo calcular el precio para el horario seleccionado.', 16, 1);

        EXEC dbo.Sp_Reservas_Crear
            @NegocioId = @NegocioId,
            @EspacioDeportivoId = @EspacioDeportivoId,
            @ClienteId = @ClienteId,
            @Fecha = @Fecha,
            @HoraInicio = @HoraInicio,
            @HoraFin = @HoraFin,
            @Total = @PrecioFinal,
            @Adelanto = 0,
            @Estado = 1,
            @RegistrarPago = 0,
            @FormaPagoId = NULL,
            @FechaPago = NULL,
            @NumeroOperacion = NULL,
            @Comentario = @ComentarioNorm,
            @CanalOrigen = N'CLIENTE_WEB',
            @ReservaId = @ReservaId OUTPUT,
            @Usuario = N'portal-web';

        IF @ReservaId IS NULL OR @ReservaId <= 0
            RAISERROR('No se pudo generar la reserva.', 16, 1);

        IF NULLIF(LTRIM(RTRIM(@UsuarioId)), N'') IS NOT NULL
           AND EXISTS (SELECT 1 FROM dbo.AspNetUsers WHERE Id = @UsuarioId)
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.ReservasUsuariosPublicos WHERE ReservaId = @ReservaId AND UsuarioId = @UsuarioId)
            BEGIN
                INSERT INTO dbo.ReservasUsuariosPublicos (ReservaId, UsuarioId, FechaCreacion, UsuarioCreacion)
                VALUES (@ReservaId, @UsuarioId, SYSDATETIME(), N'portal-web');
            END
        END

        SELECT @ReservaId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
