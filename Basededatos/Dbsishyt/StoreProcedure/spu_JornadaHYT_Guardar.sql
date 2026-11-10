-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_JornadaHYT_Guardar para el registro de asistencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_JornadaHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Entrada DATETIME2(0),
    @Salida DATETIME2(0),
    @MinutosExtrasAprobados INT,
    @JornadaId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'ASISTENCIA.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar asistencia.', 1;
        IF @Entrada IS NULL OR @MinutosExtrasAprobados IS NULL OR @MinutosExtrasAprobados < 0
           OR (@Salida IS NULL AND @MinutosExtrasAprobados <> 0)
           OR (@Salida IS NOT NULL AND (@Salida <= @Entrada
               OR CONVERT(BIGINT, @MinutosExtrasAprobados) * 60 > DATEDIFF_BIG(SECOND, @Entrada, @Salida)))
            THROW 51000, N'Revise las horas de entrada, salida y minutos extras.', 1;

        BEGIN TRANSACTION;
        DECLARE @FechaIngreso DATE;
        SELECT @FechaIngreso = trabajador.FechaIngreso
        FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
        WHERE trabajador.TrabajadorId = @TrabajadorId AND (trabajador.Activo = 1 OR @JornadaId IS NOT NULL);

        IF @FechaIngreso IS NULL OR @Entrada < @FechaIngreso
            THROW 51000, N'La jornada debe corresponder a un trabajador activo y a una fecha desde su ingreso.', 1;

        IF EXISTS (
            SELECT 1 FROM dbo.JornadaHYT AS jornada
            WHERE jornada.TrabajadorId = @TrabajadorId
              AND (@JornadaId IS NULL OR jornada.JornadaId <> @JornadaId)
              AND (@Salida IS NULL OR jornada.Entrada < @Salida)
              AND (jornada.Salida IS NULL OR jornada.Salida > @Entrada))
            THROW 51000, N'La jornada se superpone con otra jornada del trabajador.', 1;

        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @JornadaId IS NULL
        BEGIN
            INSERT INTO dbo.JornadaHYT (TrabajadorId, Entrada, Salida, MinutosExtrasAprobados)
            VALUES (@TrabajadorId, @Entrada, @Salida, @MinutosExtrasAprobados);
            SET @JornadaId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.JornadaHYT
            SET Entrada = @Entrada, Salida = @Salida, MinutosExtrasAprobados = @MinutosExtrasAprobados
            WHERE JornadaId = @JornadaId AND TrabajadorId = @TrabajadorId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La jornada cambio o no pertenece al trabajador. Vuelva a abrirla.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @JornadaId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Entrada AS Entrada, @Salida AS Salida, @MinutosExtrasAprobados AS MinutosExtrasAprobados
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'JornadaHYT', @RegistroId,
             @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @JornadaId AS JornadaId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
