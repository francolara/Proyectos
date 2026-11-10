-- Firma: FRANCO LARA | Create date: 09/11/2026 | Cargos y sueldo base de trabajadores.
USE [SistemaWebHYT];
GO
SET XACT_ABORT ON;
BEGIN TRANSACTION;
GO

-- Fuente: Tablas/CargoTrabajadorHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Catalogo de cargos del personal.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.CargoTrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CargoTrabajadorHYT
    (
        CargoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_CargoTrabajadorHYT PRIMARY KEY,
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT UQ_CargoTrabajadorHYT_Nombre UNIQUE,
        Activo BIT NOT NULL,
        VersionFila ROWVERSION NOT NULL
    );
END;
GO

-- Fuente: Script/037_CargosYSueldoHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega cargo y sueldo base mensual en soles a la ficha del trabajador.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'CargoId') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD CargoId BIGINT NULL CONSTRAINT FK_TrabajadorHYT_Cargo REFERENCES dbo.CargoTrabajadorHYT(CargoId);
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'SueldoMensual') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD SueldoMensual DECIMAL(18,2) NULL CONSTRAINT CK_TrabajadorHYT_Sueldo CHECK (SueldoMensual > 0);
GO

-- Fuente: StoreProcedure/spu_CargoTrabajadorHYT_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   spu_CargoTrabajadorHYT_Listar
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CargoTrabajadorHYT_Listar
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar cargos.', 1;

        SELECT cargo.CargoId, cargo.Nombre, cargo.Activo, cargo.VersionFila
        FROM dbo.CargoTrabajadorHYT AS cargo
        ORDER BY cargo.Nombre, cargo.CargoId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_CargoTrabajadorHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   spu_CargoTrabajadorHYT_Guardar
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CargoTrabajadorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Nombre NVARCHAR(100),
    @Activo BIT,
    @CargoId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar cargos.', 1;

        SET @Nombre = NULLIF(TRIM(@Nombre), N'');
        IF @Nombre IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique el nombre y el estado del cargo.', 1;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @CargoId IS NULL
        BEGIN
            INSERT INTO dbo.CargoTrabajadorHYT (Nombre, Activo) VALUES (@Nombre, @Activo);
            SET @CargoId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.CargoTrabajadorHYT SET Nombre = @Nombre, Activo = @Activo
            WHERE CargoId = @CargoId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'El cargo cambio o no existe.', 1;
        END;

        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CargoId);
        DECLARE @Cambios NVARCHAR(MAX) = (SELECT @Nombre AS Nombre, @Activo AS Activo FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'CargoTrabajadorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @CargoId AS CargoId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/038_PermisosCargosHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Habilita los procedimientos de cargos con permisos de Trabajadores.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CargoTrabajadorHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CargoTrabajadorHYT_Guardar TO hyt_app_executor;
GO

-- Fuente: StoreProcedure/spu_TrabajadorHYT_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta trabajadores, cargo y sueldo base en la ficha.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Listar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT = NULL,
    @Buscar NVARCHAR(180) = NULL,
    @SoloActivos BIT = 0,
    @Tipo VARCHAR(10) = NULL,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar trabajadores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA');
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        SELECT trabajador.TrabajadorId, trabajador.Documento, trabajador.NombreCompleto, trabajador.Tipo,
               trabajador.FechaIngreso, trabajador.Banco, trabajador.CargoId, cargo.Nombre AS CargoNombre,
               CASE WHEN @TrabajadorId IS NOT NULL THEN trabajador.SueldoMensual END AS SueldoMensual,
               CASE WHEN trabajador.NumeroCuenta IS NOT NULL THEN
                    CONCAT(N'****', RIGHT(trabajador.NumeroCuenta, 4)) END AS CuentaEnmascarada,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta = 1
                    THEN trabajador.NumeroCuenta END AS NumeroCuenta,
               trabajador.Activo, trabajador.VersionFila, COUNT_BIG(1) OVER () AS Total
        FROM dbo.TrabajadorHYT AS trabajador
        LEFT JOIN dbo.CargoTrabajadorHYT AS cargo ON cargo.CargoId = trabajador.CargoId
        WHERE (@TrabajadorId IS NULL OR trabajador.TrabajadorId = @TrabajadorId)
          AND (@Tipo IS NULL OR trabajador.Tipo = @Tipo)
          AND (@SoloActivos = 0 OR trabajador.Activo = 1)
          AND (@Buscar IS NULL OR CHARINDEX(@Buscar, trabajador.NombreCompleto) > 0
               OR CHARINDEX(@Buscar, trabajador.Documento) > 0)
        ORDER BY trabajador.NombreCompleto, trabajador.TrabajadorId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_TrabajadorHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Guarda trabajadores con cargo y sueldo base mensual.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Documento NVARCHAR(20),
    @NombreCompleto NVARCHAR(180),
    @Tipo VARCHAR(10),
    @FechaIngreso DATE,
    @Activo BIT,
    @TrabajadorId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @ActualizarCuenta BIT = 0,
    @Banco NVARCHAR(80) = NULL,
    @NumeroCuenta NVARCHAR(34) = NULL,
    @CargoId BIGINT = NULL,
    @SueldoMensual DECIMAL(20,4) = NULL,
    @ActualizarPuesto BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar trabajadores.', 1;

        SET @NombreCompleto = NULLIF(TRIM(@NombreCompleto), N'');
        SET @Documento = NULLIF(TRIM(@Documento), N'');
        IF @Documento IS NULL OR @Tipo IS NULL OR @Tipo NOT IN ('PLANILLA', 'EXTERNO')
           OR @FechaIngreso IS NULL
           OR @NombreCompleto IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique documento, nombre, tipo y fecha de ingreso.', 1;

        SET @Banco = NULLIF(TRIM(@Banco), N'');
        SET @NumeroCuenta = NULLIF(TRIM(@NumeroCuenta), N'');
        IF @ActualizarCuenta = 1
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA') = 0
                THROW 51000, N'No tiene permiso para cambiar datos bancarios.', 1;

            IF (@Banco IS NULL AND @NumeroCuenta IS NOT NULL)
               OR (@Banco IS NOT NULL AND @NumeroCuenta IS NULL)
                THROW 51000, N'Indique banco y numero de cuenta juntos.', 1;
        END;

        BEGIN TRANSACTION;
        IF @ActualizarPuesto = 1
        BEGIN
            IF @SueldoMensual IS NOT NULL AND (@SueldoMensual <= 0 OR @SueldoMensual <> ROUND(@SueldoMensual, 2))
                THROW 51000, N'Indique un sueldo positivo con hasta dos decimales.', 1;
            IF @CargoId IS NOT NULL AND NOT EXISTS (
                SELECT 1 FROM dbo.CargoTrabajadorHYT AS cargo WITH (HOLDLOCK)
                WHERE cargo.CargoId = @CargoId AND (cargo.Activo = 1 OR EXISTS (
                    SELECT 1 FROM dbo.TrabajadorHYT AS trabajador
                    WHERE trabajador.TrabajadorId = @TrabajadorId AND trabajador.CargoId = @CargoId)))
                THROW 51000, N'Seleccione un cargo activo.', 1;
        END;

        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @TrabajadorId IS NULL
        BEGIN
            INSERT INTO dbo.TrabajadorHYT (Documento, NombreCompleto, Tipo, FechaIngreso, Banco, NumeroCuenta, Activo, CargoId, SueldoMensual)
            VALUES (@Documento, @NombreCompleto, @Tipo, @FechaIngreso, CASE WHEN @ActualizarCuenta = 1 THEN @Banco END,
                    CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta END, @Activo,
                    CASE WHEN @ActualizarPuesto = 1 THEN @CargoId END, CASE WHEN @ActualizarPuesto = 1 THEN @SueldoMensual END);
            SET @TrabajadorId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.TrabajadorHYT
            SET Documento = @Documento, NombreCompleto = @NombreCompleto, Activo = @Activo,
                Tipo = @Tipo, FechaIngreso = @FechaIngreso,
                CargoId = CASE WHEN @ActualizarPuesto = 1 THEN @CargoId ELSE CargoId END,
                SueldoMensual = CASE WHEN @ActualizarPuesto = 1 THEN @SueldoMensual ELSE SueldoMensual END,
                Banco = CASE WHEN @ActualizarCuenta = 1 THEN @Banco ELSE Banco END,
                NumeroCuenta = CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta ELSE NumeroCuenta END
            WHERE TrabajadorId = @TrabajadorId AND VersionFila = @VersionFila;

            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador cambio o ya no existe. Vuelva a abrir su ficha.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @TrabajadorId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Documento AS Documento, @NombreCompleto AS NombreCompleto, @Activo AS Activo,
                   @Tipo AS Tipo, @FechaIngreso AS FechaIngreso, @ActualizarCuenta AS ActualizoCuenta,
                   @ActualizarPuesto AS ActualizoPuesto, @CargoId AS CargoId
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'TrabajadorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;

        SELECT @TrabajadorId AS TrabajadorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

COMMIT TRANSACTION;
GO
