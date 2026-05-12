USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/05/2026
-- Description:   Migra desafio/perfil publico para usar TiposDeporteSuperMaestro en lugar de TiposDeporte por negocio.
-- =============================================
-- Firma: Codex - 11/05/2026 | Migracion estructural del modulo Desafios hacia catalogo global de deportes (TiposDeporteSuperMaestro).

BEGIN TRY
    BEGIN TRANSACTION;

    -- Convierte datos legacy de TiposDeporte.Id hacia TiposDeporteSuperMaestro.Id en perfil publico.
    UPDATE upp
    SET upp.IdDeporteDesafio = td.TipoDeporteSuperId
    FROM dbo.UsuariosPublicosPerfil upp
    INNER JOIN dbo.TiposDeporte td
        ON td.Id = upp.IdDeporteDesafio
    WHERE upp.IdDeporteDesafio IS NOT NULL
      AND td.TipoDeporteSuperId IS NOT NULL;

    -- Convierte datos legacy de TiposDeporte.Id hacia TiposDeporteSuperMaestro.Id en desafios.
    UPDATE d
    SET d.IdDeporte = td.TipoDeporteSuperId
    FROM dbo.Desafio d
    INNER JOIN dbo.TiposDeporte td
        ON td.Id = d.IdDeporte
    WHERE d.IdDeporte IS NOT NULL
      AND td.TipoDeporteSuperId IS NOT NULL;

    -- Limpia registros invalidos que no existan en catalogo global para poder aplicar FK.
    UPDATE dbo.UsuariosPublicosPerfil
    SET IdDeporteDesafio = NULL
    WHERE IdDeporteDesafio IS NOT NULL
      AND NOT EXISTS (
          SELECT 1
          FROM dbo.TiposDeporteSuperMaestro tsm
          WHERE tsm.Id = dbo.UsuariosPublicosPerfil.IdDeporteDesafio
      );

    IF EXISTS (
        SELECT 1
        FROM dbo.Desafio d
        WHERE NOT EXISTS (
            SELECT 1
            FROM dbo.TiposDeporteSuperMaestro tsm
            WHERE tsm.Id = d.IdDeporte
        )
    )
        RAISERROR('Existen desafios con IdDeporte invalido para TiposDeporteSuperMaestro. Corrige los datos antes de aplicar la FK.', 16, 1);

    IF OBJECT_ID('dbo.FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio', 'F') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.UsuariosPublicosPerfil
        DROP CONSTRAINT FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio;
    END

    IF OBJECT_ID('dbo.FK_UsuariosPublicosPerfil_TiposDeporteSuperMaestro_IdDeporteDesafio', 'F') IS NULL
    BEGIN
        ALTER TABLE dbo.UsuariosPublicosPerfil WITH CHECK
        ADD CONSTRAINT FK_UsuariosPublicosPerfil_TiposDeporteSuperMaestro_IdDeporteDesafio
        FOREIGN KEY(IdDeporteDesafio) REFERENCES dbo.TiposDeporteSuperMaestro(Id);

        ALTER TABLE dbo.UsuariosPublicosPerfil
        CHECK CONSTRAINT FK_UsuariosPublicosPerfil_TiposDeporteSuperMaestro_IdDeporteDesafio;
    END

    IF OBJECT_ID('dbo.FK_Desafio_TiposDeporte_IdDeporte', 'F') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.Desafio
        DROP CONSTRAINT FK_Desafio_TiposDeporte_IdDeporte;
    END

    IF OBJECT_ID('dbo.FK_Desafio_TiposDeporteSuperMaestro_IdDeporte', 'F') IS NULL
    BEGIN
        ALTER TABLE dbo.Desafio WITH CHECK
        ADD CONSTRAINT FK_Desafio_TiposDeporteSuperMaestro_IdDeporte
        FOREIGN KEY(IdDeporte) REFERENCES dbo.TiposDeporteSuperMaestro(Id);

        ALTER TABLE dbo.Desafio
        CHECK CONSTRAINT FK_Desafio_TiposDeporteSuperMaestro_IdDeporte;
    END

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH
GO
