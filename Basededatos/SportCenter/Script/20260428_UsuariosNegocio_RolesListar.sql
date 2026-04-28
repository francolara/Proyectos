USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 28/04/2026 | Crea/actualiza SP de catalogo de roles de negocio para evitar listas hardcodeadas en la capa web.

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_RolesListar
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        ;WITH RolesBase AS
        (
            SELECT DISTINCT rp.RolNegocio
            FROM dbo.RolesNegocioPermiso rp
            WHERE rp.RolNegocio IS NOT NULL
        )
        SELECT
            CAST(rb.RolNegocio AS NVARCHAR(20)) AS Value,
            CASE rb.RolNegocio
                WHEN 1 THEN N'Administrador'
                WHEN 2 THEN N'Trabajador'
                WHEN 3 THEN N'Recepcion'
                WHEN 4 THEN N'Caja'
                WHEN 5 THEN N'Supervisor'
                ELSE CONCAT(N'Rol ', CONVERT(NVARCHAR(20), rb.RolNegocio))
            END AS Text
        FROM RolesBase rb
        ORDER BY rb.RolNegocio;

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
