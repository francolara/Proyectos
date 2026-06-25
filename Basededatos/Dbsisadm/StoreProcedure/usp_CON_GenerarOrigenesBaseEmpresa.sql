-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Registra los origenes contables base para una empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarOrigenesBaseEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR (N'La empresa indicada no existe.', 16, 1)
            RETURN
        END;

        MERGE dbo.CON_Origen AS destino
        USING
        (
            SELECT
                @IdEmpresa AS IdEmpresa,
                origenBase.CodigoOrigen,
                origenBase.NombreOrigen,
                origenBase.ModuloOrigen,
                origenBase.PermiteRegistroManual
            FROM
            (
                VALUES
                    ('ASI', N'Asiento manual', N'CONTABILIDAD', 1),
                    ('COM', N'Compras', N'COMPRAS', 0),
                    ('VEN', N'Ventas', N'VENTAS', 0),
                    ('BAN', N'Bancos', N'TESORERIA', 0),
                    ('47', N'Aplicaciones N/C', N'CONTABILIDAD', 1),
                    ('CIE', N'Cierre contable', N'CONTABILIDAD', 0)
            ) AS origenBase (CodigoOrigen, NombreOrigen, ModuloOrigen, PermiteRegistroManual)
        ) AS fuente
            ON destino.IdEmpresa = fuente.IdEmpresa
           AND destino.CodigoOrigen = fuente.CodigoOrigen
        WHEN MATCHED THEN
            UPDATE
            SET
                destino.NombreOrigen = fuente.NombreOrigen,
                destino.ModuloOrigen = fuente.ModuloOrigen,
                destino.PermiteRegistroManual = fuente.PermiteRegistroManual,
                destino.Estado = 1,
                destino.UsuarioRegistro = COALESCE(@UsuarioRegistro, destino.UsuarioRegistro)
        WHEN NOT MATCHED BY TARGET THEN
            INSERT
            (
                IdEmpresa,
                CodigoOrigen,
                NombreOrigen,
                ModuloOrigen,
                PermiteRegistroManual,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                fuente.IdEmpresa,
                fuente.CodigoOrigen,
                fuente.NombreOrigen,
                fuente.ModuloOrigen,
                fuente.PermiteRegistroManual,
                1,
                @UsuarioRegistro
            );

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
