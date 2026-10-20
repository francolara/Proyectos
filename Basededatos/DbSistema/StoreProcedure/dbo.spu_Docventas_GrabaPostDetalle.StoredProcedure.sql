-- =============================================
-- Author:        FRANCO LARA
-- Create date:   14/10/2026
-- Description:   Graba referencias y vendedor despues de persistir cabecera y detalle de ventas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_GrabaPostDetalle
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10),
    @IdDocumento VARCHAR(10),
    @IdSerie VARCHAR(20),
    @IdDocVentas VARCHAR(20),
    @IdCliente VARCHAR(20),
    @IdVendedorCampo VARCHAR(20),
    @GrabaFacturaGuia BIT,
    @ReferenciasXml XML
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Referencias TABLE
        (
            Item INT NOT NULL,
            TipoDocumento VARCHAR(10) NOT NULL,
            Serie VARCHAR(20) NOT NULL,
            Numero VARCHAR(20) NOT NULL,
            IndImportado VARCHAR(1) NOT NULL
        );

        INSERT INTO @Referencias (Item, TipoDocumento, Serie, Numero, IndImportado)
        SELECT
            Referencia.Nodo.value('@Item', 'INT'),
            Referencia.Nodo.value('@TipoDocumento', 'VARCHAR(10)'),
            Referencia.Nodo.value('@Serie', 'VARCHAR(20)'),
            Referencia.Nodo.value('@Numero', 'VARCHAR(20)'),
            Referencia.Nodo.value('@IndImportado', 'VARCHAR(1)')
        FROM @ReferenciasXml.nodes('/Referencias/Referencia') AS Referencia(Nodo)
        WHERE LTRIM(RTRIM(Referencia.Nodo.value('@TipoDocumento', 'VARCHAR(10)'))) <> ''
          AND LTRIM(RTRIM(Referencia.Nodo.value('@Serie', 'VARCHAR(20)'))) <> ''
          AND LTRIM(RTRIM(Referencia.Nodo.value('@Numero', 'VARCHAR(20)'))) <> '';

        DELETE FROM dbo.DocReferencia
        WHERE IdEmpresa = @IdEmpresa
          AND IdSucursal = @IdSucursal
          AND TipoDocOrigen = @IdDocumento
          AND SerieDocOrigen = @IdSerie
          AND NumDocOrigen = @IdDocVentas;

        IF @GrabaFacturaGuia = 1 AND @IdDocumento = '86'
        BEGIN
            DELETE FROM dbo.DocReferencia
            WHERE IdEmpresa = @IdEmpresa
              AND IdSucursal = @IdSucursal
              AND TipoDocOrigen = '01'
              AND TipoDocReferencia = @IdDocumento
              AND SerieDocReferencia = @IdSerie
              AND NumDocReferencia = @IdDocVentas;
        END;

        INSERT INTO dbo.DocReferencia
        (
            Item, TipoDocOrigen, SerieDocOrigen, NumDocOrigen,
            TipoDocReferencia, SerieDocReferencia, NumDocReferencia,
            IdEmpresa, IdSucursal, IdAno, IndImportado
        )
        SELECT
            Referencia.Item, @IdDocumento, @IdSerie, @IdDocVentas,
            Referencia.TipoDocumento, Referencia.Serie, Referencia.Numero,
            @IdEmpresa, @IdSucursal, '', Referencia.IndImportado
        FROM @Referencias AS Referencia;

        IF @GrabaFacturaGuia = 1 AND @IdDocumento = '86'
        BEGIN
            INSERT INTO dbo.DocReferencia
            (
                Item, TipoDocOrigen, SerieDocOrigen, NumDocOrigen,
                TipoDocReferencia, SerieDocReferencia, NumDocReferencia,
                IdEmpresa, IdSucursal, IdAno, IndImportado
            )
            SELECT
                Referencia.Item, '01', Referencia.Serie, Referencia.Numero,
                @IdDocumento, @IdSerie, @IdDocVentas,
                @IdEmpresa, @IdSucursal, '', Referencia.IndImportado
            FROM @Referencias AS Referencia
            WHERE Referencia.TipoDocumento = '01';
        END;

        IF LTRIM(RTRIM(@IdVendedorCampo)) <> ''
        BEGIN
            UPDATE dbo.Clientes
            SET IdVendedorCampo = @IdVendedorCampo
            WHERE IdEmpresa = @IdEmpresa
              AND IdCliente = @IdCliente;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(),
               @ErrorSeverity = ERROR_SEVERITY(),
               @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
