-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Traslada Código SAP al cliente, conservando los registros anteriores.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
BEGIN TRY
    IF COL_LENGTH(N'dbo.Cliente',N'CodigoSAP') IS NULL
    BEGIN
        BEGIN TRANSACTION;
        EXEC(N'ALTER TABLE dbo.Cliente ADD CodigoSAP NVARCHAR(60) NULL;');
        -- La migración se ejecuta únicamente al crear la columna. No restaura códigos borrados posteriormente.
        -- Prioriza el registro activo y después el último registrado; conserva todas las fuentes históricas.
        EXEC(N'UPDATE cliente SET CodigoSAP=fuente.CodigoSAPProveedor
            FROM dbo.Cliente AS cliente
            CROSS APPLY (SELECT TOP (1) proveedor.CodigoSAPProveedor
                FROM dbo.ProveedorCliente AS proveedor
                WHERE proveedor.ClienteId=cliente.ClienteId AND NULLIF(LTRIM(RTRIM(proveedor.CodigoSAPProveedor)),N'''') IS NOT NULL
                ORDER BY proveedor.Activo DESC,proveedor.FechaRegistro DESC,proveedor.Id DESC) AS fuente;');
        COMMIT TRANSACTION;
    END;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
    SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
    RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
END CATCH;
GO
