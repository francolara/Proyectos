-- Firma: FRANCO LARA | Create date: 10/11/2026 | Configura la cuenta de cargo en soles con control de versión y auditoría.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Registra la moneda seleccionada y valida el dígito de moneda de la cuenta.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TelecreditoHYT_Guardar
    @UsuarioId NVARCHAR(128), @TipoCuenta CHAR(1), @NumeroCuenta VARCHAR(30), @VersionFila BINARY(8)=NULL, @Moneda CHAR(3)=N'PEN'
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.CONSULTAR')=0
            OR dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.EDITAR')=0
            THROW 51000,N'No tiene permiso para configurar Telecrédito.',1;
        IF @TipoCuenta IS NULL OR @TipoCuenta NOT IN ('C','M') OR @NumeroCuenta IS NULL
            OR DATALENGTH(@NumeroCuenta)<>13 OR @NumeroCuenta COLLATE Latin1_General_100_BIN2 LIKE '%[^0-9]%'
            THROW 51000,N'Indica una cuenta de cargo BCP de 13 dígitos.',1;
        IF @Moneda IS NULL OR @Moneda NOT IN ('PEN','USD')
            THROW 51000,N'Selecciona soles o dólares.',1;
        IF SUBSTRING(@NumeroCuenta,11,1)<>CASE @Moneda WHEN 'PEN' THEN '0' ELSE '1' END
            THROW 51000,N'El dígito 11 del número de cuenta no coincide con la moneda seleccionada.',1;

        BEGIN TRANSACTION;
        DECLARE @Actual BINARY(8);
        SELECT @Actual=cuenta.VersionFila FROM dbo.TelecreditoHYT AS cuenta WITH (UPDLOCK,HOLDLOCK) WHERE cuenta.Id=1;
        IF @Actual IS DISTINCT FROM @VersionFila
            THROW 51000,N'La configuración cambió. Vuelve a abrir la ficha.',1;
        IF @Actual IS NULL
            INSERT INTO dbo.TelecreditoHYT(Id,TipoCuenta,NumeroCuenta,Moneda) VALUES (1,@TipoCuenta,@NumeroCuenta,@Moneda);
        ELSE
            UPDATE dbo.TelecreditoHYT SET TipoCuenta=@TipoCuenta,NumeroCuenta=@NumeroCuenta,Moneda=@Moneda WHERE Id=1;
        DECLARE @Correlacion UNIQUEIDENTIFIER=NEWID();
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT @TipoCuenta AS TipoCuenta,RIGHT(@NumeroCuenta,4) AS UltimosCuatro,
            @Moneda AS Moneda FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Correlacion,N'TelecreditoHYT',N'1',N'GUARDAR',@CambiosJson=@Cambios;
        COMMIT TRANSACTION;
        SELECT CONVERT(BIGINT,1);
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
