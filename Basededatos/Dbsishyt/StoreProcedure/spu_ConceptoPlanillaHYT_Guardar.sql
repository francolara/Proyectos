-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registra y actualiza conceptos con control de concurrencia y auditoría.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConceptoPlanillaHYT_Guardar
    @UsuarioId NVARCHAR(128), @ConceptoId BIGINT=NULL, @Nombre NVARCHAR(100),
    @Naturaleza VARCHAR(5), @Calculo VARCHAR(10), @Valor DECIMAL(20,4)=NULL,
    @Activo BIT, @VersionFila BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'TRABAJADORES.EDITAR')=0
            THROW 51000,N'No tiene permiso para guardar conceptos de planilla.',1;
        SET @Nombre=NULLIF(TRIM(@Nombre),N'');
        IF @Nombre IS NULL OR @Naturaleza IS NULL OR @Naturaleza NOT IN ('CARGO','ABONO')
            OR @Calculo IS NULL OR @Calculo NOT IN ('IMPORTE','PORCENTAJE') OR @Activo IS NULL
            OR @Valor<0 OR (@Calculo='PORCENTAJE' AND @Valor>100)
            OR (@Calculo='IMPORTE' AND @Valor<>ROUND(@Valor,2))
            THROW 51000,N'Revisa la naturaleza y el valor del concepto.',1;

        BEGIN TRANSACTION;
        IF @ConceptoId IS NULL
        BEGIN
            INSERT INTO dbo.ConceptoPlanillaHYT(Nombre,Naturaleza,Calculo,Valor,Activo)
            VALUES (@Nombre,@Naturaleza,@Calculo,@Valor,@Activo);
            SET @ConceptoId=SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.ConceptoPlanillaHYT SET Nombre=@Nombre,Naturaleza=@Naturaleza,Calculo=@Calculo,Valor=@Valor,Activo=@Activo
            WHERE ConceptoId=@ConceptoId AND VersionFila=@VersionFila AND (Regimen IS NULL OR @Naturaleza='CARGO');
            IF @@ROWCOUNT=0
                THROW 51000,N'El concepto cambió o el régimen previsional requiere un cargo. Vuelve a abrir la ficha.',1;
        END;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@ConceptoId), @Correlacion UNIQUEIDENTIFIER=NEWID();
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT @Nombre AS Nombre,@Naturaleza AS Naturaleza,@Calculo AS Calculo,
            @Valor AS Valor,@Activo AS Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Correlacion,N'ConceptoPlanillaHYT',@Registro,N'GUARDAR',@CambiosJson=@Cambios;
        COMMIT TRANSACTION;
        SELECT @ConceptoId AS ConceptoId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
