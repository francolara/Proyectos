
GO
/****** Object:  StoredProcedure [dbo].[Sp_Maestros_Monedas_Crear]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 36_Maestros_PorNegocio_MonedasSuper.sql (linea 317)
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   04/04/2026
-- Description:   Habilita monedas del supermaestro para la configuracion del negocio.
-- Firma:         FRANCO LARA - 06/10/2026 | Permite varias monedas por negocio sin duplicar el mismo codigo canonico.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Maestros_Monedas_Crear]
    @NegocioId INT,
    @MonedaSuperId INT,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.MonedasSuperMaestro WHERE Id = @MonedaSuperId AND Activo = 1)
            RAISERROR('La moneda del supermaestro no es valida.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE NegocioId = @NegocioId AND MonedaSuperId = @MonedaSuperId)
            RAISERROR('La moneda ya esta registrada para este negocio.', 16, 1);

        INSERT INTO dbo.Monedas (NegocioId, MonedaSuperId, Codigo, Nombre, Simbolo, Activo, FechaCreacion, UsuarioCreacion)
        SELECT @NegocioId, ms.Id, ms.Codigo, ms.Nombre, ms.Simbolo, @Activo, SYSUTCDATETIME(), @Usuario
        FROM dbo.MonedasSuperMaestro ms
        WHERE ms.Id = @MonedaSuperId;

        DECLARE @Id INT = SCOPE_IDENTITY();
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'MAESTROS', @Accion = N'CREATE', @Entidad = N'Moneda', @EntidadId = @Id, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
