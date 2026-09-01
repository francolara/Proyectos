
GO
/****** Object:  StoredProcedure [dbo].[Sp_Home_SolicitarAltaClub]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 21_Altas_Clubes.sql (linea 41)
-- Firma: FRANCO LARA - 21/07/2026 | Registra el plan comercial publico seleccionado al solicitar el alta.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_SolicitarAltaClub]
    @NombreContacto NVARCHAR(200),
    @Telefono NVARCHAR(30),
    @Correo NVARCHAR(200),
    @RelacionClub NVARCHAR(80),
    @NombreClub NVARCHAR(200),
    @Pais NVARCHAR(80),
    @ProvinciaEstado NVARCHAR(120),
    @Ciudad NVARCHAR(120),
    @Direccion NVARCHAR(250),
    @PlanComercial NVARCHAR(20) = N'PRUEBA'
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Secuencia INT;
        DECLARE @Codigo NVARCHAR(30);

        SET @PlanComercial = CASE UPPER(LTRIM(RTRIM(COALESCE(@PlanComercial, N''))))
                                  WHEN N'ESENCIAL' THEN N'ESENCIAL'
                                  WHEN N'EMPRENDEDOR' THEN N'ESENCIAL'
                                  WHEN N'PRO' THEN N'PRO'
                                  WHEN N'PROFESIONAL' THEN N'PRO'
                                  ELSE N'PRUEBA'
                              END;

        SELECT @Secuencia = COUNT(1) + 1
        FROM dbo.SolicitudesAltaClub
        WHERE CAST(FechaRegistro AS DATE) = CAST(SYSUTCDATETIME() AS DATE);

        SET @Codigo = CONCAT(
            N'CLUB-',
            CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112),
            N'-',
            RIGHT(CONCAT(N'0000', CONVERT(NVARCHAR(10), @Secuencia)), 4)
        );

        INSERT INTO dbo.SolicitudesAltaClub
        (
            CodigoSolicitud, NombreContacto, Telefono, Correo, RelacionClub, NombreClub,
            Pais, ProvinciaEstado, Ciudad, Direccion, PlanComercial, Estado, FechaRegistro
        )
        VALUES
        (
            @Codigo, @NombreContacto, @Telefono, @Correo, @RelacionClub, @NombreClub,
            @Pais, @ProvinciaEstado, @Ciudad, @Direccion, @PlanComercial, 1, SYSUTCDATETIME()
        );

        SELECT @Codigo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
