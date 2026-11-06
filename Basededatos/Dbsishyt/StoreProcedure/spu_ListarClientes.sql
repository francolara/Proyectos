-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta clientes autorizada, pagina SQL fija de 20 y total.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ListarClientes
    @UsuarioId NVARCHAR(128),
    @Pagina INT = 1,
    @Busqueda NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Pagina < 1 OR @Pagina > 100000000
            THROW 51003, N'Pagina fuera de rango.', 1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        )
            THROW 51004, N'No tiene permiso para consultar clientes.', 1;

        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL
            SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,
                N'~', N'~~'), N'%', N'~%'), N'_', N'~_'), N'[', N'~[') + N'%';

        -- Una materializacion por solicitud mantiene pagina y total coherentes.
        SELECT cliente.ClienteId, cliente.Codigo, cliente.TipoEntidad, cliente.NumeroIdentificacion,
            cliente.RazonSocial, cliente.Nombres, cliente.Apellidos, cliente.Activo,
            cliente.FechaRegistro, cliente.UsuarioRegistro, cliente.FechaActualizacion, cliente.UsuarioActualizacion
        INTO #Clientes
        FROM dbo.Cliente AS cliente
        WHERE @Patron IS NULL OR cliente.Codigo LIKE @Patron ESCAPE N'~'
            OR cliente.RazonSocial LIKE @Patron ESCAPE N'~'
            OR CONCAT(cliente.Nombres, N' ', cliente.Apellidos) LIKE @Patron ESCAPE N'~'
            OR cliente.NumeroIdentificacion LIKE @Patron ESCAPE N'~';

        SELECT listado.ClienteId, listado.Codigo, listado.TipoEntidad, listado.NumeroIdentificacion,
            listado.RazonSocial, listado.Nombres, listado.Apellidos, listado.Activo,
            listado.FechaRegistro, listado.UsuarioRegistro, listado.FechaActualizacion, listado.UsuarioActualizacion
        FROM #Clientes AS listado
        ORDER BY listado.Codigo, listado.ClienteId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS TotalRegistros FROM #Clientes AS listado;
        DROP TABLE #Clientes;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Clientes') IS NOT NULL DROP TABLE #Clientes;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
