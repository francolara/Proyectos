-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Obtiene usuario, persona, ubicación y perfil en una sola llamada.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUsuarios_Obtener
    @IdEmpresa CHAR(2),
    @IdUsuario CHAR(8),
    @CodSistema CHAR(2)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            ISNULL(U.IndCompra, 0) AS IndCompra,
            U.IndJefe,
            U.IdUsuario,
            COALESCE(PX.IdPerfil, U.IdPerfil, '') AS IdPerfil,
            U.VarUsuario,
            U.VarPass,
            U.SerieEtiquetera,
            U.ImpresoraLetras,
            P.GlsPersona,
            ISNULL(P.Direccion, '') AS Direccion,
            ISNULL(P.IdDistrito, '') AS IdDistrito,
            ISNULL(UB.IdDpto, '') AS IdDpto,
            ISNULL(UB.IdProv, '') AS IdProv,
            COALESCE(PX.GlsPerfil, PU.GlsPerfil, '') AS GlsPerfil,
            ISNULL(PA.GlsPais, '') AS GlsPais,
            ISNULL(UD.GlsUbigeo, '') AS GlsDepa,
            ISNULL(UP.GlsUbigeo, '') AS GlsProv,
            ISNULL(UB.GlsUbigeo, '') AS GlsDistrito
        FROM dbo.Usuarios AS U
        INNER JOIN dbo.Personas AS P
            ON P.IdPersona = U.IdUsuario
        LEFT JOIN dbo.Ubigeo AS UB
            ON UB.IdDistrito = P.IdDistrito
        LEFT JOIN dbo.Perfil AS PU
            ON PU.IdEmpresa = U.IdEmpresa
           AND PU.IdPerfil = U.IdPerfil
           AND PU.CodSistema = @CodSistema
        LEFT JOIN dbo.Ubigeo AS UD
            ON UD.IdPais = '02001'
           AND UD.IdDpto = UB.IdDpto
           AND UD.IdProv = '00'
           AND UD.IdDist = '00'
        LEFT JOIN dbo.Ubigeo AS UP
            ON UP.IdPais = '02001'
           AND UP.IdDpto = UB.IdDpto
           AND UP.IdProv = UB.IdProv
           AND UP.IdDist = '00'
        OUTER APPLY
        (
            SELECT TOP (1)
                PF.IdPerfil,
                PF.GlsPerfil
            FROM dbo.PerfilesPorUsuario AS PPU
            INNER JOIN dbo.Perfil AS PF
                ON PF.IdEmpresa = PPU.IdEmpresa
               AND PF.IdPerfil = PPU.IdPerfil
               AND PF.CodSistema = PPU.CodSistema
            WHERE PPU.IdEmpresa = U.IdEmpresa
              AND PPU.IdUsuario = U.IdUsuario
              AND PPU.CodSistema = @CodSistema
            ORDER BY PF.IdPerfil
        ) AS PX
        OUTER APPLY
        (
            SELECT TOP (1) D.GlsDato AS GlsPais
            FROM dbo.Datos AS D
            WHERE D.IdDato = '02001'
            ORDER BY D.IdTipoDatos
        ) AS PA
        WHERE U.IdEmpresa = @IdEmpresa
          AND U.IdUsuario = @IdUsuario;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
