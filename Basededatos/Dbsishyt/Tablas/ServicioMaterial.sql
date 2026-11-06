-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ServicioMaterial para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ServicioMaterial', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ServicioMaterial
    (
        MaterialId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ServicioMaterial PRIMARY KEY,
        Posicion NVARCHAR(60) NULL,
        CodigoInterno NVARCHAR(60) NOT NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        UnidadMedidaId INT NOT NULL,
        EspecialidadId INT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_Material_Unidad FOREIGN KEY (UnidadMedidaId) REFERENCES dbo.UnidadMedida(Id),
        CONSTRAINT FK_Material_Especialidad FOREIGN KEY (EspecialidadId) REFERENCES dbo.Especialidad(Id),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ServicioMaterial_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ServicioMaterial_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ServicioMaterial_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
