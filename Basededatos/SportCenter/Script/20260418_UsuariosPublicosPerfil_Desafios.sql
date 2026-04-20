USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Altera UsuariosPublicosPerfil para desafios y siembra catalogo NivelDesafio.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Agrega distrito y WhatsApp del equipo como informacion operativa del modulo Desafios.
-- =============================================
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'BuscarDesafios') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD BuscarDesafios BIT NOT NULL CONSTRAINT DF_UsuariosPublicosPerfil_BuscarDesafios DEFAULT ((0));
END
GO
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'IdDeporteDesafio') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD IdDeporteDesafio INT NULL;
END
GO
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'IdNivelDesafio') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD IdNivelDesafio INT NULL;
END
GO
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'ObservacionDesafio') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD ObservacionDesafio NVARCHAR(500) NULL;
END
GO
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'DetalleEquipo') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD DetalleEquipo NVARCHAR(1000) NULL;
END
GO
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'CodigoUbigeoEquipo') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD CodigoUbigeoEquipo CHAR(6) NULL;
END
GO
IF COL_LENGTH('dbo.UsuariosPublicosPerfil', 'WhatsappEquipo') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil ADD WhatsappEquipo NVARCHAR(30) NULL;
END
GO
IF OBJECT_ID('dbo.FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio', 'F') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil  WITH CHECK ADD CONSTRAINT FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio
    FOREIGN KEY(IdDeporteDesafio) REFERENCES dbo.TiposDeporte (Id);
    ALTER TABLE dbo.UsuariosPublicosPerfil CHECK CONSTRAINT FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio;
END
GO
IF OBJECT_ID('dbo.FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeoEquipo', 'F') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil  WITH CHECK ADD CONSTRAINT FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeoEquipo
    FOREIGN KEY(CodigoUbigeoEquipo) REFERENCES dbo.UbigeoDistritos (CodigoUbigeo);
    ALTER TABLE dbo.UsuariosPublicosPerfil CHECK CONSTRAINT FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeoEquipo;
END
GO
IF OBJECT_ID('dbo.FK_UsuariosPublicosPerfil_NivelDesafio_IdNivelDesafio', 'F') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosPublicosPerfil  WITH CHECK ADD CONSTRAINT FK_UsuariosPublicosPerfil_NivelDesafio_IdNivelDesafio
    FOREIGN KEY(IdNivelDesafio) REFERENCES dbo.NivelDesafio (IdNivel);
    ALTER TABLE dbo.UsuariosPublicosPerfil CHECK CONSTRAINT FK_UsuariosPublicosPerfil_NivelDesafio_IdNivelDesafio;
END
GO
IF NOT EXISTS (SELECT 1 FROM dbo.NivelDesafio WHERE Nombre = N'Basico')
    INSERT INTO dbo.NivelDesafio (Nombre, Activo, Orden) VALUES (N'Basico', 1, 1);
GO
IF NOT EXISTS (SELECT 1 FROM dbo.NivelDesafio WHERE Nombre = N'Intermedio')
    INSERT INTO dbo.NivelDesafio (Nombre, Activo, Orden) VALUES (N'Intermedio', 1, 2);
GO
IF NOT EXISTS (SELECT 1 FROM dbo.NivelDesafio WHERE Nombre = N'Competitivo')
    INSERT INTO dbo.NivelDesafio (Nombre, Activo, Orden) VALUES (N'Competitivo', 1, 3);
GO
