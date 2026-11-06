-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Inicializa roles y permisos sin usuarios ni contrasenas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-ADMIN')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-ADMIN', N'Administrador', N'ADMINISTRADOR', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-COMERCIAL')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-COMERCIAL', N'Comercial', N'COMERCIAL', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-OPERACIONES')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-OPERACIONES', N'Operaciones', N'OPERACIONES', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-SEGUIMIENTO')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-SEGUIMIENTO', N'Seguimiento documental', N'SEGUIMIENTO DOCUMENTAL', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-CONSULTA')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-CONSULTA', N'Consulta', N'CONSULTA', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CLIENTES.CONSULTAR', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CLIENTES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CLIENTES.EDITAR', N'CLIENTES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TIENDAS.CONSULTAR', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TIENDAS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TIENDAS.EDITAR', N'TIENDAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'MATERIALES.CONSULTAR', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'MATERIALES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'MATERIALES.EDITAR', N'MATERIALES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'UNIDADES.CONSULTAR', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'UNIDADES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'UNIDADES.EDITAR', N'UNIDADES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ZONAS.CONSULTAR', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ZONAS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ZONAS.EDITAR', N'ZONAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TARIFARIOS.CONSULTAR', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TARIFARIOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TARIFARIOS.EDITAR', N'TARIFARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'PRESUPUESTOS.CONSULTAR', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'PRESUPUESTOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'PRESUPUESTOS.EDITAR', N'PRESUPUESTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'OC_CLIENTE.CONSULTAR', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'OC_CLIENTE.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'OC_CLIENTE.EDITAR', N'OC_CLIENTE.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'GUIAS.CONSULTAR', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'GUIAS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'GUIAS.EDITAR', N'GUIAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIRMACIONES.CONSULTAR', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIRMACIONES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIRMACIONES.EDITAR', N'CONFIRMACIONES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'FACTURADO.CONSULTAR', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'FACTURADO.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'FACTURADO.EDITAR', N'FACTURADO.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'REPORTES.CONSULTAR', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'AUDITORIA.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'AUDITORIA.CONSULTAR', N'AUDITORIA.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'USUARIOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'USUARIOS.CONSULTAR', N'USUARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'USUARIOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'USUARIOS.EDITAR', N'USUARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIGURACION.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIGURACION.CONSULTAR', N'CONFIGURACION.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIGURACION.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIGURACION.EDITAR', N'CONFIGURACION.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'PRESUPUESTOS.APROBAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'PRESUPUESTOS.APROBAR', N'PRESUPUESTOS.APROBAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'GUIAS.ACEPTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'GUIAS.ACEPTAR', N'GUIAS.ACEPTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'DOCUMENTOS.ANULAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'DOCUMENTOS.ANULAR', N'DOCUMENTOS.ANULAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ADJUNTOS.CONSULTAR', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ADJUNTOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ADJUNTOS.EDITAR', N'ADJUNTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CLIENTES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CLIENTES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TIENDAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TIENDAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'MATERIALES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'MATERIALES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'UNIDADES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'UNIDADES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ZONAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ZONAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TARIFARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'PRESUPUESTOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'PRESUPUESTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'OC_CLIENTE.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'OC_CLIENTE.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'GUIAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'GUIAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIRMACIONES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIRMACIONES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'FACTURADO.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'FACTURADO.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'AUDITORIA.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'AUDITORIA.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'USUARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'USUARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'USUARIOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'USUARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIGURACION.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIGURACION.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIGURACION.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIGURACION.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'PRESUPUESTOS.APROBAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'PRESUPUESTOS.APROBAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'GUIAS.ACEPTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'GUIAS.ACEPTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'DOCUMENTOS.ANULAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'DOCUMENTOS.ANULAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ADJUNTOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ADJUNTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'CLIENTES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'CLIENTES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TIENDAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TIENDAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'MATERIALES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'MATERIALES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TARIFARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'PRESUPUESTOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'PRESUPUESTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'OC_CLIENTE.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'OC_CLIENTE.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'GUIAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'GUIAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'CONFIRMACIONES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'CONFIRMACIONES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'FACTURADO.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'FACTURADO.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'ADJUNTOS.CONSULTAR');
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO
