-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/06/2026
-- Description:   Crea tabla de relaciones de espacios compartidos, luego la evoluciona para distinguir bloqueo directo y espacios compuestos por componentes.
-- =============================================
-- DELETE FROM EspaciosDeportivosCompartidos
-- DROP TABLE EspaciosDeportivosCompartidos
-- SELECT * FROM EspaciosDeportivosCompartidos
IF OBJECT_ID(N'dbo.EspaciosDeportivosCompartidos', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EspaciosDeportivosCompartidos](
        [Id] [int] IDENTITY(1,1) NOT NULL,
        [EspacioDeportivoId] [int] NOT NULL,
        [EspacioRelacionadoId] [int] NOT NULL,
        [TipoRelacion] [nvarchar](40) NOT NULL CONSTRAINT [DF_EspaciosDeportivosCompartidos_TipoRelacion] DEFAULT (N'DIRECTO'),
        [Activo] [bit] NOT NULL CONSTRAINT [DF_EspaciosDeportivosCompartidos_Activo] DEFAULT ((1)),
        [FechaCreacion] [datetime2](7) NOT NULL CONSTRAINT [DF_EspaciosDeportivosCompartidos_FechaCreacion] DEFAULT (SYSUTCDATETIME()),
        [UsuarioCreacion] [nvarchar](200) NULL,
        [FechaActualizacion] [datetime2](7) NULL,
        [UsuarioActualizacion] [nvarchar](200) NULL,
        CONSTRAINT [PK_EspaciosDeportivosCompartidos] PRIMARY KEY CLUSTERED ([Id] ASC),
        CONSTRAINT [CK_EspaciosDeportivosCompartidos_TipoRelacion] CHECK ([TipoRelacion] IN (N'DIRECTO', N'COMPUESTO_COMPONENTE')),
        CONSTRAINT [FK_EspaciosDeportivosCompartidos_Espacio] FOREIGN KEY([EspacioDeportivoId]) REFERENCES [dbo].[EspaciosDeportivos]([Id]),
        CONSTRAINT [FK_EspaciosDeportivosCompartidos_EspacioRelacionado] FOREIGN KEY([EspacioRelacionadoId]) REFERENCES [dbo].[EspaciosDeportivos]([Id])
    );
END;

IF COL_LENGTH(N'dbo.EspaciosDeportivosCompartidos', N'TipoRelacion') IS NULL
BEGIN
    ALTER TABLE dbo.EspaciosDeportivosCompartidos
    ADD TipoRelacion NVARCHAR(40) NOT NULL
        CONSTRAINT DF_EspaciosDeportivosCompartidos_TipoRelacion DEFAULT (N'DIRECTO');
END;

UPDATE dbo.EspaciosDeportivosCompartidos
SET TipoRelacion = N'DIRECTO'
WHERE TipoRelacion IS NULL
   OR LTRIM(RTRIM(TipoRelacion)) = N'';

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_EspaciosDeportivosCompartidos_TipoRelacion'
      AND parent_object_id = OBJECT_ID(N'dbo.EspaciosDeportivosCompartidos', N'U')
)
BEGIN
    ALTER TABLE dbo.EspaciosDeportivosCompartidos  WITH CHECK
    ADD CONSTRAINT CK_EspaciosDeportivosCompartidos_TipoRelacion
    CHECK (TipoRelacion IN (N'DIRECTO', N'COMPUESTO_COMPONENTE'));
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_EspaciosDeportivosCompartidos_ParActivo'
      AND object_id = OBJECT_ID(N'dbo.EspaciosDeportivosCompartidos', N'U')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [UX_EspaciosDeportivosCompartidos_ParActivo]
    ON [dbo].[EspaciosDeportivosCompartidos] ([EspacioDeportivoId] ASC, [EspacioRelacionadoId] ASC)
    WHERE [Activo] = (1);
END;
