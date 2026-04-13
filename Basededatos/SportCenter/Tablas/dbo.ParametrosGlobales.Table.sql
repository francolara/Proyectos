USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[ParametrosGlobales]    Script Date: 10/04/2026 ******/
-- Firma: Codex - 10/04/2026 | Tabla de parametros globales para centralizar valores de validacion funcional con clave tecnica NombreParametro.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ParametrosGlobales](
    [ParametroId] [int] IDENTITY(1,1) NOT NULL,
    [NombreParametro] [nvarchar](100) NOT NULL,
    [Descripcion] [nvarchar](500) NOT NULL,
    [ValorParametro] [nvarchar](100) NOT NULL,
 CONSTRAINT [PK_ParametrosGlobales] PRIMARY KEY CLUSTERED
(
    [ParametroId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_ParametrosGlobales_Descripcion] UNIQUE NONCLUSTERED
(
    [Descripcion] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UQ_ParametrosGlobales_NombreParametro] UNIQUE NONCLUSTERED
(
    [NombreParametro] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
