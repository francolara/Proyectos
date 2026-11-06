-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Entrega incremental HYT etapa 3, para ejecución manual.
-- =============================================
-- Aplicar después de etapa 2. No ejecutado automáticamente.
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos internos automáticos y precios gestionados desde materiales.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Ubigeo de configuración HYT y demás fuentes de esta etapa.

-- Fuente: Tablas/UbigeoDepartamentos.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UbigeoDepartamentos',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDepartamentos
    (
        CodigoDepartamento CHAR(2) NOT NULL CONSTRAINT PK_UbigeoDepartamentos PRIMARY KEY,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT(1)
    );
END;
GO


-- Fuente: Tablas/UbigeoProvincias.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UbigeoProvincias',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoProvincias
    (
        CodigoProvincia CHAR(4) NOT NULL CONSTRAINT PK_UbigeoProvincias PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT(1),
        CONSTRAINT FK_UbigeoProvincias_CodigoDepartamento FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento)
    );
END;
GO


-- Fuente: Tablas/UbigeoDistritos.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UbigeoDistritos',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UbigeoDistritos
    (
        CodigoUbigeo CHAR(6) NOT NULL CONSTRAINT PK_UbigeoDistritos PRIMARY KEY,
        CodigoDepartamento CHAR(2) NOT NULL,
        CodigoProvincia CHAR(4) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT(1),
        CONSTRAINT FK_UbigeoDistritos_CodigoDepartamento FOREIGN KEY (CodigoDepartamento) REFERENCES dbo.UbigeoDepartamentos (CodigoDepartamento),
        CONSTRAINT FK_UbigeoDistritos_CodigoProvincia FOREIGN KEY (CodigoProvincia) REFERENCES dbo.UbigeoProvincias (CodigoProvincia)
    );
END;
GO


-- Fuente: Script/024_UbigeosCatalogo.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'01') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'01',N'AMAZONAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'02') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'02',N'ANCASH',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'03') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'03',N'APURIMAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'04') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'04',N'AREQUIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'05') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'05',N'AYACUCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'06') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'06',N'CAJAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'07') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'07',N'CALLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'08') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'08',N'CUSCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'09') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'09',N'HUANCAVELICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'10') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'10',N'HUANUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'11') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'11',N'ICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'12') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'12',N'JUNIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'13') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'13',N'LA LIBERTAD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'14') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'14',N'LAMBAYEQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'15') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'15',N'LIMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'16') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'16',N'LORETO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'17') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'17',N'MADRE DE DIOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'18') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'18',N'MOQUEGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'19') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'19',N'PASCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'20') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'20',N'PIURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'21') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'21',N'PUNO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'22') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'22',N'SAN MARTIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'23') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'23',N'TACNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'24') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'24',N'TUMBES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDepartamentos AS destino WHERE destino.CodigoDepartamento=N'25') INSERT INTO dbo.UbigeoDepartamentos (CodigoDepartamento,Nombre,Activo) VALUES (N'25',N'UCAYALI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0101') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0101',N'01',N'CHACHAPOYAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0102') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0102',N'01',N'BAGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0103') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0103',N'01',N'BONGARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0104') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0104',N'01',N'CONDORCANQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0105') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0105',N'01',N'LUYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0106') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0106',N'01',N'RODRIGUEZ DE MENDOZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0107') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0107',N'01',N'UTCUBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0201') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0201',N'02',N'HUARAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0202') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0202',N'02',N'AIJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0203') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0203',N'02',N'ANTONIO RAYMONDI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0204') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0204',N'02',N'ASUNCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0205') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0205',N'02',N'BOLOGNESI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0206') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0206',N'02',N'CARHUAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0207') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0207',N'02',N'CARLOS FERMIN FITZCARRALD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0208') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0208',N'02',N'CASMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0209') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0209',N'02',N'CORONGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0210') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0210',N'02',N'HUARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0211') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0211',N'02',N'HUARMEY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0212') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0212',N'02',N'HUAYLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0213') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0213',N'02',N'MARISCAL LUZURIAGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0214') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0214',N'02',N'OCROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0215') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0215',N'02',N'PALLASCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0216') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0216',N'02',N'POMABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0217') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0217',N'02',N'RECUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0218') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0218',N'02',N'SANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0219') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0219',N'02',N'SIHUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0220') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0220',N'02',N'YUNGAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0301') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0301',N'03',N'ABANCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0302') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0302',N'03',N'ANDAHUAYLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0303') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0303',N'03',N'ANTABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0304') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0304',N'03',N'AYMARAES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0305') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0305',N'03',N'COTABAMBAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0306') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0306',N'03',N'CHINCHEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0307') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0307',N'03',N'GRAU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0401') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0401',N'04',N'AREQUIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0402') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0402',N'04',N'CAMANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0403') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0403',N'04',N'CARAVELI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0404') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0404',N'04',N'CASTILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0405') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0405',N'04',N'CAYLLOMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0406') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0406',N'04',N'CONDESUYOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0407') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0407',N'04',N'ISLAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0408') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0408',N'04',N'LA UNION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0501') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0501',N'05',N'HUAMANGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0502') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0502',N'05',N'CANGALLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0503') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0503',N'05',N'HUANCA SANCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0504') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0504',N'05',N'HUANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0505') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0505',N'05',N'LA MAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0506') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0506',N'05',N'LUCANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0507') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0507',N'05',N'PARINACOCHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0508') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0508',N'05',N'PAUCAR DEL SARA SARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0509') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0509',N'05',N'SUCRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0510') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0510',N'05',N'VICTOR FAJARDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0511') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0511',N'05',N'VILCAS HUAMAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0601') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0601',N'06',N'CAJAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0602') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0602',N'06',N'CAJABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0603') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0603',N'06',N'CELENDIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0604') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0604',N'06',N'CHOTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0605') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0605',N'06',N'CONTUMAZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0606') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0606',N'06',N'CUTERVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0607') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0607',N'06',N'HUALGAYOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0608') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0608',N'06',N'JAEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0609') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0609',N'06',N'SAN IGNACIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0610') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0610',N'06',N'SAN MARCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0611') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0611',N'06',N'SAN MIGUEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0612') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0612',N'06',N'SAN PABLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0613') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0613',N'06',N'SANTA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0701') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0701',N'07',N'CALLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0801') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0801',N'08',N'CUSCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0802') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0802',N'08',N'ACOMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0803') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0803',N'08',N'ANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0804') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0804',N'08',N'CALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0805') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0805',N'08',N'CANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0806') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0806',N'08',N'CANCHIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0807') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0807',N'08',N'CHUMBIVILCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0808') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0808',N'08',N'ESPINAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0809') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0809',N'08',N'LA CONVENCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0810') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0810',N'08',N'PARURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0811') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0811',N'08',N'PAUCARTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0812') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0812',N'08',N'QUISPICANCHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0813') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0813',N'08',N'URUBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0901') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0901',N'09',N'HUANCAVELICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0902') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0902',N'09',N'ACOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0903') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0903',N'09',N'ANGARAES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0904') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0904',N'09',N'CASTROVIRREYNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0905') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0905',N'09',N'CHURCAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0906') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0906',N'09',N'HUAYTARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'0907') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'0907',N'09',N'TAYACAJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1001') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1001',N'10',N'HUANUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1002') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1002',N'10',N'AMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1003') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1003',N'10',N'DOS DE MAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1004') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1004',N'10',N'HUACAYBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1005') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1005',N'10',N'HUAMALIES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1006') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1006',N'10',N'LEONCIO PRADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1007') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1007',N'10',N'MARAÑON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1008') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1008',N'10',N'PACHITEA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1009') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1009',N'10',N'PUERTO INCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1010') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1010',N'10',N'LAURICOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1011') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1011',N'10',N'YAROWILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1101') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1101',N'11',N'ICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1102') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1102',N'11',N'CHINCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1103') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1103',N'11',N'NASCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1104') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1104',N'11',N'PALPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1105') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1105',N'11',N'PISCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1201') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1201',N'12',N'HUANCAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1202') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1202',N'12',N'CONCEPCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1203') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1203',N'12',N'CHANCHAMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1204') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1204',N'12',N'JAUJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1205') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1205',N'12',N'JUNIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1206') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1206',N'12',N'SATIPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1207') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1207',N'12',N'TARMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1208') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1208',N'12',N'YAULI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1209') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1209',N'12',N'CHUPACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1301') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1301',N'13',N'TRUJILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1302') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1302',N'13',N'ASCOPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1303') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1303',N'13',N'BOLIVAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1304') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1304',N'13',N'CHEPEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1305') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1305',N'13',N'JULCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1306') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1306',N'13',N'OTUZCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1307') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1307',N'13',N'PACASMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1308') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1308',N'13',N'PATAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1309') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1309',N'13',N'SANCHEZ CARRION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1310') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1310',N'13',N'SANTIAGO DE CHUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1311') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1311',N'13',N'GRAN CHIMU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1312') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1312',N'13',N'VIRU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1401') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1401',N'14',N'CHICLAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1402') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1402',N'14',N'FERREÑAFE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1403') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1403',N'14',N'LAMBAYEQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1501') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1501',N'15',N'LIMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1502') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1502',N'15',N'BARRANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1503') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1503',N'15',N'CAJATAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1504') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1504',N'15',N'CANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1505') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1505',N'15',N'CAÑETE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1506') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1506',N'15',N'HUARAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1507') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1507',N'15',N'HUAROCHIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1508') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1508',N'15',N'HUAURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1509') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1509',N'15',N'OYON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1510') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1510',N'15',N'YAUYOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1601') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1601',N'16',N'MAYNAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1602') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1602',N'16',N'ALTO AMAZONAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1603') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1603',N'16',N'LORETO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1604') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1604',N'16',N'MARISCAL RAMON CASTILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1605') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1605',N'16',N'REQUENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1606') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1606',N'16',N'UCAYALI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1607') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1607',N'16',N'DATEM DEL MARAÑON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1608') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1608',N'16',N'PUTUMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1701') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1701',N'17',N'TAMBOPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1702') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1702',N'17',N'MANU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1703') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1703',N'17',N'TAHUAMANU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1801') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1801',N'18',N'MARISCAL NIETO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1802') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1802',N'18',N'GENERAL SANCHEZ CERRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1803') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1803',N'18',N'ILO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1901') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1901',N'19',N'PASCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1902') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1902',N'19',N'DANIEL ALCIDES CARRION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'1903') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'1903',N'19',N'OXAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2001') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2001',N'20',N'PIURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2002') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2002',N'20',N'AYABACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2003') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2003',N'20',N'HUANCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2004') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2004',N'20',N'MORROPON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2005') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2005',N'20',N'PAITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2006') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2006',N'20',N'SULLANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2007') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2007',N'20',N'TALARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2008') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2008',N'20',N'SECHURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2101') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2101',N'21',N'PUNO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2102') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2102',N'21',N'AZANGARO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2103') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2103',N'21',N'CARABAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2104') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2104',N'21',N'CHUCUITO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2105') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2105',N'21',N'EL COLLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2106') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2106',N'21',N'HUANCANE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2107') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2107',N'21',N'LAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2108') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2108',N'21',N'MELGAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2109') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2109',N'21',N'MOHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2110') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2110',N'21',N'SAN ANTONIO DE PUTINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2111') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2111',N'21',N'SAN ROMAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2112') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2112',N'21',N'SANDIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2113') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2113',N'21',N'YUNGUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2201') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2201',N'22',N'MOYOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2202') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2202',N'22',N'BELLAVISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2203') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2203',N'22',N'EL DORADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2204') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2204',N'22',N'HUALLAGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2205') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2205',N'22',N'LAMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2206') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2206',N'22',N'MARISCAL CACERES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2207') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2207',N'22',N'PICOTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2208') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2208',N'22',N'RIOJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2209') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2209',N'22',N'SAN MARTIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2210') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2210',N'22',N'TOCACHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2301') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2301',N'23',N'TACNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2302') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2302',N'23',N'CANDARAVE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2303') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2303',N'23',N'JORGE BASADRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2304') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2304',N'23',N'TARATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2401') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2401',N'24',N'TUMBES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2402') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2402',N'24',N'CONTRALMIRANTE VILLAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2403') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2403',N'24',N'ZARUMILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2501') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2501',N'25',N'CORONEL PORTILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2502') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2502',N'25',N'ATALAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2503') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2503',N'25',N'PADRE ABAD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoProvincias AS destino WHERE destino.CodigoProvincia=N'2504') INSERT INTO dbo.UbigeoProvincias (CodigoProvincia,CodigoDepartamento,Nombre,Activo) VALUES (N'2504',N'25',N'PURUS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010101',N'01',N'0101',N'CHACHAPOYAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010102',N'01',N'0101',N'ASUNCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010103',N'01',N'0101',N'BALSAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010104',N'01',N'0101',N'CHETO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010105',N'01',N'0101',N'CHILIQUIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010106',N'01',N'0101',N'CHUQUIBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010107',N'01',N'0101',N'GRANADA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010108',N'01',N'0101',N'HUANCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010109',N'01',N'0101',N'LA JALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010110',N'01',N'0101',N'LEIMEBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010111',N'01',N'0101',N'LEVANTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010112',N'01',N'0101',N'MAGDALENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010113',N'01',N'0101',N'MARISCAL CASTILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010114',N'01',N'0101',N'MOLINOPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010115',N'01',N'0101',N'MONTEVIDEO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010116',N'01',N'0101',N'OLLEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010117') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010117',N'01',N'0101',N'QUINJALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010118') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010118',N'01',N'0101',N'SAN FRANCISCO DE DAGUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010119') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010119',N'01',N'0101',N'SAN ISIDRO DE MAINO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010120') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010120',N'01',N'0101',N'SOLOCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010121') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010121',N'01',N'0101',N'SONCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010201',N'01',N'0102',N'BAGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010202',N'01',N'0102',N'ARAMANGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010203',N'01',N'0102',N'COPALLIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010204',N'01',N'0102',N'EL PARCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010205',N'01',N'0102',N'IMAZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010206',N'01',N'0102',N'LA PECA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010301',N'01',N'0103',N'JUMBILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010302',N'01',N'0103',N'CHISQUILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010303',N'01',N'0103',N'CHURUJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010304',N'01',N'0103',N'COROSHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010305',N'01',N'0103',N'CUISPES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010306',N'01',N'0103',N'FLORIDA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010307',N'01',N'0103',N'JAZAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010308',N'01',N'0103',N'RECTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010309',N'01',N'0103',N'SAN CARLOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010310',N'01',N'0103',N'SHIPASBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010311',N'01',N'0103',N'VALERA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010312') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010312',N'01',N'0103',N'YAMBRASBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010401',N'01',N'0104',N'NIEVA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010402',N'01',N'0104',N'EL CENEPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010403',N'01',N'0104',N'RIO SANTIAGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010501',N'01',N'0105',N'LAMUD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010502',N'01',N'0105',N'CAMPORREDONDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010503',N'01',N'0105',N'COCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010504',N'01',N'0105',N'COLCAMAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010505',N'01',N'0105',N'CONILA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010506',N'01',N'0105',N'INGUILPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010507',N'01',N'0105',N'LONGUITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010508',N'01',N'0105',N'LONYA CHICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010509',N'01',N'0105',N'LUYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010510',N'01',N'0105',N'LUYA VIEJO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010511',N'01',N'0105',N'MARIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010512') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010512',N'01',N'0105',N'OCALLI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010513') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010513',N'01',N'0105',N'OCUMAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010514') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010514',N'01',N'0105',N'PISUQUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010515') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010515',N'01',N'0105',N'PROVIDENCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010516') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010516',N'01',N'0105',N'SAN CRISTOBAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010517') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010517',N'01',N'0105',N'SAN FRANCISCO DEL YESO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010518') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010518',N'01',N'0105',N'SAN JERONIMO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010519') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010519',N'01',N'0105',N'SAN JUAN DE LOPECANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010520') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010520',N'01',N'0105',N'SANTA CATALINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010521') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010521',N'01',N'0105',N'SANTO TOMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010522') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010522',N'01',N'0105',N'TINGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010523') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010523',N'01',N'0105',N'TRITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010601',N'01',N'0106',N'SAN NICOLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010602',N'01',N'0106',N'CHIRIMOTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010603',N'01',N'0106',N'COCHAMAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010604',N'01',N'0106',N'HUAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010605',N'01',N'0106',N'LIMABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010606',N'01',N'0106',N'LONGAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010607',N'01',N'0106',N'MARISCAL BENAVIDES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010608',N'01',N'0106',N'MILPUC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010609',N'01',N'0106',N'OMIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010610',N'01',N'0106',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010611',N'01',N'0106',N'TOTORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010612') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010612',N'01',N'0106',N'VISTA ALEGRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010701',N'01',N'0107',N'BAGUA GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010702',N'01',N'0107',N'CAJARURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010703',N'01',N'0107',N'CUMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010704',N'01',N'0107',N'EL MILAGRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010705',N'01',N'0107',N'JAMALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010706',N'01',N'0107',N'LONYA GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'010707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'010707',N'01',N'0107',N'YAMON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020101',N'02',N'0201',N'HUARAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020102',N'02',N'0201',N'COCHABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020103',N'02',N'0201',N'COLCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020104',N'02',N'0201',N'HUANCHAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020105',N'02',N'0201',N'INDEPENDENCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020106',N'02',N'0201',N'JANGAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020107',N'02',N'0201',N'LA LIBERTAD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020108',N'02',N'0201',N'OLLEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020109',N'02',N'0201',N'PAMPAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020110',N'02',N'0201',N'PARIACOTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020111',N'02',N'0201',N'PIRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020112',N'02',N'0201',N'TARICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020201',N'02',N'0202',N'AIJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020202',N'02',N'0202',N'CORIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020203',N'02',N'0202',N'HUACLLAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020204',N'02',N'0202',N'LA MERCED',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020205',N'02',N'0202',N'SUCCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020301',N'02',N'0203',N'LLAMELLIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020302',N'02',N'0203',N'ACZO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020303',N'02',N'0203',N'CHACCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020304',N'02',N'0203',N'CHINGAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020305',N'02',N'0203',N'MIRGAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020306',N'02',N'0203',N'SAN JUAN DE RONTOY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020401',N'02',N'0204',N'CHACAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020402',N'02',N'0204',N'ACOCHACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020501',N'02',N'0205',N'CHIQUIAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020502',N'02',N'0205',N'ABELARDO PARDO LEZAMETA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020503',N'02',N'0205',N'ANTONIO RAYMONDI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020504',N'02',N'0205',N'AQUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020505',N'02',N'0205',N'CAJACAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020506',N'02',N'0205',N'CANIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020507',N'02',N'0205',N'COLQUIOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020508',N'02',N'0205',N'HUALLANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020509',N'02',N'0205',N'HUASTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020510',N'02',N'0205',N'HUAYLLACAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020511',N'02',N'0205',N'LA PRIMAVERA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020512') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020512',N'02',N'0205',N'MANGAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020513') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020513',N'02',N'0205',N'PACLLON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020514') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020514',N'02',N'0205',N'SAN MIGUEL DE CORPANQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020515') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020515',N'02',N'0205',N'TICLLOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020601',N'02',N'0206',N'CARHUAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020602',N'02',N'0206',N'ACOPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020603',N'02',N'0206',N'AMASHCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020604',N'02',N'0206',N'ANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020605',N'02',N'0206',N'ATAQUERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020606',N'02',N'0206',N'MARCARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020607',N'02',N'0206',N'PARIAHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020608',N'02',N'0206',N'SAN MIGUEL DE ACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020609',N'02',N'0206',N'SHILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020610',N'02',N'0206',N'TINCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020611',N'02',N'0206',N'YUNGAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020701',N'02',N'0207',N'SAN LUIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020702',N'02',N'0207',N'SAN NICOLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020703',N'02',N'0207',N'YAUYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020801',N'02',N'0208',N'CASMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020802',N'02',N'0208',N'BUENA VISTA ALTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020803',N'02',N'0208',N'COMANDANTE NOEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020804',N'02',N'0208',N'YAUTAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020901',N'02',N'0209',N'CORONGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020902',N'02',N'0209',N'ACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020903',N'02',N'0209',N'BAMBAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020904',N'02',N'0209',N'CUSCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020905',N'02',N'0209',N'LA PAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020906',N'02',N'0209',N'YANAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'020907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'020907',N'02',N'0209',N'YUPAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021001',N'02',N'0210',N'HUARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021002',N'02',N'0210',N'ANRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021003',N'02',N'0210',N'CAJAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021004',N'02',N'0210',N'CHAVIN DE HUANTAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021005',N'02',N'0210',N'HUACACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021006',N'02',N'0210',N'HUACCHIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021007',N'02',N'0210',N'HUACHIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021008') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021008',N'02',N'0210',N'HUANTAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021009') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021009',N'02',N'0210',N'MASIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021010') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021010',N'02',N'0210',N'PAUCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021011') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021011',N'02',N'0210',N'PONTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021012') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021012',N'02',N'0210',N'RAHUAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021013') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021013',N'02',N'0210',N'RAPAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021014') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021014',N'02',N'0210',N'SAN MARCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021015') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021015',N'02',N'0210',N'SAN PEDRO DE CHANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021016') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021016',N'02',N'0210',N'UCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021101',N'02',N'0211',N'HUARMEY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021102',N'02',N'0211',N'COCHAPETI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021103',N'02',N'0211',N'CULEBRAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021104',N'02',N'0211',N'HUAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021105',N'02',N'0211',N'MALVAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021201',N'02',N'0212',N'CARAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021202',N'02',N'0212',N'HUALLANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021203',N'02',N'0212',N'HUATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021204',N'02',N'0212',N'HUAYLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021205',N'02',N'0212',N'MATO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021206',N'02',N'0212',N'PAMPAROMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021207',N'02',N'0212',N'PUEBLO LIBRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021208',N'02',N'0212',N'SANTA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021209',N'02',N'0212',N'SANTO TORIBIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021210',N'02',N'0212',N'YURACMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021301',N'02',N'0213',N'PISCOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021302',N'02',N'0213',N'CASCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021303',N'02',N'0213',N'ELEAZAR GUZMAN BARRON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021304',N'02',N'0213',N'FIDEL OLIVAS ESCUDERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021305',N'02',N'0213',N'LLAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021306',N'02',N'0213',N'LLUMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021307',N'02',N'0213',N'LUCMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021308',N'02',N'0213',N'MUSGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021401',N'02',N'0214',N'OCROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021402',N'02',N'0214',N'ACAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021403',N'02',N'0214',N'CAJAMARQUILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021404',N'02',N'0214',N'CARHUAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021405',N'02',N'0214',N'COCHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021406',N'02',N'0214',N'CONGAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021407',N'02',N'0214',N'LLIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021408',N'02',N'0214',N'SAN CRISTOBAL DE RAJAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021409',N'02',N'0214',N'SAN PEDRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021410',N'02',N'0214',N'SANTIAGO DE CHILCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021501',N'02',N'0215',N'CABANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021502',N'02',N'0215',N'BOLOGNESI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021503',N'02',N'0215',N'CONCHUCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021504',N'02',N'0215',N'HUACASCHUQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021505',N'02',N'0215',N'HUANDOVAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021506',N'02',N'0215',N'LACABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021507',N'02',N'0215',N'LLAPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021508',N'02',N'0215',N'PALLASCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021509',N'02',N'0215',N'PAMPAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021510',N'02',N'0215',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021511',N'02',N'0215',N'TAUCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021601',N'02',N'0216',N'POMABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021602',N'02',N'0216',N'HUAYLLAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021603',N'02',N'0216',N'PAROBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021604',N'02',N'0216',N'QUINUABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021701',N'02',N'0217',N'RECUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021702',N'02',N'0217',N'CATAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021703',N'02',N'0217',N'COTAPARACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021704',N'02',N'0217',N'HUAYLLAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021705',N'02',N'0217',N'LLACLLIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021706',N'02',N'0217',N'MARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021707',N'02',N'0217',N'PAMPAS CHICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021708',N'02',N'0217',N'PARARIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021709',N'02',N'0217',N'TAPACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021710') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021710',N'02',N'0217',N'TICAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021801',N'02',N'0218',N'CHIMBOTE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021802',N'02',N'0218',N'CACERES DEL PERU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021803',N'02',N'0218',N'COISHCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021804',N'02',N'0218',N'MACATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021805',N'02',N'0218',N'MORO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021806',N'02',N'0218',N'NEPEÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021807',N'02',N'0218',N'SAMANCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021808',N'02',N'0218',N'SANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021809',N'02',N'0218',N'NUEVO CHIMBOTE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021901',N'02',N'0219',N'SIHUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021902',N'02',N'0219',N'ACOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021903',N'02',N'0219',N'ALFONSO UGARTE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021904',N'02',N'0219',N'CASHAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021905',N'02',N'0219',N'CHINGALPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021906',N'02',N'0219',N'HUAYLLABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021907',N'02',N'0219',N'QUICHES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021908') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021908',N'02',N'0219',N'RAGASH',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021909') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021909',N'02',N'0219',N'SAN JUAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'021910') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'021910',N'02',N'0219',N'SICSIBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022001',N'02',N'0220',N'YUNGAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022002',N'02',N'0220',N'CASCAPARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022003',N'02',N'0220',N'MANCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022004',N'02',N'0220',N'MATACOTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022005',N'02',N'0220',N'QUILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022006',N'02',N'0220',N'RANRAHIRCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022007',N'02',N'0220',N'SHUPLUY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'022008') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'022008',N'02',N'0220',N'YANAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030101',N'03',N'0301',N'ABANCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030102',N'03',N'0301',N'CHACOCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030103',N'03',N'0301',N'CIRCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030104',N'03',N'0301',N'CURAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030105',N'03',N'0301',N'HUANIPACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030106',N'03',N'0301',N'LAMBRAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030107',N'03',N'0301',N'PICHIRHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030108',N'03',N'0301',N'SAN PEDRO DE CACHORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030109',N'03',N'0301',N'TAMBURCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030201',N'03',N'0302',N'ANDAHUAYLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030202',N'03',N'0302',N'ANDARAPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030204',N'03',N'0302',N'HUANCARAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030205',N'03',N'0302',N'HUANCARAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030206',N'03',N'0302',N'HUAYANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030207',N'03',N'0302',N'KISHUARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030208',N'03',N'0302',N'PACOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030209',N'03',N'0302',N'PACUCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030210',N'03',N'0302',N'PAMPACHIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030211',N'03',N'0302',N'POMACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030212') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030212',N'03',N'0302',N'SAN ANTONIO DE CACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030213') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030213',N'03',N'0302',N'SAN JERONIMO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030214') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030214',N'03',N'0302',N'SAN MIGUEL DE CHACCRAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030215') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030215',N'03',N'0302',N'SANTA MARIA DE CHICMO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030216') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030216',N'03',N'0302',N'TALAVERA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030217') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030217',N'03',N'0302',N'TUMAY HUARACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030218') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030218',N'03',N'0302',N'TURPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030219') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030219',N'03',N'0302',N'KAQUIABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030220') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030220',N'03',N'0302',N'JOSE MARIA ARGUEDAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030301',N'03',N'0303',N'ANTABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030302',N'03',N'0303',N'EL ORO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030303',N'03',N'0303',N'HUAQUIRCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030304',N'03',N'0303',N'JUAN ESPINOZA MEDRANO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030305',N'03',N'0303',N'OROPESA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030306',N'03',N'0303',N'PACHACONAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030307',N'03',N'0303',N'SABAINO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030401',N'03',N'0304',N'CHALHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030402',N'03',N'0304',N'CAPAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030403',N'03',N'0304',N'CARAYBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030404',N'03',N'0304',N'CHAPIMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030405',N'03',N'0304',N'COLCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030406',N'03',N'0304',N'COTARUSE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030407',N'03',N'0304',N'HUAYLLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030408',N'03',N'0304',N'JUSTO APU SAHUARAURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030409',N'03',N'0304',N'LUCRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030410',N'03',N'0304',N'POCOHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030411') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030411',N'03',N'0304',N'SAN JUAN DE CHACÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030412') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030412',N'03',N'0304',N'SAÑAYCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030413') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030413',N'03',N'0304',N'SORAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030414') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030414',N'03',N'0304',N'TAPAIRIHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030415') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030415',N'03',N'0304',N'TINTAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030416') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030416',N'03',N'0304',N'TORAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030417') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030417',N'03',N'0304',N'YANACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030501',N'03',N'0305',N'TAMBOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030502',N'03',N'0305',N'COTABAMBAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030503',N'03',N'0305',N'COYLLURQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030504',N'03',N'0305',N'HAQUIRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030505',N'03',N'0305',N'MARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030506',N'03',N'0305',N'CHALLHUAHUACHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030601',N'03',N'0306',N'CHINCHEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030602',N'03',N'0306',N'ANCO_HUALLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030603',N'03',N'0306',N'COCHARCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030604',N'03',N'0306',N'HUACCANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030605',N'03',N'0306',N'OCOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030606',N'03',N'0306',N'ONGOY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030607',N'03',N'0306',N'URANMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030608',N'03',N'0306',N'RANRACANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030609',N'03',N'0306',N'ROCCHACC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030610',N'03',N'0306',N'EL PORVENIR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030611',N'03',N'0306',N'LOS CHANKAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030612') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030612',N'03',N'0306',N'AHUAYRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030701',N'03',N'0307',N'CHUQUIBAMBILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030702',N'03',N'0307',N'CURPAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030703',N'03',N'0307',N'GAMARRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030704',N'03',N'0307',N'HUAYLLATI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030705',N'03',N'0307',N'MAMARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030706',N'03',N'0307',N'MICAELA BASTIDAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030707',N'03',N'0307',N'PATAYPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030708',N'03',N'0307',N'PROGRESO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030709',N'03',N'0307',N'SAN ANTONIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030710') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030710',N'03',N'0307',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030711') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030711',N'03',N'0307',N'TURPAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030712') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030712',N'03',N'0307',N'VILCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030713') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030713',N'03',N'0307',N'VIRUNDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'030714') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'030714',N'03',N'0307',N'CURASCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040101',N'04',N'0401',N'AREQUIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040102',N'04',N'0401',N'ALTO SELVA ALEGRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040103',N'04',N'0401',N'CAYMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040104',N'04',N'0401',N'CERRO COLORADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040105',N'04',N'0401',N'CHARACATO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040106',N'04',N'0401',N'CHIGUATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040107',N'04',N'0401',N'JACOBO HUNTER',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040108',N'04',N'0401',N'LA JOYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040109',N'04',N'0401',N'MARIANO MELGAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040110',N'04',N'0401',N'MIRAFLORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040111',N'04',N'0401',N'MOLLEBAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040112',N'04',N'0401',N'PAUCARPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040113',N'04',N'0401',N'POCSI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040114',N'04',N'0401',N'POLOBAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040115',N'04',N'0401',N'QUEQUEÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040116',N'04',N'0401',N'SABANDIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040117') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040117',N'04',N'0401',N'SACHACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040118') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040118',N'04',N'0401',N'SAN JUAN DE SIGUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040119') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040119',N'04',N'0401',N'SAN JUAN DE TARUCANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040120') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040120',N'04',N'0401',N'SANTA ISABEL DE SIGUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040121') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040121',N'04',N'0401',N'SANTA RITA DE SIGUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040122') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040122',N'04',N'0401',N'SOCABAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040123') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040123',N'04',N'0401',N'TIABAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040124') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040124',N'04',N'0401',N'UCHUMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040125') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040125',N'04',N'0401',N'VITOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040126') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040126',N'04',N'0401',N'YANAHUARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040127') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040127',N'04',N'0401',N'YARABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040128') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040128',N'04',N'0401',N'YURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040129') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040129',N'04',N'0401',N'JOSE LUIS BUSTAMANTE Y RIVERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040201',N'04',N'0402',N'CAMANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040202',N'04',N'0402',N'JOSE MARIA QUIMPER',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040203',N'04',N'0402',N'MARIANO NICOLAS VALCARCEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040204',N'04',N'0402',N'MARISCAL CACERES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040205',N'04',N'0402',N'NICOLAS DE PIEROLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040206',N'04',N'0402',N'OCOÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040207',N'04',N'0402',N'QUILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040208',N'04',N'0402',N'SAMUEL PASTOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040301',N'04',N'0403',N'CARAVELI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040302',N'04',N'0403',N'ACARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040303',N'04',N'0403',N'ATICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040304',N'04',N'0403',N'ATIQUIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040305',N'04',N'0403',N'BELLA UNION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040306',N'04',N'0403',N'CAHUACHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040307',N'04',N'0403',N'CHALA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040308',N'04',N'0403',N'CHAPARRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040309',N'04',N'0403',N'HUANUHUANU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040310',N'04',N'0403',N'JAQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040311',N'04',N'0403',N'LOMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040312') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040312',N'04',N'0403',N'QUICACHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040313') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040313',N'04',N'0403',N'YAUCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040401',N'04',N'0404',N'APLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040402',N'04',N'0404',N'ANDAGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040403',N'04',N'0404',N'AYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040404',N'04',N'0404',N'CHACHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040405',N'04',N'0404',N'CHILCAYMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040406',N'04',N'0404',N'CHOCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040407',N'04',N'0404',N'HUANCARQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040408',N'04',N'0404',N'MACHAGUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040409',N'04',N'0404',N'ORCOPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040410',N'04',N'0404',N'PAMPACOLCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040411') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040411',N'04',N'0404',N'TIPAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040412') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040412',N'04',N'0404',N'UÑON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040413') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040413',N'04',N'0404',N'URACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040414') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040414',N'04',N'0404',N'VIRACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040501',N'04',N'0405',N'CHIVAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040502',N'04',N'0405',N'ACHOMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040503',N'04',N'0405',N'CABANACONDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040504',N'04',N'0405',N'CALLALLI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040505',N'04',N'0405',N'CAYLLOMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040506',N'04',N'0405',N'COPORAQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040507',N'04',N'0405',N'HUAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040508',N'04',N'0405',N'HUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040509',N'04',N'0405',N'ICHUPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040510',N'04',N'0405',N'LARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040511',N'04',N'0405',N'LLUTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040512') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040512',N'04',N'0405',N'MACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040513') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040513',N'04',N'0405',N'MADRIGAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040514') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040514',N'04',N'0405',N'SAN ANTONIO DE CHUCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040515') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040515',N'04',N'0405',N'SIBAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040516') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040516',N'04',N'0405',N'TAPAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040517') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040517',N'04',N'0405',N'TISCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040518') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040518',N'04',N'0405',N'TUTI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040519') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040519',N'04',N'0405',N'YANQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040520') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040520',N'04',N'0405',N'MAJES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040601',N'04',N'0406',N'CHUQUIBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040602',N'04',N'0406',N'ANDARAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040603',N'04',N'0406',N'CAYARANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040604',N'04',N'0406',N'CHICHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040605',N'04',N'0406',N'IRAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040606',N'04',N'0406',N'RIO GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040607',N'04',N'0406',N'SALAMANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040608',N'04',N'0406',N'YANAQUIHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040701',N'04',N'0407',N'MOLLENDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040702',N'04',N'0407',N'COCACHACRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040703',N'04',N'0407',N'DEAN VALDIVIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040704',N'04',N'0407',N'ISLAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040705',N'04',N'0407',N'MEJIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040706',N'04',N'0407',N'PUNTA DE BOMBON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040801',N'04',N'0408',N'COTAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040802',N'04',N'0408',N'ALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040803',N'04',N'0408',N'CHARCANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040804',N'04',N'0408',N'HUAYNACOTAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040805',N'04',N'0408',N'PAMPAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040806',N'04',N'0408',N'PUYCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040807',N'04',N'0408',N'QUECHUALLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040808',N'04',N'0408',N'SAYLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040809',N'04',N'0408',N'TAURIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040810') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040810',N'04',N'0408',N'TOMEPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'040811') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'040811',N'04',N'0408',N'TORO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050101',N'05',N'0501',N'AYACUCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050102',N'05',N'0501',N'ACOCRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050103',N'05',N'0501',N'ACOS VINCHOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050104',N'05',N'0501',N'CARMEN ALTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050105',N'05',N'0501',N'CHIARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050106',N'05',N'0501',N'OCROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050107',N'05',N'0501',N'PACAYCASA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050108',N'05',N'0501',N'QUINUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050109',N'05',N'0501',N'SAN JOSE DE TICLLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050110',N'05',N'0501',N'SAN JUAN BAUTISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050111',N'05',N'0501',N'SANTIAGO DE PISCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050112',N'05',N'0501',N'SOCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050113',N'05',N'0501',N'TAMBILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050114',N'05',N'0501',N'VINCHOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050115',N'05',N'0501',N'JESUS NAZARENO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050116',N'05',N'0501',N'ANDRES AVELINO CACERES DORREGARAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050201',N'05',N'0502',N'CANGALLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050202',N'05',N'0502',N'CHUSCHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050203',N'05',N'0502',N'LOS MOROCHUCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050204',N'05',N'0502',N'MARIA PARADO DE BELLIDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050205',N'05',N'0502',N'PARAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050206',N'05',N'0502',N'TOTOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050301',N'05',N'0503',N'SANCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050302',N'05',N'0503',N'CARAPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050303',N'05',N'0503',N'SACSAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050304',N'05',N'0503',N'SANTIAGO DE LUCANAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050401',N'05',N'0504',N'HUANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050402',N'05',N'0504',N'AYAHUANCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050403',N'05',N'0504',N'HUAMANGUILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050404',N'05',N'0504',N'IGUAIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050405',N'05',N'0504',N'LURICOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050406',N'05',N'0504',N'SANTILLANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050407',N'05',N'0504',N'SIVIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050408',N'05',N'0504',N'LLOCHEGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050409',N'05',N'0504',N'CANAYRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050410',N'05',N'0504',N'UCHURACCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050411') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050411',N'05',N'0504',N'PUCACOLPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050412') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050412',N'05',N'0504',N'CHACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050413') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050413',N'05',N'0504',N'PUTIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050501',N'05',N'0505',N'SAN MIGUEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050502',N'05',N'0505',N'ANCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050503',N'05',N'0505',N'AYNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050504',N'05',N'0505',N'CHILCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050505',N'05',N'0505',N'CHUNGUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050506',N'05',N'0505',N'LUIS CARRANZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050507',N'05',N'0505',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050508',N'05',N'0505',N'TAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050509',N'05',N'0505',N'SAMUGARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050510',N'05',N'0505',N'ANCHIHUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050511',N'05',N'0505',N'ORONCCOY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050512') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050512',N'05',N'0505',N'UNION PROGRESO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050513') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050513',N'05',N'0505',N'RIO MAGDALENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050514') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050514',N'05',N'0505',N'NINABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050515') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050515',N'05',N'0505',N'PATIBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050601',N'05',N'0506',N'PUQUIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050602',N'05',N'0506',N'AUCARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050603',N'05',N'0506',N'CABANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050604',N'05',N'0506',N'CARMEN SALCEDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050605',N'05',N'0506',N'CHAVIÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050606',N'05',N'0506',N'CHIPAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050607',N'05',N'0506',N'HUAC-HUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050608',N'05',N'0506',N'LARAMATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050609',N'05',N'0506',N'LEONCIO PRADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050610',N'05',N'0506',N'LLAUTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050611',N'05',N'0506',N'LUCANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050612') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050612',N'05',N'0506',N'OCAÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050613') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050613',N'05',N'0506',N'OTOCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050614') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050614',N'05',N'0506',N'SAISA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050615') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050615',N'05',N'0506',N'SAN CRISTOBAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050616') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050616',N'05',N'0506',N'SAN JUAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050617') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050617',N'05',N'0506',N'SAN PEDRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050618') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050618',N'05',N'0506',N'SAN PEDRO DE PALCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050619') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050619',N'05',N'0506',N'SANCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050620') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050620',N'05',N'0506',N'SANTA ANA DE HUAYCAHUACHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050621') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050621',N'05',N'0506',N'SANTA LUCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050701',N'05',N'0507',N'CORACORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050702',N'05',N'0507',N'CHUMPI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050703',N'05',N'0507',N'CORONEL CASTAÑEDA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050704',N'05',N'0507',N'PACAPAUSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050705',N'05',N'0507',N'PULLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050706',N'05',N'0507',N'PUYUSCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050707',N'05',N'0507',N'SAN FRANCISCO DE RAVACAYCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050708',N'05',N'0507',N'UPAHUACHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050801',N'05',N'0508',N'PAUSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050802',N'05',N'0508',N'COLTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050803',N'05',N'0508',N'CORCULLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050804',N'05',N'0508',N'LAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050805',N'05',N'0508',N'MARCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050806',N'05',N'0508',N'OYOLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050807',N'05',N'0508',N'PARARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050808',N'05',N'0508',N'SAN JAVIER DE ALPABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050809',N'05',N'0508',N'SAN JOSE DE USHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050810') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050810',N'05',N'0508',N'SARA SARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050901',N'05',N'0509',N'QUEROBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050902',N'05',N'0509',N'BELEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050903',N'05',N'0509',N'CHALCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050904',N'05',N'0509',N'CHILCAYOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050905',N'05',N'0509',N'HUACAÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050906',N'05',N'0509',N'MORCOLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050907',N'05',N'0509',N'PAICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050908') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050908',N'05',N'0509',N'SAN PEDRO DE LARCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050909') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050909',N'05',N'0509',N'SAN SALVADOR DE QUIJE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050910') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050910',N'05',N'0509',N'SANTIAGO DE PAUCARAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'050911') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'050911',N'05',N'0509',N'SORAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051001',N'05',N'0510',N'HUANCAPI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051002',N'05',N'0510',N'ALCAMENCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051003',N'05',N'0510',N'APONGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051004',N'05',N'0510',N'ASQUIPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051005',N'05',N'0510',N'CANARIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051006',N'05',N'0510',N'CAYARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051007',N'05',N'0510',N'COLCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051008') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051008',N'05',N'0510',N'HUAMANQUIQUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051009') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051009',N'05',N'0510',N'HUANCARAYLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051010') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051010',N'05',N'0510',N'HUAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051011') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051011',N'05',N'0510',N'SARHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051012') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051012',N'05',N'0510',N'VILCANCHOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051101',N'05',N'0511',N'VILCAS HUAMAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051102',N'05',N'0511',N'ACCOMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051103',N'05',N'0511',N'CARHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051104',N'05',N'0511',N'CONCEPCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051105',N'05',N'0511',N'HUAMBALPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051106',N'05',N'0511',N'INDEPENDENCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051107',N'05',N'0511',N'SAURAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'051108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'051108',N'05',N'0511',N'VISCHONGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060101',N'06',N'0601',N'CAJAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060102',N'06',N'0601',N'ASUNCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060103',N'06',N'0601',N'CHETILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060104',N'06',N'0601',N'COSPAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060105',N'06',N'0601',N'ENCAÑADA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060106',N'06',N'0601',N'JESUS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060107',N'06',N'0601',N'LLACANORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060108',N'06',N'0601',N'LOS BAÑOS DEL INCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060109',N'06',N'0601',N'MAGDALENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060110',N'06',N'0601',N'MATARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060111',N'06',N'0601',N'NAMORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060112',N'06',N'0601',N'SAN JUAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060201',N'06',N'0602',N'CAJABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060202',N'06',N'0602',N'CACHACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060203',N'06',N'0602',N'CONDEBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060204',N'06',N'0602',N'SITACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060301',N'06',N'0603',N'CELENDIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060302',N'06',N'0603',N'CHUMUCH',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060303',N'06',N'0603',N'CORTEGANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060304',N'06',N'0603',N'HUASMIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060305',N'06',N'0603',N'JORGE CHAVEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060306',N'06',N'0603',N'JOSE GALVEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060307',N'06',N'0603',N'MIGUEL IGLESIAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060308',N'06',N'0603',N'OXAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060309',N'06',N'0603',N'SOROCHUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060310',N'06',N'0603',N'SUCRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060311',N'06',N'0603',N'UTCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060312') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060312',N'06',N'0603',N'LA LIBERTAD DE PALLAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060401',N'06',N'0604',N'CHOTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060402',N'06',N'0604',N'ANGUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060403',N'06',N'0604',N'CHADIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060404',N'06',N'0604',N'CHIGUIRIP',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060405',N'06',N'0604',N'CHIMBAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060406',N'06',N'0604',N'CHOROPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060407',N'06',N'0604',N'COCHABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060408',N'06',N'0604',N'CONCHAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060409',N'06',N'0604',N'HUAMBOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060410',N'06',N'0604',N'LAJAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060411') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060411',N'06',N'0604',N'LLAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060412') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060412',N'06',N'0604',N'MIRACOSTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060413') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060413',N'06',N'0604',N'PACCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060414') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060414',N'06',N'0604',N'PION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060415') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060415',N'06',N'0604',N'QUEROCOTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060416') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060416',N'06',N'0604',N'SAN JUAN DE LICUPIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060417') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060417',N'06',N'0604',N'TACABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060418') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060418',N'06',N'0604',N'TOCMOCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060419') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060419',N'06',N'0604',N'CHALAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060501',N'06',N'0605',N'CONTUMAZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060502',N'06',N'0605',N'CHILETE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060503',N'06',N'0605',N'CUPISNIQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060504',N'06',N'0605',N'GUZMANGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060505',N'06',N'0605',N'SAN BENITO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060506',N'06',N'0605',N'SANTA CRUZ DE TOLED',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060507',N'06',N'0605',N'TANTARICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060508',N'06',N'0605',N'YONAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060601',N'06',N'0606',N'CUTERVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060602',N'06',N'0606',N'CALLAYUC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060603',N'06',N'0606',N'CHOROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060604',N'06',N'0606',N'CUJILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060605',N'06',N'0606',N'LA RAMADA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060606',N'06',N'0606',N'PIMPINGOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060607',N'06',N'0606',N'QUEROCOTILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060608',N'06',N'0606',N'SAN ANDRES DE CUTERVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060609',N'06',N'0606',N'SAN JUAN DE CUTERVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060610',N'06',N'0606',N'SAN LUIS DE LUCMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060611',N'06',N'0606',N'SANTA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060612') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060612',N'06',N'0606',N'SANTO DOMINGO DE LA CAPILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060613') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060613',N'06',N'0606',N'SANTO TOMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060614') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060614',N'06',N'0606',N'SOCOTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060615') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060615',N'06',N'0606',N'TORIBIO CASANOVA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060701',N'06',N'0607',N'BAMBAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060702',N'06',N'0607',N'CHUGUR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060703',N'06',N'0607',N'HUALGAYOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060801',N'06',N'0608',N'JAEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060802',N'06',N'0608',N'BELLAVISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060803',N'06',N'0608',N'CHONTALI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060804',N'06',N'0608',N'COLASAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060805',N'06',N'0608',N'HUABAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060806',N'06',N'0608',N'LAS PIRIAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060807',N'06',N'0608',N'POMAHUACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060808',N'06',N'0608',N'PUCARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060809',N'06',N'0608',N'SALLIQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060810') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060810',N'06',N'0608',N'SAN FELIPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060811') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060811',N'06',N'0608',N'SAN JOSE DEL ALTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060812') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060812',N'06',N'0608',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060901',N'06',N'0609',N'SAN IGNACIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060902',N'06',N'0609',N'CHIRINOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060903',N'06',N'0609',N'HUARANGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060904',N'06',N'0609',N'LA COIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060905',N'06',N'0609',N'NAMBALLE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060906',N'06',N'0609',N'SAN JOSE DE LOURDES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'060907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'060907',N'06',N'0609',N'TABACONAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061001',N'06',N'0610',N'PEDRO GALVEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061002',N'06',N'0610',N'CHANCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061003',N'06',N'0610',N'EDUARDO VILLANUEVA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061004',N'06',N'0610',N'GREGORIO PITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061005',N'06',N'0610',N'ICHOCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061006',N'06',N'0610',N'JOSE MANUEL QUIROZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061007',N'06',N'0610',N'JOSE SABOGAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061101',N'06',N'0611',N'SAN MIGUEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061102',N'06',N'0611',N'BOLIVAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061103',N'06',N'0611',N'CALQUIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061104',N'06',N'0611',N'CATILLUC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061105',N'06',N'0611',N'EL PRADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061106',N'06',N'0611',N'LA FLORIDA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061107',N'06',N'0611',N'LLAPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061108',N'06',N'0611',N'NANCHOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061109',N'06',N'0611',N'NIEPOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061110',N'06',N'0611',N'SAN GREGORIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061111',N'06',N'0611',N'SAN SILVESTRE DE COCHAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061112',N'06',N'0611',N'TONGOD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061113',N'06',N'0611',N'UNION AGUA BLANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061201',N'06',N'0612',N'SAN PABLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061202',N'06',N'0612',N'SAN BERNARDINO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061203',N'06',N'0612',N'SAN LUIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061204',N'06',N'0612',N'TUMBADEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061301',N'06',N'0613',N'SANTA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061302',N'06',N'0613',N'ANDABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061303',N'06',N'0613',N'CATACHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061304',N'06',N'0613',N'CHANCAYBAÑOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061305',N'06',N'0613',N'LA ESPERANZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061306',N'06',N'0613',N'NINABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061307',N'06',N'0613',N'PULAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061308',N'06',N'0613',N'SAUCEPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061309',N'06',N'0613',N'SEXI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061310',N'06',N'0613',N'UTICYACU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'061311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'061311',N'06',N'0613',N'YAUYUCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070101',N'07',N'0701',N'CALLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070102',N'07',N'0701',N'BELLAVISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070103',N'07',N'0701',N'CARMEN DE LA LEGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070104',N'07',N'0701',N'LA PERLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070105',N'07',N'0701',N'LA PUNTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070106',N'07',N'0701',N'VENTANILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'070107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'070107',N'07',N'0701',N'MI PERU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080101',N'08',N'0801',N'CUSCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080102',N'08',N'0801',N'CCORCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080103',N'08',N'0801',N'POROY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080104',N'08',N'0801',N'SAN JERONIMO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080105',N'08',N'0801',N'SAN SEBASTIAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080106',N'08',N'0801',N'SANTIAGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080107',N'08',N'0801',N'SAYLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080108',N'08',N'0801',N'WANCHAQ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080201',N'08',N'0802',N'ACOMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080202',N'08',N'0802',N'ACOPIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080203',N'08',N'0802',N'ACOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080204',N'08',N'0802',N'MOSOC LLACTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080205',N'08',N'0802',N'POMACANCHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080206',N'08',N'0802',N'RONDOCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080207',N'08',N'0802',N'SANGARARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080301',N'08',N'0803',N'ANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080302',N'08',N'0803',N'ANCAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080303',N'08',N'0803',N'CACHIMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080304',N'08',N'0803',N'CHINCHAYPUJIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080305',N'08',N'0803',N'HUAROCONDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080306',N'08',N'0803',N'LIMATAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080307',N'08',N'0803',N'MOLLEPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080308',N'08',N'0803',N'PUCYURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080309',N'08',N'0803',N'ZURITE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080401',N'08',N'0804',N'CALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080402',N'08',N'0804',N'COYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080403',N'08',N'0804',N'LAMAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080404',N'08',N'0804',N'LARES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080405',N'08',N'0804',N'PISAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080406',N'08',N'0804',N'SAN SALVADOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080407',N'08',N'0804',N'TARAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080408',N'08',N'0804',N'YANATILE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080501',N'08',N'0805',N'YANAOCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080502',N'08',N'0805',N'CHECCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080503',N'08',N'0805',N'KUNTURKANKI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080504',N'08',N'0805',N'LANGUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080505',N'08',N'0805',N'LAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080506',N'08',N'0805',N'PAMPAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080507',N'08',N'0805',N'QUEHUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080508',N'08',N'0805',N'TUPAC AMARU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080601',N'08',N'0806',N'SICUANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080602',N'08',N'0806',N'CHECACUPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080603',N'08',N'0806',N'COMBAPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080604',N'08',N'0806',N'MARANGANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080605',N'08',N'0806',N'PITUMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080606',N'08',N'0806',N'SAN PABLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080607',N'08',N'0806',N'SAN PEDRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080608',N'08',N'0806',N'TINTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080701',N'08',N'0807',N'SANTO TOMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080702',N'08',N'0807',N'CAPACMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080703',N'08',N'0807',N'CHAMACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080704',N'08',N'0807',N'COLQUEMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080705',N'08',N'0807',N'LIVITACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080706',N'08',N'0807',N'LLUSCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080707',N'08',N'0807',N'QUIÑOTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080708',N'08',N'0807',N'VELILLE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080801',N'08',N'0808',N'ESPINAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080802',N'08',N'0808',N'CONDOROMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080803',N'08',N'0808',N'COPORAQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080804',N'08',N'0808',N'OCORURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080805',N'08',N'0808',N'PALLPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080806',N'08',N'0808',N'PICHIGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080807',N'08',N'0808',N'SUYCKUTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080808',N'08',N'0808',N'ALTO PICHIGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080901',N'08',N'0809',N'SANTA ANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080902',N'08',N'0809',N'ECHARATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080903',N'08',N'0809',N'HUAYOPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080904',N'08',N'0809',N'MARANURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080905',N'08',N'0809',N'OCOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080906',N'08',N'0809',N'QUELLOUNO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080907',N'08',N'0809',N'KIMBIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080908') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080908',N'08',N'0809',N'SANTA TERESA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080909') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080909',N'08',N'0809',N'VILCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080910') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080910',N'08',N'0809',N'PICHARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080911') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080911',N'08',N'0809',N'INKAWASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080912') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080912',N'08',N'0809',N'VILLA VIRGEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080913') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080913',N'08',N'0809',N'VILLA KINTIARINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080914') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080914',N'08',N'0809',N'MEGANTONI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080915') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080915',N'08',N'0809',N'KUMPIRUSHIATO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080916') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080916',N'08',N'0809',N'CIELO PUNCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080917') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080917',N'08',N'0809',N'MANITEA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'080918') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'080918',N'08',N'0809',N'UNION ASHANINKA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081001',N'08',N'0810',N'PARURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081002',N'08',N'0810',N'ACCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081003',N'08',N'0810',N'CCAPI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081004',N'08',N'0810',N'COLCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081005',N'08',N'0810',N'HUANOQUITE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081006',N'08',N'0810',N'OMACHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081007',N'08',N'0810',N'PACCARITAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081008') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081008',N'08',N'0810',N'PILLPINTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081009') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081009',N'08',N'0810',N'YAURISQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081101',N'08',N'0811',N'PAUCARTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081102',N'08',N'0811',N'CAICAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081103',N'08',N'0811',N'CHALLABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081104',N'08',N'0811',N'COLQUEPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081105',N'08',N'0811',N'HUANCARANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081106',N'08',N'0811',N'KOSÑIPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081201',N'08',N'0812',N'URCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081202',N'08',N'0812',N'ANDAHUAYLILLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081203',N'08',N'0812',N'CAMANTI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081204',N'08',N'0812',N'CCARHUAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081205',N'08',N'0812',N'CCATCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081206',N'08',N'0812',N'CUSIPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081207',N'08',N'0812',N'HUARO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081208',N'08',N'0812',N'LUCRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081209',N'08',N'0812',N'MARCAPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081210',N'08',N'0812',N'OCONGATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081211',N'08',N'0812',N'OROPESA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081212') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081212',N'08',N'0812',N'QUIQUIJANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081301',N'08',N'0813',N'URUBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081302',N'08',N'0813',N'CHINCHERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081303',N'08',N'0813',N'HUAYLLABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081304',N'08',N'0813',N'MACHUPICCHU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081305',N'08',N'0813',N'MARAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081306',N'08',N'0813',N'OLLANTAYTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'081307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'081307',N'08',N'0813',N'YUCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090101',N'09',N'0901',N'HUANCAVELICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090102',N'09',N'0901',N'ACOBAMBILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090103',N'09',N'0901',N'ACORIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090104',N'09',N'0901',N'CONAYCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090105',N'09',N'0901',N'CUENCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090106',N'09',N'0901',N'HUACHOCOLPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090107',N'09',N'0901',N'HUAYLLAHUARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090108',N'09',N'0901',N'IZCUCHACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090109',N'09',N'0901',N'LARIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090110',N'09',N'0901',N'MANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090111',N'09',N'0901',N'MARISCAL CACERES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090112',N'09',N'0901',N'MOYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090113',N'09',N'0901',N'NUEVO OCCORO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090114',N'09',N'0901',N'PALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090115',N'09',N'0901',N'PILCHACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090116',N'09',N'0901',N'VILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090117') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090117',N'09',N'0901',N'YAULI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090118') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090118',N'09',N'0901',N'ASCENSION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090119') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090119',N'09',N'0901',N'HUANDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090201',N'09',N'0902',N'ACOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090202',N'09',N'0902',N'ANDABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090203',N'09',N'0902',N'ANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090204',N'09',N'0902',N'CAJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090205',N'09',N'0902',N'MARCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090206',N'09',N'0902',N'PAUCARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090207',N'09',N'0902',N'POMACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090208',N'09',N'0902',N'ROSARIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090301',N'09',N'0903',N'LIRCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090302',N'09',N'0903',N'ANCHONGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090303',N'09',N'0903',N'CALLANMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090304',N'09',N'0903',N'CCOCHACCASA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090305',N'09',N'0903',N'CHINCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090306',N'09',N'0903',N'CONGALLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090307',N'09',N'0903',N'HUANCA-HUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090308',N'09',N'0903',N'HUAYLLAY GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090309',N'09',N'0903',N'JULCAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090310',N'09',N'0903',N'SAN ANTONIO DE ANTAPARCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090311',N'09',N'0903',N'SANTO TOMAS DE PATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090312') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090312',N'09',N'0903',N'SECCLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090401',N'09',N'0904',N'CASTROVIRREYNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090402',N'09',N'0904',N'ARMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090403',N'09',N'0904',N'AURAHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090404',N'09',N'0904',N'CAPILLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090405',N'09',N'0904',N'CHUPAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090406',N'09',N'0904',N'COCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090407',N'09',N'0904',N'HUACHOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090408',N'09',N'0904',N'HUAMATAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090409',N'09',N'0904',N'MOLLEPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090410',N'09',N'0904',N'SAN JUAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090411') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090411',N'09',N'0904',N'SANTA ANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090412') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090412',N'09',N'0904',N'TANTARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090413') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090413',N'09',N'0904',N'TICRAPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090501',N'09',N'0905',N'CHURCAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090502',N'09',N'0905',N'ANCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090503',N'09',N'0905',N'CHINCHIHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090504',N'09',N'0905',N'EL CARMEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090505',N'09',N'0905',N'LA MERCED',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090506',N'09',N'0905',N'LOCROJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090507',N'09',N'0905',N'PAUCARBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090508',N'09',N'0905',N'SAN MIGUEL DE MAYOCC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090509',N'09',N'0905',N'SAN PEDRO DE CORIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090510',N'09',N'0905',N'PACHAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090511',N'09',N'0905',N'COSME',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090601',N'09',N'0906',N'HUAYTARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090602',N'09',N'0906',N'AYAVI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090603',N'09',N'0906',N'CORDOVA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090604',N'09',N'0906',N'HUAYACUNDO ARMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090605',N'09',N'0906',N'LARAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090606',N'09',N'0906',N'OCOYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090607',N'09',N'0906',N'PILPICHACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090608',N'09',N'0906',N'QUERCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090609',N'09',N'0906',N'QUITO-ARMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090610',N'09',N'0906',N'SAN ANTONIO DE CUSICANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090611',N'09',N'0906',N'SAN FRANCISCO DE SANGAYAICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090612') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090612',N'09',N'0906',N'SAN ISIDRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090613') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090613',N'09',N'0906',N'SANTIAGO DE CHOCORVOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090614') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090614',N'09',N'0906',N'SANTIAGO DE QUIRAHUARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090615') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090615',N'09',N'0906',N'SANTO DOMINGO DE CAPILLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090616') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090616',N'09',N'0906',N'TAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090701',N'09',N'0907',N'PAMPAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090702',N'09',N'0907',N'ACOSTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090703',N'09',N'0907',N'ACRAQUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090704',N'09',N'0907',N'AHUAYCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090705',N'09',N'0907',N'COLCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090706',N'09',N'0907',N'DANIEL HERNANDEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090707',N'09',N'0907',N'HUACHOCOLPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090709',N'09',N'0907',N'HUARIBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090710') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090710',N'09',N'0907',N'ÑAHUIMPUQUIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090711') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090711',N'09',N'0907',N'PAZOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090713') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090713',N'09',N'0907',N'QUISHUAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090714') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090714',N'09',N'0907',N'SALCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090715') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090715',N'09',N'0907',N'SALCAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090716') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090716',N'09',N'0907',N'SAN MARCOS DE ROCCHAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090717') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090717',N'09',N'0907',N'SURCUBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090718') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090718',N'09',N'0907',N'TINTAY PUNCU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090719') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090719',N'09',N'0907',N'QUICHUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090720') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090720',N'09',N'0907',N'ANDAYMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090721') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090721',N'09',N'0907',N'ROBLE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090722') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090722',N'09',N'0907',N'PICHOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090723') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090723',N'09',N'0907',N'SANTIAGO DE TUCUMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090724') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090724',N'09',N'0907',N'LAMBRAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'090725') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'090725',N'09',N'0907',N'COCHABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100101',N'10',N'1001',N'HUANUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100102',N'10',N'1001',N'AMARILIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100103',N'10',N'1001',N'CHINCHAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100104',N'10',N'1001',N'CHURUBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100105',N'10',N'1001',N'MARGOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100106',N'10',N'1001',N'QUISQUI (KICHKI)',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100107',N'10',N'1001',N'SAN FRANCISCO DE CAYRAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100108',N'10',N'1001',N'SAN PEDRO DE CHAULAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100109',N'10',N'1001',N'SANTA MARIA DEL VALLE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100110',N'10',N'1001',N'YARUMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100111',N'10',N'1001',N'PILLCO MARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100112',N'10',N'1001',N'YACUS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100113',N'10',N'1001',N'SAN PABLO DE PILLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100201',N'10',N'1002',N'AMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100202',N'10',N'1002',N'CAYNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100203',N'10',N'1002',N'COLPAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100204',N'10',N'1002',N'CONCHAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100205',N'10',N'1002',N'HUACAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100206',N'10',N'1002',N'SAN FRANCISCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100207',N'10',N'1002',N'SAN RAFAEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100208',N'10',N'1002',N'TOMAY KICHWA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100301',N'10',N'1003',N'LA UNION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100307',N'10',N'1003',N'CHUQUIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100311',N'10',N'1003',N'MARIAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100313') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100313',N'10',N'1003',N'PACHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100316') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100316',N'10',N'1003',N'QUIVILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100317') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100317',N'10',N'1003',N'RIPAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100321') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100321',N'10',N'1003',N'SHUNQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100322') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100322',N'10',N'1003',N'SILLAPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100323') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100323',N'10',N'1003',N'YANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100401',N'10',N'1004',N'HUACAYBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100402',N'10',N'1004',N'CANCHABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100403',N'10',N'1004',N'COCHABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100404',N'10',N'1004',N'PINRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100501',N'10',N'1005',N'LLATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100502',N'10',N'1005',N'ARANCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100503',N'10',N'1005',N'CHAVIN DE PARIARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100504',N'10',N'1005',N'JACAS GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100505',N'10',N'1005',N'JIRCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100506',N'10',N'1005',N'MIRAFLORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100507',N'10',N'1005',N'MONZON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100508',N'10',N'1005',N'PUNCHAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100509',N'10',N'1005',N'PUÑOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100510',N'10',N'1005',N'SINGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100511',N'10',N'1005',N'TANTAMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100601',N'10',N'1006',N'RUPA-RUPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100602',N'10',N'1006',N'DANIEL ALOMIA ROBLES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100603',N'10',N'1006',N'HERMILIO VALDIZAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100604',N'10',N'1006',N'JOSE CRESPO Y CASTILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100605',N'10',N'1006',N'LUYANDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100606',N'10',N'1006',N'MARIANO DAMASO BERAUN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100607',N'10',N'1006',N'PUCAYACU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100608',N'10',N'1006',N'CASTILLO GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100609',N'10',N'1006',N'PUEBLO NUEVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100610',N'10',N'1006',N'SANTO DOMINGO DE ANDA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100701',N'10',N'1007',N'HUACRACHUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100702',N'10',N'1007',N'CHOLON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100703',N'10',N'1007',N'SAN BUENAVENTURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100704',N'10',N'1007',N'LA MORADA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100705',N'10',N'1007',N'SANTA ROSA DE ALTO YANAJANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100801',N'10',N'1008',N'PANAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100802',N'10',N'1008',N'CHAGLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100803',N'10',N'1008',N'MOLINO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100804',N'10',N'1008',N'UMARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100901',N'10',N'1009',N'PUERTO INCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100902',N'10',N'1009',N'CODO DEL POZUZO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100903',N'10',N'1009',N'HONORIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100904',N'10',N'1009',N'TOURNAVISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'100905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'100905',N'10',N'1009',N'YUYAPICHIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101001',N'10',N'1010',N'JESUS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101002',N'10',N'1010',N'BAÑOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101003',N'10',N'1010',N'JIVIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101004',N'10',N'1010',N'QUEROPALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101005',N'10',N'1010',N'RONDOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101006',N'10',N'1010',N'SAN FRANCISCO DE ASIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101007',N'10',N'1010',N'SAN MIGUEL DE CAURI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101101',N'10',N'1011',N'CHAVINILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101102',N'10',N'1011',N'CAHUAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101103',N'10',N'1011',N'CHACABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101104',N'10',N'1011',N'APARICIO POMARES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101105',N'10',N'1011',N'JACAS CHICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101106',N'10',N'1011',N'OBAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101107',N'10',N'1011',N'PAMPAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'101108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'101108',N'10',N'1011',N'CHORAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110101',N'11',N'1101',N'ICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110102',N'11',N'1101',N'LA TINGUIÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110103',N'11',N'1101',N'LOS AQUIJES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110104',N'11',N'1101',N'OCUCAJE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110105',N'11',N'1101',N'PACHACUTEC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110106',N'11',N'1101',N'PARCONA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110107',N'11',N'1101',N'PUEBLO NUEVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110108',N'11',N'1101',N'SALAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110109',N'11',N'1101',N'SAN JOSE DE LOS MOLINOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110110',N'11',N'1101',N'SAN JUAN BAUTISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110111',N'11',N'1101',N'SANTIAGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110112',N'11',N'1101',N'SUBTANJALLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110113',N'11',N'1101',N'TATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110114',N'11',N'1101',N'YAUCA DEL ROSARIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110201',N'11',N'1102',N'CHINCHA ALTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110202',N'11',N'1102',N'ALTO LARAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110203',N'11',N'1102',N'CHAVIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110204',N'11',N'1102',N'CHINCHA BAJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110205',N'11',N'1102',N'EL CARMEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110206',N'11',N'1102',N'GROCIO PRADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110207',N'11',N'1102',N'PUEBLO NUEVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110208',N'11',N'1102',N'SAN JUAN DE YANAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110209',N'11',N'1102',N'SAN PEDRO DE HUACARPANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110210',N'11',N'1102',N'SUNAMPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110211',N'11',N'1102',N'TAMBO DE MORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110301',N'11',N'1103',N'NASCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110302',N'11',N'1103',N'CHANGUILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110303',N'11',N'1103',N'EL INGENIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110304',N'11',N'1103',N'MARCONA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110305',N'11',N'1103',N'VISTA ALEGRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110401',N'11',N'1104',N'PALPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110402',N'11',N'1104',N'LLIPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110403',N'11',N'1104',N'RIO GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110404',N'11',N'1104',N'SANTA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110405',N'11',N'1104',N'TIBILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110501',N'11',N'1105',N'PISCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110502',N'11',N'1105',N'HUANCANO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110503',N'11',N'1105',N'HUMAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110504',N'11',N'1105',N'INDEPENDENCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110505',N'11',N'1105',N'PARACAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110506',N'11',N'1105',N'SAN ANDRES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110507',N'11',N'1105',N'SAN CLEMENTE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'110508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'110508',N'11',N'1105',N'TUPAC AMARU INCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120101',N'12',N'1201',N'HUANCAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120104',N'12',N'1201',N'CARHUACALLANGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120105',N'12',N'1201',N'CHACAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120106',N'12',N'1201',N'CHICCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120107',N'12',N'1201',N'CHILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120108',N'12',N'1201',N'CHONGOS ALTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120111',N'12',N'1201',N'CHUPURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120112',N'12',N'1201',N'COLCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120113',N'12',N'1201',N'CULLHUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120114',N'12',N'1201',N'EL TAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120116',N'12',N'1201',N'HUACRAPUQUIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120117') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120117',N'12',N'1201',N'HUALHUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120119') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120119',N'12',N'1201',N'HUANCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120120') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120120',N'12',N'1201',N'HUASICANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120121') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120121',N'12',N'1201',N'HUAYUCACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120122') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120122',N'12',N'1201',N'INGENIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120124') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120124',N'12',N'1201',N'PARIAHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120125') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120125',N'12',N'1201',N'PILCOMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120126') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120126',N'12',N'1201',N'PUCARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120127') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120127',N'12',N'1201',N'QUICHUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120128') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120128',N'12',N'1201',N'QUILCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120129') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120129',N'12',N'1201',N'SAN AGUSTIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120130') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120130',N'12',N'1201',N'SAN JERONIMO DE TUNAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120132') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120132',N'12',N'1201',N'SAÑO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120133') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120133',N'12',N'1201',N'SAPALLANGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120134') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120134',N'12',N'1201',N'SICAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120135') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120135',N'12',N'1201',N'SANTO DOMINGO DE ACOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120136') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120136',N'12',N'1201',N'VIQUES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120201',N'12',N'1202',N'CONCEPCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120202',N'12',N'1202',N'ACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120203',N'12',N'1202',N'ANDAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120204',N'12',N'1202',N'CHAMBARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120205',N'12',N'1202',N'COCHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120206',N'12',N'1202',N'COMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120207',N'12',N'1202',N'HEROINAS TOLEDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120208',N'12',N'1202',N'MANZANARES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120209',N'12',N'1202',N'MARISCAL CASTILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120210',N'12',N'1202',N'MATAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120211',N'12',N'1202',N'MITO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120212') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120212',N'12',N'1202',N'NUEVE DE JULIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120213') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120213',N'12',N'1202',N'ORCOTUNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120214') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120214',N'12',N'1202',N'SAN JOSE DE QUERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120215') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120215',N'12',N'1202',N'SANTA ROSA DE OCOPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120301',N'12',N'1203',N'CHANCHAMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120302',N'12',N'1203',N'PERENE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120303',N'12',N'1203',N'PICHANAQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120304',N'12',N'1203',N'SAN LUIS DE SHUARO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120305',N'12',N'1203',N'SAN RAMON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120306',N'12',N'1203',N'VITOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120401',N'12',N'1204',N'JAUJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120402',N'12',N'1204',N'ACOLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120403',N'12',N'1204',N'APATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120404',N'12',N'1204',N'ATAURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120405',N'12',N'1204',N'CANCHAYLLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120406',N'12',N'1204',N'CURICACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120407',N'12',N'1204',N'EL MANTARO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120408',N'12',N'1204',N'HUAMALI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120409',N'12',N'1204',N'HUARIPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120410',N'12',N'1204',N'HUERTAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120411') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120411',N'12',N'1204',N'JANJAILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120412') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120412',N'12',N'1204',N'JULCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120413') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120413',N'12',N'1204',N'LEONOR ORDOÑEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120414') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120414',N'12',N'1204',N'LLOCLLAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120415') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120415',N'12',N'1204',N'MARCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120416') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120416',N'12',N'1204',N'MASMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120417') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120417',N'12',N'1204',N'MASMA CHICCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120418') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120418',N'12',N'1204',N'MOLINOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120419') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120419',N'12',N'1204',N'MONOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120420') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120420',N'12',N'1204',N'MUQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120421') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120421',N'12',N'1204',N'MUQUIYAUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120422') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120422',N'12',N'1204',N'PACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120423') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120423',N'12',N'1204',N'PACCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120424') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120424',N'12',N'1204',N'PANCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120425') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120425',N'12',N'1204',N'PARCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120426') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120426',N'12',N'1204',N'POMACANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120427') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120427',N'12',N'1204',N'RICRAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120428') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120428',N'12',N'1204',N'SAN LORENZO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120429') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120429',N'12',N'1204',N'SAN PEDRO DE CHUNAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120430') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120430',N'12',N'1204',N'SAUSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120431') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120431',N'12',N'1204',N'SINCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120432') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120432',N'12',N'1204',N'TUNAN MARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120433') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120433',N'12',N'1204',N'YAULI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120434') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120434',N'12',N'1204',N'YAUYOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120501',N'12',N'1205',N'JUNIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120502',N'12',N'1205',N'CARHUAMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120503',N'12',N'1205',N'ONDORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120504',N'12',N'1205',N'ULCUMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120601',N'12',N'1206',N'SATIPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120602',N'12',N'1206',N'COVIRIALI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120603',N'12',N'1206',N'LLAYLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120604',N'12',N'1206',N'MAZAMARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120605',N'12',N'1206',N'PAMPA HERMOSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120606',N'12',N'1206',N'PANGOA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120607',N'12',N'1206',N'RIO NEGRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120608',N'12',N'1206',N'RIO TAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120609',N'12',N'1206',N'VIZCATÃN DEL ENE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120701',N'12',N'1207',N'TARMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120702',N'12',N'1207',N'ACOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120703',N'12',N'1207',N'HUARICOLCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120704',N'12',N'1207',N'HUASAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120705',N'12',N'1207',N'LA UNION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120706',N'12',N'1207',N'PALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120707',N'12',N'1207',N'PALCAMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120708',N'12',N'1207',N'SAN PEDRO DE CAJAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120709',N'12',N'1207',N'TAPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120801',N'12',N'1208',N'LA OROYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120802',N'12',N'1208',N'CHACAPALPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120803',N'12',N'1208',N'HUAY-HUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120804',N'12',N'1208',N'MARCAPOMACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120805',N'12',N'1208',N'MOROCOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120806',N'12',N'1208',N'PACCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120807',N'12',N'1208',N'SANTA BARBARA DE CARHUACAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120808',N'12',N'1208',N'SANTA ROSA DE SACCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120809',N'12',N'1208',N'SUITUCANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120810') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120810',N'12',N'1208',N'YAULI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120901',N'12',N'1209',N'CHUPACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120902',N'12',N'1209',N'AHUAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120903',N'12',N'1209',N'CHONGOS BAJO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120904',N'12',N'1209',N'HUACHAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120905',N'12',N'1209',N'HUAMANCACA CHICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120906',N'12',N'1209',N'SAN JUAN DE ISCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120907',N'12',N'1209',N'SAN JUAN DE JARPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120908') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120908',N'12',N'1209',N'TRES DE DICIEMBRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'120909') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'120909',N'12',N'1209',N'YANACANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130101',N'13',N'1301',N'TRUJILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130102',N'13',N'1301',N'EL PORVENIR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130103',N'13',N'1301',N'FLORENCIA DE MORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130104',N'13',N'1301',N'HUANCHACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130105',N'13',N'1301',N'LA ESPERANZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130106',N'13',N'1301',N'LAREDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130107',N'13',N'1301',N'MOCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130108',N'13',N'1301',N'POROTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130109',N'13',N'1301',N'SALAVERRY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130110',N'13',N'1301',N'SIMBAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130111',N'13',N'1301',N'VICTOR LARCO HERRERA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130201',N'13',N'1302',N'ASCOPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130202',N'13',N'1302',N'CHICAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130203',N'13',N'1302',N'CHOCOPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130204',N'13',N'1302',N'MAGDALENA DE CAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130205',N'13',N'1302',N'PAIJAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130206',N'13',N'1302',N'RAZURI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130207',N'13',N'1302',N'SANTIAGO DE CAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130208',N'13',N'1302',N'CASA GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130301',N'13',N'1303',N'BOLIVAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130302',N'13',N'1303',N'BAMBAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130303',N'13',N'1303',N'CONDORMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130304',N'13',N'1303',N'LONGOTEA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130305',N'13',N'1303',N'UCHUMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130306',N'13',N'1303',N'UCUNCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130401',N'13',N'1304',N'CHEPEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130402',N'13',N'1304',N'PACANGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130403',N'13',N'1304',N'PUEBLO NUEVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130501',N'13',N'1305',N'JULCAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130502',N'13',N'1305',N'CALAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130503',N'13',N'1305',N'CARABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130504',N'13',N'1305',N'HUASO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130601',N'13',N'1306',N'OTUZCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130602',N'13',N'1306',N'AGALLPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130604',N'13',N'1306',N'CHARAT',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130605',N'13',N'1306',N'HUARANCHAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130606',N'13',N'1306',N'LA CUESTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130608',N'13',N'1306',N'MACHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130610',N'13',N'1306',N'PARANDAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130611',N'13',N'1306',N'SALPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130613') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130613',N'13',N'1306',N'SINSICAP',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130614') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130614',N'13',N'1306',N'USQUIL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130701',N'13',N'1307',N'SAN PEDRO DE LLOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130702',N'13',N'1307',N'GUADALUPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130703',N'13',N'1307',N'JEQUETEPEQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130704',N'13',N'1307',N'PACASMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130705',N'13',N'1307',N'SAN JOSE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130801',N'13',N'1308',N'TAYABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130802',N'13',N'1308',N'BULDIBUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130803',N'13',N'1308',N'CHILLIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130804',N'13',N'1308',N'HUANCASPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130805',N'13',N'1308',N'HUAYLILLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130806',N'13',N'1308',N'HUAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130807',N'13',N'1308',N'ONGON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130808',N'13',N'1308',N'PARCOY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130809',N'13',N'1308',N'PATAZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130810') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130810',N'13',N'1308',N'PIAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130811') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130811',N'13',N'1308',N'SANTIAGO DE CHALLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130812') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130812',N'13',N'1308',N'TAURIJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130813') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130813',N'13',N'1308',N'URPAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130901',N'13',N'1309',N'HUAMACHUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130902',N'13',N'1309',N'CHUGAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130903',N'13',N'1309',N'COCHORCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130904',N'13',N'1309',N'CURGOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130905',N'13',N'1309',N'MARCABAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130906',N'13',N'1309',N'SANAGORAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130907',N'13',N'1309',N'SARIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'130908') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'130908',N'13',N'1309',N'SARTIMBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131001',N'13',N'1310',N'SANTIAGO DE CHUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131002',N'13',N'1310',N'ANGASMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131003',N'13',N'1310',N'CACHICADAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131004',N'13',N'1310',N'MOLLEBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131005',N'13',N'1310',N'MOLLEPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131006',N'13',N'1310',N'QUIRUVILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131007',N'13',N'1310',N'SANTA CRUZ DE CHUCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131008') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131008',N'13',N'1310',N'SITABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131101',N'13',N'1311',N'CASCAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131102',N'13',N'1311',N'LUCMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131103',N'13',N'1311',N'MARMOT',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131104',N'13',N'1311',N'SAYAPULLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131201',N'13',N'1312',N'VIRU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131202',N'13',N'1312',N'CHAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'131203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'131203',N'13',N'1312',N'GUADALUPITO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140101',N'14',N'1401',N'CHICLAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140102',N'14',N'1401',N'CHONGOYAPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140103',N'14',N'1401',N'ETEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140104',N'14',N'1401',N'ETEN PUERTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140105',N'14',N'1401',N'JOSE LEONARDO ORTIZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140106',N'14',N'1401',N'LA VICTORIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140107',N'14',N'1401',N'LAGUNAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140108',N'14',N'1401',N'MONSEFU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140109',N'14',N'1401',N'NUEVA ARICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140110',N'14',N'1401',N'OYOTUN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140111',N'14',N'1401',N'PICSI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140112',N'14',N'1401',N'PIMENTEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140113',N'14',N'1401',N'REQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140114',N'14',N'1401',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140115',N'14',N'1401',N'SAÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140116',N'14',N'1401',N'CAYALTI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140117') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140117',N'14',N'1401',N'PATAPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140118') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140118',N'14',N'1401',N'POMALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140119') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140119',N'14',N'1401',N'PUCALA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140120') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140120',N'14',N'1401',N'TUMAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140201',N'14',N'1402',N'FERREÑAFE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140202',N'14',N'1402',N'CAÑARIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140203',N'14',N'1402',N'INCAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140204',N'14',N'1402',N'MANUEL ANTONIO MESONES MURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140205',N'14',N'1402',N'PITIPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140206',N'14',N'1402',N'PUEBLO NUEVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140301',N'14',N'1403',N'LAMBAYEQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140302',N'14',N'1403',N'CHOCHOPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140303',N'14',N'1403',N'ILLIMO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140304',N'14',N'1403',N'JAYANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140305',N'14',N'1403',N'MOCHUMI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140306',N'14',N'1403',N'MORROPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140307',N'14',N'1403',N'MOTUPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140308',N'14',N'1403',N'OLMOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140309',N'14',N'1403',N'PACORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140310',N'14',N'1403',N'SALAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140311') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140311',N'14',N'1403',N'SAN JOSE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'140312') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'140312',N'14',N'1403',N'TUCUME',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150101',N'15',N'1501',N'LIMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150102',N'15',N'1501',N'ANCON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150103',N'15',N'1501',N'ATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150104',N'15',N'1501',N'BARRANCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150105',N'15',N'1501',N'BREÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150106',N'15',N'1501',N'CARABAYLLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150107',N'15',N'1501',N'CHACLACAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150108',N'15',N'1501',N'CHORRILLOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150109',N'15',N'1501',N'CIENEGUILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150110',N'15',N'1501',N'COMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150111',N'15',N'1501',N'EL AGUSTINO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150112',N'15',N'1501',N'INDEPENDENCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150113',N'15',N'1501',N'JESUS MARIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150114',N'15',N'1501',N'LA MOLINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150115',N'15',N'1501',N'LA VICTORIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150116') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150116',N'15',N'1501',N'LINCE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150117') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150117',N'15',N'1501',N'LOS OLIVOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150118') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150118',N'15',N'1501',N'LURIGANCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150119') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150119',N'15',N'1501',N'LURIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150120') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150120',N'15',N'1501',N'MAGDALENA DEL MAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150121') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150121',N'15',N'1501',N'PUEBLO LIBRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150122') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150122',N'15',N'1501',N'MIRAFLORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150123') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150123',N'15',N'1501',N'PACHACAMAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150124') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150124',N'15',N'1501',N'PUCUSANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150125') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150125',N'15',N'1501',N'PUENTE PIEDRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150126') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150126',N'15',N'1501',N'PUNTA HERMOSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150127') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150127',N'15',N'1501',N'PUNTA NEGRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150128') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150128',N'15',N'1501',N'RIMAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150129') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150129',N'15',N'1501',N'SAN BARTOLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150130') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150130',N'15',N'1501',N'SAN BORJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150131') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150131',N'15',N'1501',N'SAN ISIDRO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150132') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150132',N'15',N'1501',N'SAN JUAN DE LURIGANCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150133') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150133',N'15',N'1501',N'SAN JUAN DE MIRAFLORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150134') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150134',N'15',N'1501',N'SAN LUIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150135') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150135',N'15',N'1501',N'SAN MARTIN DE PORRES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150136') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150136',N'15',N'1501',N'SAN MIGUEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150137') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150137',N'15',N'1501',N'SANTA ANITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150138') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150138',N'15',N'1501',N'SANTA MARIA DEL MAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150139') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150139',N'15',N'1501',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150140') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150140',N'15',N'1501',N'SANTIAGO DE SURCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150141') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150141',N'15',N'1501',N'SURQUILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150142') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150142',N'15',N'1501',N'VILLA EL SALVADOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150143') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150143',N'15',N'1501',N'VILLA MARIA DEL TRIUNFO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150201',N'15',N'1502',N'BARRANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150202',N'15',N'1502',N'PARAMONGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150203',N'15',N'1502',N'PATIVILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150204',N'15',N'1502',N'SUPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150205',N'15',N'1502',N'SUPE PUERTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150301',N'15',N'1503',N'CAJATAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150302',N'15',N'1503',N'COPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150303',N'15',N'1503',N'GORGOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150304',N'15',N'1503',N'HUANCAPON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150305',N'15',N'1503',N'MANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150401',N'15',N'1504',N'CANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150402',N'15',N'1504',N'ARAHUAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150403',N'15',N'1504',N'HUAMANTANGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150404',N'15',N'1504',N'HUAROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150405',N'15',N'1504',N'LACHAQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150406',N'15',N'1504',N'SAN BUENAVENTURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150407',N'15',N'1504',N'SANTA ROSA DE QUIVES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150501',N'15',N'1505',N'SAN VICENTE DE CAÑETE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150502',N'15',N'1505',N'ASIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150503',N'15',N'1505',N'CALANGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150504',N'15',N'1505',N'CERRO AZUL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150505',N'15',N'1505',N'CHILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150506',N'15',N'1505',N'COAYLLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150507',N'15',N'1505',N'IMPERIAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150508',N'15',N'1505',N'LUNAHUANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150509',N'15',N'1505',N'MALA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150510',N'15',N'1505',N'NUEVO IMPERIAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150511',N'15',N'1505',N'PACARAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150512') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150512',N'15',N'1505',N'QUILMANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150513') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150513',N'15',N'1505',N'SAN ANTONIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150514') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150514',N'15',N'1505',N'SAN LUIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150515') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150515',N'15',N'1505',N'SANTA CRUZ DE FLORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150516') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150516',N'15',N'1505',N'ZUÑIGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150601',N'15',N'1506',N'HUARAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150602',N'15',N'1506',N'ATAVILLOS ALTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150603',N'15',N'1506',N'ATAVILLOS BAJO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150604',N'15',N'1506',N'AUCALLAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150605',N'15',N'1506',N'CHANCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150606',N'15',N'1506',N'IHUARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150607',N'15',N'1506',N'LAMPIAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150608',N'15',N'1506',N'PACARAOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150609') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150609',N'15',N'1506',N'SAN MIGUEL DE ACOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150610') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150610',N'15',N'1506',N'SANTA CRUZ DE ANDAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150611') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150611',N'15',N'1506',N'SUMBILCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150612') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150612',N'15',N'1506',N'VEINTISIETE DE NOVIEMBRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150701',N'15',N'1507',N'MATUCANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150702',N'15',N'1507',N'ANTIOQUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150703',N'15',N'1507',N'CALLAHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150704',N'15',N'1507',N'CARAMPOMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150705',N'15',N'1507',N'CHICLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150706',N'15',N'1507',N'CUENCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150707',N'15',N'1507',N'HUACHUPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150708',N'15',N'1507',N'HUANZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150709',N'15',N'1507',N'HUAROCHIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150710') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150710',N'15',N'1507',N'LAHUAYTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150711') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150711',N'15',N'1507',N'LANGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150712') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150712',N'15',N'1507',N'LARAOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150713') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150713',N'15',N'1507',N'MARIATANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150714') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150714',N'15',N'1507',N'RICARDO PALMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150715') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150715',N'15',N'1507',N'SAN ANDRES DE TUPICOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150716') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150716',N'15',N'1507',N'SAN ANTONIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150717') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150717',N'15',N'1507',N'SAN BARTOLOME',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150718') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150718',N'15',N'1507',N'SAN DAMIAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150719') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150719',N'15',N'1507',N'SAN JUAN DE IRIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150720') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150720',N'15',N'1507',N'SAN JUAN DE TANTARANCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150721') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150721',N'15',N'1507',N'SAN LORENZO DE QUINTI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150722') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150722',N'15',N'1507',N'SAN MATEO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150723') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150723',N'15',N'1507',N'SAN MATEO DE OTAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150724') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150724',N'15',N'1507',N'SAN PEDRO DE CASTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150725') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150725',N'15',N'1507',N'SAN PEDRO DE HUANCAYRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150726') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150726',N'15',N'1507',N'SANGALLAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150727') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150727',N'15',N'1507',N'SANTA CRUZ DE COCACHACRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150728') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150728',N'15',N'1507',N'SANTA EULALIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150729') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150729',N'15',N'1507',N'SANTIAGO DE ANCHUCAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150730') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150730',N'15',N'1507',N'SANTIAGO DE TUNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150731') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150731',N'15',N'1507',N'SANTO DOMINGO DE LOS OLLEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150732') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150732',N'15',N'1507',N'SURCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150801',N'15',N'1508',N'HUACHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150802',N'15',N'1508',N'AMBAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150803',N'15',N'1508',N'CALETA DE CARQUIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150804',N'15',N'1508',N'CHECRAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150805',N'15',N'1508',N'HUALMAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150806',N'15',N'1508',N'HUAURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150807',N'15',N'1508',N'LEONCIO PRADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150808',N'15',N'1508',N'PACCHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150809',N'15',N'1508',N'SANTA LEONOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150810') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150810',N'15',N'1508',N'SANTA MARIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150811') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150811',N'15',N'1508',N'SAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150812') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150812',N'15',N'1508',N'VEGUETA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150901',N'15',N'1509',N'OYON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150902',N'15',N'1509',N'ANDAJES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150903',N'15',N'1509',N'CAUJUL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150904',N'15',N'1509',N'COCHAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150905',N'15',N'1509',N'NAVAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'150906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'150906',N'15',N'1509',N'PACHANGARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151001',N'15',N'1510',N'YAUYOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151002',N'15',N'1510',N'ALÃS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151003',N'15',N'1510',N'AYAUCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151004',N'15',N'1510',N'AYAVIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151005',N'15',N'1510',N'AZANGARO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151006',N'15',N'1510',N'CACRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151007') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151007',N'15',N'1510',N'CARANIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151008') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151008',N'15',N'1510',N'CATAHUASI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151009') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151009',N'15',N'1510',N'CHOCOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151010') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151010',N'15',N'1510',N'COCHAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151011') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151011',N'15',N'1510',N'COLONIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151012') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151012',N'15',N'1510',N'HONGOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151013') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151013',N'15',N'1510',N'HUAMPARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151014') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151014',N'15',N'1510',N'HUANCAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151015') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151015',N'15',N'1510',N'HUANGASCAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151016') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151016',N'15',N'1510',N'HUANTAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151017') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151017',N'15',N'1510',N'HUAÑEC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151018') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151018',N'15',N'1510',N'LARAOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151019') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151019',N'15',N'1510',N'LINCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151020') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151020',N'15',N'1510',N'MADEAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151021') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151021',N'15',N'1510',N'MIRAFLORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151022') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151022',N'15',N'1510',N'OMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151023') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151023',N'15',N'1510',N'PUTINZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151024') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151024',N'15',N'1510',N'QUINCHES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151025') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151025',N'15',N'1510',N'QUINOCAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151026') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151026',N'15',N'1510',N'SAN JOAQUIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151027') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151027',N'15',N'1510',N'SAN PEDRO DE PILAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151028') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151028',N'15',N'1510',N'TANTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151029') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151029',N'15',N'1510',N'TAURIPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151030') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151030',N'15',N'1510',N'TOMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151031') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151031',N'15',N'1510',N'TUPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151032') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151032',N'15',N'1510',N'VIÑAC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'151033') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'151033',N'15',N'1510',N'VITIS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160101',N'16',N'1601',N'IQUITOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160102',N'16',N'1601',N'ALTO NANAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160103',N'16',N'1601',N'FERNANDO LORES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160104',N'16',N'1601',N'INDIANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160105',N'16',N'1601',N'LAS AMAZONAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160106',N'16',N'1601',N'MAZAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160107',N'16',N'1601',N'NAPO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160108',N'16',N'1601',N'PUNCHANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160110',N'16',N'1601',N'TORRES CAUSANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160112',N'16',N'1601',N'BELEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160113',N'16',N'1601',N'SAN JUAN BAUTISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160201',N'16',N'1602',N'YURIMAGUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160202',N'16',N'1602',N'BALSAPUERTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160205',N'16',N'1602',N'JEBEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160206',N'16',N'1602',N'LAGUNAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160210',N'16',N'1602',N'SANTA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160211',N'16',N'1602',N'TENIENTE CESAR LOPEZ ROJAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160301',N'16',N'1603',N'NAUTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160302',N'16',N'1603',N'PARINARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160303',N'16',N'1603',N'TIGRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160304',N'16',N'1603',N'TROMPETEROS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160305',N'16',N'1603',N'URARINAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160401',N'16',N'1604',N'RAMON CASTILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160402',N'16',N'1604',N'PEBAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160403',N'16',N'1604',N'YAVARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160404',N'16',N'1604',N'SAN PABLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160501',N'16',N'1605',N'REQUENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160502',N'16',N'1605',N'ALTO TAPICHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160503',N'16',N'1605',N'CAPELO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160504',N'16',N'1605',N'EMILIO SAN MARTIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160505',N'16',N'1605',N'MAQUIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160506',N'16',N'1605',N'PUINAHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160507',N'16',N'1605',N'SAQUENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160508',N'16',N'1605',N'SOPLIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160509',N'16',N'1605',N'TAPICHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160510',N'16',N'1605',N'JENARO HERRERA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160511',N'16',N'1605',N'YAQUERANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160601',N'16',N'1606',N'CONTAMANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160602',N'16',N'1606',N'INAHUAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160603',N'16',N'1606',N'PADRE MARQUEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160604',N'16',N'1606',N'PAMPA HERMOSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160605',N'16',N'1606',N'SARAYACU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160606',N'16',N'1606',N'VARGAS GUERRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160701',N'16',N'1607',N'BARRANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160702',N'16',N'1607',N'CAHUAPANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160703',N'16',N'1607',N'MANSERICHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160704',N'16',N'1607',N'MORONA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160705',N'16',N'1607',N'PASTAZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160706',N'16',N'1607',N'ANDOAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160801',N'16',N'1608',N'PUTUMAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160802',N'16',N'1608',N'ROSA PANDURO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160803',N'16',N'1608',N'TENIENTE MANUEL CLAVERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'160804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'160804',N'16',N'1608',N'YAGUAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170101',N'17',N'1701',N'TAMBOPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170102',N'17',N'1701',N'INAMBARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170103',N'17',N'1701',N'LAS PIEDRAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170104',N'17',N'1701',N'LABERINTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170201',N'17',N'1702',N'MANU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170202',N'17',N'1702',N'FITZCARRALD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170203',N'17',N'1702',N'MADRE DE DIOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170204',N'17',N'1702',N'HUEPETUHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170301',N'17',N'1703',N'IÑAPARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170302',N'17',N'1703',N'IBERIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'170303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'170303',N'17',N'1703',N'TAHUAMANU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180101',N'18',N'1801',N'MOQUEGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180102',N'18',N'1801',N'CARUMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180103',N'18',N'1801',N'CUCHUMBAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180104',N'18',N'1801',N'SAMEGUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180105',N'18',N'1801',N'SAN CRISTOBAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180106',N'18',N'1801',N'TORATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180107',N'18',N'1801',N'SAN ANTONIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180201',N'18',N'1802',N'OMATE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180202',N'18',N'1802',N'CHOJATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180203',N'18',N'1802',N'COALAQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180204',N'18',N'1802',N'ICHUÑA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180205',N'18',N'1802',N'LA CAPILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180206',N'18',N'1802',N'LLOQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180207',N'18',N'1802',N'MATALAQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180208',N'18',N'1802',N'PUQUINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180209',N'18',N'1802',N'QUINISTAQUILLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180210',N'18',N'1802',N'UBINAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180211',N'18',N'1802',N'YUNGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180301',N'18',N'1803',N'ILO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180302',N'18',N'1803',N'EL ALGARROBAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'180303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'180303',N'18',N'1803',N'PACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190101',N'19',N'1901',N'CHAUPIMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190102',N'19',N'1901',N'HUACHON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190103',N'19',N'1901',N'HUARIACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190104',N'19',N'1901',N'HUAYLLAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190105',N'19',N'1901',N'NINACACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190106',N'19',N'1901',N'PALLANCHACRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190107',N'19',N'1901',N'PAUCARTAMBO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190108',N'19',N'1901',N'SAN FRANCISCO DE ASIS DE YARUSYACAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190109',N'19',N'1901',N'SIMON BOLIVAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190110',N'19',N'1901',N'TICLACAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190111',N'19',N'1901',N'TINYAHUARCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190112',N'19',N'1901',N'VICCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190113',N'19',N'1901',N'YANACANCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190201',N'19',N'1902',N'YANAHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190202',N'19',N'1902',N'CHACAYAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190203',N'19',N'1902',N'GOYLLARISQUIZGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190204',N'19',N'1902',N'PAUCAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190205',N'19',N'1902',N'SAN PEDRO DE PILLAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190206',N'19',N'1902',N'SANTA ANA DE TUSI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190207',N'19',N'1902',N'TAPUC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190208',N'19',N'1902',N'VILCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190301',N'19',N'1903',N'OXAPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190302',N'19',N'1903',N'CHONTABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190303',N'19',N'1903',N'HUANCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190304',N'19',N'1903',N'PALCAZU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190305',N'19',N'1903',N'POZUZO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190306',N'19',N'1903',N'PUERTO BERMUDEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190307',N'19',N'1903',N'VILLA RICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'190308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'190308',N'19',N'1903',N'CONSTITUCION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200101',N'20',N'2001',N'PIURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200104',N'20',N'2001',N'CASTILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200105',N'20',N'2001',N'CATACAOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200107',N'20',N'2001',N'CURA MORI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200108',N'20',N'2001',N'EL TALLAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200109',N'20',N'2001',N'LA ARENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200110',N'20',N'2001',N'LA UNION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200111',N'20',N'2001',N'LAS LOMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200114',N'20',N'2001',N'TAMBO GRANDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200115',N'20',N'2001',N'VEINTISEIS DE OCTUBRE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200201',N'20',N'2002',N'AYABACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200202',N'20',N'2002',N'FRIAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200203',N'20',N'2002',N'JILILI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200204',N'20',N'2002',N'LAGUNAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200205',N'20',N'2002',N'MONTERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200206',N'20',N'2002',N'PACAIPAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200207',N'20',N'2002',N'PAIMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200208',N'20',N'2002',N'SAPILLICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200209',N'20',N'2002',N'SICCHEZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200210',N'20',N'2002',N'SUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200301',N'20',N'2003',N'HUANCABAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200302',N'20',N'2003',N'CANCHAQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200303',N'20',N'2003',N'EL CARMEN DE LA FRONTERA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200304',N'20',N'2003',N'HUARMACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200305',N'20',N'2003',N'LALAQUIZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200306',N'20',N'2003',N'SAN MIGUEL DE EL FAIQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200307',N'20',N'2003',N'SONDOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200308',N'20',N'2003',N'SONDORILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200401',N'20',N'2004',N'CHULUCANAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200402',N'20',N'2004',N'BUENOS AIRES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200403',N'20',N'2004',N'CHALACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200404',N'20',N'2004',N'LA MATANZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200405',N'20',N'2004',N'MORROPON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200406',N'20',N'2004',N'SALITRAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200407',N'20',N'2004',N'SAN JUAN DE BIGOTE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200408',N'20',N'2004',N'SANTA CATALINA DE MOSSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200409') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200409',N'20',N'2004',N'SANTO DOMINGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200410') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200410',N'20',N'2004',N'YAMANGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200501',N'20',N'2005',N'PAITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200502',N'20',N'2005',N'AMOTAPE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200503',N'20',N'2005',N'ARENAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200504',N'20',N'2005',N'COLAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200505',N'20',N'2005',N'LA HUACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200506',N'20',N'2005',N'TAMARINDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200507',N'20',N'2005',N'VICHAYAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200601',N'20',N'2006',N'SULLANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200602',N'20',N'2006',N'BELLAVISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200603',N'20',N'2006',N'IGNACIO ESCUDERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200604',N'20',N'2006',N'LANCONES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200605',N'20',N'2006',N'MARCAVELICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200606',N'20',N'2006',N'MIGUEL CHECA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200607',N'20',N'2006',N'QUERECOTILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200608',N'20',N'2006',N'SALITRAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200701',N'20',N'2007',N'PARIÑAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200702',N'20',N'2007',N'EL ALTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200703',N'20',N'2007',N'LA BREA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200704',N'20',N'2007',N'LOBITOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200705',N'20',N'2007',N'LOS ORGANOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200706',N'20',N'2007',N'MANCORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200801',N'20',N'2008',N'SECHURA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200802',N'20',N'2008',N'BELLAVISTA DE LA UNION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200803',N'20',N'2008',N'BERNAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200804',N'20',N'2008',N'CRISTO NOS VALGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200805',N'20',N'2008',N'VICE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'200806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'200806',N'20',N'2008',N'RINCONADA LLICUAR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210101',N'21',N'2101',N'PUNO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210102',N'21',N'2101',N'ACORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210103',N'21',N'2101',N'AMANTANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210104',N'21',N'2101',N'ATUNCOLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210105',N'21',N'2101',N'CAPACHICA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210106',N'21',N'2101',N'CHUCUITO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210107',N'21',N'2101',N'COATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210108',N'21',N'2101',N'HUATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210109',N'21',N'2101',N'MAÑAZO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210110',N'21',N'2101',N'PAUCARCOLLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210111',N'21',N'2101',N'PICHACANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210112') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210112',N'21',N'2101',N'PLATERIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210113') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210113',N'21',N'2101',N'SAN ANTONIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210114') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210114',N'21',N'2101',N'TIQUILLACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210115') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210115',N'21',N'2101',N'VILQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210201',N'21',N'2102',N'AZANGARO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210202',N'21',N'2102',N'ACHAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210203',N'21',N'2102',N'ARAPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210204',N'21',N'2102',N'ASILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210205',N'21',N'2102',N'CAMINACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210206',N'21',N'2102',N'CHUPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210207',N'21',N'2102',N'JOSE DOMINGO CHOQUEHUANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210208',N'21',N'2102',N'MUÑANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210209',N'21',N'2102',N'POTONI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210210',N'21',N'2102',N'SAMAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210211') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210211',N'21',N'2102',N'SAN ANTON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210212') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210212',N'21',N'2102',N'SAN JOSE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210213') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210213',N'21',N'2102',N'SAN JUAN DE SALINAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210214') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210214',N'21',N'2102',N'SANTIAGO DE PUPUJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210215') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210215',N'21',N'2102',N'TIRAPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210301',N'21',N'2103',N'MACUSANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210302',N'21',N'2103',N'AJOYANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210303',N'21',N'2103',N'AYAPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210304',N'21',N'2103',N'COASA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210305',N'21',N'2103',N'CORANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210306',N'21',N'2103',N'CRUCERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210307',N'21',N'2103',N'ITUATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210308') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210308',N'21',N'2103',N'OLLACHEA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210309') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210309',N'21',N'2103',N'SAN GABAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210310') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210310',N'21',N'2103',N'USICAYOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210401',N'21',N'2104',N'JULI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210402',N'21',N'2104',N'DESAGUADERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210403',N'21',N'2104',N'HUACULLANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210404',N'21',N'2104',N'KELLUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210405',N'21',N'2104',N'PISACOMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210406',N'21',N'2104',N'POMATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210407',N'21',N'2104',N'ZEPITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210501',N'21',N'2105',N'ILAVE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210502',N'21',N'2105',N'CAPAZO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210503',N'21',N'2105',N'PILCUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210504',N'21',N'2105',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210505',N'21',N'2105',N'CONDURIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210601',N'21',N'2106',N'HUANCANE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210602',N'21',N'2106',N'COJATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210603',N'21',N'2106',N'HUATASANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210604',N'21',N'2106',N'INCHUPALLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210605',N'21',N'2106',N'PUSI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210606') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210606',N'21',N'2106',N'ROSASPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210607') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210607',N'21',N'2106',N'TARACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210608') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210608',N'21',N'2106',N'VILQUE CHICO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210701',N'21',N'2107',N'LAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210702',N'21',N'2107',N'CABANILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210703',N'21',N'2107',N'CALAPUJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210704',N'21',N'2107',N'NICASIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210705',N'21',N'2107',N'OCUVIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210706',N'21',N'2107',N'PALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210707',N'21',N'2107',N'PARATIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210708',N'21',N'2107',N'PUCARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210709',N'21',N'2107',N'SANTA LUCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210710') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210710',N'21',N'2107',N'VILAVILA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210801',N'21',N'2108',N'AYAVIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210802',N'21',N'2108',N'ANTAUTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210803',N'21',N'2108',N'CUPI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210804',N'21',N'2108',N'LLALLI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210805',N'21',N'2108',N'MACARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210806',N'21',N'2108',N'NUÑOA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210807',N'21',N'2108',N'ORURILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210808',N'21',N'2108',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210809',N'21',N'2108',N'UMACHIRI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210901',N'21',N'2109',N'MOHO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210902',N'21',N'2109',N'CONIMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210903',N'21',N'2109',N'HUAYRAPATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'210904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'210904',N'21',N'2109',N'TILALI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211001',N'21',N'2110',N'PUTINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211002',N'21',N'2110',N'ANANEA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211003',N'21',N'2110',N'PEDRO VILCA APAZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211004',N'21',N'2110',N'QUILCAPUNCU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211005',N'21',N'2110',N'SINA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211101',N'21',N'2111',N'JULIACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211102',N'21',N'2111',N'CABANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211103',N'21',N'2111',N'CABANILLAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211104',N'21',N'2111',N'CARACOTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211105',N'21',N'2111',N'SAN MIGUEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211201',N'21',N'2112',N'SANDIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211202',N'21',N'2112',N'CUYOCUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211203',N'21',N'2112',N'LIMBANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211204',N'21',N'2112',N'PATAMBUCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211205',N'21',N'2112',N'PHARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211206',N'21',N'2112',N'QUIACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211207') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211207',N'21',N'2112',N'SAN JUAN DEL ORO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211208') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211208',N'21',N'2112',N'YANAHUAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211209') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211209',N'21',N'2112',N'ALTO INAMBARI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211210') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211210',N'21',N'2112',N'SAN PEDRO DE PUTINA PUNCO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211301',N'21',N'2113',N'YUNGUYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211302',N'21',N'2113',N'ANAPIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211303',N'21',N'2113',N'COPANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211304',N'21',N'2113',N'CUTURAPI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211305',N'21',N'2113',N'OLLARAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211306',N'21',N'2113',N'TINICACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'211307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'211307',N'21',N'2113',N'UNICACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220101',N'22',N'2201',N'MOYOBAMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220102',N'22',N'2201',N'CALZADA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220103',N'22',N'2201',N'HABANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220104',N'22',N'2201',N'JEPELACIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220105',N'22',N'2201',N'SORITOR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220106',N'22',N'2201',N'YANTALO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220201',N'22',N'2202',N'BELLAVISTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220202',N'22',N'2202',N'ALTO BIAVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220203',N'22',N'2202',N'BAJO BIAVO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220204',N'22',N'2202',N'HUALLAGA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220205',N'22',N'2202',N'SAN PABLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220206',N'22',N'2202',N'SAN RAFAEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220301',N'22',N'2203',N'SAN JOSE DE SISA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220302',N'22',N'2203',N'AGUA BLANCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220303',N'22',N'2203',N'SAN MARTIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220304',N'22',N'2203',N'SANTA ROSA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220305',N'22',N'2203',N'SHATOJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220401',N'22',N'2204',N'SAPOSOA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220402',N'22',N'2204',N'ALTO SAPOSOA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220403',N'22',N'2204',N'EL ESLABON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220404',N'22',N'2204',N'PISCOYACU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220405',N'22',N'2204',N'SACANCHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220406',N'22',N'2204',N'TINGO DE SAPOSOA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220501') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220501',N'22',N'2205',N'LAMAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220502') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220502',N'22',N'2205',N'ALONSO DE ALVARADO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220503') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220503',N'22',N'2205',N'BARRANQUITA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220504') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220504',N'22',N'2205',N'CAYNARACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220505') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220505',N'22',N'2205',N'CUÑUMBUQUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220506') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220506',N'22',N'2205',N'PINTO RECODO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220507') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220507',N'22',N'2205',N'RUMISAPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220508') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220508',N'22',N'2205',N'SAN ROQUE DE CUMBAZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220509') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220509',N'22',N'2205',N'SHANAO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220510') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220510',N'22',N'2205',N'TABALOSOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220511') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220511',N'22',N'2205',N'ZAPATERO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220601') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220601',N'22',N'2206',N'JUANJUI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220602') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220602',N'22',N'2206',N'CAMPANILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220603') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220603',N'22',N'2206',N'HUICUNGO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220604') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220604',N'22',N'2206',N'PACHIZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220605') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220605',N'22',N'2206',N'PAJARILLO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220701') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220701',N'22',N'2207',N'PICOTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220702') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220702',N'22',N'2207',N'BUENOS AIRES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220703') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220703',N'22',N'2207',N'CASPISAPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220704') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220704',N'22',N'2207',N'PILLUANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220705') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220705',N'22',N'2207',N'PUCACACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220706') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220706',N'22',N'2207',N'SAN CRISTOBAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220707') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220707',N'22',N'2207',N'SAN HILARION',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220708') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220708',N'22',N'2207',N'SHAMBOYACU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220709') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220709',N'22',N'2207',N'TINGO DE PONASA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220710') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220710',N'22',N'2207',N'TRES UNIDOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220801') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220801',N'22',N'2208',N'RIOJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220802') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220802',N'22',N'2208',N'AWAJUN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220803') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220803',N'22',N'2208',N'ELIAS SOPLIN VARGAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220804') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220804',N'22',N'2208',N'NUEVA CAJAMARCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220805') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220805',N'22',N'2208',N'PARDO MIGUEL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220806') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220806',N'22',N'2208',N'POSIC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220807') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220807',N'22',N'2208',N'SAN FERNANDO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220808') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220808',N'22',N'2208',N'YORONGOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220809') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220809',N'22',N'2208',N'YURACYACU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220901') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220901',N'22',N'2209',N'TARAPOTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220902') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220902',N'22',N'2209',N'ALBERTO LEVEAU',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220903') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220903',N'22',N'2209',N'CACATACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220904') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220904',N'22',N'2209',N'CHAZUTA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220905') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220905',N'22',N'2209',N'CHIPURANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220906') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220906',N'22',N'2209',N'EL PORVENIR',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220907') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220907',N'22',N'2209',N'HUIMBAYOC',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220908') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220908',N'22',N'2209',N'JUAN GUERRA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220909') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220909',N'22',N'2209',N'LA BANDA DE SHILCAYO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220910') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220910',N'22',N'2209',N'MORALES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220911') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220911',N'22',N'2209',N'PAPAPLAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220912') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220912',N'22',N'2209',N'SAN ANTONIO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220913') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220913',N'22',N'2209',N'SAUCE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'220914') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'220914',N'22',N'2209',N'SHAPAJA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'221001') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'221001',N'22',N'2210',N'TOCACHE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'221002') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'221002',N'22',N'2210',N'NUEVO PROGRESO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'221003') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'221003',N'22',N'2210',N'POLVORA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'221004') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'221004',N'22',N'2210',N'SHUNTE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'221005') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'221005',N'22',N'2210',N'UCHIZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'221006') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'221006',N'22',N'2210',N'SANTA LUCIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230101',N'23',N'2301',N'TACNA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230102',N'23',N'2301',N'ALTO DE LA ALIANZA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230103',N'23',N'2301',N'CALANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230104',N'23',N'2301',N'CIUDAD NUEVA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230105',N'23',N'2301',N'INCLAN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230106',N'23',N'2301',N'PACHIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230107',N'23',N'2301',N'PALCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230108') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230108',N'23',N'2301',N'POCOLLAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230109') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230109',N'23',N'2301',N'SAMA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230110') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230110',N'23',N'2301',N'CORONEL GREGORIO ALBARRACIN LANCHIPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230111') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230111',N'23',N'2301',N'LA YARADA LOS PALOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230201',N'23',N'2302',N'CANDARAVE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230202',N'23',N'2302',N'CAIRANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230203',N'23',N'2302',N'CAMILACA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230204',N'23',N'2302',N'CURIBAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230205') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230205',N'23',N'2302',N'HUANUARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230206') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230206',N'23',N'2302',N'QUILAHUANI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230301',N'23',N'2303',N'LOCUMBA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230302',N'23',N'2303',N'ILABAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230303',N'23',N'2303',N'ITE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230401',N'23',N'2304',N'TARATA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230402') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230402',N'23',N'2304',N'HEROES ALBARRACIN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230403') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230403',N'23',N'2304',N'ESTIQUE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230404') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230404',N'23',N'2304',N'ESTIQUE-PAMPA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230405') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230405',N'23',N'2304',N'SITAJARA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230406') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230406',N'23',N'2304',N'SUSAPAYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230407') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230407',N'23',N'2304',N'TARUCACHI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'230408') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'230408',N'23',N'2304',N'TICACO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240101',N'24',N'2401',N'TUMBES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240102',N'24',N'2401',N'CORRALES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240103',N'24',N'2401',N'LA CRUZ',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240104',N'24',N'2401',N'PAMPAS DE HOSPITAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240105',N'24',N'2401',N'SAN JACINTO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240106',N'24',N'2401',N'SAN JUAN DE LA VIRGEN',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240201',N'24',N'2402',N'ZORRITOS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240202',N'24',N'2402',N'CASITAS',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240203',N'24',N'2402',N'CANOAS DE PUNTA SAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240301',N'24',N'2403',N'ZARUMILLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240302',N'24',N'2403',N'AGUAS VERDES',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240303',N'24',N'2403',N'MATAPALO',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'240304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'240304',N'24',N'2403',N'PAPAYAL',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250101') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250101',N'25',N'2501',N'CALLERIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250102') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250102',N'25',N'2501',N'CAMPOVERDE',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250103') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250103',N'25',N'2501',N'IPARIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250104') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250104',N'25',N'2501',N'MASISEA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250105') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250105',N'25',N'2501',N'YARINACOCHA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250106') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250106',N'25',N'2501',N'NUEVA REQUENA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250107') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250107',N'25',N'2501',N'MANANTAY',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250201') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250201',N'25',N'2502',N'RAIMONDI',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250202') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250202',N'25',N'2502',N'SEPAHUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250203') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250203',N'25',N'2502',N'TAHUANIA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250204') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250204',N'25',N'2502',N'YURUA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250301') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250301',N'25',N'2503',N'PADRE ABAD',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250302') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250302',N'25',N'2503',N'IRAZOLA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250303') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250303',N'25',N'2503',N'CURIMANA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250304') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250304',N'25',N'2503',N'NESHUYA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250305') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250305',N'25',N'2503',N'ALEXANDER VON HUMBOLDT',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250306') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250306',N'25',N'2503',N'HUIPOCA',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250307') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250307',N'25',N'2503',N'BOQUERON',1);
IF NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS destino WHERE destino.CodigoUbigeo=N'250401') INSERT INTO dbo.UbigeoDistritos (CodigoUbigeo,CodigoDepartamento,CodigoProvincia,Nombre,Activo) VALUES (N'250401',N'25',N'2504',N'PURUS',1);
GO


-- Fuente: Script/025_DireccionesYConsideraciones.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.Cliente',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.Cliente ADD CodigoUbigeo CHAR(6) NULL CONSTRAINT FK_Cliente_Ubigeo REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
IF COL_LENGTH(N'dbo.Tienda',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.Tienda ADD CodigoUbigeo CHAR(6) NULL CONSTRAINT FK_Tienda_Ubigeo REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'ConsideracionesPresupuesto') IS NULL ALTER TABLE dbo.ConfiguracionHYT ADD ConsideracionesPresupuesto NVARCHAR(MAX) NULL;
ALTER TABLE dbo.PresupuestoRevision ALTER COLUMN ClienteDireccion NVARCHAR(600) NULL;
ALTER TABLE dbo.PresupuestoRevision ALTER COLUMN TiendaDireccion NVARCHAR(600) NULL;
GO


-- Fuente: Script/028_UbigeoConfiguracionHYT.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.ConfiguracionHYT ADD CodigoUbigeo CHAR(6) NULL;
GO
IF NOT EXISTS(SELECT 1 FROM sys.foreign_keys AS fk WHERE fk.name=N'FK_ConfiguracionHYT_Ubigeo' AND fk.parent_object_id=OBJECT_ID(N'dbo.ConfiguracionHYT')) ALTER TABLE dbo.ConfiguracionHYT ADD CONSTRAINT FK_ConfiguracionHYT_Ubigeo FOREIGN KEY(CodigoUbigeo) REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
GO


-- Fuente: Script/011_EstadosMantenimientos.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ProveedorCliente', N'Activo') IS NULL
    ALTER TABLE dbo.ProveedorCliente ADD Activo BIT NOT NULL CONSTRAINT DF_ProveedorCliente_Activo DEFAULT (1) WITH VALUES;
IF COL_LENGTH(N'dbo.TarifarioPartida', N'Activo') IS NULL
    ALTER TABLE dbo.TarifarioPartida ADD Activo BIT NOT NULL CONSTRAINT DF_TarifarioPartida_Activo DEFAULT (1) WITH VALUES;
GO


-- Fuente: Script/020_CodigosInternosAutomaticos.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF NOT EXISTS(SELECT 1 FROM sys.sequences WHERE object_id=OBJECT_ID(N'dbo.SecuenciaCodigoInterno'))
    EXEC(N'CREATE SEQUENCE dbo.SecuenciaCodigoInterno AS BIGINT START WITH 1 INCREMENT BY 1 NO CYCLE;');
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.ServicioMaterial') AND name=N'UX_ServicioMaterial_CodigoInterno')
    CREATE UNIQUE INDEX UX_ServicioMaterial_CodigoInterno ON dbo.ServicioMaterial(CodigoInterno);
GO


-- Fuente: Script/021_CodigoSAPCliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
BEGIN TRY
    IF COL_LENGTH(N'dbo.Cliente',N'CodigoSAP') IS NULL
    BEGIN
        BEGIN TRANSACTION;
        EXEC(N'ALTER TABLE dbo.Cliente ADD CodigoSAP NVARCHAR(60) NULL;');
        -- La migración se ejecuta únicamente al crear la columna. No restaura códigos borrados posteriormente.
        -- Prioriza el registro activo y después el último registrado; conserva todas las fuentes históricas.
        EXEC(N'UPDATE cliente SET CodigoSAP=fuente.CodigoSAPProveedor
            FROM dbo.Cliente AS cliente
            CROSS APPLY (SELECT TOP (1) proveedor.CodigoSAPProveedor
                FROM dbo.ProveedorCliente AS proveedor
                WHERE proveedor.ClienteId=cliente.ClienteId AND NULLIF(LTRIM(RTRIM(proveedor.CodigoSAPProveedor)),N'''') IS NOT NULL
                ORDER BY proveedor.Activo DESC,proveedor.FechaRegistro DESC,proveedor.Id DESC) AS fuente;');
        COMMIT TRANSACTION;
    END;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
    SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
    RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
END CATCH;
GO


-- Fuente: Tablas/ClienteEnlaceDocumento.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ClienteEnlaceDocumento',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ClienteEnlaceDocumento
    (
        Id BIGINT IDENTITY NOT NULL CONSTRAINT PK_ClienteEnlaceDocumento PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        Descripcion NVARCHAR(200) NOT NULL,
        Url NVARCHAR(2048) NOT NULL,
        Orden INT NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ClienteEnlaceDocumento_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        CONSTRAINT FK_ClienteEnlaceDocumento_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT FK_ClienteEnlaceDocumento_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ClienteEnlaceDocumento_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
    CREATE INDEX IX_ClienteEnlaceDocumento_Cliente ON dbo.ClienteEnlaceDocumento(ClienteId,Orden,Id);
END;
GO


-- Fuente: Script/022_EnlacesDocumentosCliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS (SELECT 1 FROM sys.extended_properties AS propiedad WHERE propiedad.major_id=OBJECT_ID(N'dbo.ClienteEnlaceDocumento') AND propiedad.minor_id=0 AND propiedad.name=N'HYT_EnlaceAnteriorMigrado')
    BEGIN
        INSERT INTO dbo.ClienteEnlaceDocumento(ClienteId,Descripcion,Url,Orden,FechaRegistro,UsuarioRegistro)
        SELECT cliente.ClienteId,N'Documentos del cliente',cliente.UrlCargaDocumentos,0,cliente.FechaRegistro,cliente.UsuarioRegistro
        FROM dbo.Cliente AS cliente
        WHERE NULLIF(LTRIM(RTRIM(cliente.UrlCargaDocumentos)),N'') IS NOT NULL
          AND NOT EXISTS(SELECT 1 FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=cliente.ClienteId AND enlace.Url=cliente.UrlCargaDocumentos);
        EXEC sys.sp_addextendedproperty @name=N'HYT_EnlaceAnteriorMigrado',@value=1,@level0type=N'SCHEMA',@level0name=N'dbo',@level1type=N'TABLE',@level1name=N'ClienteEnlaceDocumento';
    END;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
    SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
    RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
END CATCH;
GO


-- Fuente: StoreProcedure/spu_MantCatalogos.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCatalogos
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1)
            THROW 51200, N'[HYT_PERMISO] El usuario no está habilitado.', 1;
        SELECT moneda.Codigo, CONCAT(moneda.Nombre, N' (', moneda.Codigo, N') ', moneda.Simbolo) AS Nombre
        FROM dbo.Moneda AS moneda ORDER BY CASE moneda.Codigo WHEN N'PEN' THEN 0 ELSE 1 END, moneda.Codigo;
        SELECT tipo.Codigo, tipo.Nombre FROM dbo.TipoIGV AS tipo ORDER BY tipo.Codigo;
        SELECT d.CodigoDepartamento,d.Nombre FROM dbo.UbigeoDepartamentos AS d WHERE d.Activo=1 ORDER BY d.Nombre;
        SELECT p.CodigoProvincia,p.Nombre FROM dbo.UbigeoProvincias AS p WHERE p.Activo=1 ORDER BY p.Nombre;
        SELECT d.CodigoUbigeo,d.Nombre FROM dbo.UbigeoDistritos AS d WHERE d.Activo=1 ORDER BY d.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantCliente_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.ClienteId AS Id, m.Codigo AS Codigo, CASE WHEN m.TipoEntidad=N'NATURAL' THEN CONCAT(m.Nombres, N' ', m.Apellidos) ELSE m.RazonSocial END AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            CONCAT_WS(N' / ',NULLIF(m.DireccionFiscal,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=m.CodigoUbigeo)) AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.Cliente AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR CASE WHEN m.TipoEntidad=N'NATURAL' THEN CONCAT(m.Nombres, N' ', m.Apellidos) ELSE m.RazonSocial END LIKE @Patron ESCAPE N'~' OR m.CodigoSAP LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantCliente_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.ClienteId AS Id,
            m.Codigo,
            m.CodigoSAP,
            m.TipoEntidad,
            m.TipoIdentificacion,
            m.NumeroIdentificacion,
            m.RazonSocial,
            m.Nombres,
            m.Apellidos,
            m.DireccionFiscal,
            m.CodigoUbigeo,
            m.UrlCargaDocumentos,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=m.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces
        FROM dbo.Cliente AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.ClienteId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Cliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Cliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantCliente_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoSAP.',1;
        DECLARE @CodigoSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoSAP'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'TipoEntidad' AND LEN(dato.value)>20) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: TipoEntidad.',1;
        DECLARE @TipoEntidad NVARCHAR(20) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.TipoEntidad'))),N'');
        IF @TipoEntidad IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: TipoEntidad.',1;
        DECLARE @TipoIdentificacion NVARCHAR(20);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'NumeroIdentificacion' AND LEN(dato.value)>30) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: NumeroIdentificacion.',1;
        DECLARE @NumeroIdentificacion NVARCHAR(30) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroIdentificacion'))),N'');
        IF @NumeroIdentificacion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: NumeroIdentificacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'RazonSocial' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: RazonSocial.',1;
        DECLARE @RazonSocial NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RazonSocial'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombres' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombres.',1;
        DECLARE @Nombres NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombres'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Apellidos' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Apellidos.',1;
        DECLARE @Apellidos NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Apellidos'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'DireccionFiscal' AND LEN(dato.value)>300) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: DireccionFiscal.',1;
        DECLARE @DireccionFiscal NVARCHAR(300) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.DireccionFiscal'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoUbigeo' AND LEN(dato.value)>6) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoUbigeo.',1;
        DECLARE @CodigoUbigeo NVARCHAR(6) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        DECLARE @UrlCargaDocumentos NVARCHAR(2048);
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.CodigoSAP, anterior.TipoEntidad, anterior.TipoIdentificacion, anterior.NumeroIdentificacion, anterior.RazonSocial, anterior.Nombres, anterior.Apellidos, anterior.DireccionFiscal, anterior.CodigoUbigeo, anterior.UrlCargaDocumentos, anterior.Activo, JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=anterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Cliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.ClienteId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @TipoEntidad NOT IN (N'JURIDICA',N'NATURAL') THROW 51213,N'[HYT_VALIDACION] Tipo de entidad no válido.',1;
        IF (@TipoEntidad=N'JURIDICA' AND @RazonSocial IS NULL) OR (@TipoEntidad=N'NATURAL' AND (@Nombres IS NULL OR @Apellidos IS NULL)) THROW 51214,N'[HYT_VALIDACION] Completa la identificación del cliente.',1;
        SET @TipoIdentificacion=JSON_VALUE(@Anterior,N'$.TipoIdentificacion');
        IF @TipoEntidad=N'JURIDICA' BEGIN SET @Nombres=NULL; SET @Apellidos=NULL; END;
        ELSE SET @RazonSocial=NULL;
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51228,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        SET @UrlCargaDocumentos=JSON_VALUE(@Anterior,N'$.UrlCargaDocumentos');
        DECLARE @Enlaces NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Enlaces');
        IF EXISTS(SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Enlaces') AND (@Enlaces IS NULL OR LEFT(LTRIM(@Enlaces),1)<>N'[')
            THROW 51226,N'[HYT_VALIDACION] La lista de enlaces no es válida.',1;
        DECLARE @ListaEnlaces TABLE(Id BIGINT NULL,Descripcion NVARCHAR(MAX),Url NVARCHAR(MAX),Orden INT);
        IF @Enlaces IS NOT NULL
        BEGIN
            IF (SELECT COUNT_BIG(1) FROM OPENJSON(@Enlaces))>100 OR EXISTS(SELECT 1 FROM OPENJSON(@Enlaces) AS jsonEnlace WHERE jsonEnlace.type<>5)
                THROW 51226,N'[HYT_VALIDACION] Revisa la lista de enlaces (máximo cien).',1;
            INSERT INTO @ListaEnlaces(Id,Descripcion,Url,Orden)
            SELECT datos.Id,NULLIF(LTRIM(RTRIM(datos.Descripcion)),N''),NULLIF(LTRIM(RTRIM(datos.Url)),N''),CONVERT(INT,entrada.[key])
            FROM OPENJSON(@Enlaces) AS entrada
            CROSS APPLY OPENJSON(entrada.value) WITH(Id BIGINT '$.Id',Descripcion NVARCHAR(MAX) '$.Descripcion',Url NVARCHAR(MAX) '$.Url') AS datos;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS enlace WHERE enlace.Descripcion IS NULL OR LEN(enlace.Descripcion)>200 OR enlace.Url IS NULL OR LEN(enlace.Url)>2048
                OR (enlace.Url NOT LIKE N'https://%' AND enlace.Url NOT LIKE N'http://%') OR CHARINDEX(NCHAR(10),enlace.Url)>0 OR CHARINDEX(NCHAR(13),enlace.Url)>0)
                THROW 51226,N'[HYT_VALIDACION] Completa descripción y enlace HTTP o HTTPS válidos.',1;
            IF EXISTS(SELECT enlace.Id FROM @ListaEnlaces AS enlace WHERE enlace.Id IS NOT NULL GROUP BY enlace.Id HAVING COUNT_BIG(1)>1)
                THROW 51226,N'[HYT_VALIDACION] Un enlace está repetido.',1;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NOT NULL AND NOT EXISTS
                (SELECT 1 FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.Id=entrada.Id AND enlace.ClienteId=@Id))
                THROW 51226,N'[HYT_VALIDACION] Un enlace no pertenece al cliente o ya no existe.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Cliente',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'CLI-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Cliente AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Cliente (Codigo, CodigoSAP, TipoEntidad, TipoIdentificacion, NumeroIdentificacion, RazonSocial, Nombres, Apellidos, DireccionFiscal, CodigoUbigeo, UrlCargaDocumentos, Activo, UsuarioRegistro)
            VALUES (@Codigo, @CodigoSAP, @TipoEntidad, @TipoIdentificacion, @NumeroIdentificacion, @RazonSocial, @Nombres, @Apellidos, @DireccionFiscal, @CodigoUbigeo, @UrlCargaDocumentos, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Cliente SET Codigo=@Codigo, CodigoSAP=@CodigoSAP, TipoEntidad=@TipoEntidad, TipoIdentificacion=@TipoIdentificacion, NumeroIdentificacion=@NumeroIdentificacion, RazonSocial=@RazonSocial, Nombres=@Nombres, Apellidos=@Apellidos, DireccionFiscal=@DireccionFiscal, CodigoUbigeo=@CodigoUbigeo, UrlCargaDocumentos=@UrlCargaDocumentos, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ClienteId=@Id;
        END;
        IF @Enlaces IS NOT NULL
        BEGIN
            DELETE enlace FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=@Id
                AND NOT EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id=enlace.Id);
            UPDATE enlace SET Descripcion=entrada.Descripcion,Url=entrada.Url,Orden=entrada.Orden,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
            FROM dbo.ClienteEnlaceDocumento AS enlace INNER JOIN @ListaEnlaces AS entrada ON entrada.Id=enlace.Id WHERE enlace.ClienteId=@Id;
            INSERT INTO dbo.ClienteEnlaceDocumento(ClienteId,Descripcion,Url,Orden,UsuarioRegistro)
            SELECT @Id,entrada.Descripcion,entrada.Url,entrada.Orden,@UsuarioId FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NULL;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.CodigoSAP, posterior.TipoEntidad, posterior.TipoIdentificacion, posterior.NumeroIdentificacion, posterior.RazonSocial, posterior.Nombres, posterior.Apellidos, posterior.DireccionFiscal, posterior.CodigoUbigeo, posterior.UrlCargaDocumentos, posterior.Activo, JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=posterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Cliente AS posterior WHERE posterior.ClienteId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Cliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTienda_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTienda_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.TiendaId AS Id, m.Codigo AS Codigo, m.Nombre AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            CONCAT_WS(N' / ',NULLIF(m.Direccion,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=m.CodigoUbigeo)) AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion, cliente.Codigo AS ClienteCodigo, CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END AS ClienteNombre
        INTO #Pagina FROM dbo.Tienda AS m INNER JOIN dbo.Cliente AS cliente ON cliente.ClienteId=m.ClienteId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Nombre LIKE @Patron ESCAPE N'~' OR cliente.Codigo LIKE @Patron ESCAPE N'~' OR CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1) AND (@PadreId IS NULL OR m.ClienteId = @PadreId);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion, p.ClienteCodigo, p.ClienteNombre
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTienda_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTienda_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.TiendaId AS Id,
            m.ClienteId,
            m.Codigo,
            m.Nombre,
            m.Direccion,
            m.CodigoUbigeo,
            m.CentroSAPFICO,
            m.ZonaId,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            CASE WHEN ref0.TipoEntidad=N'NATURAL' THEN CONCAT(ref0.Nombres, N' ', ref0.Apellidos) ELSE ref0.RazonSocial END AS ClienteIdLabel,
            ref6.Descripcion AS ZonaIdLabel
        FROM dbo.Tienda AS m LEFT JOIN dbo.Cliente AS ref0 ON ref0.ClienteId = m.ClienteId LEFT JOIN dbo.Zona AS ref6 ON ref6.Id = m.ZonaId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.TiendaId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Tienda' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Tienda' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTienda_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTienda_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TIENDAS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF @ClienteId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: ClienteId.',1;
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombre' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombre.',1;
        DECLARE @Nombre NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombre'))),N'');
        IF @Nombre IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Nombre.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Direccion' AND LEN(dato.value)>300) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Direccion.',1;
        DECLARE @Direccion NVARCHAR(300) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Direccion'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoUbigeo' AND LEN(dato.value)>6) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoUbigeo.',1;
        DECLARE @CodigoUbigeo NVARCHAR(6) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CentroSAPFICO' AND LEN(dato.value)>50) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CentroSAPFICO.',1;
        DECLARE @CentroSAPFICO NVARCHAR(50) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CentroSAPFICO'))),N'');
        DECLARE @ZonaId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.ZonaId'));
        IF JSON_VALUE(@Datos,N'$.ZonaId') IS NOT NULL AND @ZonaId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ZonaId.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.ClienteId, anterior.Codigo, anterior.Nombre, anterior.Direccion, anterior.CodigoUbigeo, anterior.CentroSAPFICO, anterior.ZonaId, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Tienda AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.TiendaId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId'))<>@ClienteId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @ZonaId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ZonaId')),-1)<>@ZonaId) AND NOT EXISTS (SELECT 1 FROM dbo.Zona AS destino WITH (HOLDLOCK) WHERE destino.Id=@ZonaId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ZonaId.',1;
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51228,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Tienda',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'TIE-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Tienda AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Tienda (ClienteId, Codigo, Nombre, Direccion, CodigoUbigeo, CentroSAPFICO, ZonaId, Activo, UsuarioRegistro)
            VALUES (@ClienteId, @Codigo, @Nombre, @Direccion, @CodigoUbigeo, @CentroSAPFICO, @ZonaId, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Tienda SET Codigo=@Codigo, Nombre=@Nombre, Direccion=@Direccion, CodigoUbigeo=@CodigoUbigeo, CentroSAPFICO=@CentroSAPFICO, ZonaId=@ZonaId, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE TiendaId=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.ClienteId, posterior.Codigo, posterior.Nombre, posterior.Direccion, posterior.CodigoUbigeo, posterior.CentroSAPFICO, posterior.ZonaId, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Tienda AS posterior WHERE posterior.TiendaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Tienda',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantContactoCliente_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantContactoCliente_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.ContactoId AS Id, CONVERT(NVARCHAR(40), m.ContactoId) AS Codigo, m.Nombre AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion, cliente.Codigo AS ClienteCodigo, CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END AS ClienteNombre
        INTO #Pagina FROM dbo.ContactoCliente AS m INNER JOIN dbo.Cliente AS cliente ON cliente.ClienteId=m.ClienteId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR CONVERT(NVARCHAR(40), m.ContactoId) LIKE @Patron ESCAPE N'~' OR m.Nombre LIKE @Patron ESCAPE N'~' OR cliente.Codigo LIKE @Patron ESCAPE N'~' OR CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1) AND (@PadreId IS NULL OR m.ClienteId = @PadreId);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion, p.ClienteCodigo, p.ClienteNombre
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantContactoCliente_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantContactoCliente_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.ContactoId AS Id,
            m.ClienteId,
            m.TiendaId,
            m.Nombre,
            m.Telefono,
            m.Cargo,
            m.Correo,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            CASE WHEN ref0.TipoEntidad=N'NATURAL' THEN CONCAT(ref0.Nombres, N' ', ref0.Apellidos) ELSE ref0.RazonSocial END AS ClienteIdLabel,
            ref1.Nombre AS TiendaIdLabel
        FROM dbo.ContactoCliente AS m LEFT JOIN dbo.Cliente AS ref0 ON ref0.ClienteId = m.ClienteId LEFT JOIN dbo.Tienda AS ref1 ON ref1.TiendaId = m.TiendaId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.ContactoId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ContactoCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'ContactoCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantContactoCliente_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantContactoCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF @ClienteId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: ClienteId.',1;
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        DECLARE @TiendaId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TiendaId'));
        IF JSON_VALUE(@Datos,N'$.TiendaId') IS NOT NULL AND @TiendaId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: TiendaId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombre' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombre.',1;
        DECLARE @Nombre NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombre'))),N'');
        IF @Nombre IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Nombre.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Telefono' AND LEN(dato.value)>50) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Telefono.',1;
        DECLARE @Telefono NVARCHAR(50) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Telefono'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Cargo' AND LEN(dato.value)>100) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Cargo.',1;
        DECLARE @Cargo NVARCHAR(100) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Cargo'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Correo' AND LEN(dato.value)>256) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Correo.',1;
        DECLARE @Correo NVARCHAR(256) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Correo'))),N'');
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.ClienteId, anterior.TiendaId, anterior.Nombre, anterior.Telefono, anterior.Cargo, anterior.Correo, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ContactoCliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.ContactoId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId'))<>@ClienteId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @TiendaId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TiendaId')),-1)<>@TiendaId) AND NOT EXISTS (SELECT 1 FROM dbo.Tienda AS destino WITH (HOLDLOCK) WHERE destino.TiendaId=@TiendaId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: TiendaId.',1;
        IF @TiendaId IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.Tienda AS tienda WHERE tienda.TiendaId=@TiendaId AND tienda.ClienteId=@ClienteId) THROW 51217,N'[HYT_VALIDACION] La tienda no pertenece al cliente.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ContactoCliente (ClienteId, TiendaId, Nombre, Telefono, Cargo, Correo, Activo, UsuarioRegistro)
            VALUES (@ClienteId, @TiendaId, @Nombre, @Telefono, @Cargo, @Correo, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ContactoCliente SET TiendaId=@TiendaId, Nombre=@Nombre, Telefono=@Telefono, Cargo=@Cargo, Correo=@Correo, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ContactoId=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.ClienteId, posterior.TiendaId, posterior.Nombre, posterior.Telefono, posterior.Cargo, posterior.Correo, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ContactoCliente AS posterior WHERE posterior.ContactoId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ContactoCliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantProveedorCliente_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedorCliente_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.Id AS Id, CONVERT(NVARCHAR(40), m.Id) AS Codigo, COALESCE(m.ContratoMarco, N'Datos del proveedor') AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.ProveedorCliente AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR CONVERT(NVARCHAR(40), m.Id) LIKE @Patron ESCAPE N'~' OR COALESCE(m.ContratoMarco, N'Datos del proveedor') LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1) AND (@PadreId IS NULL OR m.ClienteId = @PadreId);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantProveedorCliente_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedorCliente_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.Id AS Id,
            m.ClienteId,
            m.CodigoSAPProveedor,
            m.ContratoMarco,
            m.VigenciaDesde,
            m.VigenciaHasta,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            CASE WHEN ref0.TipoEntidad=N'NATURAL' THEN CONCAT(ref0.Nombres, N' ', ref0.Apellidos) ELSE ref0.RazonSocial END AS ClienteIdLabel
        FROM dbo.ProveedorCliente AS m LEFT JOIN dbo.Cliente AS ref0 ON ref0.ClienteId = m.ClienteId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.Id=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ProveedorCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'ProveedorCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantProveedorCliente_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedorCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF @ClienteId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: ClienteId.',1;
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoSAPProveedor' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoSAPProveedor.',1;
        DECLARE @CodigoSAPProveedor NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoSAPProveedor'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'ContratoMarco' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: ContratoMarco.',1;
        DECLARE @ContratoMarco NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.ContratoMarco'))),N'');
        DECLARE @VigenciaDesde DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaDesde'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaDesde') IS NOT NULL AND @VigenciaDesde IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaDesde.',1;
        DECLARE @VigenciaHasta DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaHasta'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaHasta') IS NOT NULL AND @VigenciaHasta IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaHasta.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.ClienteId, anterior.CodigoSAPProveedor, anterior.ContratoMarco, anterior.VigenciaDesde, anterior.VigenciaHasta, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ProveedorCliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId'))<>@ClienteId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @VigenciaDesde IS NOT NULL AND @VigenciaHasta<@VigenciaDesde THROW 51219,N'[HYT_VALIDACION] Revisa las fechas de vigencia.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ProveedorCliente (ClienteId, CodigoSAPProveedor, ContratoMarco, VigenciaDesde, VigenciaHasta, Activo, UsuarioRegistro)
            VALUES (@ClienteId, @CodigoSAPProveedor, @ContratoMarco, @VigenciaDesde, @VigenciaHasta, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ProveedorCliente SET CodigoSAPProveedor=@CodigoSAPProveedor, ContratoMarco=@ContratoMarco, VigenciaDesde=@VigenciaDesde, VigenciaHasta=@VigenciaHasta, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.ClienteId, posterior.CodigoSAPProveedor, posterior.ContratoMarco, posterior.VigenciaDesde, posterior.VigenciaHasta, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ProveedorCliente AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ProveedorCliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantUnidadMedida_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantUnidadMedida_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.Id AS Id, m.Codigo AS Codigo, m.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.UnidadMedida AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Descripcion LIKE @Patron ESCAPE N'~' OR CONVERT(NVARCHAR(40),m.Id) LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantUnidadMedida_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantUnidadMedida_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.Id AS Id,
            m.Codigo,
            m.Descripcion,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion
        FROM dbo.UnidadMedida AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.Id=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'UnidadMedida' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'UnidadMedida' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantUnidadMedida_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantUnidadMedida_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'UNIDADES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Codigo' AND LEN(dato.value)>30) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Codigo.',1;
        DECLARE @Codigo NVARCHAR(30) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Codigo'))),N'');
        IF @Codigo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Codigo.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.UnidadMedida AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.UnidadMedida (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.UnidadMedida SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.UnidadMedida AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'UnidadMedida',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        IF ERROR_NUMBER() IN(2601,2627) SET @ErrorMessage=N'[HYT_VALIDACION] Ya existe una unidad con esa abreviatura.';
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantZona_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantZona_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.Id AS Id, m.Codigo AS Codigo, m.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.Zona AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Descripcion LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantZona_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantZona_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.Id AS Id,
            m.Codigo,
            m.Descripcion,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion
        FROM dbo.Zona AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.Id=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Zona' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Zona' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantZona_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantZona_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'ZONAS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(30);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Zona AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Zona',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'ZON-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Zona AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Zona (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Zona SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Zona AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Zona',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantEspecialidad_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantEspecialidad_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.Id AS Id, m.Codigo AS Codigo, m.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.Especialidad AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Descripcion LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantEspecialidad_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantEspecialidad_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.Id AS Id,
            m.Codigo,
            m.Descripcion,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion
        FROM dbo.Especialidad AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.Id=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Especialidad' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Especialidad' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantEspecialidad_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantEspecialidad_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(30);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Especialidad AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Especialidad',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'ESP-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Especialidad AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Especialidad (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Especialidad SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Especialidad AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Especialidad',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTipoServicio_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTipoServicio_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.Id AS Id, m.Codigo AS Codigo, m.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.TipoServicio AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Descripcion LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTipoServicio_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTipoServicio_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.Id AS Id,
            m.Codigo,
            m.Descripcion,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion
        FROM dbo.TipoServicio AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.Id=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'TipoServicio' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'TipoServicio' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTipoServicio_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTipoServicio_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(30);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Descripcion, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.TipoServicio AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.Id=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:TipoServicio',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'TS-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.TipoServicio AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.TipoServicio (Codigo, Descripcion, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Descripcion, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.TipoServicio SET Codigo=@Codigo, Descripcion=@Descripcion, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Descripcion, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TipoServicio AS posterior WHERE posterior.Id=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TipoServicio',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantServicioMaterial_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantServicioMaterial_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.MaterialId AS Id, m.CodigoInterno AS Codigo, m.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX), z.Descripcion), N', ') FROM dbo.MaterialZona AS mz INNER JOIN dbo.Zona AS z ON z.Id = mz.ZonaId WHERE mz.MaterialId = m.MaterialId) AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.ServicioMaterial AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.CodigoInterno LIKE @Patron ESCAPE N'~' OR m.Descripcion LIKE @Patron ESCAPE N'~' OR m.CodigoMaterialSAP LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantServicioMaterial_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantServicioMaterial_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.MaterialId AS Id,
            m.CodigoInterno,
            m.CodigoMaterialSAP,
            m.Descripcion,
            m.UnidadMedidaId,
            m.EspecialidadId,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            ref3.Codigo AS UnidadCodigo,
            ref3.Descripcion AS UnidadMedidaIdLabel,
            ref4.Descripcion AS EspecialidadIdLabel,
            JSON_QUERY((SELECT zona.ZonaId AS Id, catalogo.Codigo, catalogo.Descripcion FROM dbo.MaterialZona AS zona INNER JOIN dbo.Zona AS catalogo ON catalogo.Id=zona.ZonaId WHERE zona.MaterialId = m.MaterialId ORDER BY zona.ZonaId FOR JSON PATH)) AS Zonas
        FROM dbo.ServicioMaterial AS m LEFT JOIN dbo.UnidadMedida AS ref3 ON ref3.Id = m.UnidadMedidaId LEFT JOIN dbo.Especialidad AS ref4 ON ref4.Id = m.EspecialidadId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.MaterialId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ServicioMaterial' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'ServicioMaterial' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantServicioMaterial_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantServicioMaterial_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'MATERIALES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @CodigoInterno NVARCHAR(60);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMaterialSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMaterialSAP.',1;
        DECLARE @CodigoMaterialSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMaterialSAP'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>1500) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(1500) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        IF @Descripcion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Descripcion.',1;
        DECLARE @UnidadMedidaId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.UnidadMedidaId'));
        IF @UnidadMedidaId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: UnidadMedidaId.',1;
        IF JSON_VALUE(@Datos,N'$.UnidadMedidaId') IS NOT NULL AND @UnidadMedidaId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: UnidadMedidaId.',1;
        DECLARE @EspecialidadId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.EspecialidadId'));
        IF JSON_VALUE(@Datos,N'$.EspecialidadId') IS NOT NULL AND @EspecialidadId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: EspecialidadId.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.CodigoInterno, anterior.CodigoMaterialSAP, anterior.Descripcion, anterior.UnidadMedidaId, anterior.EspecialidadId, anterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.MaterialZona AS z WHERE z.MaterialId=anterior.MaterialId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ServicioMaterial AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.MaterialId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @UnidadMedidaId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.UnidadMedidaId')),-1)<>@UnidadMedidaId) AND NOT EXISTS (SELECT 1 FROM dbo.UnidadMedida AS destino WITH (HOLDLOCK) WHERE destino.Id=@UnidadMedidaId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: UnidadMedidaId.',1;
        IF @EspecialidadId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.EspecialidadId')),-1)<>@EspecialidadId) AND NOT EXISTS (SELECT 1 FROM dbo.Especialidad AS destino WITH (HOLDLOCK) WHERE destino.Id=@EspecialidadId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: EspecialidadId.',1;
        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
        IF @Id IS NOT NULL SET @CodigoInterno=JSON_VALUE(@Anterior,N'$.CodigoInterno');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:ServicioMaterial',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @CodigoInterno=N'MAT-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.ServicioMaterial AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.CodigoInterno=@CodigoInterno) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ServicioMaterial (CodigoInterno, CodigoMaterialSAP, Descripcion, UnidadMedidaId, EspecialidadId, Activo, UsuarioRegistro)
            VALUES (@CodigoInterno, @CodigoMaterialSAP, @Descripcion, @UnidadMedidaId, @EspecialidadId, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.ServicioMaterial SET CodigoInterno=@CodigoInterno, CodigoMaterialSAP=@CodigoMaterialSAP, Descripcion=@Descripcion, UnidadMedidaId=@UnidadMedidaId, EspecialidadId=@EspecialidadId, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE MaterialId=@Id;
        END;
        DELETE FROM dbo.MaterialZona WHERE MaterialId=@Id;
        INSERT INTO dbo.MaterialZona (MaterialId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        SELECT @Posterior=(SELECT posterior.CodigoInterno, posterior.CodigoMaterialSAP, posterior.Descripcion, posterior.UnidadMedidaId, posterior.EspecialidadId, posterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.MaterialZona AS z WHERE z.MaterialId=posterior.MaterialId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ServicioMaterial AS posterior WHERE posterior.MaterialId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ServicioMaterial',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: Script/012_PermisosEtapa3.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCatalogos TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCliente_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCliente_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCliente_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTienda_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTienda_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTienda_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantContactoCliente_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantContactoCliente_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantContactoCliente_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantProveedorCliente_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantProveedorCliente_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantProveedorCliente_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantUnidadMedida_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantUnidadMedida_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantUnidadMedida_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantZona_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantZona_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantZona_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantEspecialidad_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantEspecialidad_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantEspecialidad_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTipoServicio_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTipoServicio_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTipoServicio_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantServicioMaterial_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantServicioMaterial_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantServicioMaterial_Guardar TO hyt_app_executor;
GO

