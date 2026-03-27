-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Solicitudes publicas de reserva desde la Home (fase 1 de confirmacion).
-- =============================================

IF OBJECT_ID(N'dbo.SolicitudesReservaPublica', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SolicitudesReservaPublica
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        CodigoSolicitud NVARCHAR(20) NULL,
        EspacioDeportivoId INT NOT NULL,
        Fecha DATE NOT NULL,
        HoraInicio TIME NOT NULL,
        HoraFin TIME NOT NULL,
        NombreSolicitante NVARCHAR(200) NOT NULL,
        Telefono NVARCHAR(30) NOT NULL,
        Correo NVARCHAR(200) NULL,
        Comentario NVARCHAR(300) NULL,
        Estado INT NOT NULL CONSTRAINT DF_SolicitudesReservaPublica_Estado DEFAULT (1),
        NotificadoCliente BIT NOT NULL CONSTRAINT DF_SolicitudesReservaPublica_NotificadoCliente DEFAULT (0),
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_SolicitudesReservaPublica_FechaRegistro DEFAULT (SYSUTCDATETIME()),
        CONSTRAINT FK_SolicitudesReservaPublica_EspaciosDeportivos_EspacioDeportivoId FOREIGN KEY (EspacioDeportivoId) REFERENCES dbo.EspaciosDeportivos (Id)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.SolicitudesReservaPublica') AND name = N'IX_SolicitudesReservaPublica_FechaEspacio')
BEGIN
    CREATE INDEX IX_SolicitudesReservaPublica_FechaEspacio ON dbo.SolicitudesReservaPublica (Fecha, EspacioDeportivoId);
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_SolicitarReservaPublica
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @NombreSolicitante NVARCHAR(200),
    @Telefono NVARCHAR(30),
    @Correo NVARCHAR(200) = NULL,
    @Comentario NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            WHERE e.Id = @EspacioDeportivoId
              AND e.Estado = 1
        )
            RAISERROR('El espacio deportivo no esta disponible.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('El horario seleccionado ya no esta disponible.', 16, 1);

        INSERT INTO dbo.SolicitudesReservaPublica
        (
            EspacioDeportivoId, Fecha, HoraInicio, HoraFin,
            NombreSolicitante, Telefono, Correo, Comentario,
            Estado, NotificadoCliente, FechaRegistro
        )
        VALUES
        (
            @EspacioDeportivoId, @Fecha, @HoraInicio, @HoraFin,
            @NombreSolicitante, @Telefono, @Correo, @Comentario,
            1, 0, SYSUTCDATETIME()
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        DECLARE @CodigoSolicitud NVARCHAR(20);
        SET @CodigoSolicitud = CONCAT(N'SR', FORMAT(GETDATE(), 'yyyyMMdd'), RIGHT(CONCAT(N'00000', CONVERT(NVARCHAR(10), @Id)), 5));

        UPDATE dbo.SolicitudesReservaPublica
        SET CodigoSolicitud = @CodigoSolicitud
        WHERE Id = @Id;

        SELECT @CodigoSolicitud;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
