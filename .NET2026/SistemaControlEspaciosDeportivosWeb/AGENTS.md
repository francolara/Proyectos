# Instrucciones Codex para el repositorio Proyectos

## Alcance general
1. Revisa estas pautas antes de trabajar en cualquier archivo de este repositorio.
2. Limita los cambios a la lógica de negocio. La interfaz gráfica de los proyectos WinForms (VB6 o .NET) no debe alterarse.
3. Manten siempre el modelo del proyecto ADO.NET + SP

# ?? Regla de control de firmas por fecha

Para evitar múltiples firmas el mismo día, se deben aplicar las siguientes reglas:

1. Si el archivo *NO tiene una firma con la fecha actual, se debe **crear una nueva firma*.
2. Si el archivo *YA tiene una firma con la fecha actual, **NO se debe crear otra*.
3. En ese caso, se debe *modificar la firma existente del día actual*.
4. La descripción debe *integrar de forma resumida todos los cambios realizados durante ese mismo día*.
5. Las firmas de *días anteriores nunca deben modificarse*.
6. Solo se permite modificar la firma cuya *Create date coincida con la fecha actual*.

---

### Ejemplo de comportamiento correcto

Estado inicial del archivo:

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/03/2026
-- =============================================

Si hoy se modifica el script (06/03/2026):

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/03/2026
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/03/2026
-- =============================================

Si el mismo día se vuelve a modificar el script:

NO se debe crear otra firma.

Se debe modificar la firma existente del día:

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/03/2026
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/03/2026
-- =============================================

---

# ?? Estándar de generación de SQL Server

Cuando Codex cree o modifique procedimientos almacenados, funciones o scripts SQL Server debe seguir estas reglas.

## Creación de procedimientos

Siempre usar:

CREATE OR ALTER PROCEDURE

Nunca usar:

CREATE PROCEDURE

Esto evita errores al redeployar scripts.

---

## Estructura base obligatoria de Stored Procedure

Todo procedimiento debe seguir esta estructura:

CREATE OR ALTER PROCEDURE dbo.NombreProcedimiento

AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        -- Lógica del procedimiento

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END

---

## Buenas prácticas SQL obligatorias

1. Siempre usar SET NOCOUNT ON.
2. SELECT * está **prohibido** en Stored Procedures.
3. Siempre declarar explícitamente las columnas.
4. Usar alias claros para tablas.
5. Evitar subconsultas innecesarias si pueden reemplazarse por JOIN.
6. Cuando sea posible usar TRY/CATCH para control de errores.
7. Evitar transacciones innecesarias.
8. Si se usan tablas temporales, limpiarlas cuando corresponda.
9. Nunca incluir GO dentro de un procedimiento almacenado.
10. Los scripts deben ser *idempotentes* cuando sea posible.
## Checklist previo a confirmar
1. Ejecuta `git status` y `git diff --stat` para comprobar que solo cambiaste lo necesario.
2. Usa `file "ruta/al/archivo"` para validar la codificación de cada archivo modificado.


## Otros recordatorios
- Los procedimientos, scripts y tablas residen en la carpeta Basededatos/SportCenter.
- Si se modifica o crea un procedimiento Basededatos/SportCenter/StoreProcedure.
- Si se Altera una estructura de una tabla o se agrega un insert o update para datos de una tabla  Basededatos/SportCenter/Script.
- Para cambios masivos de codificación, apóyate en `.gitattributes` y en `git add --renormalize .`.