-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Retira persistencia de estado de onboarding y deja el flujo guiado basado solo en checklist real.
-- =============================================

SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.Sp_OnboardingEstado_Obtener', N'P') IS NOT NULL
    DROP PROCEDURE dbo.Sp_OnboardingEstado_Obtener;

IF OBJECT_ID(N'dbo.Sp_OnboardingEstado_GuardarAvance', N'P') IS NOT NULL
    DROP PROCEDURE dbo.Sp_OnboardingEstado_GuardarAvance;

IF OBJECT_ID(N'dbo.Sp_OnboardingEstado_MarcarCompletado', N'P') IS NOT NULL
    DROP PROCEDURE dbo.Sp_OnboardingEstado_MarcarCompletado;

IF OBJECT_ID(N'dbo.NegocioOnboardingEstado', N'U') IS NOT NULL
    DROP TABLE dbo.NegocioOnboardingEstado;
