USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_CRM_GetProjectInfo]    Дата сценария: 05/05/2020 14:17:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER Procedure [dbo].[spp_CRM_GetProjectInfo]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|  Получение информации о проектах за промежуток времени                               |
|  из CRM                                                                              |
|  Разработчик Новожилов А.Н. 2020                                                     |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyDateStartStr nvarchar(50),						--начало интервала
@MyDateFinStr nvarchar(50)							--конец интервала


WITH RECOMPILE
AS
set nocount on

DECLARE @MyDateStart datetime
DECLARE @MyDateFin datetime
SELECT @MyDateStart = CONVERT(datetime, @MyDateStartStr, 103)
SELECT @MyDateFin = CONVERT(datetime, @MyDateFinStr, 103)

SELECT tbl_CRM_Projects.ProjectID, 
	tbl_CRM_Projects.CompanyID, 
	Ltrim(Rtrim(ISNULL(tbl_CRM_Companies.ScalaCustomerCode, ''))) AS ScalaCustomerCode, 
	tbl_CRM_Companies.CompanyName, 
	tbl_CRM_Projects.ProjectName, 
	ISNULL(tbl_CRM_Projects.ProjectSumm, 0) AS ProjectSumm, 
	ISNULL(tbl_CRM_Projects.ProjectComment, '') AS ProjectComment, 
	tbl_CRM_Projects.FirstDate, 
	tbl_CRM_Projects.LastDate, 
	tbl_CRM_Projects.StartDate, 
	ISNULL(tbl_CRM_Projects.CloseDate, tbl_CRM_Projects.LastDate) AS CloseDate,
	ISNULL(tbl_CRM_Projects.ProposalDate, tbl_CRM_Projects.StartDate) AS ProposalDate,
    ISNULL(tbl_CRM_Projects.ProjectAddr, '') AS ProjectAddr, 
	ISNULL(tbl_CRM_Projects.Investor, '') AS Investor, 
	ISNULL(tbl_CRM_Projects.Contractor, '') AS Contractor, 
	ISNULL(tbl_CRM_Projects.ResponciblePerson, '') AS ResponciblePerson, 
	ISNULL(tbl_CRM_Projects.ManufacturersList, '') AS ManufacturersList, 
    ISNULL(tbl_CRM_Projects.AlterManufacturers, 0) AS AlterManufacturers, 
	ISNULL(tbl_CRM_Projects.Competitors, '') AS Competitors, 
	ISNULL(tbl_CRM_Projects.AdditionalExpencesPerCent, 0) AS AdditionalExpencesPerCent, 
	ISNULL(tbl_CRM_Projects.IsApproved, 0) AS IsApproved
FROM tbl_CRM_Projects INNER JOIN
    tbl_CRM_Companies ON tbl_CRM_Projects.CompanyID = tbl_CRM_Companies.CompanyID
WHERE (tbl_CRM_Projects.FirstDate <= @MyDateFin) 
	AND (tbl_CRM_Projects.LastDate >= @MyDateStart)
	AND (ISNULL(tbl_CRM_Projects.CloseDate, tbl_CRM_Projects.LastDate) >= @MyDateStart)