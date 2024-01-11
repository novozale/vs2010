USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_CRM_GetProjectActivityDetails]    Дата сценария: 05/05/2020 14:16:58 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER Procedure [dbo].[spp_CRM_GetProjectActivityDetails]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|  Получение деталей по активностям для проекта                                        |
|  в CRM                                                                               |
|  Разработчик Новожилов А.Н. 2020                                                     |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyProjectID nvarchar(50)						--код проекта


WITH RECOMPILE
AS
set nocount on

SELECT tbl_CRM_Directions.DirectionName, 
	tbl_CRM_EventTypes.EventTypeName, 
	tbl_CRM_Companies.CompanyName, 
	tbl_CRM_Events.ActionTime, 
	ScalaSystemDB.dbo.ScaUsers.FullName
FROM tbl_CRM_Events INNER JOIN
    ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.OwnerID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN
    tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN
    tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN
    tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID
WHERE (tbl_CRM_Events.ProjectID = @MyProjectID)
ORDER BY ActionTime