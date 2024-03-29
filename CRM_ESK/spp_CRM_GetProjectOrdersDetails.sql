USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_CRM_GetProjectOrdersDetails]    Дата сценария: 05/05/2020 14:17:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER Procedure [dbo].[spp_CRM_GetProjectOrdersDetails]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|  Получение деталей по заказам для проекта                                            |
|  в CRM                                                                               |
|  Разработчик Новожилов А.Н. 2020                                                     |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyProjectID nvarchar(50)						--код проекта


WITH RECOMPILE
AS
set nocount on

SELECT DocumentNumber, 
	DocumentDate, 
	SUM(DocumentSumm) AS DocumentSumm, 
	SalesmanName
FROM (SELECT OR010300.OR01001 AS DocumentNumber, 
		OR010300.OR01015 AS DocumentDate, 
		OR010300.OR01024 * SYCH0100.SYCH006 AS DocumentSumm, 
		LTRIM(RTRIM(OR010300.OR01019)) + ' ' + LTRIM(RTRIM(OR010300.OR01017)) AS SalesmanName
    FROM OR010300 INNER JOIN
		tbl_CRM_Companies ON OR010300.OR01003 = tbl_CRM_Companies.ScalaCustomerCode INNER JOIN
        SYCH0100 ON OR010300.OR01028 = SYCH0100.SYCH001 
		AND OR010300.OR01015 >= SYCH0100.SYCH004 
		AND OR010300.OR01015 < SYCH0100.SYCH005 INNER JOIN
        tbl_SalesHdr_ProjectAddInfo ON OR010300.OR01001 = tbl_SalesHdr_ProjectAddInfo.OrderID
    WHERE (tbl_SalesHdr_ProjectAddInfo.ProjectID = @MyProjectID)
    UNION ALL
    SELECT ST030300.ST03009 AS DocumentNumber, 
		ST030300.ST03015 AS DocumentDate, 
		SUM(ROUND(ST030300.ST03021 * ST030300.ST03020, 2) - ROUND(ST030300.ST03021 * ST030300.ST03020 * ST030300.ST03022 / 100, 2)) AS DocumentSumm, 
		LTRIM(RTRIM(ST030300.ST03007)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS SalesmanName
	FROM ST030300 INNER JOIN
		tbl_SalesHdr_ProjectAddInfo ON ST030300.ST03009 = tbl_SalesHdr_ProjectAddInfo.OrderID INNER JOIN
        ST010300 ON ST030300.ST03007 = ST010300.ST01001
	WHERE (tbl_SalesHdr_ProjectAddInfo.ProjectID = @MyProjectID)
	GROUP BY ST030300.ST03009, 
		ST030300.ST03015, 
		LTRIM(RTRIM(ST030300.ST03007)) + ' ' + LTRIM(RTRIM(ST010300.ST01002))
	) AS View_7
GROUP BY DocumentNumber, 
	DocumentDate, 
	SalesmanName
Order BY DocumentDate