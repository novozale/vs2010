USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_CRM_GetMonthScheduleDetails]    Дата сценария: 05/05/2020 14:16:55 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER Procedure [dbo].[spp_CRM_GetMonthScheduleDetails]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|  Получение деталей для плана / факта действий одного продавца                        |
|  в CRM за месяц                                                                      |
|  Разработчик Новожилов А.Н. 2020                                                     |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyYear Integer,								--год плана
@MyMonth Integer,								--Месяц плана
@MyCompanyID nvarchar(50)						--код компании


WITH RECOMPILE
AS
set nocount on

DECLARE @date DATETIME 
DECLARE @MyDateStart datetime
DECLARE @MyDateFin datetime
DECLARE @DaysQTY Integer

SELECT @date = convert(datetime, '01/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) 
	+ '/' + Convert(nvarchar(4), @MyYear) , 103)
SELECT @DaysQTY = Day(DATEADD(MONTH,1,@date)- day(DATEADD(MONTH,1,@date)))
SELECT @MyDateStart = @date
SELECT @MyDateFin = DATEADD(day, @DaysQTY - 1, @MyDateStart)

-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyMonthDetail') 
and xtype = N'U')
DROP TABLE #_MyMonthDetail


-----------------создание временных таблиц---------------------------------------------
CREATE TABLE #_MyMonthDetail( 
	[CompanyID] [uniqueidentifier], 
	[DocumentType] [nvarchar](50),
	[DocumentNumber] [nvarchar](50),
	[DocumentDate] datetime,
	[DocumentSumm] numeric(28,2),
	[SalesmanName] [nvarchar](100)
) 


-----------------Добавление заказов----------------------------------------------------
INSERT INTO #_MyMonthDetail
SELECT CompanyID, 
	DocumentType, 
	DocumentNumber, 
	DocumentDate, 
	SUM(DocumentSumm) AS DocumentSumm, 
	SalesmanName
FROM (SELECT tbl_CRM_Companies.CompanyID, 
		'Счет' AS DocumentType, 
		OR010300.OR01001 AS DocumentNumber, 
		OR010300.OR01015 AS DocumentDate, 
        OR010300.OR01024 * SYCH0100.SYCH006 AS DocumentSumm, 
		LTRIM(RTRIM(OR010300.OR01019)) + ' ' + LTRIM(RTRIM(OR010300.OR01017)) AS SalesmanName
    FROM OR010300 INNER JOIN
		tbl_CRM_Companies ON OR010300.OR01003 = tbl_CRM_Companies.ScalaCustomerCode INNER JOIN
        SYCH0100 ON OR010300.OR01028 = SYCH0100.SYCH001 
		AND OR010300.OR01015 >= SYCH0100.SYCH004 
		AND OR010300.OR01015 < SYCH0100.SYCH005
    WHERE (OR010300.OR01015 >= @MyDateStart) 
		AND (OR010300.OR01015 <= @MyDateFin)
		AND (tbl_CRM_Companies.CompanyID = @MyCompanyID)
    UNION ALL
    /*SELECT tbl_CRM_Companies_1.CompanyID, 
		'Счет' AS DocumentType, 
		OR200300.OR20001 AS DocumentNumber, 
		OR200300.OR20015 AS DocumentDate, 
        OR200300.OR20024 * SYCH0100_1.SYCH006 AS DocumentSumm, 
		LTRIM(RTRIM(OR200300.OR20019)) + ' ' + LTRIM(RTRIM(OR200300.OR20017)) AS SalesmanName
    FROM OR200300 INNER JOIN
		tbl_CRM_Companies AS tbl_CRM_Companies_1 ON OR200300.OR20003 = tbl_CRM_Companies_1.ScalaCustomerCode INNER JOIN
        SYCH0100 AS SYCH0100_1 ON OR200300.OR20028 = SYCH0100_1.SYCH001 
		AND OR200300.OR20015 >= SYCH0100_1.SYCH004 
		AND OR200300.OR20015 < SYCH0100_1.SYCH005
    WHERE (OR200300.OR20015 >= @MyDateStart) 
		AND (OR200300.OR20015 <= @MyDateFin)
		AND (tbl_CRM_Companies_1.CompanyID = @MyCompanyID)*/
	SELECT tbl_CRM_Companies.CompanyID, 
		'Счет' AS DocumentType, 
		ST030300.ST03009 AS DocumentNumber, 
		ST030300.ST03015 AS DocumentDate, 
		SUM(ROUND(ST030300.ST03021 * ST030300.ST03020, 2) - ROUND(ST030300.ST03021 * ST030300.ST03020 * ST030300.ST03022 / 100, 2)) AS DocumentSumm, 
		LTRIM(RTRIM(ST030300.ST03007)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS SalesmanName
	FROM ST030300 INNER JOIN
		tbl_SalesHdr_ProjectAddInfo ON ST030300.ST03009 = tbl_SalesHdr_ProjectAddInfo.OrderID INNER JOIN
        ST010300 ON ST030300.ST03007 = ST010300.ST01001 INNER JOIN
        tbl_CRM_Companies ON ST030300.ST03001 = tbl_CRM_Companies.ScalaCustomerCode
	WHERE (ST030300.ST03015 <= @MyDateFin) 
		AND (tbl_CRM_Companies.CompanyID = @MyCompanyID) 
		AND (ST030300.ST03015 >= @MyDateStart)
	GROUP BY ST030300.ST03009, 
		ST030300.ST03015, 
		tbl_CRM_Companies.CompanyID,
		LTRIM(RTRIM(ST030300.ST03007)) + ' ' + LTRIM(RTRIM(ST010300.ST01002))
	) AS View_10
GROUP BY CompanyID, 
	DocumentType, 
	DocumentNumber,
	DocumentDate,
	SalesmanName


-----------------Добавление КП---------------------------------------------------------
INSERT INTO #_MyMonthDetail
SELECT tbl_CRM_Companies.CompanyID, 
	'Коммерческое предложение' AS DocumentType, 
	tbl_OR010300.OR01001 AS DocumentNumber, 
	CONVERT(nvarchar(30), tbl_OR010300.OR01015, 103) AS DocumentDate, 
	tbl_OR010300.OR01024 * SYCH0100.SYCH006 AS DocumentSumm, 
	LTRIM(RTRIM(tbl_OR010300.OR01019)) + ' ' + LTRIM(RTRIM(tbl_OR010300.OR01017)) AS SalesmanName
FROM tbl_CRM_Companies INNER JOIN
    tbl_OR010300 ON tbl_CRM_Companies.ScalaCustomerCode = tbl_OR010300.OR01003 INNER JOIN
    SYCH0100 ON tbl_OR010300.OR01028 = SYCH0100.SYCH001 
	AND tbl_OR010300.OR01015 >= SYCH0100.SYCH004 
	AND tbl_OR010300.OR01015 < SYCH0100.SYCH005
WHERE (tbl_OR010300.OR01015 >= @MyDateStart) 
	AND (tbl_OR010300.OR01015 <= @MyDateFin) 
	AND (tbl_OR010300.OrderN = N'')
	AND (tbl_CRM_Companies.CompanyID = @MyCompanyID)


-----------------Добавление СФ, у которых выходит срок оплаты--------------------------
INSERT INTO #_MyMonthDetail
SELECT tbl_CRM_Companies.CompanyID, 
	'Счет Фактура' AS DocumentType, 
	SL030300.SL03002 AS DocumentNumber, 
	SL030300.SL03006 AS DocumentDate, 
    SL030300.SL03013 - SL030300.SL03053 AS DocumentSumm, 
	LTRIM(RTRIM(SL030300.SL03041)) + ' ' + LTRIM(RTRIM(ST010300.ST01002)) AS SalesmanName
FROM SL030300 INNER JOIN
    tbl_CRM_Companies ON SL030300.SL03001 = tbl_CRM_Companies.ScalaCustomerCode INNER JOIN
    ST010300 ON SL030300.SL03041 = ST010300.ST01001
WHERE (SL030300.SL03006 >= @MyDateStart) 
	AND (SL030300.SL03006 <= @MyDateFin) 
	AND (LEFT(SL030300.SL03017, 6) = '621010')
	AND (tbl_CRM_Companies.CompanyID = @MyCompanyID)


-----------------Вывод результата------------------------------------------------------
SELECT * 
FROM #_MyMonthDetail
ORDER BY CompanyID,
	DocumentType,
	DocumentNumber
