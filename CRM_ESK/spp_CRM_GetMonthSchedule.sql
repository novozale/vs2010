USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_CRM_GetMonthSchedule]    Дата сценария: 05/05/2020 14:16:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER Procedure [dbo].[spp_CRM_GetMonthSchedule]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|  Получение плана / факта действий одного продавца                                    |
|  в CRM за месяц                                                                      |
|  Разработчик Новожилов А.Н. 2020                                                     |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyYear Integer,								--год плана
@MyMonth Integer,								--Месяц плана
@MyUserID Integer								--код пользователя


WITH RECOMPILE
AS

set nocount on

Declare @MyWrkStr nvarchar(max)
DECLARE @MyCounter Integer
DECLARE @DaysQTY Integer
DECLARE @MyDatePart integer
DECLARE @MyDatePartStr nvarchar(3)
DECLARE @MyDateStart datetime
DECLARE @MyDateFin datetime
DECLARE @MyCurrDate datetime
DECLARE @date DATETIME 
SELECT @date = convert(datetime, '01/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) 
	+ '/' + Convert(nvarchar(4), @MyYear) , 103)
SELECT @DaysQTY = Day(DATEADD(MONTH,1,@date)- day(DATEADD(MONTH,1,@date)))
SELECT @MyDateStart = @date
SELECT @MyDateFin = DATEADD(day, @DaysQTY - 1, @MyDateStart)


-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyMonth') 
and xtype = N'U')
DROP TABLE #_MyMonth


-----------------создание временных таблиц---------------------------------------------
CREATE TABLE #_MyMonth( 
	[CompanyID] [uniqueidentifier], 
	[CompanyScalaCode] [nvarchar](10), 
	[CompanyName] [nvarchar](255),
	[CustProject] [integer],
	[OrdersQTY] [integer],
	[Status] [nvarchar](50) 
) 

SELECT @MyWrkStr = 'ALTER TABLE #_MyMonth ADD '
SELECT @MyCounter = 1
WHILE @MyCounter < @DaysQTY + 1
BEGIN
	SELECT @MyDatePart = DATEPART(dw, convert(datetime, Right('00' + Convert(nvarchar(2), 
		@MyCounter), 2) + '/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) + '/' + 
		Convert(nvarchar(4), @MyYear) , 103))
	SELECT @MyDatePartStr = 
	CASE @MyDatePart
		WHEN 1 THEN '_Пн'
		WHEN 2 THEN '_Вт'
		WHEN 3 THEN '_Ср'
		WHEN 4 THEN '_Чт'
		WHEN 5 THEN '_Пт'
		WHEN 6 THEN '_Сб'
		WHEN 7 THEN '_Вс'
	END

	SELECT @MyWrkStr = @MyWrkStr + '[d' + Convert(nvarchar(3), @MyCounter) + @MyDatePartStr + '] [nvarchar](10), '
	SELECT @MyCounter = @MyCounter + 1
END
SELECT @MyWrkStr = @MyWrkStr + '[RowTotal] [integer] '
Exec (@MyWrkStr)


-----------------Заполнение списка компаний----------------------------------------------
INSERT INTO #_MyMonth
	(CompanyID, CompanyScalaCode, CompanyName, [Status])
SELECT tbl_CRM_Companies.CompanyID, 
	Ltrim(Rtrim(tbl_CRM_Companies.ScalaCustomerCode)) as ScalaCustomerCode, 
	Ltrim(Rtrim(tbl_CRM_Companies.CompanyName)) as CompanyName,
	'План' AS [Status]
FROM tbl_CRM_Companies INNER JOIN
    SL010300 ON tbl_CRM_Companies.ScalaCustomerCode = SL010300.SL01001 INNER JOIN
    ST010300 ON SL010300.SL01035 = ST010300.ST01001 INNER JOIN
    ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName
WHERE (ScalaSystemDB.dbo.ScaUsers.UserID = @MyUserID)
UNION 
SELECT tbl_CRM_Companies.CompanyID, 
	Ltrim(Rtrim(tbl_CRM_Companies.ScalaCustomerCode)) as ScalaCustomerCode, 
	Ltrim(Rtrim(tbl_CRM_Companies.CompanyName)) as CompanyName,
	'Факт' AS [Status]
FROM tbl_CRM_Companies INNER JOIN
    SL010300 ON tbl_CRM_Companies.ScalaCustomerCode = SL010300.SL01001 INNER JOIN
    ST010300 ON SL010300.SL01035 = ST010300.ST01001 INNER JOIN
    ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName
WHERE (ScalaSystemDB.dbo.ScaUsers.UserID = @MyUserID)
UNION
SELECT tbl_CRM_Companies.CompanyID, 
	Ltrim(Rtrim(tbl_CRM_Companies.ScalaCustomerCode)) as ScalaCustomerCode, 
	Ltrim(Rtrim(tbl_CRM_Companies.CompanyName)) as CompanyName,
	'План' AS [Status]
FROM tbl_CRM_Events INNER JOIN
    tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID
WHERE (tbl_CRM_Events.ActionPlannedDate >= @MyDateStart) 
	AND (tbl_CRM_Events.ActionPlannedDate <=  @MyDateFin) 
	AND (tbl_CRM_Events.OwnerID = @MyUserID) 
	AND (tbl_CRM_Companies.ScalaCustomerCode IS NULL)
UNION
SELECT tbl_CRM_Companies.CompanyID, 
	Ltrim(Rtrim(tbl_CRM_Companies.ScalaCustomerCode)) as ScalaCustomerCode, 
	Ltrim(Rtrim(tbl_CRM_Companies.CompanyName)) as CompanyName,
	'Факт' AS [Status]
FROM tbl_CRM_Events INNER JOIN
    tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID
WHERE (tbl_CRM_Events.ActionPlannedDate >= @MyDateStart) 
	AND (tbl_CRM_Events.ActionPlannedDate <=  @MyDateFin) 
	AND (tbl_CRM_Events.OwnerID = @MyUserID) 
	AND (tbl_CRM_Companies.ScalaCustomerCode IS NULL)

-------------Заполнение запланированных активностей--------------------------------------
-----план
SELECT @MyWrkStr = 'UPDATE #_MyMonth SET '
SELECT @MyCounter = 1
WHILE @MyCounter < @DaysQTY + 1
BEGIN
	SELECT @MyDatePart = DATEPART(dw, convert(datetime, Right('00' + Convert(nvarchar(2), 
		@MyCounter), 2) + '/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) + '/' + 
		Convert(nvarchar(4), @MyYear) , 103))
	SELECT @MyDatePartStr = 
	CASE @MyDatePart
		WHEN 1 THEN '_Пн'
		WHEN 2 THEN '_Вт'
		WHEN 3 THEN '_Ср'
		WHEN 4 THEN '_Чт'
		WHEN 5 THEN '_Пт'
		WHEN 6 THEN '_Сб'
		WHEN 7 THEN '_Вс'
	END

	SELECT @MyWrkStr = @MyWrkStr + '[d' + Convert(nvarchar(3), @MyCounter) + @MyDatePartStr + '] =  View_7.['+ Convert(nvarchar(3), @MyCounter) + '], '
	SELECT @MyCounter = @MyCounter + 1
END
SELECT @MyWrkStr = @MyWrkStr + '[RowTotal] = View_7.[RowTotal] '
SELECT @MyWrkStr = @MyWrkStr + 'FROM #_MyMonth INNER JOIN '
SELECT @MyWrkStr = @MyWrkStr + '(SELECT CompanyID, '
SELECT @MyCounter = 1
WHILE @MyCounter < @DaysQTY + 1
BEGIN
	SELECT @MyCurrDate = convert(datetime, Right('00' + Convert(nvarchar(2), @MyCounter), 2) 
		+ '/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) 
		+ '/' +	Convert(nvarchar(4), @MyYear) , 103)
	SELECT @MyWrkStr = @MyWrkStr + 'Convert(nvarchar(10), CASE WHEN SUM(CASE WHEN ActionPlannedDate = CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyCurrDate, 103) + ''', 103) THEN 1 ELSE 0 END) = 0 '
		+ 'THEN NULL '
		+ 'ELSE SUM(CASE WHEN ActionPlannedDate = CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyCurrDate, 103) + ''', 103) THEN 1 ELSE 0 END) '
	+ 'END) AS [' + Convert(nvarchar(5), @MyCounter) + '], '
	SELECT @MyCounter = @MyCounter + 1
END
SELECT @MyWrkStr = @MyWrkStr + 'CASE WHEN SUM(CASE WHEN ActionPlannedDate >= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateStart, 103) + ''', 103) and ActionPlannedDate <= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateFin, 103) + ''', 103) THEN 1 ELSE 0 END) = 0 '
		+ 'THEN NULL '
		+ 'ELSE SUM(CASE WHEN ActionPlannedDate >= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateStart, 103) + ''', 103) AND ActionPlannedDate <= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateFin, 103) + ''', 103) THEN 1 ELSE 0 END) '
	+ 'END AS [RowTotal] '
SELECT @MyWrkStr = @MyWrkStr + 'FROM tbl_CRM_Events '
SELECT @MyWrkStr = @MyWrkStr + 'WHERE (OwnerID = ' + Convert(nvarchar(10), @MyUserID) + ') '
SELECT @MyWrkStr = @MyWrkStr + 'AND (ActionPlannedDate >= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateStart, 103) + ''', 103)) '
SELECT @MyWrkStr = @MyWrkStr + 'AND (ActionPlannedDate <= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateFin, 103) + ''', 103)) '
SELECT @MyWrkStr = @MyWrkStr + 'GROUP BY CompanyID) AS View_7 ON #_MyMonth.CompanyID = View_7.CompanyID '
SELECT @MyWrkStr = @MyWrkStr + 'where ([Status] = N''План'') '
Print @MyWrkStr
exec(@MyWrkStr)

-----факт
SELECT @MyWrkStr = 'UPDATE #_MyMonth SET '
SELECT @MyCounter = 1
WHILE @MyCounter < @DaysQTY + 1
BEGIN
	SELECT @MyDatePart = DATEPART(dw, convert(datetime, Right('00' + Convert(nvarchar(2), 
		@MyCounter), 2) + '/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) + '/' + 
		Convert(nvarchar(4), @MyYear) , 103))
	SELECT @MyDatePartStr = 
	CASE @MyDatePart
		WHEN 1 THEN '_Пн'
		WHEN 2 THEN '_Вт'
		WHEN 3 THEN '_Ср'
		WHEN 4 THEN '_Чт'
		WHEN 5 THEN '_Пт'
		WHEN 6 THEN '_Сб'
		WHEN 7 THEN '_Вс'
	END

	SELECT @MyWrkStr = @MyWrkStr + '[d' + Convert(nvarchar(3), @MyCounter) + @MyDatePartStr + '] =  View_7.['+ Convert(nvarchar(3), @MyCounter) + '], '
	SELECT @MyCounter = @MyCounter + 1
END
SELECT @MyWrkStr = @MyWrkStr + '[RowTotal] = View_7.[RowTotal] '
SELECT @MyWrkStr = @MyWrkStr + 'FROM #_MyMonth INNER JOIN '
SELECT @MyWrkStr = @MyWrkStr + '(SELECT CompanyID, '
SELECT @MyCounter = 1
WHILE @MyCounter < @DaysQTY + 1
BEGIN
	SELECT @MyCurrDate = convert(datetime, Right('00' + Convert(nvarchar(2), @MyCounter), 2) 
		+ '/' + Right('00' + Convert(nvarchar(2), @MyMonth), 2) 
		+ '/' +	Convert(nvarchar(4), @MyYear) , 103)
	SELECT @MyWrkStr = @MyWrkStr + 'Convert(nvarchar(10), CASE WHEN SUM(CASE WHEN ActionPlannedDate = CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyCurrDate, 103) + ''', 103) THEN 1 ELSE 0 END) = 0 '
		+ 'THEN NULL '
		+ 'ELSE SUM(CASE WHEN ActionPlannedDate = CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyCurrDate, 103) + ''', 103) THEN 1 ELSE 0 END) '
	+ 'END) AS [' + Convert(nvarchar(5), @MyCounter) + '], '
	SELECT @MyCounter = @MyCounter + 1
END
SELECT @MyWrkStr = @MyWrkStr + 'CASE WHEN SUM(CASE WHEN ActionPlannedDate >= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateStart, 103) + ''', 103) and ActionPlannedDate <= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateFin, 103) + ''', 103) THEN 1 ELSE 0 END) = 0 '
		+ 'THEN NULL '
		+ 'ELSE SUM(CASE WHEN ActionPlannedDate >= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateStart, 103) + ''', 103) AND ActionPlannedDate <= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateFin, 103) + ''', 103) THEN 1 ELSE 0 END) '
	+ 'END AS [RowTotal] '
SELECT @MyWrkStr = @MyWrkStr + 'FROM tbl_CRM_Events '
SELECT @MyWrkStr = @MyWrkStr + 'WHERE (OwnerID = ' + Convert(nvarchar(10), @MyUserID) + ') '
SELECT @MyWrkStr = @MyWrkStr + 'AND (ActionPlannedDate >= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateStart, 103) + ''', 103)) '
SELECT @MyWrkStr = @MyWrkStr + 'AND (ActionPlannedDate <= CONVERT(DATETIME, ''' + Convert(nvarchar(30), @MyDateFin, 103) + ''', 103)) '
SELECT @MyWrkStr = @MyWrkStr + 'AND (ActionResultID IS NOT NULL) '
SELECT @MyWrkStr = @MyWrkStr + 'GROUP BY CompanyID) AS View_7 ON #_MyMonth.CompanyID = View_7.CompanyID '
SELECT @MyWrkStr = @MyWrkStr + 'where ([Status] = N''Факт'') '
exec(@MyWrkStr)


----------------Заполнение количества заказов, КП и ожидаемых оплат в данный промежуток времени----------------
UPDATE #_MyMonth
SET OrdersQTY = View_11.Expr1
FROM #_MyMonth INNER JOIN
	(SELECT CompanyID, 
		COUNT(OR01001) AS Expr1
    FROM (
		---заказы из OR01
		SELECT tbl_CRM_Companies.CompanyID, 
			OR010300.OR01001
        FROM OR010300 INNER JOIN
			tbl_CRM_Companies ON OR010300.OR01003 = tbl_CRM_Companies.ScalaCustomerCode
        WHERE (OR010300.OR01015 >= @MyDateStart) 
			AND (OR010300.OR01015 <= @MyDateFin)
        UNION
		---заказы из OR20
        SELECT tbl_CRM_Companies_1.CompanyID, 
			OR200300.OR20001
        FROM OR200300 INNER JOIN
			tbl_CRM_Companies AS tbl_CRM_Companies_1 ON OR200300.OR20003 = tbl_CRM_Companies_1.ScalaCustomerCode
        WHERE (OR200300.OR20015 >= @MyDateStart) 
			AND (OR200300.OR20015 <= @MyDateFin)
		UNION
		---КП, не переведенные в заказы
		SELECT tbl_CRM_Companies.CompanyID, 
			tbl_OR010300.OR01001
		FROM tbl_CRM_Companies INNER JOIN
			tbl_OR010300 ON tbl_CRM_Companies.ScalaCustomerCode = tbl_OR010300.OR01003
		WHERE (tbl_OR010300.OR01015 >= @MyDateStart) 
			AND (tbl_OR010300.OR01015 <= @MyDateFin)
			AND (tbl_OR010300.OrderN = N'')
		UNION
		---заказы с истекающим сроком оплаты
		SELECT tbl_CRM_Companies.CompanyID, 
			SL030300.SL03002
		FROM SL030300 INNER JOIN
			tbl_CRM_Companies ON SL030300.SL03001 = tbl_CRM_Companies.ScalaCustomerCode
		WHERE (SL030300.SL03006 >= @MyDateStart) 
			AND (SL030300.SL03006 <= @MyDateFin) 
			AND (LEFT(SL030300.SL03017, 6) = '621010')
		) AS View_10
    GROUP BY CompanyID) AS View_11 ON 
	#_MyMonth.CompanyID = View_11.CompanyID


----------------Заполнение количества проектов в данный промежуток времени---------------
UPDATE #_MyMonth
SET CustProject = View_13.cc
FROM #_MyMonth INNER JOIN
	(SELECT CompanyID, 
		COUNT(ProjectID) AS cc
    FROM tbl_CRM_Projects
    WHERE (FirstDate <= @MyDateStart) 
		AND (LastDate >= @MyDateFin)
    GROUP BY CompanyID) AS View_13 ON 
	#_MyMonth.CompanyID = View_13.CompanyID


----------------Получение итоговой информации--------------------------------------------
Select *
from #_MyMonth
ORDER BY OrdersQTY Desc,
	CompanyName,
	[Status]