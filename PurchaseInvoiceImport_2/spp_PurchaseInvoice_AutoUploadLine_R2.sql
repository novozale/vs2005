USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_PurchaseInvoice_AutoUploadLine_R2]    Дата сценария: 05/25/2012 08:54:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_PurchaseInvoice_AutoUploadLine_R2]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Приемка одного запаса (строки) по СФ поставщика                                   |
|    присланного в формате XML                                                         |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@ID int,                                            -- ID строки в таблице  
@Invoice nvarchar(35),                              -- Номер СФ
@InvoiceDateSTR nvarchar(50),                       -- Дата СФ
@CurrCode int,                                      -- код валюты в инвойсе поставщика
@MySalesmanCode nvarchar(6),                        -- код продавца
@MySalesmanName nvarchar(50),                       -- ФИО продавца
@PurchInvoiceExRate numeric(28,8),                  -- курс валюты поставщика, указанный им в инвойсе (на дату СФ)
@ConsPOrder nvarchar(10),                           -- Номер консолидированного заказа на закупку
@ItemSuppCode nvarchar(35),                         -- код товара поставщика
@QTY numeric(28,8),                                 -- количество в СФ в единицах закупки
@Price numeric(28,8),                               -- сумма строки без НДС
@Country nvarchar(50),                              -- страна производителя
@GTD nvarchar (255),                                -- ГТД
@MyRezStr nvarchar(4000) output,                    -- результирующая строка
@MyRestQTY numeric(28,8)output                      -- непринятый остаток

WITH RECOMPILE
AS

Declare @InvoiceQTY numeric(28,8);                  -- количество в единицах закупки, кот. осталось принять по СФ
Declare @InvoiceDate datetime;                      -- Дата СФ
Declare @InvoiceInputDate datetime                  -- дата прогрузки СФ в систему
declare @MyFinDate datetime;                        -- 31/12/9999 
declare @MyStartDate datetime;                      -- 01/01/1900 
Declare @MyInvoiceExRate numeric(28,8);             -- курс валюты поставщика из Scala (на дату СФ)
Declare @MyPOrder nchar(10);                        -- номер заказа на закупку (неконсолидированный) 
Declare @MyCount1 int;                              -- счетчик
Declare @MyCurrCode int;                            -- код валюты в нашем заказе на закупку
Declare @MyOrderExRate numeric(28,8);               -- курс валюты заказа из Scala на дату СФ
Declare @MyInvoiceExRateCD numeric(28,8);           -- курс валюты поставщика из Scala (на дату приемки)
DECLARE @MyOrderExRateCD numeric(28,8);             -- курс валюты заказа из Scala на дату приемки
Declare @MyItemCode nvarchar(35);                   -- код запаса (Скальский) в заказе на закупку 
Declare @OrderQTY numeric(28,8);                    -- количество в единицах закупки, кот. осталось принять в заказе
Declare @OrderWHQTY numeric(28,8);                  -- кол-во в единицах измерения хранения, кот. осталось принять в заказе
Declare @OrderQTYRcv numeric(28,8);                 -- количество в единицах закупки, принимаемое в заказ
Declare @OrderWHQTYRcv numeric(28,8);               -- кол-во в единицах измерения хранения, принимаемое в заказ
Declare @MySC01140 numeric(28,8);                   -- коэфф. единиц закупки
Declare @MySC01141 numeric(28,8);                   -- коэфф. единиц хранения
Declare @MyCustomPc numeric(28,8);                  -- % таможенных сборов
Declare @WHQTY numeric(28,8);                       -- Количество в СФ в единицах измерения хранения
Declare @IntInvoice nvarchar(35);                   -- автоматический внутренний номер накладной
Declare @MyStrNum nvarchar(6);                      -- номер строки заказа
Declare @MySubStrNum nvarchar(6);                   -- номер подстроки заказа 
Declare @MyAutStrNum nvarchar(6);                   -- авт. номер подстроки заказа 
Declare @MyBatchNum nvarchar(12);                   -- номер партии
Declare @MyExchRate2 [numeric](18, 8);              -- Второй курс валюты (рубли - 1, Евро - 0)
Declare @MyExchAlgorithm [nchar](3);                -- Алгоритм обмена (рубли - *, Евро - 1)
Declare @MyNetWeight numeric(28,8);                 -- Вес нетто товара
Declare @MyGrossWeight numeric(28,8);               -- Вес Gross товара
Declare @MyVolume numeric(28,8);                    -- объем товара
declare @MySuppCode nvarchar(10);                   -- код поставщика
declare @MyWHNum nvarchar(10);                      -- номер склада, на который идет приемка
Declare @MyGTDSuppDate nvarchar(35);                -- часть ГТД - поставщик и дата
Declare @MyGTDNum nvarchar(35);                     -- часть ГТД - N п/п
Declare @MyGTD_GTD nvarchar(35);                    -- часть ГТД - собственно ГТД
Declare @MyFirstPos int;                            -- начало 2 части ГТД
Declare @MySecondPos int;                           -- начало 3 части ГТД
Declare @MyCountryCode nvarchar(35);                -- код страны - поставщика
Declare @MyCell [nvarchar](6);                      -- ячейка по умолчанию для хранения запаса на конкретном складе
Declare @MySYVT002 nvarchar(400);                   -- Значение поля SYVT002
Declare @MyAccountDef nvarchar(12);                 -- Account поставщика по умолчанию
Declare @MySearchIndex nvarchar(12);                -- SearchIndex
Declare @MyUserID nvarchar(255);                    -- идентификатор пользователя
Declare @MyFIFOValue numeric(28,8);                 -- FIFO цена по запасу по складу
Declare @MyAVGCostPrice2 numeric(28,8);             -- Средняя цена по запасу по складу
Declare @MyAVGCostPrice1 numeric(28,8);             -- Средняя цена по запасу
Declare @MySalesOrder nvarchar(10);                 -- Номер заказа на продажу

SELECT @InvoiceQTY = @QTY
SELECT @Invoice = Ltrim(Rtrim(@Invoice))
SELECT @InvoiceDate = CONVERT(datetime,Ltrim(Rtrim(@InvoiceDateSTR)),103)
SELECT @InvoiceInputDate = dateadd( day, datediff(day, 0, GETDATE()), 0)
SELECT @MyFinDate = CONVERT(datetime,'31/12/9999',103)
SELECT @MyStartDate = CONVERT(datetime,'01/01/1900',103)
SELECT @MySalesmanCode = Ltrim(Rtrim(@MySalesmanCode))
SELECT @MySalesmanName = Ltrim(Rtrim(@MySalesmanName))
SELECT @ConsPOrder = Ltrim(Rtrim(@ConsPOrder))
SELECT @ItemSuppCode = Ltrim(Rtrim(@ItemSuppCode))
SELECT @Country = Ltrim(Rtrim(@Country))
SELECT @GTD = Ltrim(Rtrim(@GTD))
SELECT @MyRezStr = ''
--SELECT @MyRezStr = 'ID строки ' + CONVERT(nvarchar(255),@ID) +  ', Номер СФ ' + @Invoice + ' от ' + CONVERT(nvarchar(255),@InvoiceDate,103) + ', код валюты ' + CONVERT(nvarchar(255),@CurrCode) + ', код продавца ' + @MySalesmanCode + ', ФИО продавца ' + @MySalesmanName + ', курс валюты поставщика ' + CONVERT(nvarchar(255),@PurchInvoiceExRate) + ', Консолидированный заказ на закупку ' + @ConsPOrder + ', код товара поставщика ' + @ItemSuppCode + ', кол - во ' + CONVERT(nvarchar(255), @QTY) + ', сумма без НДС ' + CONVERT(nvarchar(255),@Price) +  ', страна производитель ' + @Country + ', ГТД ' + @GTD  + CHAR(13)


-----------Проверка, что курс валюты, указанный поставщиком, совпадает с курсом валюты,-------
-----------полученным из Scala на эту дату----------------------------------------------------
SELECT @MyInvoiceExRate = SYCH006
FROM SYCH0100
WHERE (SYCH001 = @CurrCode) AND 
	(SYCH004 <= @InvoiceDate) AND 
	(SYCH005 > @InvoiceDate)

IF (ABS(@PurchInvoiceExRate - @MyInvoiceExRate) >= 0.0001) AND (@PurchInvoiceExRate <> 0) --не проверяем, если курса нет
BEGIN
	SELECT @MyRezStr = @MyRezStr + 'В СФ поставщика на дату ' + CONVERT(nvarchar(255),@InvoiceDate,103) + ' указан курс валюты с кодом ' + CONVERT(nvarchar(255),@CurrCode) + ' равный ' + CONVERT(nvarchar(255),@PurchInvoiceExRate) + '. В Scala на эту дату курс для данной валюты равен ' + CONVERT(nvarchar(255),@MyInvoiceExRate) + '.' + CHAR(13)
	GOTO EndProc
END



-----------Курсор по заказам не 0 типа, входящим в состав обобщенного заказа------------------
-----------и содержащих данный код товара поставщика------------------------------------------
DECLARE PO_Cursor CURSOR FOR
SELECT PC010300.PC01001
FROM PC010300 INNER JOIN
	PC030300 ON PC010300.PC01001 = PC030300.PC03001 INNER JOIN
    SC010300 ON PC030300.PC03005 = SC010300.SC01001
WHERE (PC010300.PC01002 <> 0) AND 
	(SC010300.SC01060 = @ItemSuppCode) AND 
	(PC010300.PC01052 = @ConsPOrder) AND 
	((PC030300.PC03010 - PC030300.PC03011) > 0) 
ORDER BY (PC030300.PC03010 - PC030300.PC03011) Desc

OPEN PO_Cursor
FETCH NEXT FROM PO_Cursor INTO @MyPOrder
WHILE @@FETCH_STATUS = 0
BEGIN
	-----------------В цикле обрабатываем все заказы на закупку
	--SELECT @MyRezStr = 'ID строки ' + CONVERT(nvarchar(255),@ID) +  ', Номер СФ ' + @Invoice + ' от ' + CONVERT(nvarchar(255),@InvoiceDate,103) + ', код валюты ' + CONVERT(nvarchar(255),@CurrCode) + ', код продавца ' + @MySalesmanCode + ', ФИО продавца ' + @MySalesmanName + ', курс валюты поставщика ' + CONVERT(nvarchar(255),@PurchInvoiceExRate) + ', Консолидированный заказ на закупку ' + @ConsPOrder + ', код товара поставщика ' + @ItemSuppCode + ', кол - во ' + CONVERT(nvarchar(255), @QTY) + ', сумма без НДС ' + CONVERT(nvarchar(255),@Price) +  ', страна производитель ' + @Country + ', ГТД ' + @GTD  + ', Номер заказа на продажу ' + @MyPOrder + CHAR(13)
	-----------Проверка, что в заказе нет разных запасов с таким кодом запаса поставщика---------
	SELECT @MyCount1 = COUNT(PC03005)
	FROM (SELECT PC030300.PC03005
		FROM PC030300 INNER JOIN
			SC010300 ON PC030300.PC03005 = SC010300.SC01001
		WHERE (PC030300.PC03001 = @MyPOrder) AND 
			(SC010300.SC01060 = @ItemSuppCode)
		GROUP BY PC030300.PC03005) AS View_1

	IF @MyCount1 > 1
	BEGIN
		SELECT @MyRezStr = @MyRezStr + 'В заказе на закупку ' + @MyPOrder + ' запасу с кодом товара поставщика N ' + @ItemSuppCode + ' соответствуют несколько запасов. Произведите выбор и ввод данного запаса в заказ вручную.' + CHAR(13)
		GOTO EndProc1
	END

	-----------Проверка, что нет неотфактурованных накладных по данному заказу--------------------
	SELECT @MyCount1 = COUNT(PC19001)
	FROM PC190300
	WHERE (PC19007 <> 0) AND 
		(PC19012 <> @Invoice) AND 
		(PC19001 = @MyPOrder)

	IF @MyCount1 > 1
	BEGIN
		SELECT @MyRezStr = @MyRezStr + 'Для заказа на закупку ' + @MyPOrder + ' есть неотфактурованные приемки. Сначала отфактуруйте предыдущую приемку. ' + CHAR(13)
		GOTO EndProc1
	END

	------------Определение валюты заказа и курса этой валюты на дату СФ--------------------------
	SELECT @MyCurrCode = PC01022
	FROM PC010300
	WHERE (PC01001 = @MyPOrder)

	SELECT @MyOrderExRate = SYCH006
	FROM SYCH0100
	WHERE (SYCH001 = @MyCurrCode) AND 
		(SYCH004 <= @InvoiceDate) AND 
		(SYCH005 > @InvoiceDate)

	------------Определение курса валюты СФ на дату приемки---------------------------------------
	SELECT @MyInvoiceExRateCD = SYCH006
	FROM SYCH0100
	WHERE (SYCH001 = @CurrCode) AND 
		(SYCH004 <= @InvoiceInputDate) AND 
	(	SYCH005 > @InvoiceInputDate)

	------------Определение курса валюты заказа на дату приемки-----------------------------------
	SELECT @MyOrderExRateCD = SYCH006
	FROM SYCH0100
	WHERE (SYCH001 = @MyCurrCode) AND 
		(SYCH004 <= @InvoiceInputDate) AND 
	(	SYCH005 > @InvoiceInputDate)

	--=====================Изменение закупочной цены==============================================
	------------Проставление закупочной цены за кол - во, указанное в инвойсе---------------------
	------------Получение нашего кода запаса, весов и объема
	SELECT @MyItemCode = PC030300.PC03005
	FROM PC030300 INNER JOIN
		SC010300 ON PC030300.PC03005 = SC010300.SC01001
	WHERE (PC030300.PC03001 = @MyPOrder) AND 
		(SC010300.SC01060 = @ItemSuppCode)
	GROUP BY PC030300.PC03005

	---количество, которое осталось принять в заказ
	SELECT @OrderQTY = PC030300.PC03010 - PC030300.PC03011, 
		@OrderWHQTY = PC030300.PC03044 - PC030300.PC03043, 
		@MySC01140 = CASE WHEN SC010300.SC01140 = 0 THEN 1 ELSE SC010300.SC01140 END, 
		@MySC01141 = CASE WHEN SC010300.SC01141 = 0 THEN 1 ELSE SC010300.SC01141 END,
		@MyCustomPc = SC010300.SC01057 / 100
	FROM PC030300 INNER JOIN
		SC010300 ON PC030300.PC03005 = SC010300.SC01001
	WHERE (PC030300.PC03001 = @MyPOrder) AND 
		(PC030300.PC03005 = @MyItemCode)

	---принимаемое в заказ количество
	SELECT @WHQTY = @InvoiceQTY * @MySC01140 / @MySC01141

	IF @InvoiceQTY > @OrderQTY      --кол - во в СФ больше, чем надо принять в заказ
	BEGIN
		SELECT @OrderQTYRcv = @OrderQTY
		SELECT @OrderWHQTYRcv = @OrderWHQTY
	END
	ELSE
	BEGIN
		SELECT @OrderQTYRcv = @InvoiceQTY
		SELECT @OrderWHQTYRcv = @WHQTY
	END

	------------Проверка, что в заказе только одна строка с таким запасом
	SELECT @MyCount1 = COUNT(PC03005)
	FROM PC030300
	WHERE (PC03001 = @MyPOrder) AND 
		(PC03005 = @MyItemCode)

	IF @MyCount1 > 1
	BEGIN
		SELECT @MyRezStr = @MyRezStr + 'В заказе на закупку ' + @MyPOrder + ' запас с кодом N ' + @MyItemCode + ' присутствует более чем в одной строке. Такой запас необходимо принять на склад вручную.' + CHAR(13)
		GOTO EndProc1
	END

	------------Закупочная цена
	UPDATE PC030300
	SET PC03008 = @Price * @MyInvoiceExRate / @MyOrderExRate / @QTY * @OrderQTYRcv,  --цена за принимаемое кол - во
		PC03019 = @OrderQTYRcv,                                                      --кол-во, за которое проставлена цена
		PC03029 = N'0'                                                               --флаг подтверждения заказа
	WHERE (PC03001 = @MyPOrder) AND 
		(PC03005 = @MyItemCode)

	------------Пересчет общей суммы заказа-------------------------------------------------------
	UPDATE PC010300
	SET PC01020 = View_1.Expr1
	FROM PC010300 INNER JOIN
		(SELECT PC03001, 
			SUM(PC03008 / PC03019 * (PC03010 - PC03011)) AS Expr1
		FROM PC030300
		WHERE (PC03001 = @MyPOrder)
		GROUP BY PC03001) AS View_1 ON 
		PC010300.PC01001 = View_1.PC03001
	WHERE (PC010300.PC01001 = @MyPOrder)

	--=====================Приемка запаса=========================================================

	---Номер внутренней накладной
	SELECT @MyCount1 = COUNT(PC19041)
	FROM PC190300
	WHERE (PC19001 = @MyPOrder) AND 
		(PC19012 = @Invoice)

	IF @MyCount1 = 0
	BEGIN     --берем следующий номер из счетчика
		SELECT @IntInvoice = RIGHT('0000000000' + CONVERT(nvarchar(10), SY68002), 10) 
		FROM SY6803XX
		WHERE (SY68001 = N'PC53')

		UPDATE SY6803XX
		SET SY68002 = SY68002 + 1
		WHERE (SY68001 = N'PC53')
		
	END
	ELSE
	BEGIN     --используем уже существующий для данной СФ для данного заказа
		SELECT @IntInvoice = PC19041
		FROM PC190300
		WHERE (PC19001 = @MyPOrder) AND 
			(PC19012 = @Invoice)
		GROUP BY PC19041
	END

	---Номер строки и подстроки заказа на закупку
	SELECT @MyStrNum = PC03002, 
		@MySubStrNum = PC03003
	FROM PC030300
	WHERE (PC03001 = @MyPOrder) AND 
		(PC03005 = @MyItemCode)

	---Номер автостроки (счетчик строк)
	SELECT @MyAutStrNum = MAX(PC19004)
	FROM PC190300
	WHERE (PC19001 = @MyPOrder) AND 
		(PC19002 = @MyStrNum) AND 
		(PC19003 = @MySubStrNum)

	IF @MyAutStrNum IS NULL
	BEGIN     --первая
		SELECT @MyAutStrNum = '0010'
	END
	ELSE
	BEGIN     --увеличиваем на 10
		SELECT @MyAutStrNum = RIGHT('0000' + CONVERT(nvarchar(4),(CONVERT(int,@MyAutStrNum) / 10 + 1) * 10),4)
	END

	---Номер новой партии
	SELECT @MyBatchNum = RIGHT('000000000000' + CONVERT(nvarchar(12), SY68002), 12) 
	FROM SY6803XX
	WHERE (SY68001 = N'SC28')

	UPDATE SY6803XX
	SET SY68002 = SY68002 + 1
	WHERE (SY68001 = N'SC28')

	---Курс обмена - 2 и алгоритм обмена
	SELECT @MyExchRate2     = CASE WHEN @MyCurrCode = 0 THEN 1   WHEN @MyCurrCode = 12 THEN 0   ELSE 0 END
	SELECT @MyExchAlgorithm = CASE WHEN @MyCurrCode = 0 THEN '*' WHEN @MyCurrCode = 12 THEN '1' ELSE '1' END 
											--RUR                         --EUR                     --остальное - USD

	---Вес и объем в единицах закупки
	SELECT  @MyNetWeight = SC01069 * @MySC01140 / @MySC01141, 
		@MyGrossWeight = SC01070 * @MySC01140 / @MySC01141, 
		@MyVolume = SC01071 * @MySC01140 / @MySC01141
	FROM SC010300
	WHERE (SC01001 = @MyItemCode)

	-----------PC190300---------------------------------------------------------------------------
	INSERT INTO PC190300 (PC19001,   PC19002,   PC19003,      PC19004,      PC19005,     PC19006,    PC19007,      PC19008,    PC19009,           PC19010,           PC19011,           PC19012,  PC19013,        PC19014,                                                          PC19015, PC19016, PC19017, PC19018, PC19019, PC19020, PC19021, PC19022, PC19023, PC19024, PC19025,      PC19026,        PC19027,   PC19028, PC19029, PC19030, PC19031, PC19032, PC19033,                                                                                                                                       PC19034,     PC19035, PC19036,         PC19037, PC19038,      PC19039,          PC19040,     PC19041,     PC19042, PC19043, PC19044, PC19045, PC19046,          PC19047, PC19048, PC19049, PC19050, PC19051, PC19052, PC19053, PC19054, PC19055, PC19056, PC19057, PC19058, PC19059, PC19060, PC19061, PC19062)
	VALUES               (@MyPOrder, @MyStrNum, @MySubStrNum, @MyAutStrNum, @MyBatchNum, 0.00000000, @OrderQTYRcv, 0.00000000, @InvoiceInputDate, @InvoiceInputDate, @InvoiceInputDate, @Invoice, @MyOrderExRate, @Price * @MyInvoiceExRate / @MyOrderExRate / @QTY * @OrderQTYRcv, 0,       0,       0,       0,       0,       0,       0,       0,       0,       0,       @MyNetWeight, @MyGrossWeight, @MyVolume, N'',     N'',     N'',     N'',     0,       CONVERT(nvarchar(255),DATEPART(hh,GETDATE())) + CONVERT(nvarchar(255),DATEPART(mm,GETDATE())) + CONVERT(nvarchar(255),DATEPART(ss,GETDATE())), @MyCustomPc, 0,       @OrderWHQTYRcv,  0,       @MyExchRate2, @MyExchAlgorithm, @MyCurrCode, @IntInvoice, N'',     N'',     N'',     12,      @MyOrderExRateCD, 0,       N'1',    0,       0,       0,       N'',     0,       N'',     N'',     N'',     N'',     N'',     N'',     N'',     N'',     0)

	-----------маленькое шаманство - подправляем цены во всех записях по данной СФ по данному запасу
	-----------иначе Scala бузит и генерит кучу проводок
	UPDATE PC190300
	SET PC19014 = @Price * @MyInvoiceExRate / @MyOrderExRate / @QTY * @OrderQTYRcv
	WHERE (PC19001 = @MyPOrder) AND 
		(PC19002 = @MyStrNum) AND 
		(PC19012 = @Invoice)

	-----------SC330300---------------------------------------------------------------------------
	---код поставщика и номер склада
	SELECT @MySuppCode = PC01003,
		@MyWHNum = PC01023
	FROM PC010300
	WHERE (PC01001 = @MyPOrder)

	---Разбор ГТД
	SELECT @MyFirstPos = CHARINDEX('/', @GTD, 0) + 1
	SELECT @MyFirstPos = CHARINDEX('/', @GTD, @MyFirstPos)
	SELECT @MySecondPos = CHARINDEX('/', @GTD, @MyFirstPos + 1)

	SELECT @MyGTDSuppDate = LEFT (@GTD,@MyFirstPos - 1)
	SELECT @MyGTDNum = SUBSTRING(@GTD,@MyFirstPos, @MySecondPos - @MyFirstPos)
	SELECT @MyGTD_GTD = RIGHT(@GTD,LEN(@GTD) - @MySecondPos + 1)

	---страна производитель
	---из SY240300
	SELECT @MyCountryCode = SY24002
	FROM SY240300
	WHERE (SY24001 = N'BM') AND 
		(SY24003 = @Country)

	IF @MyCountryCode IS NULL
	BEGIN
		---из tbl_CountryNameAndCode
		SELECT @MyCountryCode = Code
		FROM tbl_CountryNameAndCode
		WHERE (Name = @Country)

		IF @MyCountryCode IS NULL
		BEGIN
			SELECT @MyCountryCode = ''
			SELECT @MyRezStr = @MyRezStr + 'В заказе на закупку ' + @MyPOrder + ' запас с кодом N ' + @MyItemCode + ' страна производитель ' + @Country + ' не найдена ни в Scala, ни в дополнительной таблице со странами производителями. ' + CHAR(13)
		END
	END

	---ячейка хранения по умолчанию
	SELECT @MyCell = SC03013
	FROM SC030300
	WHERE (SC03001 = @MyItemCode) AND 
		(SC03002 = @MyWHNum)

	INSERT INTO SC330300  (SC33001,     SC33002,  SC33003,     SC33004, SC33005,         SC33006, SC33007, SC33008,         SC33009,     SC33010, SC33011, SC33012, SC33013,           SC33014,      SC33015,    SC33016,    SC33017, SC33018,     SC33019,   SC33020, SC33021, SC33022, SC33023,                                                    SC33024,        SC33025,        SC33026,    SC33027,   SC33028, SC33029, SC33030, SC33031, SC33032, SC33033, SC33034,                                                                    SC33035,      SC33036, SC33037, SC33038, SC33039, SC33040, SC33041, SC33042, SC33043, SC33044, SC33045,    SC33046, SC33047, SC33048, SC33049, SC33050,           SC33051, SC33052,      SC33053, SC33054)
	VALUES                (@MyItemCode, @MyWHNum, @MyBatchNum, @MyCell, @OrderWHQTYRcv,  0,       0,       @OrderWHQTYRcv,  @MyBatchNum, 0,       N'',     N'0',    @InvoiceInputDate, @MyStartDate, @MyFinDate, @MyFinDate, 0x0,     @MySuppCode, @MyPOrder, N'',     N'',     N'',     @Price * @MyInvoiceExRate / @QTY / @MySC01140 * @MySC01141, @MyGTDSuppDate, @MyCountryCode, @MyGTD_GTD, @MyGTDNum, N'',     N'',     N'',     N'',     N'',     N'',     @Price * @MyInvoiceExRate /@MyOrderExRate / @QTY / @MySC01140 * @MySC01141, @MyStartDate, N'',     0,       N'00',   0,       N'1',    N'',     1,       0,       N'',     @MyFinDate, N'',     N'00',   N'',     0,       @InvoiceInputDate, N'',     @MyStartDate, N'',     N'')

	-----------SYVT0300---------------------------------------------------------------------------
	---Поле SYVT002
	SELECT @MySYVT002 = Left(@MyItemCode + '                                   ',35) + Left(@MyWHNum + '      ',6) + @MyBatchNum

	INSERT INTO SYVT0300 (SYVT001,     SYVT002,    SYVT003,    SYVT004,      SYVT005,      SYVT006, SYVT007, SYVT008, SYVT009, SYVT010, SYVT011, SYVT012)
	VALUES               (N'SC330300', @MySYVT002, N'SC33055', @MyStartDate, @MyStartDate, N'',     0,       0,       0,       2,       0,       0)

	-----------SC070300---------------------------------------------------------------------------
	-----------Основная запись
	---Account поставщика по умолчанию
	SELECT @MyAccountDef = PL01043
	FROM PL010300
	WHERE (PL01001 = @MySuppCode)

	---SearchIndex
	SELECT @MySearchIndex = Right('00000000' + CONVERT(nvarchar(12), SY68002),8)
	FROM SY6803XX
	WHERE (SY68001 = N'SC07AUTO')

	UPDATE SY6803XX
	SET SY68002 = SY68002 + 1
	WHERE (SY68001 = N'SC07AUTO')

	---пользователь
	SELECT @MyUserID = ScalaSystemDB.dbo.ScaUserProperty.Value
	FROM ScalaSystemDB.dbo.ScaUsers INNER JOIN
		ScalaSystemDB.dbo.ScaUserProperty ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserProperty.UserID
	WHERE (ScalaSystemDB.dbo.ScaUsers.FullName = @MySalesmanName) AND 
		(ScalaSystemDB.dbo.ScaUserProperty.PropertyID = 5)

	INSERT INTO SC070300 (SC07001, SC07002,           SC07003,     SC07004,         SC07005,                                                    SC07006,     SC07007,   SC07008, SC07009,  SC07010,                                                    SC07011, SC07012,       SC07013, SC07014, SC07015, SC07016, SC07017, SC07018, SC07019,      SC07020, SC07021,     SC07022,        SC07023,     SC07024,        SC07025, SC07026, SC07027, SC07028,                                                                     SC07029,     SC07030, SC07031,     SC07032, SC07033, SC07034, SC07035,           SC07036,                                                                                                                                       SC07037,   SC07038, SC07039, SC07040, SC07041, SC07042, SC07043, SC07044,      SC07045,          SC07046, SC07047, SC07048, SC07049, SC07050, SC07051, SC07052, SC07053, SC07054, SC07055, SC07056, SC07057, SC07058, SC07059)
	VALUES               (N'00',   @InvoiceInputDate, @MyItemCode, @OrderWHQTYRcv,  @Price * @MyInvoiceExRate / @QTY / @MySC01140 * @MySC01141, @MySuppCode, @MyPOrder, N'01',   @MyWHNum, @Price * @MyInvoiceExRate / @QTY / @MySC01140 * @MySC01141, N'00',   @MyAccountDef, 0x0,     N'',     N'',     N'',     N'0',    N'',     @MyStartDate, N'',     @MyBatchNum, @MySearchIndex, @MyCurrCode, @MyOrderExRate, N'0',    @MyCell, N'',     @Price * @MyInvoiceExRate / @MyOrderExRate / @QTY / @MySC01140 * @MySC01141, @MyBatchNum, N'',     @MyBatchNum, N'',     0,       N'',     @InvoiceInputDate, CONVERT(nvarchar(255),DATEPART(hh,GETDATE())) + CONVERT(nvarchar(255),DATEPART(mm,GETDATE())) + CONVERT(nvarchar(255),DATEPART(ss,GETDATE())), @MyUserID, N'',     N'',     N'',     N'',     N'',     N'',     @MyExchRate2, @MyExchAlgorithm, N'',     N'00',   N'',     N'0',    N'0',    N'',     N'',     N'',     N'',     N'',     N'',     N'',     N'',     N'')


	IF @MyCustomPc <> 0
	BEGIN
		-----------таможенный сбор
		---SearchIndex
		SELECT @MySearchIndex = CONVERT(nvarchar(12), SY68002)
		FROM SY6803XX
		WHERE (SY68001 = N'SC07AUTO')

		UPDATE SY6803XX
		SET SY68002 = SY68002 + 1
		WHERE (SY68001 = N'SC07AUTO')

		INSERT INTO SC070300 (SC07001, SC07002,           SC07003,     SC07004, SC07005,                                                         SC07006,     SC07007,   SC07008, SC07009,   SC07010, SC07011, SC07012,       SC07013, SC07014, SC07015, SC07016, SC07017, SC07018, SC07019,      SC07020, SC07021,     SC07022,        SC07023,     SC07024,        SC07025, SC07026,  SC07027, SC07028,                                                                        SC07029,     SC07030, SC07031,     SC07032, SC07033, SC07034, SC07035,           SC07036,                                                                                                                                       SC07037,   SC07038, SC07039, SC07040, SC07041, SC07042, SC07043, SC07044,      SC07045,          SC07046, SC07047, SC07048, SC07049, SC07050, SC07051, SC07052, SC07053, SC07054, SC07055, SC07056, SC07057, SC07058, SC07059)
		VALUES               (N'03',   @InvoiceInputDate, @MyItemCode, 0,       @Price * @MyInvoiceExRate / @QTY * @OrderQTYRcv * @MyCustomPc,   @MySuppCode, @MyPOrder, N'01',   @MyWHNum,  0,       N'00',   @MyAccountDef, 0x0,     N'',     N'',     N'',     N'1',    N'',     @MyStartDate, N'',     @MyBatchNum, @MySearchIndex, @MyCurrCode, @MyOrderExRate, N'0',    @MyCell,  N'',     @Price * @MyInvoiceExRate / @QTY * @OrderQTYRcv * @MyCustomPc / @MyOrderExRate, @MyBatchNum, N'',     @MyBatchNum, N'',     0,       N'',     @InvoiceInputDate, CONVERT(nvarchar(255),DATEPART(hh,GETDATE())) + CONVERT(nvarchar(255),DATEPART(mm,GETDATE())) + CONVERT(nvarchar(255),DATEPART(ss,GETDATE())), @MyUserID, N'',     N'',     N'',     N'',     N'',     N'',     @MyExchRate2, @MyExchAlgorithm, N'',     N'00',   N'',     N'0',    N'0',    N'',     N'',     N'',     N'',     N'',     N'',     N'',     N'',     N'')

	END


	-----------SC340300---------------------------------------------------------------------------
	INSERT INTO SC340300 (SC34001,     SC34002,     SC34003,  SC34004,                                                    SC34005)
	VALUES               (@MyItemCode, @MyBatchNum, @MyWHNum, @Price * @MyInvoiceExRate / @QTY / @MySC01140 * @MySC01141, @Price * @MyInvoiceExRate / @MyOrderExRate / @QTY / @MySC01140 * @MySC01141)


	-----------SC110300---------------------------------------------------------------------------
	---Уменьшаем на количество принятого в ед. измерения хранения
	UPDATE SC110300
	SET SC11005 = SC11005 - @OrderWHQTYRcv
	WHERE (SC11003 = @MyPOrder) AND 
		(SC11001 = @MyItemCode)

	---Удаляем, если в результате 0
	DELETE FROM SC110300
	WHERE (SC11003 = @MyPOrder) AND 
		(SC11001 = @MyItemCode) AND 
		(SC11005 = 0)


	-----------SC030300---------------------------------------------------------------------------
	---FIFO Value
	SELECT @MyFIFOValue = CASE WHEN SUM(SC33005) = 0 THEN 0 ELSE SUM(SC33023 * SC33005) / SUM(SC33005) END 
	FROM SC330300
	WHERE  (SC33001 = @MyItemCode) AND 
		(SC33002 = @MyWHNum) AND 
		(SC33005 <> 0)

	---AVGCostPrice2
	SELECT @MyAVGCostPrice2 = CASE WHEN SUM(SC07004) = 0 THEN 0 ELSE SUM(CASE WHEN SC07004 = 0 THEN CASE SC07001 WHEN 09 THEN 0 WHEN 07 THEN - SC07005 ELSE SC07005 END ELSE SC07004 * SC07005 END) / SUM(SC07004) END
	FROM SC070300
	WHERE (SC07003 = @MyItemCode) AND 
		(SC07009 = @MyWHNum)

	UPDATE SC030300
	SET  SC03003 = SC03003 + @OrderWHQTYRcv,      --новый баланс
		SC03006 = SC03006 - @OrderWHQTYRcv,       --остаток в заказах (непринято)
		SC03029 = @MyFIFOValue,                   --FIFO Value
		SC03054 = @InvoiceInputDate,              --дата последней поставки
		SC03057 = @MyAVGCostPrice2                --AVGCostPrice2
	WHERE (SC03001 = @MyItemCode) AND 
		(SC03002 = @MyWHNum)

	-----------SC010300---------------------------------------------------------------------------
	---@MyAVGCostPrice1 (по 1 складу)
	SELECT @MyAVGCostPrice1 = CASE WHEN SUM(SC07004) = 0 THEN 0 ELSE SUM(CASE WHEN SC07004 = 0 THEN CASE SC07001 WHEN 09 THEN 0 WHEN 07 THEN - SC07005 ELSE SC07005 END ELSE SC07004 * SC07005 END) / SUM(SC07004) END
	FROM SC070300
	WHERE (SC07003 = @MyItemCode) AND 
		(SC07009 = N'01')

	UPDATE SC010300
	SET SC01042 = SC01042 + @OrderWHQTYRcv,       --новый баланс
		SC01045 = SC01045 - @OrderWHQTYRcv,       --остаток в заказах (непринято)
		SC01049 = @InvoiceInputDate,              --дата последней поставки
		SC01052 = @MyAVGCostPrice1,               --AVGCostPrice1
		SC01054 = SC01054 + @OrderWHQTYRcv        --QTYRecSuppl
	WHERE     (SC01001 = @MyItemCode)

	-----------PC030300---------------------------------------------------------------------------
	UPDATE PC030300
	SET PC03011 = PC03011 + @OrderQTYRcv,         --увеличение на кол-во принятого в ед. закупки 
		PC03038 = PC03038 + @OrderQTYRcv,		  --увеличение на кол-во принятого в ед. закупки 
		PC03043 = PC03043 + @OrderWHQTYRcv,       --увеличение на кол-во принятого в ед. хранения
		PC03017 = @InvoiceInputDate,
		PC03024 = @InvoiceInputDate
	WHERE (PC03001 = @MyPOrder) AND 
		(PC03005 = @MyItemCode)

	------------Пересчет общей суммы заказа-------------------------------------------------------
	UPDATE PC010300
	SET PC01020 = View_1.Expr1,
		PC01016 = @InvoiceInputDate,
		PC01026 = '1',
		PC01010 = '1',
		PC01028 = 1
	FROM PC010300 INNER JOIN
		(SELECT PC03001, 
			SUM(PC03008 / PC03019 * (PC03010 - PC03011)) AS Expr1
		FROM PC030300
		WHERE (PC03001 = @MyPOrder)
		GROUP BY PC03001) AS View_1 ON 
		PC010300.PC01001 = View_1.PC03001
	WHERE (PC010300.PC01001 = @MyPOrder)

	--=====================вызов ф-ции автораспределения==========================================
	---Номер заказа на продажу для данного заказа на закупку
	SELECT @MySalesOrder = PC01060
	FROM PC010300
	WHERE (PC01001 = @MyPOrder)

	---проверяем, что естть такой заказ на продажу ненулевого типа
	SELECT @MyCount1 = COUNT(OR01001)
	FROM OR010300
	WHERE (OR01002 <> 0) AND 
		(OR01001 = @MySalesOrder)

	IF @MyCount1 = 0 
	BEGIN  ---Вызываем ф-цию автораспределения для данного заказа на закупку
		EXEC spp_AutoAllocation_AllocateItems_MASS @MyPOrder
	END
	ELSE
	BEGIN  ---Вызываем ф-цию автораспределения для данного з-за на закупку, данного з-за на продажу, данного инвойса
		EXEC spp_AutoAllocation_AllocateItems2 @MySalesOrder, @MyPOrder, @Invoice
	END

	-------Уменьшаем количество на принятую величину
	SELECT @InvoiceQTY = @InvoiceQTY - @OrderQTYRcv

	EndProc1:
	FETCH NEXT FROM PO_Cursor INTO @MyPOrder
END
CLOSE PO_Cursor
DEALLOCATE PO_Cursor

EndProc:

Select @MyRestQTY = @InvoiceQTY
