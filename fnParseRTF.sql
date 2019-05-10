USE [<DBname>]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/******************************  
** Name: fnParseRTF
** Desc: Modification of https://www.codeproject.com/Tips/821281/Convert-RTF-to-Plain-Text-Revised-Again
to use a tally table instead of numbers in [master]..[spt_values]
** Auth: NightWizzard?
** Date: 4/9/2016
**************************/
CREATE FUNCTION [dbo].[fnParseRTF] (@rtf NVARCHAR(max))
RETURNS NVARCHAR(max)
AS
BEGIN
	DECLARE @Pos1 INT;
	DECLARE @Pos2 INT;
	DECLARE @hex VARCHAR(316);
	DECLARE @Stage TABLE (
		[Char] CHAR(1)
		,[Pos] INT
		);

	WITH Tally (Number)
	AS (
		-- 10000 rows
		SELECT ROW_NUMBER() OVER (
				ORDER BY (
						SELECT NULL
						)
				)
		FROM (
			VALUES (0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
			) a(Number)
		CROSS JOIN (
			VALUES (0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
			) b(Number)
		CROSS JOIN (
			VALUES (0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
			) c(Number)
		CROSS JOIN (
			VALUES (0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
				,(0)
			) d(Number)
		)
	INSERT @Stage (
		[Char]
		,[Pos]
		)
	/*
	ORIGINAL METHOD FOR STAGING CHARS USES STATIC NUMBERING IN MASTER..SPT_VALUES
	THIS MAXES OUT AT 2047 CHARACTERS, SO I UPDATED WITH A 10K ROW TALLY TABLE.
	ADDITIONAL CROSS JOINS CAN BE ADDED DEPENDING ON THE MAX LENGTH OF THE RTF DATA
	*/
	SELECT SUBSTRING(@rtf, Number, 1)
		,Number
	FROM Tally
	WHERE SUBSTRING(@rtf, Number, 1) IN (
			'{'
			,'}'
			)

	SELECT @Pos1 = MIN([Pos])
		,@Pos2 = MAX([Pos])
	FROM @Stage;

	DELETE
	FROM @Stage
	WHERE (
			[Pos] IN (
				@Pos1
				,@Pos2
				)
			);

	/* 
	check if the doc is already plaintext. if so, return the body. 
	if not, jump to the branch containing parsing code
	*/
	BEGIN
		SET @Pos1 = PATINDEX('{\rtf1%', @rtf)

		IF @Pos1 > 0
			GOTO Conversion_Branch;
		ELSE
			RETURN @rtf;
	END

	Conversion_Branch:

	WHILE (1 = 1)
	BEGIN
		SELECT TOP 1 @Pos1 = s1.[Pos]
			,@Pos2 = s2.[Pos]
		FROM @Stage s1
		INNER JOIN @Stage s2 ON s2.[Pos] > s1.[Pos]
		WHERE (s1.[Char] = '{')
			AND (s2.[Char] = '}')
		ORDER BY s2.[Pos] - s1.[Pos];

		IF @@ROWCOUNT = 0
			BREAK

		DELETE
		FROM @Stage
		WHERE (
				[Pos] IN (
					@Pos1
					,@Pos2
					)
				);

		UPDATE @Stage
		SET [Pos] = [Pos] - @Pos2 + @Pos1 - 1
		WHERE ([Pos] > @Pos2);

		SET @rtf = STUFF(@rtf, @Pos1, @Pos2 - @Pos1 + 1, '');
	END

	SET @rtf = REPLACE(@rtf, '\pard', ' ');
	SET @rtf = REPLACE(@rtf, '\par', ' ');
	SET @rtf = STUFF(@rtf, 1, CHARINDEX(' ', @rtf), '');

	--WHILE (Right(@rtf, 1) IN (' ', CHAR(13), CHAR(10), '}'))
	WHILE (Right(@rtf, 1) IN ('}'))
	BEGIN
		SELECT @rtf = SUBSTRING(@rtf, 1, (LEN(@rtf + 'x') - 2));

		IF LEN(@rtf) = 0
			BREAK
	END

	-- \pict processing code
	SET @Pos1 = PATINDEX('%{\pict%', @rtf);-- find the position of a picture in the rtf data

	WHILE @Pos1 > 0
	BEGIN
		IF @Pos1 > 0
		BEGIN
			SET @Pos2 = CHARINDEX('}', @rtf, @Pos1);-- find the closing brace for the \pict data
			SET @rtf = STUFF(@rtf, @Pos1, (@Pos2 - @Pos1) + 1, '');
			SET @Pos1 = PATINDEX('%{\pict%', @rtf);-- reset the index of the opening brace for additional \pict data 
		END
	END

	SET @Pos1 = CHARINDEX('\''', @rtf);

	WHILE @Pos1 > 0
	BEGIN
		IF @Pos1 > 0
		BEGIN
			SET @hex = '0x' + SUBSTRING(@rtf, @Pos1 + 2, 2);
			SET @rtf = REPLACE(@rtf, SUBSTRING(@rtf, @Pos1, 4), CHAR(CONVERT(INT, CONVERT(BINARY (1), @hex, 1))));
			SET @Pos1 = CHARINDEX('\''', @rtf);
		END
	END

	SET @rtf = @rtf + ' ';
	SET @Pos1 = PATINDEX('%\%[0123456789][\ ]%', @rtf);

	WHILE @Pos1 > 0
	BEGIN
		SET @Pos2 = CHARINDEX(' ', @rtf, @Pos1 + 1);

		IF @Pos2 < @Pos1
			SET @Pos2 = CHARINDEX('\', @rtf, @Pos1 + 1);

		IF @Pos2 < @Pos1
		BEGIN
			SET @rtf = SUBSTRING(@rtf, 1, @Pos1 - 1);
			SET @Pos1 = 0;
		END
		ELSE
		BEGIN
			SET @rtf = STUFF(@rtf, @Pos1, @Pos2 - @Pos1 + 1, '');
			SET @Pos1 = PATINDEX('%\%[0123456789][\ ]%', @rtf);
		END
	END

	RETURN @rtf;
END
GO