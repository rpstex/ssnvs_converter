SET NOCOUNT ON;
DECLARE @PATH NVARCHAR(MAX) = 'C:\users\user\Desktop\ssnvs_tool\lastInput.xlsx'
DECLARE @SQL NVARCHAR(MAX) = 
'SELECT
--9 char, pos 1-9: SSN including leading zeros.
REPLACE(SSN,''-'','''')+
--3 char, pos 10-12: ENTRY CODE "TPV"
''TPV''+
--3 char, position 13-15: Processing Code 214
    ''214''+
    --13 char, pos 16-28: Last Name. Do not use hyphens, apostrophes, spaces, periods, suffixes (Jr) or prefixes (Dr). Must contain at least one character.
    LEFT(([Last Name]) + space(13), 13)+
    --10 char, pos 29-38: First Name. Do not use hyphens, apostrophes, spaces, periods, suffixes (Jr) or prefixes (Dr). Must contain at least one character.
    LEFT(([First Name]) + space(10), 10)+
    --7 char, pos 39-45: Middle Name, optional.
    space(7)+  
    --8 char, pos 46-53: Date of birth, optional. MMDDYYYY format. (remove the ''faked'' dates and replace them with blank spaces)
	--refactor the line below! it was commented out because the birth date values were not always valid.
    --case when [Birth Date] = ''19000101'' THEN space(8) ELSE REPLACE(CONVERT (VARCHAR(10), [Birth Date], 101), ''/'', '''')END+
	space(8)+
    --36 char, pos 54-89: BLANK SSA use only.
    space(36)+
    --14 char, pos 90-103: User Data Control. Freeform alphanumeric test for employer use.
    LEFT(cast([employee number] as varchar(14)) + space(14), 14)+
    --20 char, pos 104-123: BLANK SSA use only
    space(20)+
    --4 char, pos 124-127: Requester Identification Code. Enter OEVS.
    ''OEVS''+
    --3 chr, pos 128-130: Multiple Request Indicator (must insert "000").
    ''000''
	--the path below is hard coded. this needs to be refactored.
	--also the excel *sheet* name is hard coded to AdvancedSearch which needs to be refactored.
FROM OPENROWSET(
				''Microsoft.ACE.OLEDB.12.0'',
				''Excel 12.0 Xml;HDR=YES;Database=' + @PATH + ''',
				''SELECT * FROM [AdvancedSearch$]''
				)'
EXEC(@SQL)
--:out "c:/folder/somefile.txt"
--Unclosed quotation mark after the character string '','SELECT * FROM [AdvancedSearch$]')'.
--Msg 102, Level 15, State 1, Server TLN-CH-147, Line 34
--Incorrect syntax near '','SELECT * FROM [AdvancedSearch$]')'.
--FROM OPENROWSET(
--	'Microsoft.ACE.OLEDB.12.0',
--	'Excel 12.0 Xml;HDR=YES;Database=C:\users\user\Desktop\ssnvs_tool\lastInput.xlsx',
--	'SELECT * FROM [AdvancedSearch$]'
--)
