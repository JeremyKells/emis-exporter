DECLARE @Year int;
SET @Year = {0};

IF OBJECT_ID('tempdb.dbo.#StudentsBaseTable',    'U') IS NOT NULL DROP TABLE dbo.#StudentsBaseTable;
IF OBJECT_ID('tempdb.dbo.#TeacherBaseTable', 'U') IS NOT NULL DROP TABLE dbo.#TeacherBaseTable;

select 
	d.ISCED_TOP,
	d.ISCED,
	d.SCHOOLTYPE,
	d.GENDER,
	d.AGE,
	d.CLASS,
	SUM(d.REPEATERS) as REPEATER,  -- Subset of Count
	SUM(d.ECE) as ECE,             -- Subset of Count
	SUM(d.TOTAL) as 'COUNT'

into dbo.#StudentsBaseTable 

from (
	SELECT
		CASE 
			WHEN E.ClassLevel like 'P%' THEN 'ISCED 1'
			WHEN E.ClassLevel like 'JS%' THEN 'ISCED 2'
			WHEN E.ClassLevel like 'SS%' THEN 'ISCED 3'
		END as ISCED_TOP,
		CASE 
			WHEN E.ClassLevel like 'P%' THEN 'ISCED 1'
			WHEN E.ClassLevel like 'JS%' THEN 'ISCED 24'
			WHEN E.ClassLevel like 'SS%' THEN 'ISCED 34'
		END as ISCED, 
		CASE WHEN S.schAuth = 'MoE' THEN 'Public' ELSE 'Private' END as SCHOOLTYPE,
		E.GenderCode as GENDER,
		E.Age as AGE,
		L.lvlYear as 'CLASS',
		CASE WHEN Rep is Null THEN 0 else Rep END as REPEATERS,

		--E.Enrol - ISNULL(Rep, 0) as NOTREPEATERS
		E.Enrol as TOTAL
		,ISNULL(PTG.pt, 0) as ECE
		from warehouse.enrol E
		left join Schools S on E.schNo = S.schNo 
		left join lkpLevels L on E.ClassLevel = L.codeCode
		left join warehouse.PupilTablesG PTG 
			on  E.ClassLevel = PTG.ClassLevel 
			and E.schNo = PTG.schNo 
			and E.Age = PTG.Age 
			and E.GenderCode = PTG.genderCode
			and E.schNo = PTG.schNo 
			and E.surveyYear = PTG.SurveyYear
		where E.surveyYear = @Year
		and E.Enrol is not NULL
		and (PTG.ptCode = 'PSA' or PTG.ptCode is Null)
) d
group by 	d.ISCED_TOP,
	d.ISCED,
	d.SCHOOLTYPE,
	d.GENDER,
	d.AGE,
	d.CLASS

--Teachers



select 
	CASE 
		WHEN TSI.TSISCED = 'ISCED1' THEN 'ISCED 1'
		WHEN TSI.TSISCED = 'ISCED2A' THEN 'ISCED 24'
		WHEN TSI.TSISCED = 'ISCED3A' THEN 'ISCED 34'
	END as ISCED, 
	CASE 
		WHEN DA.AuthorityGroup = 'Government' THEN 'PUBLIC'
		WHEN DA.AuthorityGroup = 'Non-government' THEN 'PRIVATE'
	END as SCHOOLTYPE, 
	CASE WHEN TS.tchGender = 'M' THEN 'M' ELSE 'F' END as GENDER, -- Nulls become 'F'
	--TS.tchFullPart, 
	TQC.Certified as TRAINED, 
	TQC.Qualified as QUALIFIED,
	COUNT(1) as COUNT

into dbo.#TeacherBaseTable 

from dbo.tfnESTIMATE_BestSurveyEnrolments() EE
	inner join SchoolSurvey SS on EE.bestssID = SS.ssID
	inner join DimensionAuthority DA on SS.ssAuth = DA.AuthorityCode 
	join [TeacherSurvey] TS on EE.bestssID = TS.ssID
	join [TeacherSurveyISCED] TSI on TS.tchsID = TSI.tchsID
	join [tchsIDQualifiedCertified] TQC on TQC.tchsID = TS.tchsID
where lifeyear  = @Year
group by 
	DA.AuthorityGroup, 
	TS.tchFullPart, 
	TS.tchGender, 
	TSI.TSISCED, 
	TQC.Certified, 
	TQC.Qualified