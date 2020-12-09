﻿DECLARE @Year int;
SET @Year = {0};

IF OBJECT_ID('tempdb.dbo.#StudentsBaseTable',    'U') IS NOT NULL DROP TABLE dbo.#StudentsBaseTable;
IF OBJECT_ID('tempdb.dbo.#TeacherBaseTable', 'U') IS NOT NULL DROP TABLE dbo.#TeacherBaseTable;

select 
'ISCED ' + SUBSTRING(DL.[ISCED Level], 6, 1) as ISCED_TOP, 
CASE DL.[ISCED SubClass] 
  WHEN 'ISCED01' THEN 'ISCED 01'
  WHEN 'ISCED02' THEN 'ISCED 02'
  WHEN 'ISCED1' THEN 'ISCED 1'
  WHEN 'ISCED24' THEN 'ISCED 24'
  WHEN 'ISCED34' THEN 'ISCED 34'
END as ISCED,
CASE DA.AuthorityGroup 	WHEN 'Government' THEN 'Public'	ELSE 'Private' END as SCHOOLTYPE,
E.GenderCode as GENDER,
ISNULL(E.Age, -1) as AGE,
DL.[Year of Education] as CLASS,
sum(ISNULL(E.Rep, 0)) as REPEATER,
sum(ISNULL(E.PSA, 0)) as ECE,
sum(E.Enrol) as 'COUNT'

into dbo.#StudentsBaseTable 

from warehouse.tableEnrol E
left join DimensionAuthority DA on E.AuthorityCode = DA.AuthorityCode
left join DimensionLevel DL on E.ClassLevel = DL.LevelCode
where E.SurveyYear = @year AND E.Enrol IS NOT NULL

group by DL.[ISCED Level], DL.[ISCED SubClass], DA.AuthorityGroup, E.GenderCode, E.Age, DL.[Year of Education]
--Teachers

select
	CASE t.ISCEDSubClass
	    WHEN 'ISCED01' THEN 'ISCED 01'
		WHEN 'ISCED02' THEN 'ISCED 02'
		WHEN 'ISCED1' THEN 'ISCED 1'
		WHEN 'ISCED24' THEN 'ISCED 24'
		WHEN 'ISCED34' THEN 'ISCED 34'
	END as ISCED,
	CASE  
		WHEN da.AuthorityGroupCode = 'G' THEN 'PUBLIC'
		WHEN da.AuthorityGroupCode = 'N' THEN 'PRIVATE'
	END as SCHOOLTYPE,
	t.GenderCode as GENDER,
	'FULLTIME' as tchFullPart,
	sum(t.Certified) as TRAINED,
	sum(t.Qualified) as QUALIFIED,
	sum(t.NumTeachers) as COUNT

into dbo.#TeacherBaseTable 

from warehouse.TeacherCountSchool t
left join DimensionAuthority da on da.AuthorityCode = t.AuthorityCode
where t.SurveyYear = @year AND t.GenderCode IS NOT NULL
group by t.ISCEDSubClass, da.AuthorityGroupCode, t.GenderCode
