IF OBJECT_ID('tempdb.dbo.#TeacherBaseTable', 'U') IS NOT NULL DROP TABLE dbo.#TeacherBaseTable;
go
CREATE TABLE dbo.#TeacherBaseTable (
                                    ISCED varchar(300), 
                                    SCHOOLTYPE Varchar(1000), 
                                    GENDER varchar(200), 
                                    QUALIFIED varchar(10), 
                                    TRAINED varchar(10), 
                                    COUNT int)
go

insert into dbo.#TeacherBaseTable (ISCED, SCHOOLTYPE, GENDER, QUALIFIED, TRAINED, COUNT) 
select
	LEVEL.ISCED,
	SCHOOLTYPE = CASE WHEN school_type = 1 THEN 'PUBLIC' WHEN school_type = 2 THEN 'PRIVATE' END,
	GENDER,
	ISNULL(teaching_qual, 'N') as QUALIFIED,
	'N' as TRAINED,
	count(1) as COUNT
from TGCA
left outer join STAFF on TGCA.staff_id = STAFF.staff_id
left outer join SCHOOLS on TGCA.school_id = SCHOOLS.school_id
left outer join (select distinct class,
		'ISCED' =
			CASE
				WHEN class < 1 THEN 'ISCED 02'
				WHEN class >=1 and class <=6 THEN 'ISCED 1'
				WHEN class >=8.1 and class <=8.2 THEN 'ISCED 25'
				WHEN class >=7 and class <=8 THEN 'ISCED 24'
				WHEN class >=10.1 and class <=10.2 THEN 'ISCED 35'
				WHEN class >=9 and class <=13 THEN 'ISCED 34'
			END from TGCA where year = 2015 ) 	LEVEL on LEVEL.class = TGCA.class
where year = 2015
and gender in ('F', 'M')
group by ISCED, gender, school_type, teaching_qual
go

select ISCED, SCHOOLTYPE, GENDER, QUALIFIED, TRAINED, sum(COUNT) as COUNT from #TeacherBaseTable
                    group by ISCED, SCHOOLTYPE, GENDER, QUALIFIED, TRAINED

