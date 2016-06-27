using System;
using System.Collections.Generic;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;

namespace EmisExporter
{
    public partial class Exporter
    {
        private List<string> solomonIslandSQL()
        {
            List<string> sql = new List<string> { };

            sql.Add(""); // 0
            sql.Add(""); // 1
            sql.Add(@"select 
	                    'ISCED' = CASE DimensionLevel.[ISCED SubClass]
				                    WHEN 'ISCED0'  THEN 'ISCED 02'
				                    WHEN 'ISCED1'  THEN 'ISCED 1'
				                    WHEN 'ISCED2A' THEN 'ISCED 24'
				                    WHEN 'ISCED3A' THEN 'ISCED 34'
				                    WHEN 'ISCED3C' THEN 'ISCED 35'
			                      END, 
	                    'schoolType' = CASE DimensionSchoolSurveyNoYear.AuthorityGroup WHEN 'Government' THEN 'Public' ELSE 'Private' END,
	                    DG.GenderCode as gender,
	                    'count' = CASE DG.GenderCode WHEN 'F' then sum(ISNULL(enF, 0)) else sum(ISNULL(enM, 0)) END

                    from       dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                    INNER JOIN dbo.Enrollments  ON EE.bestssID = dbo.Enrollments.ssID 
                    INNER JOIN DimensionLevel on enLevel = DimensionLevel.LevelCode 
                    INNER JOIN DimensionSchoolSurveyNoYear on DimensionSchoolSurveyNoYear.[Survey ID] = EE.surveyDimensionssID
                    Cross JOIN DimensionGender DG

                    where EE.LifeYear = {0}

                    group by 
	                    DimensionLevel.[ISCED SubClass], 
	                    DimensionSchoolSurveyNoYear.AuthorityGroup, 
	                    GenderCode"); // 2

            sql.Add(@"select 
	                    'ISCED' = CASE DimensionLevel.[ISCED SubClass]
				                    WHEN 'ISCED0'  THEN 'ISCED 02'
				                    WHEN 'ISCED1'  THEN 'ISCED 1'
				                    WHEN 'ISCED2A' THEN 'ISCED 24'
				                    WHEN 'ISCED3A' THEN 'ISCED 34'
				                    WHEN 'ISCED3C' THEN 'ISCED 35'
			                      END, 
	                    CAST(enAge as int) as AGE,
	                    DG.GenderCode as gender,
	                    'count' = CASE DG.GenderCode WHEN 'F' then sum(ISNULL(enF, 0)) else sum(ISNULL(enM, 0)) END

                    from       dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                    INNER JOIN dbo.Enrollments  ON EE.bestssID = dbo.Enrollments.ssID 
                    INNER JOIN DimensionLevel on enLevel = DimensionLevel.LevelCode 
                    INNER JOIN DimensionSchoolSurveyNoYear on DimensionSchoolSurveyNoYear.[Survey ID] = EE.surveyDimensionssID
                    Cross JOIN DimensionGender DG

                    where EE.LifeYear = {0}

                    group by 
	                    enAge, 
	                    DimensionLevel.[ISCED SubClass], 
	                    GenderCode"); //3


            sql.Add(""); //4
            sql.Add(@"  select
	                        CAST(DimensionLevel.[Year of Education] as decimal) as class,
	                        CAST(enAge as varchar(3)) as AGE,
	                        DG.GenderCode as gender,
	                        'count' = CASE DG.GenderCode WHEN 'F' then sum(ISNULL(enF, 0)) else sum(ISNULL(enM, 0)) END
                        
                        from       dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                        INNER JOIN dbo.Enrollments  ON EE.bestssID = dbo.Enrollments.ssID 
                        INNER JOIN DimensionLevel on enLevel = DimensionLevel.LevelCode 
                        INNER JOIN DimensionSchoolSurveyNoYear on DimensionSchoolSurveyNoYear.[Survey ID] = EE.surveyDimensionssID
                        Cross JOIN DimensionGender DG

                        where EE.LifeYear = {0}
                        and DimensionLevel.[ISCED SubClass] = 'ISCED1'

                        group by 
	                        DimensionLevel.[Year of Education], 
	                        enAge, 
	                        GenderCode"); //5

            sql.Add(@"  select 
	                        CAST(DimensionLevel.[Year of Education] - 7 as decimal) as class,
	                        CAST(enAge as varchar(3)) as AGE,
	                        DG.GenderCode as gender,
	                        'count' = CASE DG.GenderCode WHEN 'F' then sum(ISNULL(enF, 0)) else sum(ISNULL(enM, 0)) END

                        from       dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                        INNER JOIN dbo.Enrollments  ON EE.bestssID = dbo.Enrollments.ssID 
                        INNER JOIN DimensionLevel on enLevel = DimensionLevel.LevelCode 
                        INNER JOIN DimensionSchoolSurveyNoYear on DimensionSchoolSurveyNoYear.[Survey ID] = EE.surveyDimensionssID
                        Cross JOIN DimensionGender DG

                        where EE.LifeYear = {0}
                        and DimensionLevel.[ISCED SubClass] in ('ISCED2A') --, 'ISCED3A', 'ISCED3C')

                        group by 
	                        DimensionLevel.[Year of Education], 
	                        enAge, 
	                        GenderCode"); //6

            sql.Add(@"select 
                        ISCED, 
                        Level, 
                        gender = genderCode, 
                        sum(count) as count from 

                        (
	                        select 
		                        'ISCED' = CASE ilsCode
				                        WHEN 'ISCED1'  THEN 'ISCED 1'
				                        WHEN 'ISCED2A' THEN 'ISCED 2'
				                        WHEN 'ISCED3A' THEN 'ISCED 3'
				                        END, 
	                        --ptLevel,
		                        'Level' = CASE ptLevel
				                        WHEN 'Prep'   THEN 1.0
				                        WHEN 'Std 1'  THEN 2.0
				                        WHEN 'Std 2'  THEN 3.0
				                        WHEN 'Std 3'  THEN 4.0
				                        WHEN 'Std 4'  THEN 5.0
				                        WHEN 'Std 5'  THEN 6.0
				                        WHEN 'Std 6'  THEN 7.0

				                        WHEN 'Form 1'  THEN 8.0
				                        WHEN 'Form 2'  THEN 9.0
				                        WHEN 'Form 3'  THEN 10.0
				                        WHEN 'Form 4'  THEN 11.0
				                        WHEN 'Form 5'  THEN 12.0
				                        WHEN 'Form 6'  THEN 13.0
				                        END, 

	                        DG.genderCode,
	                        'count' = CASE DG.GenderCode WHEN 'F' then sum(ISNULL(ptF, 0)) else sum(ISNULL(ptM, 0)) END
	                        from dbo.tfnESTIMATE_BestSurveyEnrolments() EE
	                        inner join vtblRepeaters r on EE.bestssID = r.ssID
	                        inner join lkpLevels l on l.codeCode = r.ptLevel
	                        Cross JOIN DimensionGender DG
	                        where EE.LifeYear = {0}
	                        group by ptLevel, DG.genderCode, ilsCode
                        ) a
	                        group by ISCED, Level, genderCode
	                        order by ISCED, Level, genderCode"); //7
            sql.Add(@"select 
                        'ISCED' = CASE enLevel WHEN 'Prep' THEN 'ISCED 1' ELSE 'ISCED 2' END,
                        enAge as AGE,
                        gender,
                        sum(count) as count
                        from(
                        --all Prep, Form 1 enrollments
                        select
                        enLevel,
                        gender = CASE DG.gender WHEN 'Female' THEN 'F' ELSE 'M' END,
                        enAge,
                        'count' = CASE DG.gender WHEN 'F' then sum(ISNULL(enF, 0)) else sum(ISNULL(enM, 0)) END
                        from dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                        INNER JOIN dbo.Enrollments  ON EE.bestssID = dbo.Enrollments.ssID
                        Cross JOIN DimensionGender DG
                        where lifeyear = {0}
                                and enLevel in ('Prep', 'Form 1')
	                    group by enLevel, gender, enAge

                    union select
                        --all repeaters in Prep, Form 1
                        ptLevel as enLevel,
	                    gender = CASE DG.gender WHEN 'Female' THEN 'F' ELSE 'M' END,
	                    ptAge as enAge,
	                    'count' = -1 * CASE DG.gender WHEN 'F' then sum(ISNULL(ptF, 0)) else sum(ISNULL(ptM, 0)) END
                         from dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                         inner join vtblRepeaters r on EE.bestssID = r.ssID
                         inner join lkpLevels l on l.codeCode = r.ptLevel
                         Cross JOIN DimensionGender DG
                         where EE.LifeYear = {0}
                                and ptLevel in ('Prep', 'Form 1')	
	                    group by ptLevel, DG.gender, ilsCode, ptAge
	                    ) a
                    group by enLevel, gender, enAge

                    union select
                        --Prep level new entrants that attended ECE
                        'ISCED 1-ECE' as 'ISCED',
	                    AGE, 
	                    gender = CASE DG.gender WHEN 'Female' THEN 'F' ELSE 'M' END,
	                    'count' = CASE DG.gender WHEN 'F' then sum(ISNULL([PSAF], 0)) else sum(ISNULL([PSAM], 0)) END
                        from  dbo.tfnESTIMATE_BestSurveyEnrolments() EE
                        inner join vtblPSA psa on EE.bestssID = psa.ssID
                        Cross JOIN DimensionGender DG
                        where lifeyear = {0}
                        group by AGE, DG.gender
                "); //8
            sql.Add(""); //9
            sql.Add(""); //10
            sql.Add(""); //11
            sql.Add(""); //12

            return sql;
        }
    }
}

