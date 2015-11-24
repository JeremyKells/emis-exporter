﻿using System;
using System.Collections.Generic;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;

namespace EmisExporter
{
    public partial class Exporter
    {
        public string sheet2_SQL = @"select
                                    LEVEL.ISCED,
                                    'schoolType' = case when sg.school_type = 1 then 'Public' else 'Private' end,
                                    sg.gender,
                                    count(1) as count
                                    from v_sgca sg
                                    join (  select distinct sg.class,
                                            'ISCED' = 
		                                            CASE
			                                            -- Currently K1, K2, K3 all ISCED 02     THEN 'ISCED 01'
			                                            WHEN sg.class < 1 THEN 'ISCED 02'
			                                            WHEN sg.class >=1 and sg.class <=6 THEN 'ISCED 1'
			                                            WHEN sg.class >=7 and sg.class <=8 THEN 'ISCED 24'
			                                            WHEN sg.class = 8.1 or sg.class = 8.2 THEN 'ISCED 25'
			                                            WHEN sg.class = 10.1 or sg.class = 10.2 THEN 'ISCED 35'
			                                            WHEN sg.class >=9 and sg.class <=13 THEN 'ISCED 34'
			                                            -- Currently no data for ISCED 44 / ISCED 45
		                                            END 
                                             from v_sgca sg ) LEVEL on LEVEL.class = sg.class
                                    where year = {0}
                                    group by ISCED, school_type, gender
                                    ";

        public string sheet3_SQL = @"select ISCED, AGE, gender, count(1) as count from (
                                    select 
                                    'ISCED' = 
                                        CASE
	                                    -- Currently K1, K2, K3 all ISCED 02     THEN 'ISCED 01'
	                                    WHEN sg.class < 1 THEN 'ISCED 02'
	                                    WHEN sg.class >=1 and sg.class <=6 THEN 'ISCED 1'
	                                    WHEN sg.class >=7 and sg.class <=8 THEN 'ISCED 24'
	                                    WHEN sg.class = 8.1 or sg.class = 8.2 THEN 'ISCED 25'
	                                    WHEN sg.class = 10.1 or sg.class = 10.2 THEN 'ISCED 35'
	                                    WHEN sg.class >=9 and sg.class <=13 THEN 'ISCED 34'
	                                    -- Currently no data for ISCED 44 / ISCED 45
                                        END,
                                    AGE,
                                    gender 
                                    from v_sgca sg
                                    join (select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg where year = {0}) AGES on AGES.student_id = sg.student_id
                                    where year = {0}
                                    ) v group by ISCED, AGE, gender
                                    ";

        public string sheet5_SQL = @"select
                                    sg.class, 
                                    isnull(convert(nvarchar(3),AGE),'N/A') AGE,
                                    sg.gender,
                                    count(1) as count
                                    from v_sgca sg
                                    join(select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg where year = {0}) AGES on AGES.student_id = sg.student_id
                                    where sg.year = {0}
                                    --and AGE is not null
                                    group by sg.class, AGES.AGE, gender
                              ";
        public string sheet6_SQL = @"select
                                    sg.class, 
                                    isnull(convert(nvarchar(3),AGE),'N/A') AGE,
                                    sg.gender,
                                    count(1) as count
                                    from v_sgca sg
                                    join(select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg where year = {0}) AGES on AGES.student_id = sg.student_id
                                    where sg.year = {0}
                                    --and AGE is not null
                                    group by sg.class, AGES.AGE, gender
                              ";
        public string sheet7_SQL = @"select  
                                    LEVEL.ISCED,
                                    sg.class,
                                    gender,
                                    count(1) 
                                    from v_sgca sg
                                    join (select distinct sg.student_id as id,
	                                    'ISCED' = 
		                                    CASE
			                                    WHEN sg.class >=1 and sg.class <=6  THEN 'ISCED 1'
			                                    WHEN sg.class >=7 and sg.class <=8  THEN 'ISCED 2'
			                                    WHEN sg.class >=9 and sg.class <=13 THEN 'ISCED 3'
		                                    END from v_sgca sg where year = {0}) 
	                                    LEVEL on LEVEL.id = sg.student_id
                                    where year = {0}
                                    and status = 'R'
                                    and ISCED is not null
                                    group by ISCED, sg.class, gender
                                  ";

        public string sheet8_SQL = @"select 
                                    LEVEL.ISCED,
                                    Ages.AGE, 
                                    gender,
                                    count(1) as count
                                    from V_SGCA sg
                                    join (select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg) AGES on AGES.student_id = sg.student_id
                                    join (select distinct sg.student_id as id,
	                                    'ISCED' = 
		                                    CASE
			                                    WHEN sg.class >=1 and sg.class <=6 THEN 'ISCED 1'
			                                    WHEN sg.class >=7 and sg.class <=8 THEN 'ISCED 2'
		                                    END from v_sgca sg where year = {0}) 
	                                    LEVEL on LEVEL.id = sg.student_id
                                    where year = {0}
                                    and status = 'N' -- New Entrant
                                    and ISCED is not null
                                    group by AGE, gender, LEVEL.ISCED

                                    UNION 

                                    select 
                                    LEVEL.ISCED,
                                    Ages.AGE, 
                                    gender,
                                    sum(ece) as count
                                    from V_SGCA sg
                                    join (select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg) AGES on AGES.student_id = sg.student_id
                                    join (select distinct sg.student_id as id,
	                                    'ISCED' = 
		                                    CASE
			                                    WHEN sg.class >=1 and sg.class <=6 THEN 'ISCED 1-ECE'
		                                    END from v_sgca sg where year = {0}) 
	                                    LEVEL on LEVEL.id = sg.student_id
                                    left outer join (select student_id, CASE WHEN sum(ece) >= 1 THEN 1 ELSE 0 END as ece 
		                                    from (
			                                    select student_id, ece = CASE WHEN class < 1 THEN 1 ELSE 0 END
			                                    from v_sgca where year < {0}
			                                    ) ece
		                                    group by student_id) 
	                                    ece on ece.student_id = sg.student_id

                                    where year = {0}
                                    and status = 'N' -- New Entrant
                                    and ISCED is not null
                                    group by AGE, gender, LEVEL.ISCED
                                    ";

        public string sheet10_SQL = @"select 
                                    LEVEL.ISCED, 
                                    school_type = CASE WHEN school_type = 1 THEN 'PUBLIC' WHEN school_type = 2 THEN 'PRIVATE' END,
                                    gender, 
                                    count(1) as count
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
		                                    END from TGCA where year = 2014 ) 
	                                    LEVEL on LEVEL.class = TGCA.class
                                    where year = {0}
                                    and gender in ('F', 'M')
                                    group by ISCED, gender, school_type
                                    ";


        public string sheet12_SQL = @"select 
                                    LEVEL.ISCED, 
                                    gender, 
                                    count(1)
                                    from TGCA
                                    left outer join STAFF on TGCA.staff_id = STAFF.staff_id
                                    left outer join (select distinct class,
	                                    'ISCED' = 
		                                    CASE
			                                    WHEN class < 1 THEN 'ISCED 02'
			                                    WHEN class >=1 and class <=6 THEN 'ISCED 1'
			                                    WHEN class >=8.1 and class <=8.2 THEN 'ISCED 25'
			                                    WHEN class >=7 and class <=8 THEN 'ISCED 24'
			                                    WHEN class >=10.1 and class <=10.2 THEN 'ISCED 35'
			                                    WHEN class >=9 and class <=13 THEN 'ISCED 34'
		                                    END from TGCA where year = {0}) 
	                                    LEVEL on LEVEL.class = TGCA.class
                                    where year = {0}
                                    and teaching_qual = 'Y'
                                    group by ISCED, gender
                                  ";
    }
}
