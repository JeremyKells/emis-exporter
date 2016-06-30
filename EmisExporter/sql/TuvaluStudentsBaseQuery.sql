select
    LEVEL.ISCED_TOP,
	LEVEL.ISCED,
    SCHOOLTYPE = case when sg.school_type = 1 then 'Public' else 'Private' end,
    sg.GENDER,
	AGES.AGE,
	REPEATERS.REPEATER,
	sg.class as CLASS,
	ISNULL(ece.ece, '') as ECE,
    count(1) as COUNT
    from v_sgca sg
    join (  select distinct sg.class,
			ISCED_TOP = CASE
			            WHEN sg.class < 1 THEN 'ISCED 0'
						WHEN sg.class >=1 and sg.class <=6  THEN 'ISCED 1'
						WHEN sg.class >=7 and sg.class <=8  THEN 'ISCED 2'
						WHEN sg.class >=9 and sg.class <=13 THEN 'ISCED 3'
					END,
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



					from v_sgca sg where year = {0}) LEVEL
		on LEVEL.class = sg.class
	join (select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg where year = {0}) AGES on AGES.student_id = sg.student_id
    join (select student_id, case when sg.status = 'R' then 'Repeaters' else '' end as REPEATER from v_sgca sg where year = {0}) REPEATERS on REPEATERS.student_id = sg.student_id
	left outer join (select student_id, CASE WHEN sum(ece) >= 1 THEN 'ECE' ELSE '' END as ece
		from (
			select student_id, ece = CASE WHEN class < 1 THEN 1 ELSE 0 END
			from v_sgca where year < {0}
			) ece
		group by student_id)
	ece on ece.student_id = sg.student_id

	where year = {0}

    group by ISCED, ISCED_TOP, school_type, gender, AGE, REPEATERS.REPEATER, sg.class, ece.ece
