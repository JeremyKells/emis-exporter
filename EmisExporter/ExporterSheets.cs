using System;
using System.Collections.Generic;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;

namespace EmisExporter
{
  public partial class Exporter
  {

    // A2: Number of students by level of education, intensity of participation, type of institution and sex
    void sheetA2(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {

      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 1;     //row offset
      const int PUBLIC = 14;           //row 
      const int PRIVATE = 17;          //row

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A2"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      SqlCommand cmd = new SqlCommand(
          @"select ISCED, SCHOOLTYPE, gender, sum(count) as COUNT from #StudentsBaseTable group by ISCED, SCHOOLTYPE, gender",
          sqlConn);

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          if (rdr.IsDBNull(2))
          {
            Console.WriteLine("Skipping row, count: " + rdr.GetInt32(3).ToString());
            continue;
          }
          string isced = rdr.GetString(0);
          string schoolType = rdr.GetString(1);
          string gender = rdr.GetString(2);
          int count = rdr.GetInt32(3);

          int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
          int row = (schoolType == "Public" ? PUBLIC : PRIVATE) + rowOffset;
          int column = getCol(isced);

          workSheet.Cells[row, column] = count;
        }
      }
    }

    // A3: Number of students by level of education, age and sex
    void sheetA3(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {

      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 29;           //row offset
      const int UNDER_TWO = 14;               //row 
      const int TWENTYFIVE_TWENTYNINE = 38;   //row
      const int OVER_TWENTYNINE = 39;         //row
      const int AGE_UNKNOWN = 40;             //row
      const int ZERO = 13;                    //row
      const int MISSING_AGE = -1;

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A3"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      SqlCommand cmd = new SqlCommand(
          @"select ISCED, AGE, gender, sum(count) as COUNT from #StudentsBaseTable group by ISCED, AGE, gender",
          sqlConn);

      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          string isced;
          int age;
          string gender;
          int count;

          isced = rdr.GetString(0);
          try
          {
            age = rdr.GetInt32(1);
          }
          catch (System.Data.SqlTypes.SqlNullValueException)
          {
            age = MISSING_AGE;
          }
          gender = rdr.GetString(2);
          count = rdr.GetInt32(3);

          int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
          int row;
          if (age >= 2 && age <= 24)
          {
            row = ZERO + age + rowOffset;
          }
          else if (age == MISSING_AGE)
          {
            row = AGE_UNKNOWN + rowOffset;
          }
          else if (age < 2)
          {
            row = UNDER_TWO + rowOffset;
          }
          else if (age >= 25 && age <= 29)
          {
            row = TWENTYFIVE_TWENTYNINE + rowOffset;
          }
          else if (age > 29)
          {
            row = OVER_TWENTYNINE + rowOffset;
          }
          else
          {
            Console.WriteLine("Invalid Age: " + age);
            continue;
          }

          int column = getCol(isced);
          workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + count;
        }
      }
    }


    // A5: Number of students in initial primary education by age, grade and sex																													
    void sheetA5(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {
      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 27;     //row offset
      const int AGE_UNKNOWN = 37;       //row
      const int UNDER_FOUR = 14;        //row
      const int OVER_TWENTYFOUR = 36;   //row
      const int ZERO = 11;              //row offset 
      const int REPEATERS = 39;         //row

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A5"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      SqlCommand cmd = new SqlCommand(
         @"select CLASS, AGE, gender, REPEATER, sum(count) as COUNT from #StudentsBaseTable
                    where class >= 1 and class <= 6
                    group by CLASS, AGE, gender, REPEATER",
         sqlConn);

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          short _class = (short)rdr["class"];
          string strAge = rdr["AGE"].ToString();
          string gender = (string)rdr["gender"];
          int repeaters = (int)rdr["REPEATER"];
          int count = (int)rdr["count"];

          int column = getCol("Grade " + ((int)_class).ToString());
          int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
          int row;

          if (strAge == "N/A")
          {
            row = AGE_UNKNOWN + rowOffset;
          }
          else
          {
            int age = Convert.ToInt16(strAge);
            if (age < 4)
            {
              row = UNDER_FOUR + rowOffset;
            }
            else if (age > 24)
            {
              row = OVER_TWENTYFOUR + rowOffset;
            }
            else //if (age >= 4 && age <= 24)
            {
              row = ZERO + age + rowOffset;
            }
          }
          workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + count;

          row = REPEATERS + rowOffset;
          workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + repeaters;
        }
      }
    }

    // A6: Number of students and repeaters in initial lower and upper secondary general education by age, grade and sex																										
    void sheetA6(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {

      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 21;     //row offset
      const int AGE_UNKNOWN = 31;       //row
      const int UNDER_TEN = 14;         //row
      const int OVER_TWENTYFOUR = 30;   //row
      const int ZERO = 5;               //row offset 
      const int REPEATERS = 33;         //row
      const int ISCED_34 = 46;          //column

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A6"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      SqlCommand cmd = new SqlCommand(
          @"select CLASS, AGE, GENDER, REPEATER, sum(COUNT) as COUNT from (
	            select 
		            CASE ISCED  
			            WHEN 'ISCED 24' THEN CLASS - 6
			            WHEN 'ISCED 34' THEN -1
		            END as CLASS, 
		            AGE, gender, REPEATER, COUNT 
	            from #StudentsBaseTable where ISCED in ( 'ISCED 24', 'ISCED 34' ) 
            ) s group by CLASS, AGE, GENDER, REPEATER",
          sqlConn);

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          int _class = (int)rdr["class"];
          string strAge = rdr["AGE"].ToString();
          string gender = (string)rdr["gender"];
          int repeaters = (int)rdr["REPEATER"];
          int count = (int)rdr["count"];

          int column = _class == -1 ? ISCED_34 : getCol("Grade " + ((int)_class).ToString());

          int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
          int row;
          if (strAge == "N/A")
          {
            row = AGE_UNKNOWN + rowOffset;
          }
          else
          {
            int age = Convert.ToInt16(strAge);
            if (age < 10)
            {
              row = UNDER_TEN + rowOffset;
            }
            else if (age > 24)
            {
              row = OVER_TWENTYFOUR + rowOffset;
            }
            else //if (age >= 10 && age <= 24)
            {
              row = ZERO + age + rowOffset;
            }
          }
          //Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", _class, strAge, gender, count));
          workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + count;

          row = REPEATERS + rowOffset;
          workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + repeaters;
        }
      }
    }

    // A7: Number of new entrants to Grade 1 in initial education and prior enrolment by age and sex											
    void sheetA7(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {

      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 21;           //row offset
      const int UNDER_FOUR = 14;              //row 
      const int OVER_EIGHTEEN = 30;           //row
      const int ZERO = 11;                    //row offset
      const int AGE_UNKNOWN = 31;             //row
      const int PRIMARY_COL = 22;             //col
      const int ECE = 33;                     //row
      const int LOWER_SECONDARY_COL = 25;     //col

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A7"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      SqlCommand cmd = new SqlCommand(
          @"select ISCED_TOP as ISCED, AGE, gender, ECE, sum(REPEATER) as REPEATER, sum(count) as COUNT from #StudentsBaseTable
                    where class in (1.0, 7.0)
                    group by ISCED_TOP, gender, AGE, ECE",
          sqlConn);


      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          string isced;
          int age;
          string gender;
          int ece;
          int count;
          int repeaters;

          isced = (string)rdr["ISCED"];
          age = (int)rdr["AGE"];
          gender = (string)rdr["gender"];
          ece = (int)rdr["ECE"];
          count = (int)rdr["count"];
          repeaters = (int)rdr["REPEATER"];

          //Console.WriteLine(String.Format("{0}, {1}, {2}, {3}, {4}, {5}", isced, gender, age, ece, count, repeaters));

          int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
          int row;
          if (age >= 2 && age <= 24)
          {
            row = ZERO + age + rowOffset;
          }
          else if (age < 4)
          {
            row = UNDER_FOUR + rowOffset;
          }
          else if (age > 18)
          {
            row = OVER_EIGHTEEN + rowOffset;
          }
          else
          {
            row = AGE_UNKNOWN + rowOffset;
          }

          if (isced == "ISCED 1")
          {
            workSheet.Cells[row, PRIMARY_COL] = workSheet.get_Range(helpers.GetCellAddress(PRIMARY_COL, row)).Value2 + count - repeaters;
            row = ECE + rowOffset;
            workSheet.Cells[row, PRIMARY_COL] = workSheet.get_Range(helpers.GetCellAddress(PRIMARY_COL, row)).Value2 + ece;
          }
          else if (isced == "ISCED 2")
          {
            workSheet.Cells[row, LOWER_SECONDARY_COL] = workSheet.get_Range(helpers.GetCellAddress(LOWER_SECONDARY_COL, row)).Value2 + count - repeaters;
          }
          else
          {
            Console.WriteLine("Didn't match ISCED");
          }
        }
      }
    }

    // A9: Number of classroom teachers by teaching level of education, employment status, type of institution and sex																																	
    void sheetA9(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {
      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 1;     //row offset
      const int PUBLIC = 14;
      const int PRIVATE = 17;

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A9"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      SqlCommand cmd = new SqlCommand(
          @"select ISCED, SCHOOLTYPE, GENDER, sum(COUNT) as COUNT 
                    from #TeacherBaseTable
                    group by ISCED, SCHOOLTYPE, GENDER",
          sqlConn);


      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          string isced = rdr.GetString(0);
          string schoolType = rdr.GetString(1);
          string gender = rdr.GetString(2);
          int count = rdr.GetInt32(3);

          int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
          int row = (schoolType == "PUBLIC" ? PUBLIC : PRIVATE) + rowOffset;

          //Console.WriteLine(String.Format("{0}, {1}, {2}, {3}, {4}, {5}", isced, schoolType, gender, count, rowOffset, row));

          List<string> columns = new List<string>();

          if (isced == "ISCED 24" || isced == "ISCED 34")
          {
            columns.Add("ISCED 24+34");
            columns.Add(isced.Substring(0, 7));
          }
          else if (isced == "ISCED 25" || isced == "ISCED 35")
          {
            columns.Add("ISCED 25+35");
            columns.Add(isced.Substring(0, 7));
          }
          else
          {
            columns.Add(isced);
          }
          foreach (string column in columns)
          {
            workSheet.Cells[row, getCol(column)] = workSheet.get_Range(helpers.GetCellAddress(getCol(column), row)).Value2 + count;
          }
        }
      }
    }

    // A10: Number of classroom teachers by qualified and trained status, teaching level of education, type of institution and sex
    void sheetA10(Excel.Application excelApp, SqlConnection sqlConn, string year, string country)
    {

      //Constant references for columns and rows            
      const int FEMALE_OFFSET = 1;     //row offset
      const int TRAINED_OFFSET = 10;     //row offset
      const int PUBLIC = 14;
      const int PRIVATE = 17;

      Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A10"];
      workSheet.Activate();
      Excel.Range usedRange = workSheet.UsedRange;

      Func<string, int> getCol = null;
      getCol = n => usedRange.Find(n).Column;
      getCol.Memoize();

      SqlCommand cmd = new SqlCommand(
          @"select ISCED, SCHOOLTYPE, GENDER, QUALIFIED, TRAINED from #TeacherBaseTable", sqlConn);

      using (SqlDataReader rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          string isced = rdr.GetString(0);
          string schooltype = rdr.GetString(1);
          string gender = rdr.GetString(2);
          int qualified = rdr.GetInt32(3);
          int trained = rdr.GetInt32(4);

          int row = (schooltype == "PUBLIC" ? PUBLIC : PRIVATE)
                            + (gender == "F" ? FEMALE_OFFSET : 0);

          Console.WriteLine(String.Format("{0}, {1}, {2}, {3}, {4}, {5}", isced, schooltype, gender, qualified, trained, row));

          List<string> columns = new List<string>();

          if (isced == "ISCED 24" || isced == "ISCED 34")
          {
            columns.Add("ISCED 24+34");
            columns.Add(isced.Substring(0, 7));
          }
          else if (isced == "ISCED 25" || isced == "ISCED 35")
          {
            columns.Add("ISCED 25+35");
            columns.Add(isced.Substring(0, 7));
          }
          else
          {
            columns.Add(isced);
          }
          foreach (string column in columns)
          {
            // This val below was added to both qualified and trained values below. I'm not sure why
            // but it was providing incorrect trained teachers (those figure were actually trained + qualified
            // inflating numbers to impossible figures (e.g. 300 trained and only 200 total teachers)
            //double val = workSheet.get_Range(helpers.GetCellAddress(getCol(column), row)).Value2;
            workSheet.Cells[row, getCol(column)] = qualified;
            workSheet.Cells[row + TRAINED_OFFSET, getCol(column)] = trained;
          }
        }
      }
    }
  }
}
