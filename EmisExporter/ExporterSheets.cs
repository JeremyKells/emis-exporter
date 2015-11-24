using System;
using System.Collections.Generic;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;

namespace EmisExporter
{
    public partial class Exporter
    {
        //private static readonly log4net.ILog log = log4net.LogManager.GetLogger (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // A2: Number of students by level of education, intensity of participation, type of institution and sex
        void sheetA2(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 4;     //row offset
            const int PUBLIC = 17;           //row 
            const int PRIVATE = 18;          //row
            //const int PART_TIME = 28;        //row
            //log.Error("sheetA2");
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A2"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand(
              string.Format(sheet2_SQL, year),
                                temis);
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
            }

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
                    Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, schoolType, count));

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row = (schoolType == "Public" ? PUBLIC : PRIVATE) + rowOffset;
                    //int column = usedRange.Find(isced).Column;
                    int column = getCol(isced);

                    workSheet.Cells[row, column] = count;
                    Console.WriteLine(row.ToString() + " : " + column.ToString());
                }
            }
        }

        // A3: Number of students by level of education, age and sex
        void sheetA3(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 29;     //row offset
            const int UNDER_TWO = 17;           //row 
            const int TWENTYFIVE_TWENTYNINE = 41;          //row
            const int OVER_TWENTYNINE = 42;        //row
            const int ZERO = 16;        //row

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A3"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(sheet3_SQL, year),
                              temis);
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
            }

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced;
                    int age;
                    string gender;
                    int count;

                    try
                    {
                        isced = rdr.GetString(0);
                        age = rdr.GetInt32(1);
                        gender = rdr.GetString(2);
                        count = rdr.GetInt32(3);
                    }
                    catch
                    {
                        continue;
                    }

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row;
                    if (age >= 2 && age <= 24)
                    {
                        row = ZERO + age + rowOffset;
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
                    //Console.WriteLine(row.ToString() + " : " + column.ToString());
                }
            }
        }


        // A5: Number of students in initial primary education by age, grade and sex																													
        void sheetA5(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 26;     //row offset
            const int AGE_UNKNOWN = 40;       //row
            const int UNDER_FOUR = 17;        //row
            const int OVER_TWENTYFOUR = 39;   //row
            const int ZERO = 14;              //row offset 
            //const int UNSPECIFIED_GRADE = 38; //column AL

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A5"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;


            SqlCommand cmd = new SqlCommand(
              string.Format(sheet5_SQL, year),
                              temis);
            string classColName = "class";
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
                classColName = "grade";
            }


            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    decimal _class = (decimal)rdr[classColName];
                    string strAge = (string)rdr["AGE"];
                    string gender = (string)rdr["gender"];
                    int count = (int)rdr["count"];
                    
                    if (_class >= 1 && _class <= 6){
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
                    }
                }
            }
        }

        // A6: Number of students in initial lower secondary general education by age, grade and sex																										
        void sheetA6(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 20;     //row offset
            const int AGE_UNKNOWN = 34;       //row
            const int UNDER_TEN = 17;         //row
            const int OVER_TWENTYFOUR = 33;   //row
            const int ZERO = 8;               //row offset 
            //const int UNSPECIFIED_GRADE = 35; // column AI

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A6"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand(
               string.Format(sheet6_SQL, year),
                               temis);
            string classColName = "class";
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
                classColName = "grade";
            }


            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    decimal _class = (decimal)rdr[classColName] - 6; 
                    string strAge = (string)rdr["AGE"];
                    string gender = (string)rdr["gender"];
                    int count = (int)rdr["count"];

                    if (_class >= 1 && _class <= 6)
                    {
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
                        workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + count;
                    }
                }
            }
        }

        // A7: Number of repeaters in initial primary and general secondary education by level, grade and sex																																																							
        void sheetA7(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {
            //Constant references for columns and rows            
            const int MALE_ROW = 17;      //row
            const int FEMALE_ROW = 18;      //row
            const int ZERO = 14;            //column
            const int SECONDARY_OFFSET = 9; //column offset
            const int UPPER_SECONDARY = 68; //column

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A7"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand(
              string.Format(sheet7_SQL, year),
                              temis);
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
            }


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    Decimal _class = rdr.GetDecimal(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);

                    int row = gender == "M" ? MALE_ROW : FEMALE_ROW;
                    int column = ZERO + (int)(_class * 3);
                    column = (isced == "ISCED 2") ? column + SECONDARY_OFFSET : column;
                    column = (isced == "ISCED 3") ? UPPER_SECONDARY : column;
                    workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + count;
                }
            }
        }

        // A8: Number of new entrants to Grade 1 in initial education and prior enrolment by age and sex											
        void sheetA8(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 20;           //row offset
            const int UNDER_FOUR = 17;              //row 
            const int OVER_EIGHTEEN = 33;           //row
            const int ZERO = 14;                    //row offset
            const int AGE_UNKNOWN = 34;             //row
            const int PRIMARY_COL = 17;             //col
            const int PRIMARY_ECE_COL = 20;         //col
            const int LOWER_SECONDARY_COL = 23;     //col

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A8"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(sheet8_SQL, year),
                              temis);
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
            }

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced;
                    int age;
                    string gender;
                    int count;

                    try
                    {
                        isced = rdr.GetString(0);
                        age = rdr.GetInt32(1);
                        gender = rdr.GetString(2);
                        count = rdr.GetInt32(3);
                        Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, age, count));
                    }
                    catch
                    {
                        continue;  // Data needs to be clean
                    }

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

                    int column;
                    switch (isced)
                    {
                        case "ISCED 1":
                            column = PRIMARY_COL;
                            break;
                        case "ISCED 1-ECE":
                            column = PRIMARY_ECE_COL;
                            break;
                        case "ISCED 2":
                            column = LOWER_SECONDARY_COL;
                            break;
                        default:
                            column = 0;  // ERROR
                            break;
                    }

                    workSheet.Cells[row, column] = workSheet.get_Range(helpers.GetCellAddress(column, row)).Value2 + count;
                    Console.WriteLine(age + " " + isced + " " + gender);
                    Console.WriteLine(row.ToString() + " : " + column.ToString());
                }
            }
        }

        // A10: Number of classroom teachers by teaching level of education, employment status, type of institution and sex																																	
        void sheetA10(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {
            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 4;     //row offset
            const int PUBLIC = 17;
            const int PRIVATE = 18;

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A10"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(sheet10_SQL, year),
                              temis);
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
            }


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    string schoolType = rdr.GetString(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);
                    Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, schoolType, count));

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row = schoolType == "PUBLIC" ? PUBLIC : PRIVATE + rowOffset;

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

        // A12: Number of qualified classroom teachers by teaching level of education and sex																															
        void sheetA12(Excel.Application excelApp, SqlConnection temis, string year, string country)
        {

            //Constant references for columns and rows            
            const int FEMALE = 18;     //row offset
            const int MALE = 17;

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A12"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(sheet12_SQL, year),
                              temis);
            if (country == "NAURU")
            {
                cmd.CommandText = cmd.CommandText.Replace("class", "grade");
                cmd.CommandText = cmd.CommandText.Replace("teaching_qual = 'Y'", "qual_type is not null");
            }

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    string gender = rdr.GetString(1);
                    int count = rdr.GetInt32(2);

                    int row = gender == "M" ? MALE : FEMALE;
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
    }
}
