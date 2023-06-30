using System.IO;
using System.Data;
using ExcelDataReader;
using System;
using System.Collections.Generic;

namespace Lab5Games.UnityExcelReader
{
    public class DataReader
    {
        public static DataRowCollection ReadExcelFile(string filePath, int sheetIndex = 0)
        {
            using(FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader dataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

                DataSet dataSet = dataReader.AsDataSet();

                return dataSet.Tables[sheetIndex].Rows;
            }
        }

        public static DataRowCollection ReadCsvFile(string filePath)
        {
            using(StreamReader sr = new StreamReader(filePath))
            {
                return ReadCsv(sr.ReadToEnd());
            }
        }

        public static DataRowCollection ReadCsv(string content)
        {
            using (StringReader sr = new StringReader(content))
            {
                DataTable table = new DataTable();

                string line;
                List<string> lines = new List<string>();
                
                while ((line = sr.ReadLine()) != null)
                {
                    lines.Add(line);
                }

                // columns
                int numColumns = lines[0].Split(',').Length;
                for(int i=0; i<numColumns; i++)
                {
                    table.Columns.Add($"col_{i}");
                }

                // rows
                for(int i=0; i<lines.Count; i++)
                {
                    DataRow row = table.NewRow();
                    row.ItemArray = lines[i].Split(','); 

                    table.Rows.Add(row);
                }

                return table.Rows;
            }
        }
    }
}
