using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream stream = File.Open("性能测试feedback.xls", FileMode.Open, FileAccess.Read);
            IExcelDataReader exc = ExcelReaderFactory.CreateBinaryReader(stream);
            DataSet mResultSets = exc.AsDataSet();

            Console.WriteLine(mResultSets.Tables[0].Columns.Count);
            Console.WriteLine(mResultSets.Tables[0].Rows.Count);


            for (int i = 0; i < mResultSets.Tables[0].Columns.Count; i++)
            {
                for (int j = 0; j < mResultSets.Tables[0].Rows.Count; j++)
                {
                    Console.WriteLine(mResultSets.Tables[0].Rows[j][i]);
                }

            }
        }
    }
}
