using System;
using Microsoft.Office.Interop.Excel;

namespace Hahaton
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application excelApp = new Application();
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }
            Workbook excelBook = excelApp.Workbooks.Open(@"D:\Libraries\Desktop\Hahaton\Хакатон3.csv"); // Заменить на заданный путь
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            //System.IO.FileStream fs = new System.IO.FileStream("Путь", System.IO.FileMode.Create);
            //System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, Encoding.Unicode);

            excelBook.SaveAs(@"D:\Libraries\Desktop\Hahaton\" + "Хакатон tmp.xlsx", XlFileFormat.xlExcel4Workbook, XlSaveAsAccessMode.xlNoChange);
            //for (int i = 0; i < rowCount; i++)
            //{
            //    string[] line = excelRange.Cells[i,1].Split(',');

            //}


            Console.WriteLine("I am all");




            excelBook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadLine();

            void PrintBranchDown()
            {
                string tmp = excelRange.Cells[2, 2].Text.ToString(); //исправить колонки на заданную
                if (tmp != "" && tmp.Length > 6)
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (excelRange.Cells[i, 3].Text.ToString() == tmp)
                        {
                            Console.WriteLine(excelRange.Cells[i, 1].Value2.ToString());
                            tmp = excelRange.Cells[i, 2].Text.ToString();
                            i = 1;
                        }
                        //if (excelRange.Cells[i, 1].Text.ToString() == "Component 00000000001000") Console.WriteLine("\n1000");
                        //if (excelRange.Cells[i, 1].Text.ToString() == "Component 00000000010000") Console.WriteLine("\n10000");
                        //if (excelRange.Cells[i, 1].Text.ToString() == "Component 00000000100000") Console.WriteLine("\n100000");
                    }
                }
                Console.WriteLine("End");
            }

            //void PrintChildrens()
            //{
            //    //Выводит потомков по значению 2 колонки
            //    string tmp = excelRange.Cells[1, 2].Text.ToString(); //исправить колонки на заданную
            //    if (tmp != "" && tmp.Length > 6)
            //    {
            //        for (int i = 1; i < rowCount; i++)
            //        {
            //            if (excelRange.Cells[i, 3].Text.ToString() == tmp) Console.WriteLine(excelRange.Cells[i, 1].Value2.ToString());
            //        }

            //    }
            //}
        }
    }
}
