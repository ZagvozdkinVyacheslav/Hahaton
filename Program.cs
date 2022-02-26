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
            void PrintBranchUp()
            {
                string tmp = excelRange.Cells[5, 3].Text.ToString(); //исправить колонки на заданную


                if (tmp != "" && tmp.Length > 6)
                {
                    for (int j = 1; j <= rowCount; j++)
                    {

                        if (excelRange.Cells[j, 2].Text.ToString() == tmp)
                        {
                            Console.WriteLine(excelRange.Cells[j, 1].Value2.ToString());
                            tmp = excelRange.Cells[j, 3].Text.ToString();
                            j = 1;

                        }


                    }

                }




                Console.WriteLine("End");
            }
            void Svoistva()
            {
                int n = Console.Read(); ;
                string tmp = excelRange.Cells[n, 2].Text.ToString(); //исправить колонки на заданную


                if (tmp != "" && tmp.Length > 6)
                {
                    Console.WriteLine("Номер элемента " + excelRange.Cells[n, 1].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 4].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 5].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 6].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 7].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 8].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 9].Value2.ToString());
                    Console.WriteLine(excelRange.Cells[n, 10].Value2.ToString());
                }
                else Console.WriteLine("Битый элемент");




                Console.WriteLine("End");
            }
            void EXCEPTIONS()
            {
                List<string> allbad = new List<string>();
                string[] elem = new string[colCount];


                for (int j = 1; j <= rowCount; j++)
                {
                    string b = excelRange.Cells[j, 2].Value2.ToString();
                    string tmp = excelRange.Cells[j, 3].Text.ToString();

                    if ((tmp == "" && b != "1436259974") || (tmp.Length > 0 && tmp.Length <= 6))
                    {
                        for (int i = 0; i < colCount; i++)
                        {
                            elem[i] = excelRange.Cells[j, i + 1].Text.ToString();

                        }
                        allbad.Add(elem.ToString());
                    }
                }
                foreach (var item in allbad)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        Console.Write(elem[i]);
                    }
                    Console.WriteLine();
                }
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
