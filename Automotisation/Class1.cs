using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using Org.BouncyCastle.Crypto.Agreement;
using NPOI.HSSF.UserModel;



namespace Automotisation;

public class Class1
{
    public void Execute(string readyAnswers, string studentFiles, string outResult,string param = "*.xlsx")
    {
        string[] strings = { }; // Имена файлов 
        string answer = readyAnswers;
        IWorkbook resultWorkbook = new XSSFWorkbook();
        IWorkbook answbook = new XSSFWorkbook(answer, true);
        ISheet resultSheet = resultWorkbook.CreateSheet("Sheet1");
        ISheet answSheet = answbook.GetSheetAt(0);
        IRow answRow = answSheet.GetRow(0);
        ICell answerCell = answRow.GetCell(0);
        resultSheet.DefaultColumnWidth = 80;    
        try
        {
            strings = Directory.GetFiles(studentFiles,param);
        }
        catch (Exception)
        {
            Console.WriteLine("Err");
        }
        //Таблица с ответами
        int answerCount = 0;
        int ANSWERCOUNT = 0;
        int maxI = 0;
           
        foreach (var file in strings)
        {
            bool skip = false;
            IWorkbook studentsBook = new XSSFWorkbook(file, true);
            ISheet StudentsSheet = studentsBook.GetSheetAt(0);
            IRow StudentsRow = StudentsSheet.GetRow(0);
            ICell studentsCell = StudentsRow.GetCell(0);
            ICellStyle st = resultWorkbook.CreateCellStyle();
            IFont font = resultWorkbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeight = 650;
            XSSFColor color = new(IndexedColors.Orange);

            st.Alignment = HorizontalAlignment.Center;
            st.LeftBorderColor = color.Indexed;
            if (answSheet == null && StudentsSheet == null)
                Console.WriteLine("Check files for mistakes");
            answRow = answSheet.GetRow(0);

            if (answRow == null)
            {
                Console.WriteLine("End of file");
                skip = true;
            }
            StudentsRow = StudentsSheet.GetRow(0);
            if (StudentsRow == null)
            {
                Console.WriteLine("End of answers (error)");
                skip = true;
            }
            if (skip != true)
            {
                IRow resultRow;
                string resultString = "";
                string res = "";

                int i = 0;
                while (true)
                {
                    if (i == ANSWERCOUNT && ANSWERCOUNT != 0)
                        break;
                    if(ANSWERCOUNT == 0)
                        resultRow = resultSheet.CreateRow(i);
                    else
                        resultRow = resultSheet.CreateRow(i + ANSWERCOUNT * answerCount);
                    ICell resultCell = resultRow.CreateCell(0);
                    if (i == 0)
                    {
                        ICell cell = resultRow.CreateCell(1);
                        cell.SetCellValue(file);
                    }
                    answRow = answSheet.GetRow(i);
                    StudentsRow = StudentsSheet.GetRow(i);
                    if (answRow == null)
                    {
                        break;
                    }
                    else if (StudentsRow == null)
                    {
                        Console.WriteLine("Empty value");
                    }
                    else
                    {
                        if (answRow == null) break;
                        answerCell = answRow.GetCell(0);
                       
                        studentsCell = StudentsRow.GetCell(0);
                        if (answerCell != null)
                        {
                           // if (answerCell.ToString() == studentsCell.ToString())
                           //     resultCell.SetCellValue($"{i+1}, +");
                           // else resultCell.SetCellValue($"{i+1}, -");
                           for(int j = 0; j < answerCell.ToString().Length; j++)
                            {
                                if (j > studentsCell.ToString().Length - 1)
                                    resultString += "-, ";
                                else if (studentsCell.ToString()[j] == answerCell.ToString()[j])
                                    resultString += "+, ";
                                else
                                    resultString += "-, ";
                                res = resultString.Remove(resultString.Length - 1);
                                resultCell.SetCellValue(res);
                                
                            }
                            res = "";
                            resultString = "";

                        }
                        
                        resultRow.RowStyle = st;
                        resultRow.RowStyle.SetFont(font);

                    }
                    ++i;
                    maxI = i;
                }
                ++answerCount;
                ++answerCount;
            }
            if(ANSWERCOUNT == 0)
                ANSWERCOUNT = maxI;
            
        }
        using (FileStream fileStream = new FileStream(outResult, FileMode.Create))
        {
            resultWorkbook.Write(fileStream, false);
        }
    }
}

