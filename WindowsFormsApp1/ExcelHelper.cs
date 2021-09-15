using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace WindowsFormsApp1
{
    interface IExcel
    {
        string ReadCell(int range, int column);
        void WriteCell(int range, int column, string value);
        void DeletCell(int range, int column);
        void SelectWorkSheet(int SheetNumber);
        void DeletWorkSheet(int SheetNumber);
        string ReadRange(int rangNum, int count);
        string ReadRange(int rangNum, int count, int offset);
        string ReadRange(int rangNum, int count, int offset, string delimiter);
        void WriteRange(int rangNum, string args, string delimiter);
        void WriteRange(int rangNum, string args, string delimiter, int offset);
        void WriteColumn(int ColumnNum, string args, string delimiter);
        void WriteColumn(int ColumnNum, string args, string delimiter, int offset);
        void WriteMatrixByRange(int startPointRange, int startPointColumn, string args, string elementDelimiter, string rangeDelimiter);
        void WriteMatrixByColumn(int startPointRange, int startPointColumn, string args, string elementDelimiter, string columnDelimiter);
        void CleanByMatrix(int startPointRange, int startPointColumn, int rangeOffset, int columnOffset);
    }
    class ExcelHelper : IExcel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public bool border;
        public int substringcount;

        public ExcelHelper(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = excel.Worksheets[Sheet];
        }
        public ExcelHelper(string path, int Sheet, bool border)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = excel.Worksheets[Sheet];
            this.border = border;
        }
        #region program function
        public string ReadCell(int range, int column)
        {
            string val = Convert.ToString(ws.Cells[range, column].Value != null ? ws.Cells[range, column].Value : "");
            return val;
        }
        public void WriteCell(int range, int column, string value)
        {
            ws.Cells.NumberFormat = "@";
            ws.Cells[range, column].Value = value;
            if (border)
            {
                ws.Cells[range, column].Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            }
            else
            {
                ws.Cells[range, column].Borders.LineStyle = _Excel.XlLineStyle.xlLineStyleNone;
            }
        }
        public void WriteCell(int range, int column, string value, int subSymCount)
        {
            int index = value.IndexOf(",");
            string res = index > 0 ? value.Substring(0, index + subSymCount) : value;
            ws.Cells.NumberFormat = "@";
            ws.Cells[range, column].Value = res;
            if (border)
            {
                ws.Cells[range, column].Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            }
            else
            {
                ws.Cells[range, column].Borders.LineStyle = _Excel.XlLineStyle.xlLineStyleNone;
            }
        }
        public void DeletCell(int range, int column)
        {
            ws.Cells[range, column].Value = "";
        }
        public void SelectWorkSheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }
        public void DeletWorkSheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delet();
        }

        #region ReadRange
        public string ReadRange(int rangNum, int count)
        {
            count++;
            string result = "";
            for (int i = 1; i < count; i++)
            {
                char explodeSymbol = i < count - 1 ? ',' : '\0';
                result += Convert.ToString(ReadCell(rangNum, i)) + explodeSymbol;
            }
            return result;
        }
        public string ReadRange(int rangNum, int count, int offset)
        {
            count++;
            count = count + offset;
            string result = "";
            for (int i = offset + 1; i < count; i++)
            {
                char explodeSymbol = i < count - 1 ? ',' : '\0';
                result += Convert.ToString(ReadCell(rangNum, i)) + explodeSymbol;
            }
            return result;
        }
        public string ReadRange(int rangNum, int count, int offset, string delimiter)
        {
            count++;
            count = count + offset;
            string result = "";
            for (int i = offset + 1; i < count; i++)
            {
                string explodeSymbol = i < count - 1 ? delimiter : "";
                result += Convert.ToString(ReadCell(rangNum, i)) + explodeSymbol;
            }
            return result;
        }
        #endregion
        #region WriteRange
        public void WriteRange(int rangNum, string args, string delimiter)
        {

            BaseHelper baseH = new BaseHelper();
            string[] arrAgrs = baseH.explode(delimiter, args);
            int count = arrAgrs.Count() + 1;
            for (int i = 1; i < count; i++)
            {
                WriteCell(rangNum, i, arrAgrs[i - 1]);
            }

        }
        public void WriteRange(int rangNum, string args, string delimiter, int offset)
        {

            BaseHelper baseH = new BaseHelper();
            string[] arrAgrs = baseH.explode(delimiter, args);
            int count = arrAgrs.Count() + 1 + offset;
            for (int i = offset + 1; i < count; i++)
            {
                WriteCell(rangNum, i, arrAgrs[(i - 1) - offset]);
            }

        }
        #endregion
        #region WriteColumn
        public void WriteColumn(int ColumnNum, string args, string delimiter)
        {

            BaseHelper baseH = new BaseHelper();
            string[] arrAgrs = baseH.explode(delimiter, args);
            int count = arrAgrs.Count() + 1;
            for (int i = 1; i < count; i++)
            {
                WriteCell(i, ColumnNum, arrAgrs[i - 1]);
            }
        }
        public void WriteColumn(int ColumnNum, string args, string delimiter, int offset)
        {
            BaseHelper baseH = new BaseHelper();
            string[] arrAgrs = baseH.explode(delimiter, args);
            int count = arrAgrs.Count() + 1 + offset;
            for (int i = offset + 1; i < count; i++)
            {
                WriteCell(i, ColumnNum, arrAgrs[(i - 1) - offset]);
            }
        }
        #endregion
        public void WriteMatrixByRange(int startPointRange, int startPointColumn, string args, string elementDelimiter, string rangeDelimiter)
        {
            int offset = startPointColumn - 1;
            BaseHelper baseH = new BaseHelper();
            string[] Ranges = baseH.explode(rangeDelimiter, args);
            for (int i = 0; i < Ranges.Count(); i++)
            {
                WriteRange(startPointRange + i, Ranges[i], elementDelimiter, offset);

            }

        }
        public void WriteMatrixByColumn(int startPointRange, int startPointColumn, string args, string elementDelimiter, string columnDelimiter)
        {
            int offset = startPointRange - 1;
            BaseHelper baseH = new BaseHelper();
            string[] Columns = baseH.explode(columnDelimiter, args);
            for (int i = 0; i < Columns.Count(); i++)
            {
                WriteColumn(startPointColumn + i, Columns[i], elementDelimiter, offset);
            }
        }
        public void CleanByMatrix(int startPointRange, int startPointColumn, int rangeOffset, int columnOffset)
        {
            for (int i = 0; i < columnOffset + 1; i++)
            {
                for (int j = 0; j < rangeOffset + 1; j++)
                {
                    WriteCell(startPointRange + j, startPointColumn + i, "");
                }
            }
        }
        #endregion
        #region file setting
        public void Save()
        {
            wb.Save();
        }
        public void Close()
        {
            wb.Close(true, this.path, null);
            excel.Quit();
            this.CloseProcess();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        public void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }
        #endregion
    }
}
