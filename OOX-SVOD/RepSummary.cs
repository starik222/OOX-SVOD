using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelPrint;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace OOX_SVOD
{
    public class RepSummary
    {
        public List<NamedRangeInfo> Ranges { get; set; }
        private string svodFilePath;
        public RepSummary()
        {
            Ranges = new List<NamedRangeInfo>();
        }
        
        public void OpenSvodExcel(string path)
        {
            svodFilePath = path;
            PrintToExcel excelManager = new PrintToExcel();
            excelManager.Open(path);
            try
            {
                for (int nameIndex = 1; nameIndex < excelManager.Workbook.Names.Count; nameIndex++)
                {
                    if (excelManager.Workbook.Names.Item(nameIndex, Type.Missing, Type.Missing).Name.StartsWith("razdel_"))
                    {
                        var cellRangeName = (string)excelManager.Workbook.Names.Item(nameIndex, Type.Missing, Type.Missing).RefersToLocal;
                        excelManager.CurrentRange = excelManager.Application.get_Range(cellRangeName);
                        NamedRangeInfo rInfo = new NamedRangeInfo();
                        rInfo.Name = excelManager.Workbook.Names.Item(nameIndex, Type.Missing, Type.Missing).Name;
                        rInfo.LocalRangeName = cellRangeName;
                        excelManager.Rows = excelManager.CurrentRange.Rows;
                        excelManager.Columns = excelManager.CurrentRange.Columns;
                        for (int row = 2; row <= excelManager.Rows.Count; row++)
                        {
                            for (int col = 1; col <= excelManager.Columns.Count; col++)
                            {
                                excelManager.CurrentCell = excelManager.CurrentRange[row, col];
                                //cell.Value2 = 1;
                                if (!excelManager.CurrentCell.Locked)
                                {
                                    var nF = (string)excelManager.CurrentCell.NumberFormat;
                                    if (nF.EndsWith("0"))
                                    {
                                        rInfo[row, col] = (float?)excelManager.CurrentCell.Value2;
                                    }
                                }
                                excelManager.ReleaseObject(excelManager.CurrentCell);//Замедляет работу, но без этого excel зависает в процессах...
                            }
                        }
                        Ranges.Add(rInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                excelManager.Finally(true);
                throw;
            }
            excelManager.Finally(true);
        }

        public async Task OpenSvodExcelAsync(string path)
        {
            await Task.Run(() => OpenSvodExcel(path));
        }

        public void AddToSvod(string path)
        {
            PrintToExcel excelManager = new PrintToExcel();
            excelManager.Open(path);
            try
            {
                foreach (var item in Ranges)
                {
                    excelManager.CurrentRange = excelManager.Application.get_Range(item.LocalRangeName);
                    foreach (var svodCell in item.Cells)
                    {

                        excelManager.CurrentCell = excelManager.CurrentRange[svodCell.RowIndex, svodCell.ColumnIndex];
                        if (!excelManager.CurrentCell.Locked)
                        {
                            var nF = (string)excelManager.CurrentCell.NumberFormat;
                            if (nF.EndsWith("0"))
                            {
                                var v = (float?)excelManager.CurrentCell.Value2;
                                if (v != null)
                                {
                                    if (svodCell.Value == null)
                                        svodCell.Value = v;
                                    else
                                        svodCell.Value += v;
                                }
                            }
                            else
                            {
                                throw new Exception("Ошидаемая ячейка заблокирована, возможно, структура файла повреждена или это другая версия файла.");
                            }
                        }
                        else
                        {
                            throw new Exception("Ошидаемая ячейка заблокирована, возможно, структура файла повреждена или это другая версия файла.");
                        }
                        excelManager.ReleaseObject(excelManager.CurrentCell);//Замедляет работу, но без этого excel зависает в процессах...
                    }

                }
            }
            catch (Exception ex)
            {
                excelManager.Finally(true);
                throw;
            }
            excelManager.Finally(true);
        }

        public async Task AddToSvodAsync(string path)
        {
            await Task.Run(() => AddToSvod(path));
        }

        public void SaveSvod(string path)
        {
            PrintToExcel excelManager = new PrintToExcel();
            excelManager.Open(svodFilePath);
            try
            {
                foreach (var item in Ranges)
                {
                    excelManager.CurrentRange = excelManager.Application.get_Range(item.LocalRangeName);
                    foreach (var svodCell in item.Cells)
                    {
                        if (svodCell.Value != null)
                        {
                            excelManager.CurrentCell = excelManager.CurrentRange[svodCell.RowIndex, svodCell.ColumnIndex];
                            excelManager.CurrentCell.Value2 = svodCell.Value;
                            excelManager.ReleaseObject(excelManager.CurrentCell);//Замедляет работу, но без этого excel зависает в процессах...
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                excelManager.Finally(true);
                throw;
            }
            excelManager.Application.DisplayAlerts = false;
            excelManager.Workbook.SaveAs(path);
            excelManager.Finally(true);
        }

        public async Task SaveSvodAsync(string path)
        {
            await Task.Run(() => SaveSvod(path));
        }

        //public class SheetData
        //{
        //    public string Name { get; set; }
        //    public int Index { get; set; }
        //    public List<CellData> Cells { get; set; }
        //}
        //public class CellData
        //{
        //    public string Location { get; set; }
        //    public float Value { get; set; }
        //}

        public class NamedRangeInfo
        {
            public string Name { get; set; }
            public string LocalRangeName { get; set; }
            public List<CellInfo> Cells { get; set; }

            public NamedRangeInfo()
            {
                Cells = new List<CellInfo>();
            }

            public float? this[int row, int column]
            {
                get
                {
                    return Cells?.Find(a => a.RowIndex == row && a.ColumnIndex == column)?.Value;
                }
                set
                {
                    var cell = Cells.Find(a => a.RowIndex == row && a.ColumnIndex == column);
                    if (cell != null)
                    {
                        cell.Value = value;
                    }
                    else
                    {
                        Cells.Add(new CellInfo(row, column, value));
                    }
                }
            }
        }

        public class CellInfo
        {
            public int RowIndex { get; set; }
            public int ColumnIndex { get; set; }
            public float? Value { get; set; }

            public CellInfo(int rowIndex, int columnIndex, float? value)
            {
                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
                Value = value;
            }
        }
    }
}
