using OfficeOpenXml;
using System;

namespace TqkLibrary.ExcelSupport.Attributes
{
    public class SheetIndexAttribute : Attribute
    {
        public SheetIndexAttribute(int index)
        {
            this.Index = index;
        }
        public SheetIndexAttribute(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            this.Name = name;
        }

        public int? Index { get; }
        public string? Name { get; }

        public override string ToString()
        {
            return Index.HasValue ? Index.Value.ToString() : Name!;
        }

        public ExcelWorksheet GetSheet(ExcelWorksheets worksheets)
        {
            ExcelWorksheet? excelWorksheet = null;
            if (!string.IsNullOrWhiteSpace(Name))
            {
                excelWorksheet = worksheets[Name];
            }
            else
            {
                excelWorksheet = worksheets[Index!.Value];//IndexOutOfRangeException
            }
            if (excelWorksheet is null)
                throw new InvalidOperationException($"Sheet '{this.ToString()}' not found");

            return excelWorksheet;
        }
    }
}
