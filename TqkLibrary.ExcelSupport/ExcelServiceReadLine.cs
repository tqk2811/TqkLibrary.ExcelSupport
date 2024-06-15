using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OfficeOpenXml;
using System.Reflection;
using TqkLibrary.ExcelSupport.Attributes;

namespace TqkLibrary.ExcelSupport
{
    public class ExcelServiceReadLine : BaseExcelService
    {
        protected readonly Dictionary<Type, int> _dict_startLineIndex = new();
        public ExcelServiceReadLine(string filePath) : base(filePath)
        {
        }


        public override async Task ResetAsync(CancellationToken cancellationToken = default)
        {
            using var l = await _asyncLock.LockAsync(cancellationToken);
            _dict_startLineIndex.Clear();
        }

        public virtual async Task<T?> GetDataAsync<T>(CancellationToken cancellationToken = default) where T : BaseData, new()
        {
            using var l = await _asyncLock.LockAsync(cancellationToken);
            return await _RunInTask(() => _GetDataAsync<T>(cancellationToken));
        }


        protected virtual T? _GetDataAsync<T>(CancellationToken cancellationToken = default) where T : BaseData, new()
        {
            SheetIndexAttribute? sheetIndexAttribute = typeof(T).GetCustomAttribute<SheetIndexAttribute>();
            if (sheetIndexAttribute is null)
                throw new InvalidOperationException($"'{typeof(T).FullName}' must contain attribute {nameof(SheetIndexAttribute)}");

            using ExcelPackage package = new ExcelPackage(_filePath);
            ExcelWorksheet excelWorksheet = sheetIndexAttribute.GetSheet(package.Workbook.Worksheets);

            if (!_dict_startLineIndex.ContainsKey(typeof(T)))
                _dict_startLineIndex[typeof(T)] = -1;

            for (int i = Math.Max(_dict_startLineIndex[typeof(T)], excelWorksheet.Rows.StartRow + 1); i < excelWorksheet.Rows.EndRow; i++)
            {
                _dict_startLineIndex[typeof(T)] = i + 1;

                T? instance = _ReadRow<T>(excelWorksheet, i, false, out bool isEmptyLine);
                if (instance is null)
                    continue;
                return instance;
            }

            return null;
        }
    }
}
