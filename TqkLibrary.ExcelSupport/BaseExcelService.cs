using Nito.AsyncEx;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TqkLibrary.ExcelSupport.Attributes;
using TqkLibrary.ExcelSupport.Enums;

namespace TqkLibrary.ExcelSupport
{
    public abstract partial class BaseExcelService
    {
        protected readonly AsyncLock _asyncLock = new AsyncLock();
        protected readonly string _filePath;
        public BaseExcelService(string filePath)
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException(filePath);
            this._filePath = filePath;
        }

        public virtual Task ResetAsync(CancellationToken cancellationToken = default)
        {
            return Task.CompletedTask;
        }

        public virtual async Task<IReadOnlyList<T>> GetDatasAsync<T>(
            bool isReadAll = false,
            bool stopAtEmptyLine = false,
            CancellationToken cancellationToken = default) where T : BaseData, new()
        {
            using var l = await _asyncLock.LockAsync(cancellationToken);
            return _GetDatas<T>(isReadAll, stopAtEmptyLine);
        }

        public virtual async Task SaveDataAsync<T>(T data, CancellationToken cancellationToken = default) where T : BaseData, new()
        {
            using var l = await _asyncLock.LockAsync(cancellationToken);
            SheetIndexAttribute? sheetIndexAttribute = typeof(T).GetCustomAttribute<SheetIndexAttribute>();
            if (sheetIndexAttribute is null)
                throw new InvalidOperationException($"'{typeof(T).FullName}' must contain attribute {nameof(SheetIndexAttribute)}");

            using ExcelPackage package = new ExcelPackage(_filePath);
            ExcelWorksheet excelWorksheet = sheetIndexAttribute.GetSheet(package.Workbook.Worksheets);

            bool isChanged = false;
            if (excelWorksheet is not null)
            {
                foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
                {
                    ColAttribute? colAttribute = propertyInfo.GetCustomAttribute<ColAttribute>();
                    if (colAttribute is not null && colAttribute.Flag.HasFlag(ColFlag.IsUpdateBack))
                    {
                        object? pData = propertyInfo.GetValue(data);
                        if (pData is not null)
                        {
                            excelWorksheet.Cells[$"{colAttribute.Col}{data.LineIndex}"].Value = pData;
                            isChanged = true;
                        }
                    }
                }
            }

            if (isChanged)
                await package.SaveAsync();
        }



        protected virtual IReadOnlyList<T> _GetDatas<T>(
            bool isReadAll = false,
            bool stopAtEmptyLine = false
            ) where T : BaseData, new()
        {
            SheetIndexAttribute? sheetIndexAttribute = typeof(T).GetCustomAttribute<SheetIndexAttribute>();
            if (sheetIndexAttribute is null)
                throw new InvalidOperationException($"'{typeof(T).FullName}' must contain attribute {nameof(SheetIndexAttribute)}");

            using ExcelPackage package = new ExcelPackage(_filePath);
            ExcelWorksheet excelWorksheet = sheetIndexAttribute.GetSheet(package.Workbook.Worksheets);

            List<T> values = new List<T>();
            if (excelWorksheet is not null)
            {
                for (int i = excelWorksheet.Rows.StartRow + 1; i < excelWorksheet.Rows.EndRow; i++)
                {
                    T? instance = ReadRow<T>(excelWorksheet, i, isReadAll, out bool isEmptyLine);
                    if (instance is not null)
                        values.Add(instance);
                    else if (stopAtEmptyLine && isEmptyLine)
                        break;
                }
            }
            return values;
        }
        protected virtual T? ReadRow<T>(ExcelWorksheet excelWorksheet, int lineIndex, bool isReadAll, out bool isEmptyLine) where T : BaseData, new()
        {
            bool isSkip = false;
            isEmptyLine = true;

            T instance = new T();
            instance.LineIndex = lineIndex;
            instance.ExcelFilePath = _filePath;
            foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
            {
                ColAttribute? colAttribute = propertyInfo.GetCustomAttribute<ColAttribute>();
                if (colAttribute is not null)
                {
                    string? data = excelWorksheet.Cells[$"{colAttribute.Col}{lineIndex}"].Value?.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(data))
                    {
                        if (!isReadAll && colAttribute.Flag.HasFlag(ColFlag.SkipReadLineIfCell_Empty))
                        {
                            isSkip = true;
                            break;
                        }
                    }
                    else
                    {
                        if (!isReadAll && colAttribute.Flag.HasFlag(ColFlag.SkipReadLineIfCell_NotEmpty))
                        {
                            isSkip = true;
                            break;
                        }
                        propertyInfo.SetValue(instance, data);
                        isEmptyLine = false;
                    }
                }
            }

            if (isSkip)
                return null;
            if (isEmptyLine)
                return null;

            return instance;
        }

        public abstract class BaseData
        {
            public virtual int LineIndex { get; set; }
            public virtual string? ExcelFilePath { get; set; }
        }
    }
}
