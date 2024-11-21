using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using ToolAutoBuild.Model;

namespace ToolAutoBuild.CUtils
{
    public static class File
    {
        /// <summary>
        /// Read data from excel file
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="mode"></param>
        /// <returns></returns>
        public static Dictionary<string, ItemModel> ReadExcelToList(string filePath, short mode)
        {
            string key = string.Empty, value = string.Empty, comment = string.Empty;
            try
            {
                Dictionary<string, ItemModel> listItems = new Dictionary<string, ItemModel>();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = mode == 0 ? package.Workbook.Worksheets[1] : package.Workbook.Worksheets[3];
                    var rowCount = worksheet.Dimension.Rows;

                    if (mode == 0)
                    {
                        for (int row = 2; row <= rowCount; row++)
                        {
                            key = worksheet.Cells[row, 2]?.Text;
                            value = worksheet.Cells[row, 3]?.Text;
                            comment = worksheet.Cells[row, 4]?.Text;

                            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                            {
                                listItems.Add(key, new ItemModel(key, value, comment));
                            }
                        }
                    }
                    else
                    {
                        for (int row = 3; row <= rowCount; row++)
                        {
                            key = worksheet.Cells[row, 2]?.Text;
                            value = worksheet.Cells[row, 3]?.Text;
                            comment = worksheet.Cells[row, 4]?.Text;

                            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                            {
                                listItems.Add(key, new ItemModel(key, value, comment));
                            }
                        }
                    }
                }
                return listItems;
            }
            catch
            {
                throw new Exception($"Key: {key} and Value: {value} already exist");
            }
        }

        /// <summary>
        /// Read data from .resx file
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<ItemModel> ReadResxData(string filePath)
        {
            List<ItemModel> listItems = new List<ItemModel>();

            XDocument xmlDoc = XDocument.Load(filePath);
            var dataElements = xmlDoc.Descendants("data");

            foreach (var element in dataElements)
            {
                string key = element.Attribute("name")?.Value;
                string value = element.Element("value")?.Value;
                string comment = element.Element("comment")?.Value;

                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                {
                    listItems.Add(new ItemModel(key, value, comment));
                }
            }
            return listItems;
        }

    }
}