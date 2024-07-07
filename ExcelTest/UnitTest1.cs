using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelOp.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace ExcelTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Assert.AreEqual(5, 5);

            // 読み込み (パス指定)
            XLWorkbook readByPathWb = new XLWorkbook("test.xlsx");
            // 読み込み (実データ指定)
            IXLWorksheet readByNameWs = readByPathWb.Worksheet("日付テスト");
            // 保存 (パス保存)
            // 保存 (stream保存)
            // rowは数値の部分、 columnはアルファベットの部分 2=B。
            IXLCell cell = readByNameWs.Cell(5, 2);
            Assert.AreEqual(true, cell.Value.IsDateTime);

            DateTime targetDate = DateTime.ParseExact("2024/3/1", "yyyy/M/d", CultureInfo.CurrentCulture, DateTimeStyles.None);
            Assert.AreEqual(targetDate.Date, cell.Value.GetDateTime().Date);


            Assert.AreEqual(true,ExcelOp.Data.Span.TryGetSpanList(readByNameWs, 4,2,3, out List<ExcelOp.Data.Span> spanList ));
            Assert.AreEqual(19, spanList.Count);

            Assert.AreEqual(spanList[0].startDate.Date, DateTime.Parse("2024/1/1"));
            Assert.AreEqual(spanList[0].endDate.Date, DateTime.Parse("2024/1/2"));
            Assert.AreEqual(spanList[1].startDate.Date, DateTime.Parse("2024/3/1"));
            Assert.AreEqual(spanList[1].endDate.Date, DateTime.Parse("2024/4/2"));

            Assert.AreEqual(true, spanList[0].TryGetSpacDay(spanList[2],out int spanDay) );
            Assert.AreEqual(2, spanDay);


            Assert.AreEqual(true, spanList[0].TryGetSpacDay(spanList[1], out  spanDay));
            Assert.AreEqual(0, spanDay);


            // 書き込み

            IXLWorksheet writeTest = readByPathWb.Worksheet("書き込みテスト");

            writeTest.Clear();

            ExcelOp.Data.Span testSpan = new Span(DateTime.Parse("2024/01/10"), DateTime.Parse("2024/02/10"));
            int writeIndex = 5;
            int columnIndex = 3;

            IXLCell writeCell = writeTest.Cell(writeIndex, columnIndex);
            writeCell.SetValue("テストデータ書き込み開始");

            foreach (  ExcelOp.Data.Span span in spanList)
            {
                writeIndex++;
                writeCell = writeTest.Cell(writeIndex, columnIndex);
                testSpan.TryGetSpacDay(span,out  spanDay);
                writeCell.SetValue(spanDay);

            }
            readByPathWb.Save();

            readByPathWb.Dispose();

            string location = System.Reflection.Assembly.GetEntryAssembly().Location;

            string dirPath = Path.GetDirectoryName(location);


            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                Arguments = dirPath,
                FileName = "explorer.exe"
            };

            Process.Start(startInfo);
        }

    
    }
}
