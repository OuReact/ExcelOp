using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelOp.Data
{
    public class Span
    {
        const string EXCELE_DATE_TIME_FORMAT = "yyyy-MM-dd";

        const DateTimeStyles EXCELE_DATE_TIME_STYLE = DateTimeStyles.None;

        private readonly static CultureInfo  EXCELE_CULTURE_INFO = CultureInfo.CurrentCulture;

        public DateTime startDate { get; private set; } = DateTime.Now;


        public DateTime endDate { get; private set; } = DateTime.Now;

        public string labelText { get; private set; } = "";


        /// <summary>
        /// 基本的には使用することがないが、テストなどに使用する。
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        public Span(DateTime startDate , DateTime endDate) 
        {
            
            this.startDate = startDate;
            this.endDate = endDate;
        
        }

        private Span()
        {


        }

        public static bool TryGetSpan(IXLWorksheet sheet, int targetRowIndex, int startDateColumnIndex, int endDateColumnIndex, string spanLabelText, out Span span)
        {
            bool ret = false;
            span = new Span();

            if(sheet == null)
            {
                Log.logger.Error("sheetがnullのため、取得できません。");
                return false;

            }

            try
            {

                // rowは数値の部分、 columnはアルファベットの部分 2=B。
                IXLCell cell = sheet.Cell(targetRowIndex, startDateColumnIndex);

                if (cell.Value.IsBlank)
                {
                    Log.logger.Info(string.Format("シート名 {0}の {1}:{2}  は空データです", sheet?.Name ?? "", targetRowIndex, startDateColumnIndex));
                    return false;
                }

                if (cell.Value.IsDateTime == false)
                {
                    Log.logger.Error( string.Format("シート名 {0}の {1}:{2}  は日付のデータではありません", sheet?.Name??"" , targetRowIndex, startDateColumnIndex));
                    return false;
                }

                span.startDate = cell.Value.GetDateTime();


                cell = sheet.Cell(targetRowIndex, endDateColumnIndex);

                if (cell.Value.IsDateTime == false)
                {
                    Log.logger.Error(string.Format("シート名 {0}の {1}:{2}  は日付のデータではありません", sheet?.Name ?? "", targetRowIndex, endDateColumnIndex));
                    return false;
                }
                span.endDate = cell.Value.GetDateTime();

                if (span.endDate < span.startDate)
                {
                    Log.logger.Error(string.Format("シート名 {0}の {1}:{2}  の開始日データが　終了日よりも未来の日付になってます", sheet?.Name ?? "", targetRowIndex, startDateColumnIndex));
                    return false;
                }
                span.labelText = spanLabelText;

                ret = true;
            }
            catch(Exception ex)
            {
                Log.logger.Error("日付取得でエラーが発生しました。 "+ex.Message);

            }


            return ret;
        }


        public static bool TryGetSpanList(IXLWorksheet sheet, int startRowIndex ,int startDateColumnIndex, int endDateColumnIndex, out List<Span> spanList)
        {

            spanList = new List<Span>();

            if (sheet == null)
            {
                Log.logger.Error("sheetがnullのため、取得できません。");
                return false;

            }

            for ( int i= startRowIndex; ;i++)
            {

                if(TryGetSpan(sheet,i, startDateColumnIndex, endDateColumnIndex,"",out Span span) == false)
                {
                    break;
                }
                spanList.Add(span);
            }
            
            return true;
        }

        /// <summary>
        /// 開始日と終了日が設けてある　spacクラスの　重複する期間日を計算する。
        /// 自インスタンスが開始日： 1/1、終了日：1/20 、　targetSpan が 開始日:1/10 終了日：1/30 であれば、 1/10 ～1/20のため、 11という数値になる。
        /// </summary>
        /// <param name="targetSpan"></param>
        /// <param name="span2"></param>
        /// <param name="spanDay"></param>
        /// <returns></returns>
        public bool TryGetSpacDay(Span targetSpan,out int spanDay)
        {

            spanDay = 0;
            if (targetSpan == null )
            {
                return false;
            }

            DateTime maxStart =　this.startDate > targetSpan.startDate ? startDate : targetSpan.startDate;
            DateTime minEnd = this.endDate < targetSpan.endDate ? this.endDate : targetSpan.endDate;

            if (maxStart < minEnd)
            {
                TimeSpan overlap = minEnd - maxStart;
                spanDay = (int)overlap.TotalDays + 1;
            }
            else
            {
                //期間外でもエラーにならない
                spanDay = 0;
            }


            return true;


        }
       

    }
}
