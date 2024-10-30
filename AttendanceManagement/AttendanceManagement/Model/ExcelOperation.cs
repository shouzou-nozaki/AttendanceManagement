using System;
using System.IO;

using AttendanceManagement.dao;
using OfficeOpenXml;

namespace AttendanceManagement.Model
{
    /// <summary>
    /// インターフェイス
    /// </summary>
    public interface IExcelOperation
    {

        void SaveToExcel(SettingInfo settingInfo, AttendanceInfo attendanceInfo, string excelFileName);
    }

    /// <summary>
    /// Excel操作ビジネスクラス(親クラス)
    /// </summary>
    public abstract class ExcelOperation : IExcelOperation
    {
        public string SheetName_ThisMonth { get; set; } = $"{DateTime.Now:yyyy}_{DateTime.Now:MM}";

        /// <summary>
        /// クラス別実装部
        /// </summary>
        public abstract void SaveToExcel(SettingInfo settingInfo, AttendanceInfo attendanceInfo, string excelFileName);

        /// <summary>
        /// テンプレートExcelコピー処理
        /// </summary>
        /// <param name="settingInfo">設定情報</param>
        public void CopyTemplate(SettingInfo settingInfo, string excelFileName)
        {
            // TODO:本番用に変更
            //var templateFolder = @"C:\Users\soro0\work\program\AttendanceManagement\AttendanceManagement\AttendanceManagement\AttendanceManagement\Template";
            var templateFolder = AppDomain.CurrentDomain.BaseDirectory;
            var templateFile = "template.xlsx";
            File.Copy($@"{Path.Combine(templateFolder, templateFile)}", $@"{Path.Combine(settingInfo.ExcelFilePath, excelFileName)}");
        }



        /// <summary>
        /// Excel基本フレーム作成処理
        /// </summary>
        /// <param name="package">Excelパッケージ</param>
        /// <param name="sheetName_thisMonth">今月分シート</param>
        /// <param name="settingInfo">設定情報</param>
        /// <returns>曜日等の基本フレームを記述したExcelWorkSheet</returns>
        public ExcelWorksheet MakeExcelFrame(ExcelPackage package, SettingInfo settingInfo)
        {
            // 今月分ワークシート読み込み
            var workSheets_thisMonth = package.Workbook.Worksheets[this.SheetName_ThisMonth];

            // テンプレートシートを取得
            var workSheets_template = package.Workbook.Worksheets["template"];

            // 今月分としてコピー
            workSheets_thisMonth = package.Workbook.Worksheets.Add(this.SheetName_ThisMonth, workSheets_template);


            var year = DateTime.Now.ToString("yyyy");
            var month = DateTime.Now.ToString("MM");

            // フレーム作成
            workSheets_thisMonth.Cells["A3"].Value = year + "年";          // 年
            workSheets_thisMonth.Cells["C3"].Value = month + "月";         // 月
            workSheets_thisMonth.Cells["H3"].Value = settingInfo.UserName; // 名前
            workSheets_thisMonth.Cells["G37"].Formula = "SUM(G6:G36)*24";  // 合計
            // 曜日情報の入力
            for (int day = 1; day <= DateTime.DaysInMonth(int.Parse(year), int.Parse(month)); day++)
            {
                var dayOfWeek = new DateTime(int.Parse(year), int.Parse(month), day);

                workSheets_thisMonth.Cells[$"B{5 + day}"].Value = dayOfWeek.ToString("ddd");
            }

            return workSheets_thisMonth;

        }


        /// <summary>
        /// 勤怠情報書き込み処理
        /// </summary>
        /// <param name="package">Excelパッケージ</param>
        /// <param name="settingInfo">設定情報</param>
        /// <param name="attendanceInfo">勤怠情報</param>
        /// <returns>Excelワークシート</returns>
        public ExcelWorksheet WriteAttendanceInfo(ExcelPackage package, SettingInfo settingInfo, AttendanceInfo attendanceInfo, string excelFileName)
        {
            // ワークシート読み込み
            var workSheets_thisMonth = package.Workbook.Worksheets[this.SheetName_ThisMonth] ?? MakeExcelFrame(package, settingInfo);

            // 実稼働時間を取得(基本は、退勤時間 - 始業時間)
            var actualWorkTime = (DateTime.Parse(attendanceInfo.EndTime) - DateTime.Parse(settingInfo.StartTime_Comp)).ToString(@"hh\:mm");
            
            // 遅刻したときは、退勤時間 - 出勤時間
            if(DateTime.Parse(settingInfo.StartTime_Comp) < DateTime.Parse(attendanceInfo.StartTime)) actualWorkTime = attendanceInfo.WorkTime;


            // 出勤時間と退勤時間の間に、休憩時間があるときは実稼働時間として休憩分を引く
            var breakTime = (DateTime.Parse(settingInfo.BreakTo) - DateTime.Parse(settingInfo.BreakFrom)).ToString(@"hh\:mm");

            if (DateTime.Parse(attendanceInfo.StartTime) <= DateTime.Parse(settingInfo.BreakFrom) &&
                DateTime.Parse(settingInfo.BreakTo) <= DateTime.Parse(attendanceInfo.EndTime))
            {
                actualWorkTime = (DateTime.Parse(actualWorkTime) - DateTime.Parse(breakTime)).ToString(@"hh\:mm");
            }
            

            // 始業時刻より前の打刻はしない
            //var startTime = (DateTime.Parse(attendanceInfo.StartTime) < DateTime.Parse(settingInfo.StartTime_Comp)) ? settingInfo.StartTime_Comp : attendanceInfo.StartTime;

            // 勤怠情報書き込み
            workSheets_thisMonth.Cells[$"C{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = DateTime.Parse(attendanceInfo.StartTime); // 出勤時間
            workSheets_thisMonth.Cells[$"D{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = DateTime.Parse(attendanceInfo.EndTime);   // 退勤時間
            workSheets_thisMonth.Cells[$"E{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = DateTime.Parse(attendanceInfo.WorkTime);  // 勤務時間
            workSheets_thisMonth.Cells[$"F{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = DateTime.Parse(breakTime);                // 休憩時間
            workSheets_thisMonth.Cells[$"G{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = DateTime.Parse(actualWorkTime);           // 実稼働時間

            return workSheets_thisMonth;

        }


        /// <summary>
        /// Excel保存処理
        /// </summary>
        /// <param name="package">Excelパッケージ</param>
        /// <param name="excelFileName">Excelファイル名</param>
        public void Save(ExcelPackage package, SettingInfo settingInfo, string excelFileName)
        {
            package.SaveAs(Path.Combine(settingInfo.ExcelFilePath, excelFileName));
        }
    }


    /// <summary>
    /// 新規Excel作成クラス(子クラス)
    /// </summary>
    public class CreateNewExcel : ExcelOperation
    {

        public override void SaveToExcel(SettingInfo settingInfo, AttendanceInfo attendanceInfo, string excelFileName)
        {
            try
            {
                // テンプレートコピー
                CopyTemplate(settingInfo, excelFileName);

                using (ExcelPackage package = new ExcelPackage(Path.Combine(settingInfo.ExcelFilePath, excelFileName)))
                {
                    // Excel基本フレーム作成
                    var workSheet = MakeExcelFrame(package, settingInfo);
                    // 勤怠情報書き込み
                    workSheet = WriteAttendanceInfo(package, settingInfo, attendanceInfo, excelFileName);
                    // 保存
                    Save(package, settingInfo, excelFileName);

                }
            }
            catch (Exception ex) 
            { 
               Console.WriteLine(ex.Message);
            }
            

        }
    }

    /// <summary>
    /// 既存Excel更新クラス(子クラス)
    /// </summary>
    public class UpdateExcel : ExcelOperation
    {
        public override void SaveToExcel(SettingInfo settingInfo, AttendanceInfo attendanceInfo, string excelFileName)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(Path.Combine(settingInfo.ExcelFilePath, excelFileName)))
                {
                    // 勤怠情報書き込み
                    var workSheet = WriteAttendanceInfo(package, settingInfo, attendanceInfo, excelFileName);
                    // 保存
                    Save(package, settingInfo, excelFileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           
        }
    }

}
