using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AttendanceManagement.dao
{
    public class SettingInfo
    {
        public string UserName       { get; set; } = ""; // 利用者名
        public string StartTime_Comp { get; set; } = ""; // 始業時間
        public string EndTime_Comp   { get; set; } = ""; // 終業時間
        public string BreakFrom      { get; set; } = ""; // 休憩時間(カラ)
        public string BreakTo        { get; set; } = ""; // 休憩時間(マデ)
        public string ExcelFilePath  { get; set; } = ""; // Excel出力先

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SettingInfo()
        {
            this.UserName = UserName;
            this.StartTime_Comp = StartTime_Comp;
            this.EndTime_Comp = EndTime_Comp;
            this.BreakFrom = BreakFrom;
            this.BreakTo = BreakTo;
            this.ExcelFilePath = ExcelFilePath;
        }
    }
}
