using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.dao
{
    public class AttendanceInfo
    {
        public string StartTime { get; set; } = ""; // 出勤時間
        public string EndTime   { get; set; } = ""; // 退勤時間
        public string WorkTime  { get; set; } = ""; // 勤務時間 
        public string Message { get; set; } = "";  // メッセージ 

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public AttendanceInfo()
        {
            this.StartTime = StartTime;
            this.EndTime = EndTime;
            this.WorkTime = WorkTime;
            this.Message = Message;
        }
    }
}
