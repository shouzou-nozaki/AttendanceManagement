using System;
using System.IO;
using System.Windows;

namespace AttendanceApp
{
    public partial class ScheduleSettings : Window
    {
        public ScheduleSettings()
        {
            InitializeComponent();
        }

        private void SaveSchedule_Click(object sender, RoutedEventArgs e)
        {
            string startTime = txtStartTime.Text;
            string endTime = txtEndTime.Text;

            // ファイルにスケジュールを保存する例（簡単なテキストファイルとして保存）
            string schedulePath = "schedule.txt";
            File.WriteAllText(schedulePath, $"勤務開始時間: {startTime}\n勤務終了時間: {endTime}");

            MessageBox.Show("スケジュールが保存されました！");
        }
    }
}
