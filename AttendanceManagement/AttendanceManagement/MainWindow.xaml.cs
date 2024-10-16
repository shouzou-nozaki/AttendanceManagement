using System;
using System.Windows;

namespace AttendanceApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            // 出勤時間を表示
            lblStartTime.Content = "出勤時間: " + DateTime.Now.ToString("HH:mm:ss");
        }

        private void btnEnd_Click(object sender, RoutedEventArgs e)
        {
            // 退勤時間を表示
            lblEndTime.Content = "退勤時間: " + DateTime.Now.ToString("HH:mm:ss");
        }
    }
}
