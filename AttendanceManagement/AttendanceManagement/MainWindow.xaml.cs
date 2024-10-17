using System;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using System.Linq;



namespace AttendanceApp
{
    public partial class MainWindow : Window
    {
        private DateTime? startTime;  // 出勤時間を保持
        private DateTime? endTime;    // 退勤時間を保持


        public MainWindow()
        {
            InitializeComponent();
        }

        // 出勤ボタンがクリックされたときの処理
        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            startTime = DateTime.Now;
            lblStartTime.Text = startTime?.ToString("HH:mm:ss");
        }

        // 退勤ボタンがクリックされたときの処理
        private void btnEnd_Click(object sender, RoutedEventArgs e)
        {
            if (startTime == null)
            {
                MessageBox.Show("まず出勤を打刻してください！");
                return;
            }

            endTime = DateTime.Now;
            lblEndTime.Text = endTime?.ToString("HH:mm:ss");

            // 勤務時間を計算して表示
            TimeSpan workDuration = endTime.Value - startTime.Value;
            lblWorkHours.Text = workDuration.TotalHours.ToString("F2") + " 時間";

            // エクセル出力
            SaveToExcel();
        }


        /// <summary>
        /// エクセル出力メソッド
        /// </summary>
        private void SaveToExcel()
        {
            // Excelファイルパス
            string filePath = "勤怠データ.xlsx";
            FileInfo file = new FileInfo(filePath);

            // ライセンス設定
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                // シートがなければ作成
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("勤怠データ");

                int row = worksheet.Dimension?.Rows + 1 ?? 1;

                // 出勤・退勤時間を書き込み
                worksheet.Cells[row, 1].Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Cells[row, 2].Value = startTime?.ToString("HH:mm:ss");
                worksheet.Cells[row, 3].Value = endTime?.ToString("HH:mm:ss");
                worksheet.Cells[row, 4].Value = lblWorkHours.Text;

                // Excelファイルの保存
                package.Save();
            }

            MessageBox.Show("勤怠データがExcelに保存されました。");
        }

    }
}
