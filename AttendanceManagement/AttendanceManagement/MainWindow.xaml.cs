using System;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Threading;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using AttendanceManagement;



namespace AttendanceApp
{
    public partial class MainWindow : Window
    {
        private DateTime? startTime;  // 出勤時間を保持
        private DateTime? endTime;    // 退勤時間を保持

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            txtSystemTime.Text = DateTime.Now.ToString("yyyy/MMMM/dd HH:mm");

            // 現在日時タイマー
            DispatcherTimer timer = new DispatcherTimer();
            timer.Tick += Timer_Tick;
            timer.Interval = new TimeSpan(1); // 1秒ごとに行う

            // タイマースタート
            timer.Start();
        }

        /// <summary>
        //　現在時刻更新イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer_Tick(object sender, EventArgs e)
        {
            // 日付ラベル表示を変更
            this.txtSystemTime.Text = DateTime.Now.ToString("yyyy/MMMM/dd HH:mm");

            // 日付変更時処理
            if(DateTime.Now.ToString("HH:mm").Equals("00:00"))
            {
                // ボタン制御、ラベル表示を初期化
                btnStart.IsEnabled = true;
                btnEnd.IsEnabled = false;
                lblMessage.Content = "";
            }
        }

        /// <summary>
        /// 出勤ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            // 出勤時間表示
            startTime = DateTime.Now;
            lblStartTime.Text = startTime?.ToString("HH:mm:ss");

            // 退勤時間クリア
            lblEndTime.Text = "";

            // 出勤・退勤ボタン制御
            this.btnStart.IsEnabled = false;
            this.btnEnd.IsEnabled = true;

            // メッセージクリア
            lblMessage.Content = "";

            // 勤務時間クリア
            lblWorkHours.Text = "";
            
        }

        /// <summary>
        /// 退勤ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEnd_Click(object sender, RoutedEventArgs e)
        {
            // memo:ボタン制御したからこの部分は必要ないはず・・・
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

            // 出勤・退勤ボタン制御
            this.btnStart.IsEnabled = true;
            this.btnEnd.IsEnabled = false;

            // メッセージ表示
            lblMessage.Content = "お疲れ様でした！また明日も頑張りましょう！";
        }

        /// <summary>
        /// 設定ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSetting_Click(object sender, RoutedEventArgs e)
        {
            // 設定画面を開く
            SettingWindow setting = new SettingWindow();
            setting.ShowDialog();
        }
        

        /// <summary>
        /// エクセル出力メソッド
        /// </summary>
        private void SaveToExcel()
        {
            // エクセルファイルは年ごとに作られる
            // シートは月ごと

            // Excelファイル名
            string filePath = $"勤怠データ_{DateTime.Now:yyyy}.xlsx";
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

            //MessageBox.Show("勤怠データがExcelに保存されました。");
        }

    }
}
