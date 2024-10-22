using System;
using System.Windows;
using OfficeOpenXml;
using System.IO;

using System.Windows.Threading;
using AttendanceManagement;
using AttendanceManagement.dao;

using AttendanceManagement.Model;


namespace AttendanceApp
{
    public partial class MainWindow : Window
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            // 現在日時タイマー
            txtSystemTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            var timer = new DispatcherTimer();
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
            this.txtSystemTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            // 日付変更時処理
            if (DateTime.Now.ToString("HH:mm").Equals("00:00"))
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
            var startTime = DateTime.Now;
            lblStartTime.Text = startTime.ToString("HH:mm");

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
            // 現在時刻を表示
            var endTime = DateTime.Now;
            lblEndTime.Text = endTime.ToString("HH:mm");

            // 勤務時間を計算して表示
            TimeSpan workDuration = DateTime.Parse(lblEndTime.Text) - DateTime.Parse(lblStartTime.Text);
            lblWorkHours.Text = workDuration.Hours.ToString("00") + ":" + workDuration.Minutes.ToString("00");

            // 勤怠情報クラスの作成
            var attendanceInfo = new AttendanceInfo();
            attendanceInfo.StartTime = lblStartTime.Text;
            attendanceInfo.EndTime = lblEndTime.Text;
            attendanceInfo.WorkTime = lblWorkHours.Text;

            // エクセル出力
            var settingInfoSerializer = new SettingInfoSerializer();
            var settingInfo = settingInfoSerializer.GetSettingInfo();

            var excelFileName = $"勤怠データ_{DateTime.Now:yyyy}({settingInfo.UserName}).xlsx";

            // ライセンス設定
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            IExcelOperation excelOperation = null;

            if (File.Exists(Path.Combine(settingInfo.ExcelFilePath, excelFileName)))
            {
                excelOperation = new UpdateExcel();
            }
            else
            {
                excelOperation = new CreateNewExcel();
            }
            excelOperation.SaveToExcel(settingInfo, attendanceInfo, excelFileName);

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

    }
}
