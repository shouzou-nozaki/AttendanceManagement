using System;
using System.Windows;
using OfficeOpenXml;
using System.IO;

using System.Windows.Threading;
using AttendanceManagement;
using AttendanceManagement.dao;

using AttendanceManagement.Model;
using Xceed.Wpf.Toolkit.Core.Converters;


namespace AttendanceApp
{
    public partial class MainWindow : Window
    {
        private AttendanceInfo AttendanceInfo { get; set; } = new AttendanceInfo();

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

            // 勤怠情報取得
            var attendanceInfoSerialier = new AttendanceInfoSerializer();
            this.AttendanceInfo = attendanceInfoSerialier.GetAttendanceInfo();

            RePaint();
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
                // 勤怠情報を初期化
                var attendanceInfoSerializer = new AttendanceInfoSerializer();
                attendanceInfoSerializer.SetAttendanceInfo(new AttendanceInfo());
            }
        }

        /// <summary>
        /// 出勤ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            // 勤怠情報更新
            this.AttendanceInfo.StartTime = DateTime.Now.ToString("HH:mm");
            this.AttendanceInfo.EndTime = "";
            this.AttendanceInfo.WorkTime = "";
            this.AttendanceInfo.Message = "";

            //画面再描画
            RePaint();

            // 勤怠情報登録
            var attendanceInfoSerialier = new AttendanceInfoSerializer();
            attendanceInfoSerialier.SetAttendanceInfo(this.AttendanceInfo);
        }

        /// <summary>
        /// 退勤ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEnd_Click(object sender, RoutedEventArgs e)
        {
            // 勤怠情報更新
            this.AttendanceInfo.EndTime = DateTime.Now.ToString("HH:mm");

            TimeSpan workDuration = DateTime.Parse(this.AttendanceInfo.EndTime) - DateTime.Parse(this.AttendanceInfo.StartTime);
            this.AttendanceInfo.WorkTime = workDuration.Hours.ToString("00") + ":" + workDuration.Minutes.ToString("00");
            this.AttendanceInfo.Message = "お疲れ様でした！また明日も頑張りましょう！";

            // 画面再描画
            RePaint();

            // エクセル出力
            var settingInfoSerializer = new SettingInfoSerializer();
            var settingInfo = settingInfoSerializer.GetSettingInfo();

            var excelFileName = $"勤怠データ_{DateTime.Now:yyyy}({settingInfo.UserName}).xlsx";

            // ライセンス設定
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel操作オブジェクト作成
            IExcelOperation excelOperation = new CreateNewExcel();

            if (File.Exists(Path.Combine(settingInfo.ExcelFilePath, excelFileName))) excelOperation = new UpdateExcel();

            // Excel作成
            excelOperation.SaveToExcel(settingInfo, this.AttendanceInfo, excelFileName);

            // 勤怠情報初期化
            var attendanceInfoSerializer = new AttendanceInfoSerializer();
            attendanceInfoSerializer.SetAttendanceInfo(new AttendanceInfo());
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
        /// 画面再描画
        /// </summary>
        private void RePaint()
        {
            lblStartTime.Text = this.AttendanceInfo.StartTime;
            lblEndTime.Text = this.AttendanceInfo.EndTime;
            lblWorkHours.Text = this.AttendanceInfo.WorkTime;
            lblMessage.Content = this.AttendanceInfo.Message;

            // 出勤・退勤時間ともになし
            if (this.AttendanceInfo.StartTime.Equals("") && this.AttendanceInfo.EndTime.Equals(""))
            {
                btnStart.IsEnabled = true;
                btnEnd.IsEnabled = false;
                return;
            }
            // 出勤時間のみあり
            if (!this.AttendanceInfo.StartTime.Equals("") && this.AttendanceInfo.EndTime.Equals(""))
            {
                btnStart.IsEnabled = false;
                btnEnd.IsEnabled = true;
                return;
                
            }
            // 出勤・退勤時間ともにあり
            if(!this.AttendanceInfo.StartTime.Equals("") && !this.AttendanceInfo.EndTime.Equals(""))
            {
                btnStart.IsEnabled = true;
                btnEnd.IsEnabled = false;
                return;
            }
        }

    }
}
