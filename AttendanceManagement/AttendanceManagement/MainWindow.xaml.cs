using System;
using System.Windows;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Threading;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using AttendanceManagement;
using AttendanceManagement.dao;


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
            var startTime = DateTime.Now;
            lblStartTime.Text = startTime.ToString("HH:mm:ss");

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
            lblEndTime.Text = endTime.ToString("HH:mm:ss");

            // 勤務時間を計算して表示
            TimeSpan workDuration = DateTime.Parse(lblEndTime.Text) - DateTime.Parse(lblStartTime.Text);
            lblWorkHours.Text = workDuration.Hours.ToString("00") + "時間" + workDuration.Minutes.ToString("00") + "分";

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


            // 設定情報ファイル名
            string settingFile = @"settingInfo.xml";
            SettingInfo settingInfo = new SettingInfo();

            // 設定情報ファイルがあればデシリアライズ
            if (File.Exists(settingFile))
            {
                // 設定情報デシリアライズ
                // XmlSerializerオブジェクトを作成
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingInfo));
                // 読み込むファイルを開く
                var sr = new StreamReader(settingFile, new System.Text.UTF8Encoding(false));
                // デシリアライズ内容を設定情報にセット
                settingInfo = (SettingInfo)serializer.Deserialize(sr);
            }

            // Excelファイル名
            var fileName = $"勤怠データ_{DateTime.Now:yyyy}({settingInfo.UserName}).xlsx";

            // Excelファイルがないとき
            if(!File.Exists(Path.Combine(settingInfo.ExcelFilePath, fileName)))
            {
                // テンプレートファイル TODO:本番用にパスを変更
                fileName = "C:\\Users\\soro0\\work\\program\\AttendanceManagement\\AttendanceManagement\\AttendanceManagement\\AttendanceManagement\\Template\\template.xlsx";
                
            }

            // ファイル情報読み込み
            FileInfo file = new FileInfo(fileName);

            // ライセンス設定
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                // ワークシートを取得
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                // シートがなければ作成
                //worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add($"{DateTime.Now:yyyy}_{DateTime.Now:MM}");

                // 出勤・退勤時間を書き込み

                fileName = $"勤怠データ_{DateTime.Now:yyyy}({settingInfo.UserName}).xlsx";

                // Excelファイルの保存
                package.SaveAs(Path.Combine(settingInfo.ExcelFilePath, fileName));
            }

        }

    }
}
