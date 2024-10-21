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
using System.Data;
using System.Windows.Data;
using Xceed.Wpf.Toolkit.PropertyGrid.Editors;
using Xceed.Wpf.Toolkit.Primitives;


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
            try
            {
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
                var excelFileName = $"勤怠データ_{DateTime.Now:yyyy}({settingInfo.UserName}).xlsx";

                // ライセンス設定
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // 対象Excelファイルが存在するとき
                if (File.Exists(Path.Combine(settingInfo.ExcelFilePath, excelFileName)))
                {
                    // 既存Excelファイルの更新
                    UpdateExcelFile(settingInfo, excelFileName);
                    return;
                }

                // 新規Excelファイル作成
                CreateNewExcelFile(settingInfo, excelFileName);

            }
            catch(Exception ex)
            {
                // ログ出力
            }

        }

        /// <summary>
        /// 既存Excelファイル更新メソッド
        /// </summary>
        /// <param name="settingInfo"></param>
        /// <param name="excelFileName"></param>
        private void UpdateExcelFile(SettingInfo settingInfo, string excelFileName)
        {
            using (ExcelPackage package = new ExcelPackage(Path.Combine(settingInfo.ExcelFilePath, excelFileName)))
            {
                // 今月分のシート名
                var sheetName_thisMonth = $"{DateTime.Now:yyyy}_{DateTime.Now:MM}";

                // ワークシート読み込み
                var workSheets_thisMonth = package.Workbook.Worksheets[sheetName_thisMonth];

                // 今月分のシートがないときは新しく作成　
                if (workSheets_thisMonth == null)
                {
                    // テンプレートシートを取得
                    var workSheets_template = package.Workbook.Worksheets["template"];

                    // 今月分としてコピー
                    workSheets_thisMonth = package.Workbook.Worksheets.Add(sheetName_thisMonth, workSheets_template);

                    // Excelの基本フレームの設定
                    workSheets_thisMonth = MakeExcelFrame(package, sheetName_thisMonth, settingInfo);
                }

                // 休憩時間を取得
                var breakTime = (DateTime.Parse(settingInfo.BreakTo) - DateTime.Parse(settingInfo.BreakFrom)).ToString(@"hh\:mm");

                // 実稼働時間を取得
                // 基本は、出勤時間 - 退勤時間
                var actualWorkTime = lblWorkHours.Text;

                // 出勤時間と退勤時間の間に、休憩時間があるときは実稼働時間として休憩分を引く
                if(DateTime.Parse(lblStartTime.Text) <= DateTime.Parse(settingInfo.BreakFrom) && DateTime.Parse(settingInfo.BreakTo) <= DateTime.Parse(lblEndTime.Text))
                {
                    actualWorkTime = (DateTime.Parse(lblWorkHours.Text) - DateTime.Parse(breakTime.ToString())).ToString();
                }

                // 勤怠情報書き込み
                workSheets_thisMonth.Cells[$"C{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = lblStartTime.Text;    // 出勤時間
                workSheets_thisMonth.Cells[$"D{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = lblEndTime.Text;      // 退勤時間
                workSheets_thisMonth.Cells[$"E{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = lblWorkHours.Text;    // 勤務時間(計)
                workSheets_thisMonth.Cells[$"F{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = breakTime.ToString(); // 休憩時間
                workSheets_thisMonth.Cells[$"G{5 + int.Parse(DateTime.Now.ToString("dd"))}"].Value = actualWorkTime;       // 実稼働時間


                // Excelファイルの保存
                package.SaveAs(Path.Combine(settingInfo.ExcelFilePath, excelFileName));
            }
        }

        /// <summary>
        /// 新規Excelファイル作成メソッド
        /// </summary>
        /// <param name="settingInfo"></param>
        /// <param name="excelFileName"></param>
        private void CreateNewExcelFile(SettingInfo settingInfo, string excelFileName)
        {
            using (ExcelPackage package = new ExcelPackage(Path.Combine(settingInfo.ExcelFilePath, excelFileName)))
            {
                // 各パターンごとにExcelパッケージ編集

                // ワークシートを取得
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                // シートがなければ作成
                //worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add($"{DateTime.Now:yyyy}_{DateTime.Now:MM}");



                // Excelファイルの保存
                package.SaveAs(Path.Combine(settingInfo.ExcelFilePath, excelFileName));
            }
        }

        /// <summary>
        /// Excel基本フレーム作成
        /// </summary>
        /// <param name="package">Excelパッケージ</param>
        /// <param name="sheetName_thisMonth">今月分シート</param>
        /// <param name="settingInfo">設定情報</param>
        /// <returns>曜日等の基本フレームを記述したExcelWorkSheet</returns>
        private ExcelWorksheet MakeExcelFrame(ExcelPackage package, string sheetName_thisMonth, SettingInfo settingInfo)
        {
            var workSheets_thisMonth = package.Workbook.Worksheets[sheetName_thisMonth];

            var year = DateTime.Now.ToString("yyyy");
            var month = DateTime.Now.ToString("MM");

            // フレーム作成
            workSheets_thisMonth.Cells["A3"].Value = year + "年";          // 年
            workSheets_thisMonth.Cells["C3"].Value = month + "月";         // 月
            workSheets_thisMonth.Cells["H3"].Value = settingInfo.UserName; // 名前
            workSheets_thisMonth.Cells["G37"].Formula = "SUM(G6:G36)*24";  // 合計
            // 曜日情報の入力
            for(int day = 1; day < DateTime.DaysInMonth(int.Parse(year),int.Parse(month)); day++)
            {
                var dayOfWeek = new DateTime(int.Parse(year), int.Parse(month), day);

                workSheets_thisMonth.Cells[$"B{5+ day}"].Value = dayOfWeek.ToString("ddd"); 
            }


            return workSheets_thisMonth;
        }
    }
}
