using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Windows;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using AttendanceManagement.dao;
using OfficeOpenXml.SystemDrawing.Text;

namespace AttendanceManagement
{
    public partial class SettingWindow : Window
    {
        public SettingInfo SettingInfo { get; private set; }
        public string UserName { get; private set; }       // 利用者名
        public string StartTime { get; private set; }      // 始業時間
        public string EndTime { get; private set; }        // 終業時間
        public int BreakTime { get; private set; }         // 休憩時間（分）
        public string ExcelFilePath { get; private set; }  // Excel出力先

        public SettingWindow()
        {
            try
            {
                InitializeComponent();

                // 設定情報ファイル名
                string settingFile = @"settingInfo.xml";

                // 設定情報がない場合は処理を抜ける
                if (!File.Exists(settingFile)) return;

                // 設定情報デシリアライズ
                // XmlSerializerオブジェクトを作成
                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingInfo));
                // 読み込むファイルを開く
                StreamReader sr = new StreamReader(settingFile, new System.Text.UTF8Encoding(false));
                // XMLファイルから読み込み、デシリアライズする
                this.SettingInfo = (SettingInfo)serializer.Deserialize(sr);

                // 画面に値を入れる
                txtUserName.Text = this.SettingInfo.UserName;
                txtStartTime.Text = this.SettingInfo.StartTime;
                txtEndTime.Text = this.SettingInfo.EndTime;
                txtBreakFrom.Text = this.SettingInfo.BreakFrom;
                txtBreakTo.Text = this.SettingInfo.BreakTo;
                txtExcelPath.Text = this.SettingInfo.ExcelFilePath;

                //ファイルを閉じる
                sr.Close();
            }
            catch(Exception ex)
            {
                // ログ出力
                
            }
           
        }

        /// <summary>
        /// 参照ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (CommonOpenFileDialog cofd = new CommonOpenFileDialog())
            {
                // フォルダを選択できるようにする
                cofd.IsFolderPicker = true;

                if (cofd.ShowDialog() == CommonFileDialogResult.Ok)
                {

                    txtExcelPath.Text = cofd.FileName;
                }
            }

        }

        /// <summary>
        /// 保存ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                // 休憩時間が整数値であるか確認


                // 設定情報をシリアライズ
                String settingFile = @"settingInfo.xml";

                SettingInfo obj = new SettingInfo();
                obj.UserName = txtUserName.Text;        // 利用者名
                obj.StartTime = txtStartTime.Text;      // 始業時間
                obj.EndTime = txtEndTime.Text;          // 終業時間
                obj.BreakFrom = txtBreakFrom.Text;      // 休憩時間(カラ)
                obj.BreakTo = txtBreakTo.Text;          // 休憩時間(マデ)
                obj.ExcelFilePath = txtExcelPath.Text;  // Excel出力先

                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingInfo));

                //書き込むファイルを開く（UTF-8 BOM無し）
                System.IO.StreamWriter sw = new System.IO.StreamWriter(settingFile, false, new System.Text.UTF8Encoding(false));
                //シリアル化し、XMLファイルに保存する
                serializer.Serialize(sw, obj);
                //ファイルを閉じる
                sw.Close();

                // グローバル変数に設定
                this.SettingInfo = obj;

                this.DialogResult = true;  // ウィンドウを閉じるときに結果を返す
                this.Close();
            }
            catch(Exception ex)
            {
                // ログ出力
            }
            
        }
    }
}
