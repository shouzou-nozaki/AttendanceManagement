using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Windows;

using AttendanceManagement.dao;
using AttendanceManagement.Model;


namespace AttendanceManagement
{
    public partial class SettingWindow : Window
    {

        public SettingWindow()
        {
            try
            {
                InitializeComponent();

                // 設定情報取得
                var settingInfoSerializer = new SettingInfoSerializer();
                var settingInfo = settingInfoSerializer.GetSettingInfo();

                // 画面に値を入れる
                txtUserName.Text  = settingInfo.UserName;
                txtStartTime.Text = settingInfo.StartTime_Comp;
                txtEndTime.Text   = settingInfo.EndTime_Comp;
                txtBreakFrom.Text = settingInfo.BreakFrom;
                txtBreakTo.Text   = settingInfo.BreakTo;
                txtExcelPath.Text = settingInfo.ExcelFilePath;

            }
            catch(Exception ex)
            {
                // ログ出力
                Console.WriteLine(ex.Message);
            }
           
        }

        /// <summary>
        /// 参照ボタン押下イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (CommonOpenFileDialog cofd = new CommonOpenFileDialog())
                {
                    // フォルダを選択できるようにする
                    cofd.IsFolderPicker = true;

                    // 選択されたフォルダパスを画面に表示
                    if (cofd.ShowDialog() == CommonFileDialogResult.Ok) txtExcelPath.Text = cofd.FileName;
                }
            }
            catch (Exception ex) 
            { 
                // ログ出力
                Console.WriteLine(ex.Message);
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
                // 設定情報をシリアライズ
                var settingInfo = new SettingInfo();
                settingInfo.UserName      = txtUserName.Text;  // 利用者名
                settingInfo.StartTime_Comp     = txtStartTime.Text; // 始業時間
                settingInfo.EndTime_Comp 　    = txtEndTime.Text;   // 終業時間
                settingInfo.BreakFrom     = txtBreakFrom.Text; // 休憩時間(カラ)
                settingInfo.BreakTo       = txtBreakTo.Text;   // 休憩時間(マデ)
                settingInfo.ExcelFilePath = txtExcelPath.Text; // Excel出力先

                var settingInfoSerializer = new SettingInfoSerializer();
                settingInfoSerializer.SetSettingInfo(settingInfo);

                this.DialogResult = true;  // ウィンドウを閉じるときに結果を返す
                this.Close();
            }
            catch(Exception ex)
            {
                // ログ出力
                Console.WriteLine(ex.Message);
            }
            
        }
    }
}
