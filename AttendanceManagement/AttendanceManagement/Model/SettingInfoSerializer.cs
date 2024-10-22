using AttendanceManagement.dao;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceManagement.Model
{
    public class SettingInfoSerializer
    {

        private string SettingFile { get; } = "settingInfo.xml"; // 設定ファイル名

        /// <summary>
        /// 設定情報取得
        /// </summary>
        /// <returns>設定情報</returns>
        public SettingInfo GetSettingInfo()
        {
            try
            {
                // 設定情報がない場合は、設定情報を新規作成
                if (!File.Exists(this.SettingFile)) return new SettingInfo();

                // XmlSerializerオブジェクトを作成
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingInfo));
                // 読み込むファイルを開く
                var sr = new StreamReader(this.SettingFile, new UTF8Encoding(false));

                // デシリアライズした設定情報を返す
                return (SettingInfo)serializer.Deserialize(sr);

            }catch(Exception ex)
            {
                // ログ出力
                Console.WriteLine(ex.Message);

                // 新規設定情報を返す
                return new SettingInfo();
            }
        }

        /// <summary>
        /// 設定情報保存
        /// </summary>
        /// <param name="settingInfo">設定情報</param>
        public void SetSettingInfo(SettingInfo settingInfo)
        {
            // 設定情報をシリアライズ
            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingInfo));
            //書き込むファイルを開く（UTF-8 BOM無し）
            var sw = new StreamWriter(this.SettingFile, false, new UTF8Encoding(false));

            try
            {
                //シリアル化し、XMLファイルに保存する
                serializer.Serialize(sw, settingInfo);
            }
            catch(Exception ex)
            {
                // ログ出力
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //ファイルを閉じる
                sw.Close();
            }
            
        }
    }
}
