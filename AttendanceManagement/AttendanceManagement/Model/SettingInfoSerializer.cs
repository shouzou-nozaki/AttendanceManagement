using AttendanceManagement.dao;
using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;

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
            XmlSerializer serializer = null;
            StreamReader sr = null;
            try
            {
                // 設定情報がない場合は、設定情報を新規作成
                if (!File.Exists(this.SettingFile)) return new SettingInfo();

                // XmlSerializerオブジェクトを作成
                serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingInfo));
                // 読み込むファイルを開く
                sr = new StreamReader(this.SettingFile, new UTF8Encoding(false));

                // デシリアライズした設定情報を返す
                return (SettingInfo)serializer.Deserialize(sr);

            }catch(Exception ex)
            {
                // ログ出力
                Console.WriteLine(ex.Message);

                // 新規設定情報を返す
                return new SettingInfo();
            }
            finally
            {
                if (sr != null) sr.Close();

            }
        }

        /// <summary>
        /// 設定情報保存
        /// </summary>
        /// <param name="settingInfo">設定情報</param>
        public void SetSettingInfo(SettingInfo settingInfo)
        {
            XmlSerializer serializer = null;
            StreamWriter sw = null;

            try
            {
                // 設定情報をシリアライズ
                serializer = new XmlSerializer(typeof(SettingInfo));
                //書き込むファイルを開く（UTF-8 BOM無し）
                sw = new StreamWriter(this.SettingFile, false, new UTF8Encoding(false));
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
