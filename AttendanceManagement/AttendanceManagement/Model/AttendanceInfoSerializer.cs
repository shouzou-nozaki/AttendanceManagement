using AttendanceManagement.dao;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace AttendanceManagement.Model
{
    public class AttendanceInfoSerializer
    {

        private string AttendanceFile { get; } = "attendanceInfo.xml"; // 設定ファイル名

        /// <summary>
        /// 設定情報取得
        /// </summary>
        /// <returns>設定情報</returns>
        public AttendanceInfo GetAttendanceInfo()
        {
            XmlSerializer serializer = null;
            StreamReader sr = null;
            try
            {
                // 設定情報がない場合は、設定情報を新規作成
                if (!File.Exists(this.AttendanceFile)) return new AttendanceInfo();

                // XmlSerializerオブジェクトを作成
                serializer = new System.Xml.Serialization.XmlSerializer(typeof(AttendanceInfo));
                // 読み込むファイルを開く
                sr = new StreamReader(this.AttendanceFile, new UTF8Encoding(false));

                // デシリアライズした設定情報を返す
                return (AttendanceInfo)serializer.Deserialize(sr);

            }
            catch (Exception ex)
            {
                // ログ出力
                Console.WriteLine(ex.Message);

                // 新規設定情報を返す
                return new AttendanceInfo();
            }
            finally
            {
                if(sr != null) sr.Close();

            }
        }

        /// <summary>
        /// 設定情報保存
        /// </summary>
        /// <param name="settingInfo">設定情報</param>
        public void SetAttendanceInfo(AttendanceInfo attendanceInfo)
        {
            // 設定情報をシリアライズ
            XmlSerializer serializer = null;
            //書き込むファイルを開く（UTF-8 BOM無し）
            StreamWriter sw = null;

            try
            {
                // 設定情報をシリアライズ
                serializer = new XmlSerializer(typeof(AttendanceInfo));
                //書き込むファイルを開く（UTF-8 BOM無し）
                sw = new StreamWriter(this.AttendanceFile, false, new UTF8Encoding(false));
                //シリアル化し、XMLファイルに保存する
                serializer.Serialize(sw, attendanceInfo);
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
