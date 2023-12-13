//  * @file 檔案名稱    : 
//  * @author 作者名稱  : 
//  * @date 初版日期    : 
//  * @date 更新日期    : 20190111
//  * @版本             : 1.0.1
//  * @brief 檔案描述   : LOG工具&显示工具
//  * @                 : 更新只显示不写LOG功能

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace PLC_PC_Signal_TEST
{
    public class LogTool
    {
        /*========================================================================*/
        /*    變數宣告                                                             */
        /*========================================================================*/
        private System.Windows.Forms.ListBox m_ListBox;
        private System.Windows.Forms.MethodInvoker m_MethodInvoker; //跨线程委托
        private string      m_StrNowFileName;
        private StreamWriter m_ObjFile;
        private string      m_DirectorPath;
        private TimeSpan    m_ReservedDate;
        private Boolean     m_TimeFlag;
        private string      m_LogMessage;
        private object      m_LockResourse;
        private int m_ItemsCount;

        /*========================================================================*/
        /*    Constructor                                                         */
        /*========================================================================*/

        public LogTool(string DirectPath, int reservedDate)
        {
            string fileName;
            fileName = DateTime.Now.ToString("yyyyMMdd");
            if (!(Directory.Exists(DirectPath)))
            {
                Directory.CreateDirectory(DirectPath);
            }
            m_DirectorPath = DirectPath;
            m_ReservedDate = new TimeSpan(reservedDate, 0, 0, 0);
            m_StrNowFileName = fileName;
            m_TimeFlag = true;
            m_ObjFile = new StreamWriter(DirectPath + "\\" + m_StrNowFileName + ".log", true);
            m_MethodInvoker = new System.Windows.Forms.MethodInvoker(UpdateGUI);
            m_LockResourse = new object();
            m_ItemsCount = 300;
        }


        public LogTool()
        {          
            m_TimeFlag = true;
            m_ItemsCount = 300;
            m_MethodInvoker = new System.Windows.Forms.MethodInvoker(UpdateGUI);
        }

        /*========================================================================*/
        /*    變數存取                                                             */
        /*========================================================================*/
        public Boolean TimerTag
        {
            get { return this.m_TimeFlag; }
            set { this.m_TimeFlag = value; }
        }
        public int ReservedDate
        {
            get { return this.m_ReservedDate.Days; }
            //set { this.m_ReservedDate.TotalDays  = value; }
        }
        public int ItemsCount
        {
            get { return m_ItemsCount; }
            set { m_ItemsCount = value; }
        }

        /*========================================================================*/
        /*    Function                                                            */
        /*========================================================================*/
        public void ListViewLog(ListBox ListBoxLog)
        {
            m_ListBox = ListBoxLog;
        }
        private void UpdateGUI() 
        {
            if (m_ListBox != null)
            {
                if (m_ListBox.Items.Count == 0)
                {
                    m_ListBox.Items.Add(m_LogMessage);
                }
                else
                {
                    if (m_ListBox.Items.Count > m_ItemsCount)
                    {
                        m_ListBox.Items.RemoveAt(m_ListBox.Items.Count - 1);
                    }
                    m_ListBox.Items.Insert(0, m_LogMessage);
                }
            }
        } 
        private void CheckFileName() 
        {
            string fileName;
            fileName = DateTime.Now.ToString("yyyyMMdd");
            if (m_StrNowFileName.Equals(fileName)==false) 
            {
                
                m_StrNowFileName = fileName;
                m_ObjFile.Close();
                m_ObjFile.Dispose();
                CheckReservedDate();
                m_ObjFile = new StreamWriter(m_DirectorPath + "\\" + m_StrNowFileName + ".log", true); 
            } 
        }
        private void CheckReservedDate() 
        {
            DateTime DeleteDate = DateTime.Now.Subtract(m_ReservedDate);
            if (File.Exists(m_DirectorPath + "\\" + DeleteDate.ToString("yyyyMMdd") + ".log")) 
            {
                File.Delete(m_DirectorPath + "\\" + DeleteDate.ToString("yyyyMMdd") + ".log");
            }
 
        }


        public void LogPrintf(string msg, params string[] args)  //传多个值
        {
            CheckFileName();
            lock (m_LockResourse)
            {
                if (this.m_TimeFlag)
                {
                    m_LogMessage = String.Format("[{0}] ", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss:ff"));
                }
                //m_ObjFile.WriteLine(msg, args);
                m_LogMessage = m_LogMessage + String.Format(msg, args);
                m_ObjFile.WriteLine(m_LogMessage);
                m_ObjFile.Flush();
                if (m_ListBox != null) 
                {
                    m_ListBox.Invoke(m_MethodInvoker);
                }
            }
        }
        public void Printf(string msg, params string[] args)  //传多个值
        {
         
                if (this.m_TimeFlag)
                {
                    m_LogMessage = String.Format("[{0}] ", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                }

                 m_LogMessage = m_LogMessage + String.Format(msg, args);
                if (m_ListBox != null)
                {
                    m_ListBox.Invoke(m_MethodInvoker);
                }
            
        }


        public void Show(string msg, params string[] args)
        {
            
            CheckFileName();
            lock (m_LockResourse)
            {
                if (this.m_TimeFlag)
                {
                    m_LogMessage = String.Format("[{0}] ", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                }
                //m_ObjFile.WriteLine(msg, args);
                m_LogMessage = m_LogMessage + String.Format(msg, args);
                m_ObjFile.WriteLine(m_LogMessage);
                m_ObjFile.Flush();
                if (m_ListBox != null)
                {
                    m_ListBox.Invoke(m_MethodInvoker);
                }
                MessageBox.Show(m_LogMessage);
            }
        }
        
        public void Close() 
        {
            m_ObjFile.Close();
        }
    }
}
