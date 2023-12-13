using HslCommunication.Profinet.FATEK.Helper;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using Org.BouncyCastle.Utilities;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Metadata.Profiles.Iptc;
using HslCommunication;
using HslCommunication.Core;
using NPOI.XSSF.Streaming.Values;

namespace PLC_PC_Signal_TEST
{
    public partial class Form1 : Form
    {
        DataTable d1 = new DataTable();
        private LogTool m_LogTool = new LogTool(Application.StartupPath + "\\LogFile", 90); //创建主界面的log 

        public Form1()
        {
            InitializeComponent();
        }
        private HslCommunication.Profinet.Melsec.MelsecMcServer mcNetServer;

        private void button3_Click(object sender, EventArgs e)
        {
            mcNetServer?.ServerClose();
            m_LogTool.LogPrintf("服务器已关闭");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                mcNetServer = new HslCommunication.Profinet.Melsec.MelsecMcServer(true);                       // 实例化对象
                mcNetServer.ActiveTimeSpan = TimeSpan.FromHours(1);
                mcNetServer.OnDataReceived += MelsecMcServer_OnDataReceived;
                mcNetServer.ServerStart(Convert.ToInt32(textBox2.Text));
               // mcNetServer.OnDataSend += MelsecMcServer_OnDataReceived;
                byte[] byteArray = System.Text.Encoding.Default.GetBytes("1234");
                byte[] bytes = BitConverter.GetBytes(123);//将int32转换为字节数组
                mcNetServer.Write("M100", bytes);
                m_LogTool.LogPrintf("服务器已开启");



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void MelsecMcServer_OnDataReceived(object sender, object source, byte[] receive)
        {
            // 我们可以捕获到接收到的客户端的modbus报文
            // 如果是TCP接收的
            if (source is HslCommunication.Core.Net.AppSession session)
            {
                // 获取当前客户的IP地址
                string ip = session.IpAddress;



            }

            // 如果是串口接收的
            if (source is System.IO.Ports.SerialPort serialPort)
            {
                // 获取当前的串口的名称
                string portName = serialPort.PortName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m_LogTool.ListViewLog(listBox1);
            string filenanme = Application.StartupPath + "\\" + "Test.xlsx";
            d1 = ReadExcelToDataTable(filenanme, "Sheet1", true);

            dataGridView1.DataSource = d1;
        }


        /// <summary>
        /// 将excel文件内容读取到DataTable数据表中
        /// </summary>
        /// <param name="fileName">文件完整路径名</param>
        /// <param name="sheetName">指定读取excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名：true=是，false=否</param>
        /// <returns>DataTable数据表</returns>
        public static DataTable ReadExcelToDataTable(string fileName, string sheetName = null, bool isFirstRowColumn = true)
        {
            //定义要返回的datatable对象
            DataTable data = new DataTable();
            //excel工作表
            ISheet sheet = null;
            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                if (!File.Exists(fileName))
                {
                    throw new Exception("文件不存在");
                }
                //根据指定路径读取文件
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                //根据文件流创建excel数据结构
                IWorkbook workbook = null;
                var fileType = Path.GetExtension(fileName).ToLower();
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                #region 判断Excel版本
                switch (fileType)
                {
                    //.XLSX是07版(或者07以上的)的Office Excel
                    case ".xlsx":
                        workbook = new XSSFWorkbook(fs);
                        break;
                    //.XLS是03版的Office Excel
                    case ".xls":
                        workbook = new HSSFWorkbook(fs);
                        break;
                    default:
                        throw new Exception("Excel文档格式有误");
                }
                #endregion

                //IWorkbook workbook = new HSSFWorkbook(fs);
                //如果有指定工作表名称
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheet = workbook.GetSheet(sheetName);
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (sheet == null)
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    //如果没有指定的sheetName，则尝试获取第一个sheet
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;
                    //如果第一行是标题列名
                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }
                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null || row.FirstCellNum < 0) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();

                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {

                            //if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            //    dataRow[j] = row.GetCell(j);

                            #region 格式转换 NPOI获取Excel单元格中不同类型的数据
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {

                                //获取指定的单元格信息

                                switch (cell.CellType)
                                {
                                    //首先在NPOI中数字和日期都属于Numeric类型
                                    //通过NPOI中自带的DateUtil.IsCellDateFormatted判断是否为时间日期类型
                                    case CellType.Numeric when DateUtil.IsCellDateFormatted(cell):
                                        dataRow[j] = cell.DateCellValue;
                                        break;
                                    case CellType.Numeric:
                                        //其他数字类型
                                        dataRow[j] = cell.NumericCellValue;
                                        break;
                                    //空数据类型
                                    case CellType.Blank:
                                        dataRow[j] = "";
                                        break;
                                    //公式类型
                                    case CellType.Formula:
                                        {
                                            HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(workbook);
                                            dataRow[j] = eva.Evaluate(cell).StringValue;
                                            break;
                                        }
                                    //布尔类型
                                    case CellType.Boolean:
                                        dataRow[j] = row.GetCell(j).BooleanCellValue;
                                        break;
                                    //错误
                                    case CellType.Error:
                                        // dataRow[j] = HSSF Constants.GetText(row.GetCell(j).ErrorCellValue);
                                        break;
                                    //其他类型都按字符串类型来处理（未知类型CellType.Unknown，字符串类型CellType.String）
                                    default:
                                        dataRow[j] = cell.StringCellValue;
                                        break;
                                }

                            }
                            #endregion
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!bg_Main.IsBusy)
            {


                bg_Main.RunWorkerAsync();

                this.button2.Text = "运行中";
            }
            else
            {
                bg_Main.CancelAsync();
                this.button2.Text = "开始运行";
            }
        }
        public static byte[] ObjectToBytes(object obj)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                return ms.GetBuffer();
            }
        }
        private void bg_Main_DoWork(object sender, DoWorkEventArgs e)
        {
            int A = 1;
            //while (!bg_Main.CancellationPending)
            //{
            Thread.Sleep(10);


            for (int i = 0; i < d1.Rows.Count; i++)
            {

                if (d1.Rows[i][0].ToString() == A.ToString())//确认运行步骤
                {
                    if (Convert.ToBoolean(d1.Rows[i][1].ToString()))//确认是输入（true）还是输出
                    {
                        bool waitsignal = false;
                            Type type = Type.GetType(d1.Rows[i][5].ToString());
                        while (!bg_Main.CancellationPending&& !waitsignal)
                        {
                            Thread.Sleep(100);
                            if (type == typeof(int))
                            {
                                if (int.TryParse(d1.Rows[i][3].ToString(), out int value))
                                {
                                    OperateResult<int>  m_Result = mcNetServer.ReadInt32(d1.Rows[i][2].ToString());

                                    if (m_Result.IsSuccess)
                                    {
                                        if (m_Result.Content == value)
                                        {
                                            waitsignal = true;
                                            //读取后复位
                                            if (int.TryParse(d1.Rows[i][6].ToString(), out int initvalue))
                                                mcNetServer.Write(d1.Rows[i][2].ToString(), initvalue);
                                            m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "读取" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");
                                            //Thread.Sleep(100);
                                        }
                                        else 
                                        {
                                            bg_Main.ReportProgress(i,"等待"+ d1.Rows[i][2].ToString()+"值为"+ value.ToString());
                                        }
                                        //Thread.Sleep(100);
                                    }
                                }
                            }
                            else if (type == typeof(string))
                            {
                                OperateResult<string> m_Result = mcNetServer.ReadString(d1.Rows[i][2].ToString(),10);
                                string value = d1.Rows[i][3].ToString();
                                if (m_Result.IsSuccess)
                                {
                                    if (m_Result.Content == value)
                                    {
                                        waitsignal = true;
                                        //读取后复位
                                        mcNetServer.Write(d1.Rows[i][2].ToString(), d1.Rows[i][6].ToString());
                                        m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "读取" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");

                                    }
                                    else
                                    {
                                        bg_Main.ReportProgress(i, "等待" + d1.Rows[i][2].ToString() + "值为" + value.ToString());
                                    }
                                    //Thread.Sleep(100);
                                }
                            }
                            else if (type == typeof(double))
                            {
                                if (double.TryParse(d1.Rows[i][3].ToString(), out double value))
                                {
                                    OperateResult<double> m_Result = mcNetServer.ReadDouble(d1.Rows[i][2].ToString());

                                    if (m_Result.IsSuccess)
                                    {
                                        if (m_Result.Content == value)
                                        {
                                            waitsignal = true;
                                            //读取后复位
                                            if (double.TryParse(d1.Rows[i][6].ToString(), out double initvalue))
                                                mcNetServer.Write(d1.Rows[i][2].ToString(), initvalue);
                                            m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "读取" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");

                                        }
                                        else
                                        {
                                            bg_Main.ReportProgress(i, "等待" + d1.Rows[i][2].ToString() + "值为" + value.ToString());
                                        }
                                        //Thread.Sleep(100);
                                    }
                                }

                            }
                            else if (type == typeof(bool))
                            {

                                if (bool.TryParse(d1.Rows[i][3].ToString(), out bool value))
                                {

                                    OperateResult<Boolean> m_Result = mcNetServer.ReadBool(d1.Rows[i][2].ToString());

                                    if (m_Result.IsSuccess)
                                    {
                                        if (m_Result.Content == value)
                                        {
                                            waitsignal = true;
                                            //读取后复位
                                            if (bool.TryParse(d1.Rows[i][6].ToString(), out bool initvalue))
                                                mcNetServer.Write(d1.Rows[i][2].ToString(), initvalue);
                                            m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "读取" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");


                                        }
                                        else
                                        {
                                            bg_Main.ReportProgress(i, "等待" + d1.Rows[i][2].ToString() + "值为" + value.ToString());
                                        }
                                        //Thread.Sleep(100);
                                    }
                                }

                            }

                        }
                    }
                    else
                    {
                 
                        Type type = Type.GetType(d1.Rows[i][5].ToString());

                        if (type == typeof(int))
                        {
                            if (int.TryParse(d1.Rows[i][3].ToString(), out int value))
                            {
                                OperateResult m_Result = mcNetServer.Write(d1.Rows[i][2].ToString(), value);

                                if (m_Result.IsSuccess)
                                {
                                    m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "写入" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");
                                    //Thread.Sleep(100);
                                }
                            }
                        }
                        else if (type == typeof(string))
                        {
                            string stringValue = d1.Rows[i][3].ToString();
                            byte[] byteArray = System.Text.Encoding.Default.GetBytes(stringValue);
                            OperateResult m_Result = mcNetServer.Write(d1.Rows[i][2].ToString(), byteArray);
                            if (m_Result.IsSuccess)
                            {
                                m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "写入" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");
                                //Thread.Sleep(100);
                            }
                            //  Console.WriteLine("Converted string value: " + stringValue);
                        }
                        else if (type == typeof(double))
                        {
                            if (double.TryParse(d1.Rows[i][3].ToString(), out double value))
                            {
                                OperateResult m_Result = mcNetServer.Write(d1.Rows[i][2].ToString(), value);
                                if (m_Result.IsSuccess)
                                {
                                    m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "写入" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");
                                 //   Thread.Sleep(100);
                                }
                            }

                        }
                        else if (type == typeof(bool))
                        {

                            if (bool.TryParse(d1.Rows[i][3].ToString(), out bool value))
                            {
                                OperateResult m_Result = mcNetServer.Write(d1.Rows[i][2].ToString(), value);
                                if (m_Result.IsSuccess)
                                {
                                    m_LogTool.LogPrintf("PLC在地址" + d1.Rows[i][2].ToString() + "写入" + d1.Rows[i][5].ToString() + "类型值" + d1.Rows[i][3].ToString() + "成功");
                                   // Thread.Sleep(100);
                                }

                            }

                        }
                    }
                }
                if (A < d1.Rows.Count)
                {
                    A++;
                }
                else
                {
                    A = 1;
                }

            }

            bg_Main.ReportProgress(100, "");
            //}

        }

        private void bg_Main_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
          
            if (e.ProgressPercentage == 100)
            {



                label1.Text = "流程结束";
                this.button2.Text = "开始运行";
            }
            else
            {
                label1.Text = e.UserState.ToString();
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void 保存表格ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filenanme = Application.StartupPath + "\\" + "Test.xlsx";
        //    d1 = ReadExcelToDataTable(filenanme, "Sheet1", true);
          //  string filenanme = DCpath + "\\" + listBox1.SelectedItem.ToString();
            TableToExcel(d1, "Sheet1", filenanme);
        }

        /// <summary>
        /// Datable导出成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file">导出路径(包括文件名与扩展名)</param>
        public static void TableToExcel(DataTable dt1, string sheetName1, string file)
        {
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else if (fileExt == ".xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                workbook = null;
            }
            if (workbook == null)
            {
                return;
            }



            ISheet sheet1 = workbook.CreateSheet(sheetName1);
          //  ISheet sheet2 = workbook.CreateSheet(sheetName2);

            //表头  
            IRow row1 = sheet1.CreateRow(0);
            for (int i = 0; i < dt1.Columns.Count; i++)
            {
                ICell cell = row1.CreateCell(i);
                cell.SetCellValue(dt1.Columns[i].ColumnName);

            }

            //数据  
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                IRow row2 = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dt1.Columns.Count; j++)
                {
                    ICell cell = row2.CreateCell(j);
                    cell.SetCellValue(dt1.Rows[i][j].ToString());
                }
            }

         

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            //Thread.Sleep(1000); 
            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }
    }
}
