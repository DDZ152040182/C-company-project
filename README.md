# C-company-project
company project 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Threading.Tasks;

using System.Threading;
using Fancy西湖醉鱼君;

//using System.Threading.Tasks;
using System.IO.Ports;
using ZedGraph;
using System.Runtime.InteropServices;
using System.Collections;

namespace LCR
{
    public partial class Form1 : Form
    {
        #region 全局变量

        private float xvalues;//窗体宽度、高度变量
        private float yvalues;

        private int[] baud = new int[8] { 4800, 9600, 14400, 19200, 38400, 56000, 576000, 115200 };//波特率表

        //LCR符号
        string LSign = "H", CSign = "F", RSign = "Ω", FSign ="Hz";  //C代表C  V代表电容 R代表电阻

        //private string[] dataLCR = { };
        private string dataFre = "1000"; //预设值频率值

        private double dataPulse = 0;
        private double calibrationPulse = 0;

        // 数据切花输出
        public static int chooseNm = 0;

        /*数据库*/
        private string fileName;
        private string fileInitName = "Init.txt";
        private string[] fileInit;

        /*记录电机合拢张开位置*/
        private int motorState = 0;

        /*ini 文件配置*/
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);//系统dll导入ini写函数
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);//系统dll导入ini读函数
        string FileName = System.AppDomain.CurrentDomain.BaseDirectory + "data.ini";
        StringBuilder temp = new StringBuilder(255);//存储读出ini内容变量
        string PathFirst = "";

        

        Fancy西湖醉鱼君.Modbus modbusMotor = new Fancy西湖醉鱼君.Modbus();

        //数据库文件
        Fancy西湖醉鱼君.TextOperate txtOperate = new Fancy西湖醉鱼君.TextOperate();
        Fancy西湖醉鱼君.TextOperate txtRecord = new Fancy西湖醉鱼君.TextOperate();

        //access 数据库

        private Fancy西湖醉鱼君.AccessOperate access = new AccessOperate();
        private string tableName;
        private ArrayList cloumsName = new ArrayList();
        private ArrayList dataList = new ArrayList();




        public  string fileRecordNamePath = "C:\\";
        public string fileRecordName = "1";
        public string fileRecord = "C:\\1.txt";
        bool RecordEnble = false;

        GraphPane myPanel;
        
        PointPairList list_L = new PointPairList();
        PointPairList list_C = new PointPairList();
        PointPairList listResistance = new PointPairList();

        LineItem myCurve_L ;
        LineItem myCurve_C; 
        LineItem myCurveResistance;
        // ZedGraph设置      
        double Current=0, Voltage=0, Resistance=0,Frequence= 0;  //C代表C  V代表电容 R代表电阻
        double CurrentMin, CurrentMax, VoltageMin, VoltageMax, ResistanceMin, ResistanceMax;

        //存入原始数据access
        string strLength = "0", strResistanc = "0", strInductance = "0", strCapacitance = "0", strFrequence = "0";



        #endregion
        public Form1()
        {
            InitializeComponent();

            serialPortLCR.DataReceived += new SerialDataReceivedEventHandler(DataReceivedLCR);       //添加串口中断事件
            //当收到数据时，调用DataReceivedLCR函数； serialPortLCR.DataReceived是一个事件（Event）；
            //DataReceivedLCR为编写的事件处理函数，当事件发生时调用；SerialDataReceivedEventHandler 是事件处理上的Delegate（委派）;
            //+= 将事件处理函数，挂接到事件DataReceived 中

            //serialPortMotor.DataReceived += new SerialDataReceivedEventHandler(DataReceivedMotor);       //添加串口中断事件
            //this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(From1_FromClosing);
            
        }


        #region 串口扫描
        private void SearchAndAddSerialToComboBox(SerialPort MyPort, ComboBox MyBox)
        {
            string Buffer;
            MyBox.Items.Clear();
            for (int i = 1; i < 20; i++)
            {
                try
                {
                    Buffer = "COM" + i.ToString();
                    MyPort.PortName = Buffer;
                    MyPort.Open();
                    MyBox.Items.Add(Buffer);
                    MyPort.Close();
                }
                catch
                {
                }
            }

        }

        private void SearchAndAddSerialToComboBox(SerialPort MyPort, ComboBox MyBox, ComboBox MyBox2, ComboBox MyBox3)
        {
            string Buffer;
            MyBox.Items.Clear();
            MyBox2.Items.Clear();
            for (int i = 1; i < 20; i++)
            {
                try
                {
                    Buffer = "COM" + i.ToString();
                    MyPort.PortName = Buffer;
                    MyPort.Open();
                    MyBox.Items.Add(Buffer);
                    MyBox2.Items.Add(Buffer);
                    MyBox3.Items.Add(Buffer);
                    MyPort.Close();
                }
                catch
                {
                }
            }

        }
        #endregion

        #region 自适应屏幕
        private void MainForm_Resize(object sender, EventArgs e)//重绘事件
        {
            float newX = this.Width / xvalues;
            float newY = this.Height / yvalues;
            SetControls(newX, newY, this);
        }


        /// <summary>
        /// 遍历窗体中控件函数
        /// </summary>
        /// <param name="cons"></param>
        private void SetTag(Control cons)
        {
            foreach (Control con in cons.Controls)  //遍历窗体中的控件
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                {
                    SetTag(con);
                }
            }
        }

        /// <summary>
        /// 根据窗体大小调整控件大小
        /// </summary>
        /// <param name="newX"></param>
        /// <param name="newY"></param>
        /// <param name="cons"></param>
        private void SetControls(float newX, float newY, Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                string[] mytag = con.Tag.ToString().Split(new char[] { ':' });
                float a = Convert.ToSingle(mytag[0]) * newX;
                con.Width = (int)a;
                a = Convert.ToSingle(mytag[1]) * newY;
                con.Height = (int)a;
                a = Convert.ToSingle(mytag[2]) * newX;
                con.Left = (int)a;
                a = Convert.ToSingle(mytag[3]) * newY;
                con.Top = (int)a;
                Single currentSize = Convert.ToSingle(mytag[4]) * newY;

                con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                if (con.Controls.Count > 0)
                {
                    SetControls(newX, newY, con);
                }
            }
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Resize += new EventHandler(MainForm_Resize); //添加窗体拉伸重绘事件
            xvalues = this.Width;
            yvalues = this.Height;
            SetTag(this);   // 遍历窗体中控件函数  

            //数据库文件
            textBox3.Text = fileRecordName;
           // label13.Text ="保存路径： "+fileRecordNamePath;

            //创建数据库
            access.FolderName = "databaseAccess";
            access.DatabaseName = "DEP.accdb";//数据库名
            access.CreateFolder();
            if (!access.CreateDatabase())
            {
                MessageBox.Show("数据库创建失败");
                //return;
            }
            cloumsName.Add("时间");
            cloumsName.Add("频率Hz");
            cloumsName.Add("长度mm");
            cloumsName.Add("电阻Ω");
            cloumsName.Add("电感H");
            cloumsName.Add("电容F");

            //串口端口扫描
            SearchAndAddSerialToComboBox(serialPortLCR, comboBoxComLCR, comboBoxMotorCom, comboBoxMotor); 

            GetPrivateProfileString("SerialLCR", "COM", "COM1", temp, 1024, FileName);
            comboBoxComLCR.Text = temp.ToString();
            GetPrivateProfileString("SerialLCR", "Baud", "9600", temp, 1024, FileName);
            comboBoxBaudRate.Text = temp.ToString();

            GetPrivateProfileString("SerialPulse", "COM", "COM1", temp, 1024, FileName);
            comboBoxMotorCom.Text = temp.ToString();
            GetPrivateProfileString("SerialPulse", "Baud", "9600", temp, 1024, FileName);
            comboBoxMotorBaudRate.Text = temp.ToString();

            GetPrivateProfileString("SerialMotor", "COM", "COM1", temp, 1024, FileName);
            comboBoxMotor.Text = temp.ToString();
            GetPrivateProfileString("SerialMotor", "Baud", "9600", temp, 1024, FileName);
            comboBoxMotorRate.Text = temp.ToString();

            GetPrivateProfileString("Pulse", "Frequence", "1000", temp, 1024, FileName);
            comboBoxFre.Text = temp.ToString();



            #region  图表设置

            myPanel = zedGraphControl.GraphPane;             //得到绘图面板

            #region //设置图表标题
            myPanel.Title.IsVisible = true;             //设置图表标题显示            
            myPanel.Title.Text = "冰芯交流固体电导率测定仪（DEP）";
            myPanel.Title.FontSpec.Size = 16;             //设置注释字体大小
            myPanel.Title.FontSpec.Family = "微软雅黑";       //设置字体
            myPanel.Title.FontSpec.FontColor = Color.Teal;//获取系统定义的颜色
            myPanel.Legend.IsVisible = true;            //注释框显示  注释为电阻
            myPanel.Legend.FontSpec.Size = 12;            //设置字体大小
            myPanel.Legend.FontSpec.Family = "微软雅黑"; //设置字体              
            #endregion
            //   AddAxis(); //添加坐标
            #region  AddAxis();XY坐标轴设置
            // 添加xY轴的名称
            myPanel.XAxis.Title.Text = "位移mm"; //设置X轴文字
            myPanel.YAxis.Title.Text = "";//0
            myPanel.AddYAxis(""); //电流值//1
            myPanel.AddYAxis(""); //电压值//2
            myPanel.AddYAxis("主轴，不显示"); //电阻值  //3
            // 添加xY轴的名称的字体大小
            myPanel.XAxis.Scale.FontSpec.Size = 8;        //设置X轴字体大小
            myPanel.YAxisList[0].Scale.FontSpec.Size = 8;  //设置Y轴数字字体大小
            myPanel.YAxisList[1].Scale.FontSpec.Size = 8;
            myPanel.YAxisList[2].Scale.FontSpec.Size = 8;
            myPanel.YAxisList[3].Scale.FontSpec.Size = 8;

            // 设置Y轴显示以及大小 属性
            myPanel.YAxisList[0].IsVisible = true;          //显示电感坐标轴
            myPanel.YAxisList[0].Scale.FontSpec.Size = 9;
            myPanel.YAxisList[0].Scale.FontSpec.Family = "微软雅黑";
            myPanel.YAxisList[1].IsVisible = true;         //显示电压坐标轴
            myPanel.YAxisList[1].Scale.FontSpec.Size = 9;
            myPanel.YAxisList[1].Scale.FontSpec.Family = "微软雅黑";
            myPanel.YAxisList[2].IsVisible = true;         //显示电阻坐标轴以及坐标名
            myPanel.YAxisList[2].Scale.FontSpec.Size = 9;
            myPanel.YAxisList[2].Scale.FontSpec.Family = "微软雅黑";
            myPanel.YAxisList[3].IsVisible = false;        //不显示Y主轴

            //设置XY轴虚线显示
            myPanel.XAxis.MajorGrid.IsVisible = true;    // 显示X轴虚线 
            myPanel.YAxis.MajorGrid.IsVisible = true; // 显示Y轴虚线 //0
            myPanel.YAxisList[1].MajorGrid.IsVisible = true;//1
            myPanel.YAxisList[2].MajorGrid.IsVisible = true;//2
            myPanel.YAxisList[3].MajorGrid.IsVisible = false;//3

            //设置XY轴中间为0的线
            myPanel.XAxis.MajorGrid.IsZeroLine = true; // 显示X轴数据为0的线条
            myPanel.YAxis.MajorGrid.IsZeroLine = true;  // 显示Y1轴数据为0的线条//0
            myPanel.YAxisList[1].MajorGrid.IsZeroLine = false;//1
            myPanel.YAxisList[2].MajorGrid.IsZeroLine = false;//2
            myPanel.YAxisList[3].MajorGrid.IsZeroLine = false;//3

            #endregion

            //  SetColor(); //颜色设置
            #region   SetColor(); //颜色设置
            myPanel.YAxis.Scale.FontSpec.FontColor = buttonLColor.BackColor;
            myPanel.YAxisList[1].Scale.FontSpec.FontColor = buttonCColor.BackColor;
            myPanel.YAxisList[2].Scale.FontSpec.FontColor = buttonRColor.BackColor;

            //myPanel.YAxisList[1].Scale.FontSpec.FontColor = buttonVoltageColor.BackColor;
            // myPanel.YAxisList[2].Scale.FontSpec.FontColor = Color.Blue;
            #endregion

            //  SetAxis();  //Y坐标刻度范围设置
            #region SetAxis();  //Y坐标刻度范围设置
            myPanel.YAxis.Scale.Min = CurrentMin;            //电流范围设置0
            myPanel.YAxis.Scale.Max = CurrentMax;
            myPanel.YAxisList[1].Scale.Min = VoltageMin;     //电压范围设置1
            myPanel.YAxisList[1].Scale.Max = VoltageMax;
            myPanel.YAxisList[2].Scale.Min = ResistanceMin;     //电阻范围设置2
            myPanel.YAxisList[2].Scale.Max = ResistanceMax;

            zedGraphControl.AxisChange();
            zedGraphControl.Refresh();
            #endregion

            //  AutoYAxis();//自动调整Y轴
            #region  AutoYAxis();//自动调整Y轴
            CurrentMin = CurrentMax = VoltageMin = VoltageMax = ResistanceMin = ResistanceMax = 0;
            myPanel.YAxis.Scale.MaxAuto = true; //最大最小刻度自动调整0
            myPanel.YAxis.Scale.MinAuto = true;
            myPanel.YAxisList[1].Scale.MaxAuto = true;//1
            myPanel.YAxisList[1].Scale.MinAuto = true;
            myPanel.YAxisList[2].Scale.MaxAuto = true;//2
            myPanel.YAxisList[2].Scale.MinAuto = true;

            zedGraphControl.Invalidate();
            zedGraphControl.AxisChange();
            #endregion

            #region //添加说明书菜单
            TextObj text_obs = new TextObj("缩放：鼠标滚轮\n移动：鼠标中键\n菜单：鼠标右键", 0.899f, 0.001f, CoordType.ChartFraction, AlignH.Center, AlignV.Bottom);
            text_obs.FontSpec.StringAlignment = StringAlignment.Center;
            text_obs.FontSpec.Fill = new Fill(Color.Transparent);
            text_obs.FontSpec.FontColor = Color.Black;
            myPanel.GraphObjList.Add(text_obs);
            #endregion






            // AddCurve(); //添加曲线
            #region AddCurve(); //添加曲线
            //myCurve_L = myPanel.AddCurve("电感", list_L, Color.Orange, SymbolType.Circle);
            myCurve_L = myPanel.AddCurve("电感"+LSign, list_L, buttonLColor.BackColor, SymbolType.Circle);
            myCurve_L.Symbol.Fill = new Fill(Color.White);
            myCurve_L.YAxisIndex = 0;
            myCurve_L.Line.IsSmooth = true;//Make the curve smooth

            //myCurve_C = myPanel.AddCurve("电容", list_C, Color.LimeGreen, SymbolType.Triangle);
            myCurve_C = myPanel.AddCurve("电容"+CSign, list_C, buttonCColor.BackColor, SymbolType.Triangle);
            myCurve_C.Symbol.Fill = new Fill(Color.White);
            myCurve_C.YAxisIndex = 1;
            myCurve_C.Line.IsSmooth = true;

            //myCurveResistance = myPanel.AddCurve("电阻", listResistance, Color.Blue, SymbolType.Diamond);
            myCurveResistance = myPanel.AddCurve("电阻"+RSign, listResistance, buttonRColor.BackColor, SymbolType.Diamond);
            myCurveResistance.Symbol.Fill = new Fill(Color.White);
            myCurveResistance.YAxisIndex = 2;
            myCurveResistance.Line.IsSmooth = true;

            //myCurveVoltage = myPanel.AddCurve("电压", listVoltage, buttonVoltageColor.BackColor, SymbolType.Triangle);
            //myCurveVoltage.Symbol.Fill = new Fill(Color.White);
            //myCurveVoltage.YAxisIndex = 1;
            //myCurveVoltage.Line.IsSmooth = true;


            //public LineItem(string label, IPointList points, Color color, SymbolType symbolType);
            //myCurveResistance = myPanel.AddCurve("电阻", listResistance, Color.Blue, SymbolType.Diamond);
            //myCurveResistance.Symbol.Fill = new Fill(Color.White);
            //myCurveResistance.YAxisIndex = 2;
            //myCurveResistance.Line.IsSmooth = true;
            #endregion

            #region 背景填充
            //  myPanel.Chart.Fill = new Fill(Color.White, Color.LightBlue, 100.0f); //图表区域的背景填充
            myPanel.Chart.Fill = new Fill(Color.White, Color.LightGreen, 100.0f); //图表区域的背景填充
            // myPanel.Fill = new Fill(Control.DefaultBackColor);                   //窗体区域的背景填充
            myPanel.Fill = new Fill(Color.White, Color.LightGreen, 100.0f);                   //窗体区域的背景填充
            #endregion

            zedGraphControl.PointValueEvent += new ZedGraphControl.PointValueHandler(MyPointValueHandler);
            zedGraphControl.IsShowPointValues = true; //显示坐标点值
            zedGraphControl.ZoomStepFraction = 0.1; //设置鼠标滚轮缩放的比例大小，值越大缩放越灵敏
            zedGraphControl.IsShowHScrollBar = false;// 
            zedGraphControl.IsEnableZoom = true;

            #endregion 


        }

        private string MyPointValueHandler(ZedGraphControl control, GraphPane pane, CurveItem curve, int iPt)
        {
            PointPair pt = curve[iPt];
            string unit = "";
            //if (curve == myCurve_L)
            //{
            //    unit = "H";
            //}
            //if (curve == myCurve_C)
            //{
            //    unit = "f";
            //}
            //if (curve == myCurveResistance)
            //{
            //    unit = "Ω";
            //}

            return curve.Label.Text + "是" + pt.Y.ToString("f4") + unit + " 在位置为 " + pt.X.ToString("f3");
        }



        #region 按钮选择颜色，选择曲线显示
        private void buttonLColor_Click(object sender, EventArgs e)
        {
            ColorDialog my = new ColorDialog();
            if (DialogResult.OK == my.ShowDialog())
            {
                buttonLColor.BackColor = my.Color;
                myCurve_L.Color = my.Color;
                myPanel.YAxis.Scale.FontSpec.FontColor = my.Color;
                zedGraphControl.Refresh();
            }
        }

        private void buttonCColor_Click(object sender, EventArgs e)
        {
            ColorDialog my = new ColorDialog();
            if (DialogResult.OK == my.ShowDialog())
            {
                buttonCColor.BackColor = my.Color;
                myCurve_C.Color = my.Color;
                myPanel.YAxisList[1].Scale.FontSpec.FontColor = my.Color;
                zedGraphControl.Refresh();
            }

        }

        private void buttonRColor_Click(object sender, EventArgs e)
        {
            ColorDialog my = new ColorDialog();
            if (DialogResult.OK == my.ShowDialog())
            {
                buttonRColor.BackColor = my.Color;
                myCurveResistance.Color = my.Color;
                myPanel.Y2Axis.Scale.FontSpec.FontColor = my.Color;
                zedGraphControl.Refresh();
            }
        }


        private void checkBoxL_CheckedChanged(object sender, EventArgs e)
        {
            myCurve_L.IsVisible = checkBoxL.Checked;
            zedGraphControl.Refresh();
        }

        private void checkBoxC_CheckedChanged(object sender, EventArgs e)
        {
            myCurve_C.IsVisible = checkBoxC.Checked;
            zedGraphControl.Refresh();
        }

        private void checkBoxR_CheckedChanged(object sender, EventArgs e)
        {
            myCurveResistance.IsVisible = checkBoxR.Checked;
            zedGraphControl.Refresh();
        }
        #endregion


        #region UI图形刷新控制程序

        public int num; //演示绘图效果所需变量 //时间控制
        private void UIRefresh()
        {
            //Current = Voltage = Resistance = 0;
            double x = 0;
            //x = (x + num) / 10;
            x = dataPulse;
            //Current = Math.Sin(num * 0.2 - 1.57) * 10 + 10; //设置曲线坐标随机正弦波曲线
            //Voltage = Math.Cos(num * 0.04) * 12 + 15;
            //Resistance = Math.Sin(num * 0.4 - 2.57) * 20 + 40;


            list_L.Add(x, Current); //添加坐标
            list_C.Add(x, Voltage);
            listResistance.Add(x, Resistance);

            //myCurve_L = myPanel.AddCurve("电感" + LSign, list_L, buttonLColor.BackColor, SymbolType.Circle);
            myCurve_L.Label.Text = "电感" + LSign;
            myCurve_C.Label.Text = "电容" + CSign;
            myCurveResistance.Label.Text = "电阻" + RSign;
            //gaiwjh
            //textBoxL.Text = Current.ToString("f4") + "H"; //显示电gan值
            //textBoxC.Text = Voltage.ToString("f4") + "F"; //显示电rong值
            //textBoxR.Text = Resistance.ToString("f4") + "Ω"; //显示电阻值

            //if (chooseNm == 1)
            //{
            //    textBoxCurrent.Text = out_num.ToString("f4"); //显示电gan值
            //}
            //else if (chooseNm == 2)
            //{
            //    textBoxVoltage.Text = out_num.ToString("f4"); //显示电rong值
            //}
            //else
            //{
            //    textBoxResistance.Text = out_num.ToString("f4"); //显示电阻值
            //}

            //textBoxCurrent.Text = Current.ToString("f4"); //显示电流值
            //textBoxVoltage.Text = Voltage.ToString("f4"); //显示电压值
            //textBoxResistance.Text = Resistance.ToString("f4"); //显示电阻值

            zedGraphControl.AxisChange();
            zedGraphControl.Refresh();
            //num++;
        }

        #endregion


        #region 单位转化

        string[] LSignSet = { "H", "mH", "uH", "nH", "pH", "fH","kH","MH" };
        string[] CSignSet = { "F", "mF", "uF", "nF", "pF", "fF" ,"kF","MF"};
        string[] RSignSet = {  "Ω", "mΩ", "uΩ", "nΩ", "pΩ", "fΩ" ,"kΩ", "MΩ"};
        string[] FSignSet = { "Hz", "mHz", "uHz", "nHz", "pHz", "fHz", "kHz", "MHz" };

        private void UnitChange(string Lstr,string[] sign ,ref double changeValue, ref string changeSign )
        {

            if (Lstr.IndexOf('E') != -1)
            {
                string[] data = Lstr.Split('E');//将读到的字符串存入数组data中
                double datafirst = Convert.ToDouble(data[0]);//data[0]判断有效数字
                int datasecond = Convert.ToInt16(data[1]);  //判断数量级

                if (datasecond >= 0 && datasecond < 3) //判断电阻值Ω 级别
                 {
                        changeSign = sign[0];
                        changeValue = datafirst * Math.Pow(10, datasecond);
                  }
                  else if (datasecond >= 3 && datasecond < 6) //判断电阻值 kΩ级别
                  {
                        changeSign = sign[6];
                        datasecond = datasecond - 3;
                        changeValue = datafirst * Math.Pow(10, datasecond);
                  }
                  else if (datasecond >= 6) //判断电阻值MΩ 级别
                  {
                        changeSign = sign[7];
                        datasecond = datasecond - 6;
                        changeValue = datafirst * Math.Pow(10, datasecond);
                   }
                    //级别datasecond < 0
                   else if (datasecond < 0 && datasecond >= -3) //mH级别
                   {
                        datasecond = datasecond + 3;
                        changeSign = sign[1];
                        changeValue = datafirst * Math.Pow(10, datasecond);
                    }
                    else if (datasecond < -3 && datasecond >= -6) //uH级别
                    {
                        datasecond = datasecond + 6;
                        changeSign = sign[2];
                        changeValue = datafirst * Math.Pow(10, datasecond);
                    }
                    else if (datasecond < -6 && datasecond >= -9) //nH级别
                    {
                        datasecond = datasecond + 9;
                        changeSign = sign[3];
                        changeValue = datafirst * Math.Pow(10, datasecond);
                    }
                    else if (datasecond < -9 && datasecond >= -12) //nH级别
                    {
                        datasecond = datasecond + 12;
                        changeSign = sign[4];
                        changeValue = datafirst * Math.Pow(10, datasecond);
                    }
                    else if (datasecond < -12) //fH级别 和更小的
                    {
                        datasecond = datasecond + 15;
                        changeSign = sign[5];
                        changeValue = datafirst * Math.Pow(10, datasecond);
                    }
            }
            else //if (Lstr.IndexOf('E') != -1) 非科学计数法
            {
                changeValue = Convert.ToDouble(Lstr);
                changeSign = sign[0];
            }


        }



        #endregion


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult.No == MessageBox.Show("确认要退出吗?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                e.Cancel = true;
            }
            else
            {

                if (timerUI.Enabled)
                {
                    timerUI.Enabled = false;
                }

                if (timerLCR.Enabled)
                {
                    timerLCR.Enabled = false;
                }

                if (serialPortLCR.IsOpen)
                {
                    try {
                        serialPortLCR.Close();
                    }
                    catch { }
                }

                if (serialPortMotor.IsOpen)
                {
                    try
                    {
                        serialPortMotor.Close();
                    }
                    catch { }
                }

                //GetPrivateProfileString("SerialLCR", "COM", "COM1", temp, 1024, FileName);
                //comboBoxComLCR.Text = temp.ToString();
                //GetPrivateProfileString("SerialLCR", "Baud", "9600", temp, 1024, FileName);
                //comboBoxBaudRate.Text = temp.ToString();

                //GetPrivateProfileString("SerialPulse", "COM", "COM1", temp, 1024, FileName);
                //comboBoxMotorCom.Text = temp.ToString();
                //GetPrivateProfileString("SerialPulse", "Baud", "9600", temp, 1024, FileName);
                //comboBoxMotorBaudRate.Text = temp.ToString();

                //GetPrivateProfileString("SerialMotor", "COM", "COM1", temp, 1024, FileName);
                //comboBoxMotor.Text = temp.ToString();
                //GetPrivateProfileString("SerialMotor", "Baud", "9600", temp, 1024, FileName);
                //comboBoxMotorRate.Text = temp.ToString();

                //GetPrivateProfileString("Pulse", "Frequence", "1000", temp, 1024, FileName);
                //comboBoxFre.Text = temp.ToString();


                WritePrivateProfileString("SerialLCR", "COM", comboBoxComLCR.Text, FileName);
                WritePrivateProfileString("SerialLCR", "Baud", comboBoxBaudRate.Text, FileName);

                WritePrivateProfileString("SerialPulse", "COM", comboBoxMotorCom.Text, FileName);
                WritePrivateProfileString("SerialPulse", "Baud", comboBoxMotorBaudRate.Text, FileName);

                WritePrivateProfileString("SerialMotor", "COM", comboBoxMotor.Text, FileName);
                WritePrivateProfileString("SerialMotor", "Baud", comboBoxMotorRate.Text, FileName);

                WritePrivateProfileString("Pulse", "Frequence", comboBoxFre.Text, FileName); 

              
            }

        }

        private void buttonZero_Click(object sender, EventArgs e)
        {
            calibrationPulse = dataPulse;
            dataPulse = 0;
            textBoxPulse.Text = "0";
        }

        #region 开始监控
    
        private void buttonMain_Click(object sender, EventArgs e)
        {
            if ( buttonMain.Text == "开始监控" 
                && serialPortLCR.IsOpen &&serialPortMotor.IsOpen)
            {
                
                buttonMain.Text = "停止监控";
                buttonMain.BackColor = Color.YellowGreen;

                //LCR设置
                dataFre = comboBoxFre.Text;
                serialPortLCR.WriteLine(":FREQuency:CW " + dataFre + "\r\n"); //发送频率设置
                Thread.Sleep(300);
                chooseNm = 1;//切换电感
                serialPortLCR.Write(":FUNC:IMP:TYPE LPD\r\n"); //查询
                Thread.Sleep(300);
                flagSendcmd = false;//切换


                strFrequence = comboBoxFre.Text;

                timerLCR.Enabled = true; //超时控制
                timerPulse.Enabled = true;
                timerUI.Enabled = true;
                

                
            }
            else 
            {
                if (RecordEnble)
                {
                    MessageBox.Show("请先关闭记录数据");
                    return;
                }
                timerLCR.Enabled = false; 
                timerPulse.Enabled = false;
                timerUI.Enabled = false;

                buttonMain.Text = "开始监控";
                buttonMain.BackColor = Color.SkyBlue;
               
            }
        }

        #endregion 

       

        #region 定时器motor 调节慢点
        private void timerMotor_tick(object sender, EventArgs e)
        {
            if (serialPortMotor.IsOpen) //零才发送
            {
                try
                {
                  
                    modbusMotor.Write(serialPortMotor, 1, 3, 16, 4); //发送脉冲数据 
                    
                }
                catch
                {
                    MessageBox.Show("serialPortMotor失败");
                }

            }
        }
        #endregion

        #region Motor串口打开关闭函数
        private void buttonSerialMotor_Click(object sender, EventArgs e)
        {

            if (serialPortMotor.IsOpen == true)
            {
                if (buttonMain.Text == "停止监控")
                {
                    MessageBox.Show("先关闭监控");
                    return;
                }

                //串口处于连接状态则断开连接
                serialPortMotor.Close();
                //timerLCR.Enabled = false;
                buttonSerialMotor.BackColor = Color.LightCoral;
                buttonSerialMotor.Text = "连接";
                toolStripLabel2.Text = "位移串口未打开！";
            }
            else
            {//串口处于断开状态则连接
               
                try
                {
                    //serialPortMotor.PortName = "COM" + (comboBoxMotorCom.SelectedIndex + 1).ToString();
                    serialPortMotor.PortName = comboBoxMotorCom.Text;
                    //serialPortMotor.BaudRate = baud[comboBoxMotorBaudRate.SelectedIndex];
                    serialPortMotor.BaudRate = Convert.ToInt32(comboBoxMotorBaudRate.Text);
                    serialPortMotor.Parity = Parity.None;
                    serialPortMotor.DataBits = 8;
                    serialPortMotor.StopBits = StopBits.One;

                    serialPortMotor.Open();
                   
                }
                catch
                {
                    MessageBox.Show("打开位移串口失败，请检查串口是否存在或已被占用", "警告");
                }
                if (serialPortMotor.IsOpen == true)
                {
                    //成功打开串口
                    buttonSerialMotor.Text = "断开";
                    buttonSerialMotor.BackColor = Color.LightGreen;
                    toolStripLabel2.Text = "位移串口已经打开！"; 
                    
                }
            }
        }
        #endregion

        #region Motor串口接收

        public static void CalculateCRC(List<byte> pByte, int nNumberOfBytes, out ushort pChecksum)
        {
            int nBit;
            ushort nShiftedBit;
            pChecksum = 0xFFFF;

            for (int nByte = 0; nByte < nNumberOfBytes; nByte++)
            {
                pChecksum ^= pByte[nByte];
                for (nBit = 0; nBit < 8; nBit++)
                {
                    if ((pChecksum & 0x1) == 1)
                    {
                        nShiftedBit = 1;
                    }
                    else
                    {
                        nShiftedBit = 0;
                    }
                    pChecksum >>= 1;
                    if (nShiftedBit != 0)
                    {
                        pChecksum ^= 0xA001;
                    }
                }
            }
        }

        public static void CalculateCRC(List<byte> pByte, int nNumberOfBytes, out byte hi, out byte lo)
        {
            ushort sum;
            CalculateCRC(pByte, nNumberOfBytes, out sum);
            lo = (byte)(sum & 0xFF);
            hi = (byte)((sum & 0xFF00) >> 8);
        }

        private List<byte> buffer = new List<byte>(4096);//默认分配1页内存，并始终限制不允许超过  
        private byte[] binary_data_1 = new byte[13];//01 03 08 01 00 00 00 00 00 00 00 54 1B
        private long received_count = 0;//接收计数  
        
        private void DataReceivedMotor(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                int n = serialPortMotor.BytesToRead;//先记录下来，避免某种原因，人为的原因，操作几次之间时间长，缓存不一致  
                byte[] buf = new byte[n];//声明一个临时数组存储当前来的串口数据  
                int cooo = serialPortMotor.Read(buf, 0, n);//读取缓冲数据

                //测试接收数据情况，成功后注释掉
                //this.Invoke((EventHandler)(delegate  
                //{
                //  textBox2.AppendText(Convert.ToString(cooo));
                // textBox2.AppendText("--\r\n");
                // foreach (byte mydata in buf)
                //{
                //    string str = Convert.ToString(mydata, 16).ToUpper();
                //    textBox2.AppendText("0x" + (str.Length == 1 ? "0" + str : str) + " ");
                //}
                // textBox2.AppendText("--\r\n");
                // //textBox2.Text = Convert.ToString(n);

                //}));



                // //因为要访问ui资源，所以需要使用invoke方式同步ui。  
                //this.Invoke((EventHandler)(delegate  
                //{  
                //    //判断是否是显示为16禁止  
                //    if (checkBoxHexView.Checked)  
                //    {  
                //        //依次的拼接出16进制字符串  
                //        foreach (byte b in buf)  
                //        {  
                //            builder.Append(b.ToString("X2") + " ");  
                //        }  
                //    }  
                //    else  
                //    {  
                //        //直接按ASCII规则转换成字符串  
                //        builder.Append(Encoding.ASCII.GetString(buf));  
                //    }  
                //    //追加的形式添加到文本框末端，并滚动到最后。  
                //    this.txGet.AppendText(builder.ToString());  
                //    //修改接收计数  
                //    labelGetCount.Text = "Get:" + received_count.ToString();  
                //}));





                //<协议解析>  
                bool data_1_catched = false;//缓存记录数据是否捕获到 
                //1.缓存数据  
                buffer.AddRange(buf);
                //2.完整性判断  
                while (buffer.Count >= 5)//至少要包含头（2字节）+长度（1字节）+校验（2字节） 
                {
                    //2.1 查找数据头
                    if (buffer[0] == 0x01 && buffer[1] == 0x03)
                    {
                        //2.2 探测缓存数据是否有一条数据的字节，如果不够，就不用费劲的做其他验证了  
                        //前面已经限定了剩余长度>=4，那我们这里一定能访问到buffer[2]这个长度 
                        int len = buffer[2];//数据长度  
                        //数据完整判断第一步，长度是否足够  
                        //len是数据段长度,4个字节是while行注释的3部分长度  
                        if (buffer.Count < len + 5) break;//数据不够的时候什么都不做  
                        //这里确保数据长度足够，数据头标志找到，我们开始计算校验  
                        //2.3 校验数据，确认数据正确  

                        //计算CRC
                        byte crch = 0;
                        byte crcl = 0;
                        CalculateCRC(buffer, len + 3, out crch, out crcl);
                        if (buffer[len + 3] != crcl || buffer[len + 4] != crch) //如果数据校验失败，丢弃这一包数据 
                        {
                            buffer.RemoveRange(0, len + 5);//从缓存中删除错误数据  
                            continue;//继续下一次循环  
                        }
                        //rebyte[6] = crcl;
                        //rebyte[7] = crch;
                        //至此，已经被找到了一条完整数据。我们将数据直接分析，或是缓存起来一起分析  
                        //我们这里采用的办法是缓存一次，好处就是如果你某种原因，数据堆积在缓存buffer中  
                        //已经很多了，那你需要循环的找到最后一组，只分析最新数据，过往数据你已经处理不及时  
                        //了，就不要浪费更多时间了，这也是考虑到系统负载能够降低。 

                        buffer.CopyTo(0, binary_data_1, 0, len + 5);//复制一条完整数据到具体的数据缓存
                        data_1_catched = true;
                        buffer.RemoveRange(0, len + 5);//正确分析一条数据，从缓存中移除数据。  

                    }
                    else
                    {
                        //这里是很重要的，如果数据开始不是头，则删除数据  
                        buffer.RemoveAt(0);
                    }

                    //分析数据  
                    if (data_1_catched)
                    {
                        //我们的数据都是定好格式的，所以当我们找到分析出的数据1，就知道固定位置一定是这些数据，我们只要显示就可以了  
                        string data = binary_data_1[0].ToString("X2") + " " + binary_data_1[1].ToString("X2") + " " +
                                        binary_data_1[2].ToString("X2") + " " + binary_data_1[3].ToString("X2") + " " +
                                        binary_data_1[4].ToString("X2") + " " + binary_data_1[5].ToString("X2") + " " + binary_data_1[6].ToString("X2") + " " +
                                        binary_data_1[7].ToString("X2") + " " + binary_data_1[8].ToString("X2") + " " +
                                        binary_data_1[9].ToString("X2") + " " + binary_data_1[10].ToString("X2") + " " + binary_data_1[11].ToString("X2") + " " +
                                        binary_data_1[12].ToString("X2");

                        this.Invoke((EventHandler)(delegate { textBox1.Text = data; }));

                        Int32 u = BitConverter.ToInt32(new byte[] { binary_data_1[6], binary_data_1[5], binary_data_1[4], binary_data_1[3] }, 0);
                        //byte[] check = CrcCheck(data);
                        //dataPulse = -u * 0.084;
                        dataPulse = -u * 0.084 - calibrationPulse;

                        strLength = dataPulse.ToString("f2");
                    }



                }



            }
            catch { }




            //try
            //{
            //    byte[] data = new byte[1024];
            //    int n=serialPortMotor.Read(data, 0, data.Length);


            //    //foreach (byte mydata in data)
            //    //{
            //    //    string str = Convert.ToString(mydata, 16).ToUpper();
            //    //    textBox1.AppendText("0x" + (str.Length == 1 ? "0" + str : str) + " ");
            //    //}
            //    //textBox1.AppendText("--\r\n");
            //    //textBox2.Text = Convert.ToString(n);


            //    //bool a = modbusMotor.ReadIsRight(data);
            //    //if (data[0] == 0x01 && data[1] == 0x03 && data[2] == 0x08 )//modbusMotor.ReadIsRight(data))
            //    if (n==12 && data[0] == 0x03 && data[1] == 0x08)//modbusMotor.ReadIsRight(data))
            //    {
            //        Int32 u = BitConverter.ToInt32(new byte[] { data[5], data[4], data[3], data[2] }, 0);
            //        //byte[] check = CrcCheck(data);
            //        dataPulse = -u * 0.084 - calibrationPulse;
            //       // textBoxPulse.Text = (dataPulse).ToString("f2") + "mm";
            //    }

            //    if (n == 13 && data[0] == 0x01 && data[1] == 0x03 && data[2] == 0x08) //&& modbusMotor.ReadIsRight(data))
            //    //if (data[0] == 0x03 && data[1] == 0x08 )
            //    {
            //        //Int32 u = BitConverter.ToInt32(new byte[] { data[5], data[4], data[3], data[2] }, 0);
            //        Int32 u = BitConverter.ToInt32(new byte[] { data[6], data[5], data[4], data[3] }, 0);
            //        //byte[] check = CrcCheck(data);
            //        dataPulse = -u * 0.084;
            //        //textBox3.Text = (dataPulse).ToString("f2") + "mm";
            //        //textBox3.AppendText();

            //    }



            //}
            //catch
            //{
            //   // MessageBox.Show("Pulse接收S失败");

            //}
        }
        public delegate void DelD(string message2); //Pulse接收串口委托
        public delegate void DelDdb(double data); //Pulse接收串口委托

        #endregion 

        #region LCR串口开关
        private void buttonSerialLCR_Click(object sender, EventArgs e)
        {
            if (serialPortLCR.IsOpen == true)
            {
                //串口处于连接状态则断开连接
                if (buttonMain.Text == "停止监控")
                {
                    MessageBox.Show("先关闭监控");
                    return;
                }
                serialPortLCR.Close();
                buttonSerialLCR.BackColor = Color.LightCoral;
                buttonSerialLCR.Text = "连接";
                toolStripLabel1.Text = "数据监测串口未打开！";

            }
            else
            {//串口处于断开状态则连接
                try
                {
                    serialPortLCR.PortName = comboBoxComLCR.Text;
                    serialPortLCR.BaudRate = Convert.ToInt32(comboBoxBaudRate.Text);
                    serialPortLCR.Parity = Parity.None;
                    serialPortLCR.DataBits = 8;
                    serialPortLCR.StopBits = StopBits.One;
                    serialPortLCR.Open();
                }
                catch
                {
                    MessageBox.Show("打开数据监测串口失败，请检查串口是否存在或已被占用", "警告");
                }
                if (serialPortLCR.IsOpen == true)
                {
                    //成功打开串口
                    buttonSerialLCR.Text = "断开";
                    buttonSerialLCR.BackColor = Color.LightGreen;
                    toolStripLabel1.Text = "数据监测串口已打开！";

                }
            }
        }
        #endregion

        #region LCR数据接收
        private void DataReceivedLCR(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string message2 = serialPortLCR.ReadLine();
                string strR = message2.Replace('\r', ' ');
                string[] data = strR.Split(',');//将读到的字符串存入数组data中
                if (data.Length >= 3)
                {
                    if (chooseNm == 1)  //电感
                    {
                        //strInductance = ChangeEUnit(data[0]);
                        strInductance = data[0];
                        UnitChange(data[0], LSignSet, ref Current, ref LSign);
                        chooseNm = 2;
                    }
                    else if (chooseNm == 2)//电容
                    {
                        //strCapacitance = ChangeEUnit(data[0]);
                        strCapacitance  = data[0];
                        UnitChange(data[0], CSignSet, ref Voltage, ref CSign);
                        chooseNm = 3;

                    }
                    else if (chooseNm == 3) //电阻
                    {
                        //strResistanc = ChangeEUnit(data[0]);
                        strResistanc = data[0];
                        UnitChange(data[0], RSignSet, ref Resistance, ref RSign);
                        chooseNm = 1;

                    }
                    else
                    {
                        chooseNm = 1;
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
               // MessageBox.Show("接送LCR");
            }
        }

        #endregion LCR数据接收

        #region LCR发送发数据请求 //只发送请求查询
        private bool workLCR = false;
        private int countLCR = 0;
        bool flagSendcmd = false;//一个定时器交换发送LCR 频率 与查询码
        private void timerLCR_Tick(object sender, EventArgs e)
        {
            if (serialPortLCR.IsOpen)
            {
                try
                {
                    if (flagSendcmd) //给一个查询一个
                    {
                        serialPortLCR.Write(":FETCH?\r\n"); //查询
                        flagSendcmd = false;
                    }
                    else
                    {
                        string cmd = ""; //发送LCR数据请求信号
                        if (chooseNm == 1) //
                        {
                            cmd = ":FUNC:IMP:TYPE LPD\r\n";
                        }
                        else if (chooseNm == 2) //
                        {
                            cmd = ":FUNC:IMP:TYPE CPD\r\n";
                        }
                        else if (chooseNm == 3)
                        {
                            cmd = ":FUNC:IMP:TYPE RX\r\n";
                        }
                        else
                        {
                            chooseNm = 1;
                            cmd = ":FUNC:IMP:TYPE LPD\r\n";
                        }

                        serialPortLCR.Write(cmd);
                        flagSendcmd = true;
                    }
                }
                catch
                {
                    flagSendcmd = false;//错了几滑档
                    MessageBox.Show("LCR发送失败");
                }

            }
        }

        #endregion

        #region 总UI刷新
        private void timerUI_Tick(object sender, EventArgs e)
        {
            //RLCFTextdataRefresh();//主界面刷新
            //UI刷新函数还没写
            if (timerLCR.Enabled)
            {
                //if (serialPortLCR.IsOpen && serialPortMotor.IsOpen && chooseNm != 0 && chooseNm != 1 )
                //{

                //    UIRefresh();         //UI刷新  // 改了
                //}
                textBoxL.Text = Current.ToString("f4") + LSign; //显示电gan值   
                textBoxC.Text = Voltage.ToString("f4") + CSign; //显示电rong值
                textBoxR.Text = Resistance.ToString("f4") + RSign; //显示电阻值
            }
            if (timerPulse.Enabled)
            {
                textBoxPulse.Text = (dataPulse).ToString("f2") + "mm";
            }
            if (RecordEnble)
            {
                UIRefresh();         //UI刷新  // 改了
                //数据导入数据库
                //string data = (dataPulse).ToString("f2") + "," + Current.ToString("f5") + "," + LSign + "," + Voltage.ToString("f5") + "," + CSign + "," + Resistance.ToString("f5") + "," + RSign;
                //txtRecord.Write(fileRecord, data);
                SaveFile();
            }



        }
        private void RLCFTextdataRefresh()
        {
            //textBoxL.Text = Current.ToString("f2") + "H"; //显示电gan值   
            textBoxL.Text = Current.ToString("f5") + LSign; //显示电gan值   
            textBoxC.Text = Voltage.ToString("f5") + CSign; //显示电rong值
            textBoxR.Text = Resistance.ToString("f5") + RSign; //显示电阻值
            //textBoxF.Text = Frequence.ToString("f0") + FSign;//频率
            
           textBoxPulse.Text = (dataPulse ).ToString("f2") + "mm";
            //textBoxPulse1.Text = dataPulse.ToString("f2") + "mm";
        }

        #endregion


        #region 图像复原
        private void buttonPicZero_Click(object sender, EventArgs e)
        {

            list_L.Clear();
            list_C.Clear();
            listResistance.Clear();

            zedGraphControl.Refresh();
            textBox2.Text = " ";
        }
        #endregion

        #region 电机控制区

        private void buttonLeft_Click(object sender, EventArgs e)
        {

            if (serialPortMotorReal.IsOpen == true)
            {
                //串口处于连接状态则断开连接
                //serialPortMotorReal.WriteLine("Left");
                //byte[] data = { 0xBA, 0x01, 0x00, 0xFF, 0x00, 0x03, 0x01, 0x30, 0x00, 0x63, 0x9C, 0x00, 0xC8, 0x00, 0xC8, 0x89, 0xFE };
                 byte[] data = { 0xBA,0x01,0xEA,0x60,0x00,0x03,0x01,0x30,0x00,0x63,0x9C,0x01,0xF4,0x00,0xC8,0xC1,0xFE };
                serialPortMotorReal.Write(data, 0, 17);

            }
            else
            {
                MessageBox.Show("请打开电机串口");
            }



        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            if (serialPortMotorReal.IsOpen == true)
            {
                //串口处于连接状态则断开连接
                //serialPortMotorReal.WriteLine("stop");

                byte[] data = { 0xBA, 0x01, 0x01, 0x00, 0x00, 0x03, 0x03, 0x30, 0x00, 0x63, 0x9C, 0x00, 0xC8, 0x00, 0xC8, 0x75, 0xFE };
                serialPortMotorReal.Write(data, 0, 17);

            }
            else 
            {
                MessageBox.Show("请打开电机串口");
            }

        }

        private void buttonRight_Click(object sender, EventArgs e)
        {
            //if (serialPortMotorReal.IsOpen == true)
            //{
            //    //串口处于连接状态则断开连接
            //    //serialPortMotorReal.WriteLine("Right");
            //    byte[] data = { 0xBA, 0x01, 0xFF, 0xFF, 0x00, 0x03, 0x02, 0x30, 0x00, 0x63, 0x9C, 0x00, 0xC8, 0x00, 0xC8, 0x75, 0xFE };
            //    serialPortMotorReal.Write(data, 0, 17);

            //}
            //else
            //{
            //    MessageBox.Show("请打开电机串口");
            //}

            //快速右转
            if (serialPortMotorReal.IsOpen == true)
            {
                //串口处于连接状态则断开连接
                //serialPortMotorReal.WriteLine("Right");
                //byte[] data = { 0xBA, 0x01, 0x00, 0xFF, 0x00, 0x03, 0x02, 0x30, 0x00, 0x63, 0x9C, 0x00, 0xC8, 0x00, 0xC8, 0x8A, 0xFE };
                byte[] data = { 0xBA, 0x01, 0xEA, 0x60, 0x00, 0x03, 0x02, 0x30, 0x00, 0x63, 0x9C, 0x01, 0xF4, 0x00, 0xC8, 0xC2, 0xFE };
                serialPortMotorReal.Write(data, 0, 17);

            }
            else
            {
                MessageBox.Show("请打开电机串口");
            }

        }

        private void serialMotorOpenClose_Click(object sender, EventArgs e)
        {
            if (serialPortMotorReal.IsOpen == true)
            {
                //串口处于连接状态则断开连接
                serialPortMotorReal.Close();
                serialMotorOpenClose.BackColor = Color.LightCoral;
                serialMotorOpenClose.Text = "连接";
                toolStripLabel3.Text = "电机串口未打开！";

            }
            else
            {//串口处于断开状态则连接
                try
                {
                    serialPortMotorReal.PortName = comboBoxMotor.Text;
                    serialPortMotorReal.BaudRate = Convert.ToInt32(comboBoxMotorRate.Text);
                    serialPortMotorReal.Parity = Parity.None;
                    serialPortMotorReal.DataBits = 8;
                    serialPortMotorReal.StopBits = StopBits.One;
                    serialPortMotorReal.Open();
                }
                catch
                {
                    MessageBox.Show("打开电机串口失败，请检查串口是否存在或已被占用", "警告");
                }
                if (serialPortMotorReal.IsOpen == true)
                {
                    //成功打开串口
                    serialMotorOpenClose.Text = "断开";
                    serialMotorOpenClose.BackColor = Color.LightGreen;
                    toolStripLabel3.Text = "电机串口已打开！";

                }
            }

        }

        #endregion

        private void buttonData_Click(object sender, EventArgs e)
        {
            if (buttonMain.Text == "停止监控")
            {

                if (buttonData.Text == "记录数据")
                {

                    //清空数组
                    list_L.Clear();
                    list_C.Clear();
                    listResistance.Clear();
                    zedGraphControl.Refresh();
                    textBox2.Text = " ";

                    //开启慢运行
                    if (serialPortMotorReal.IsOpen == true)
                    {
                        //串口处于连接状态则断开连接
                        //serialPortMotorReal.WriteLine("Right");
                        //byte[] data = { 0xBA, 0x01, 0xFF, 0xFF, 0x00, 0x03, 0x02, 0x30, 0x00, 0x63, 0x9C, 0x00, 0xC8, 0x00, 0xC8, 0x75, 0xFE };
                        byte[] data = {0xBA,0x01,0x05,0xDD,0x00,0x03,0x02,0x30,0x00,0x63,0x9C,0x00,0xC8,0x00,0xC8,0xAD,0xFE};
                        serialPortMotorReal.Write(data, 0, 17);

                    }
                    else
                    {
                        MessageBox.Show("请打开电机串口");
                        return;
                    }


                    //acces 数据库建表
                    try
                    {

                        //access.FolderName = "databaseAccess";
                        //access.DatabaseName = "DEP.accdb";//数据库名
                        //access.CreateFolder();
                        //if (!access.CreateDatabase())
                        //{
                        //    MessageBox.Show("数据库创建失败");
                        //    return;
                        //}


                        //LSign = "H", CSign = "F", RSign = "Ω", FSign ="Hz";  //C代表C  V代表电容 R代表电阻


                        

                        
                        tableName = textBox3.Text+"__"+ DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                       bool fac = access.CreateDataTable(tableName, cloumsName);
                       if (fac)
                       { }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }


                    buttonData.Text = "停止记录数据";  
                    RecordEnble = true;

                }
                else
                {

                    RecordEnble = false;
                    buttonData.Text = "记录数据";
                    //停止电机
                    if (serialPortMotorReal.IsOpen == true)
                    {
                        //串口处于连接状态则断开连接
                        byte[] data = { 0xBA, 0x01, 0x01, 0x00, 0x00, 0x03, 0x03, 0x30, 0x00, 0x63, 0x9C, 0x00, 0xC8, 0x00, 0xC8, 0x75, 0xFE };
                        serialPortMotorReal.Write(data, 0, 17);

                    }
                    else
                    {
                        MessageBox.Show("请打开电机串口");
                    }
 
                }

            }
            else
            {
                MessageBox.Show("请打开监控");
                RecordEnble = false;

            }

            

        }

        private void buttonFileRoad_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = fileRecordNamePath;
            //if (DialogResult.No == MessageBox.Show("该功能用于系统出厂配置，非厂方人员不得修改，是否强制更改?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            //{
            //    // System_Operation("取消退出系统");
            //    return;
            //}
            //else
            //{
            //    if (DialogResult.No == MessageBox.Show("再次确认是否强制更改?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            //    {
            //        // System_Operation("取消退出系统");
            //        return;
            //    }

            //}
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                fileRecordNamePath = fbd.SelectedPath;
            }

            //label13.Text = "保存路径： " + fileRecordNamePath;



        }



        public void SaveFile()
        {
            System.DateTime currentTime = new System.DateTime();
            currentTime = DateTime.Now;
            string Year = currentTime.Year.ToString();
            string Month = currentTime.Month.ToString();
            string Day = currentTime.Day.ToString();
            string Hour = currentTime.Hour.ToString();
            string Minute = currentTime.Minute.ToString();
            string Second = currentTime.Second.ToString();
            string time = string.Format("{0}-{1}-{2}/{3}:{4}:{5}",
                Year, Month, Day, Hour, Minute, Second);


            dataList.Add(time);
            dataList.Add(strFrequence);
            dataList.Add(strLength);
            dataList.Add(strResistanc);
            dataList.Add(strInductance);
            dataList.Add(strCapacitance);
            bool a = access.InsertData(tableName, cloumsName, dataList);
            dataList.Clear();

        }

        public string ChangeEUnit(string str)
        {
            string[] data = str.Split('E');//将读到的字符串存入数组data中
            double datafirst = Convert.ToDouble(data[0]);//data[0]判断有效数字
            int datasecond = Convert.ToInt16(data[1]);  //判断数量级
            datafirst = datafirst * Math.Pow(10, datasecond);
            return datafirst.ToString("f6");
        }








    }



}
