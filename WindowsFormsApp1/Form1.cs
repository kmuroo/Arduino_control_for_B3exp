using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Windows.Forms.DataVisualization.Charting;
using System.Threading;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        //private void add_serial_portnames();
        int n=4096, dt=1;  //パラメーター初期値
        string[] recieved_str = new string[4096];
        float[] fa0 = new float[4096];
        float[] fa1 = new float[4096];
        bool cancel = false;


        public Form1()
        {
            InitializeComponent();
            add_serial_portname();
         }

        private void button1_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen == false)
            {
                serialPort1.PortName = comboBox1.SelectedItem.ToString();// comboBox1.SelectedText;
                serialPort1.Open();
                while (serialPort1.IsOpen == false)
                {
                    // ポートがオープンするまで待つ
                }
                serialPort1.ReadExisting();

                try
                {
                    serialPort1.ReadExisting(); //バッファを空に
                    serialPort1.Write("c");     // 接続確認、「c」を送る
                    serialPort1.ReadTimeout = 5000;
                    int b = serialPort1.ReadByte();
                    if(b.ToString() != "67")
                    {
                        throw new Exception("エコーバックがCでではありません。");
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("COMポートから反応ない、または誤った反応をしています。もう一度「Open COM」ボタンを押して、反応がおかしいようであれば他のCOMポートを試してください。");
                    serialPort1.Close();
                    return;
                }
                
                button1.Text = "Connected";
                button1.BackColor = Color.DarkGray;
                MessageBox.Show("COMポートを接続しました");
                arduino_send_recv("s"); //念のためArduinoのTime割り込み停止指示


                //パラメーター初期値設定
                arduino_send_recv("t" + dt.ToString());
                arduino_send_recv("n" + n.ToString());

                label4.Text = dt.ToString();
                label5.Text = n.ToString();
                label6.Text = "STATUS: Ready to RUN";
                label6.BackColor = Color.LightCyan;

            }
            else
            {
                comclose();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)//パラメーター設定
        {
            if(serialPort1.IsOpen == true)
            {
                label6.Text = "STATUS: Not Ready";
                label6.BackColor = Color.Yellow;
                try
                {   
                    //現在の設定を設定ダイアログに格納
                    //Form2_Tex1は、Publicのように、外部からアクセスできるModifiersである必要がある
                    Form2.Instance.Form2_Text1 = dt.ToString();
                    Form2.Instance.Form2_Text2 = n.ToString();
                    //Form2を表示する
                    Form2.Instance.ShowDialog();
                    
                    //設定ダイアログでの設定値について
                    //パラメーター範囲外トリミング処理
                    if ( int.Parse(Form2.Instance.Form2_Text1) > 5000)
                    {
                        Form2.Instance.Form2_Text1 = "5000";
                    }
                    else if( int.Parse(Form2.Instance.Form2_Text1) < 1)
                    {
                        Form2.Instance.Form2_Text1 = "1";
                    }
                    if ( int.Parse(Form2.Instance.Form2_Text2) > 4096)
                    {
                        Form2.Instance.Form2_Text2 = "4096";
                    }
                    else if (int.Parse(Form2.Instance.Form2_Text2) < 1)
                    {
                        Form2.Instance.Form2_Text2 = "1";
                    }
                    //ダイアログでの設定値を現在の設定値にする
                    dt = int.Parse(Form2.Instance.Form2_Text1);
                    n = int.Parse(Form2.Instance.Form2_Text2);
                    //新設定値をArduinoに書き込み
                    arduino_send_recv("t" + dt.ToString());
                    arduino_send_recv("n" + n.ToString());

                    label4.Text = dt.ToString();
                    label5.Text = n.ToString();
                    label6.Text = "STATUS: Ready to RUN";
                    label6.BackColor = Color.LightCyan;

                }
                catch (Exception err)
                {
                    MessageBox.Show("設定: " + err.Message);
                }

            }
            else
            {
                MessageBox.Show("接続されていません。先にOpen COMボタンを押して接続してください。");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void add_serial_portname()
        {
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                 comboBox1.Items.Add(port);
            }
            comboBox1.SelectedIndex = comboBox1.Items.Count-1;
        }

        private void comclose()
        {
            if(serialPort1.IsOpen == true) 
            {
                arduino_send_recv("s");

                serialPort1.Close();
                MessageBox.Show("COMポートを切断しました");
                button1.Text = "Open COM";
                button1.BackColor = SystemColors.Control;
                label6.Text = "STATUS: Not Ready";
                label6.BackColor = Color.Yellow;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen == true)
            {
                Task<int> task = Task.Run(() => {
                    return arduino_run();
                });
            }
            else
            {
                MessageBox.Show("接続されていません。先にOpen COMボタンを押して接続してください。");
            }
        }

        private int arduino_run()
        {
            try
            {
                button1.Enabled = false;    // Open COM ボタンを無効に
                button2.Enabled = false;    // 設定 ボタンを無効に
                button3.Enabled = false;    // Run ボタンを無効に
                button4.Enabled = false;    // Save DATA ボタンを無効に
                button5.Enabled = true;     // 中断 ボタンを有効に

                serialPort1.Write("r");
                button3.Text = "Sampling...";
                button3.BackColor = Color.LightCyan;
                string data;

                while (serialPort1.BytesToRead == 0)
                {
                    // バッファにデータが溜まるまで待つ（rのエコーバック待ち）
                }
                for (int i = 0; i < n; i++)
                    {
                    if(cancel == true)
                    {
                        break;
                    }
                    recieved_str[i] = serialPort1.ReadLine();
                }

                if(cancel == true)
                {
                    arduino_send_recv("s"); 
                    chart1.Series.Clear();
                    MessageBox.Show("中止しました");
                    cancel = false;
                    serialPort1.ReadExisting(); //バッファを空に
                }
                else
                {
                    for (int i = 0; i < n; i++)
                    {
                        data = recieved_str[i];
                        fa0[i] = float.Parse(data.Substring(0, data.IndexOf(","))) * 5 / 1023;
                        fa1[i] = float.Parse(data.Substring(data.IndexOf(",") + 1)) * 5 / 1023;
                    }
                    draw_graph();
                }
                

                button3.Text = "Run";
                button3.BackColor = SystemColors.Control;
                label6.Text = "STATUS: Ready to RUN";
                label6.BackColor = Color.LightCyan;


                
                button1.Enabled = true;    // Open COM ボタンを有効に
                button2.Enabled = true;    // 設定 ボタンを有効に
                button3.Enabled = true;    // Run ボタンを有効に
                button4.Enabled = true;    // Save DATA ボタンを有効に
                button5.Enabled = false;   // 中断ボタンを無効に

            }
            catch (Exception err)
            {
                MessageBox.Show("Run エラー:" + err.Message);
            }
            return 1;
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            int i, t;
            string data;

            saveFileDialog1.FileName = "data01.csv";
            saveFileDialog1.Filter = "csv型式ファイル(*.csv)|*.csv";
            saveFileDialog1.Title = "Save an DATA File";
            saveFileDialog1.ShowDialog();

            System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
            System.Text.Encoding.GetEncoding("shift_jis");
            System.IO.StreamWriter sw = new System.IO.StreamWriter(fs);


            data = recieved_str[128];

            sw.WriteLine("# time (ms), a0 (V), a1 (V)");
            for (i=0; i < n ; i++)
            {
                t = i * dt;
                data = recieved_str[i];
                sw.WriteLine(t.ToString() + ", " + fa0[i].ToString("F4") + ", " + fa1[i].ToString("F4"));
            }
           
            sw.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
                       // 質問ダイアログを表示する
            DialogResult result = MessageBox.Show("終了しますか？", "質問", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(result == DialogResult.No)
            {
                // はいボタンをクリックしたときはウィンドウを閉じる
                e.Cancel = true;
            }
            else
            {
                comclose(); //COMポートを閉じて終了
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            cancel = true;
        }

        private void draw_graph()
        {
            Series analog0 = new Series();
            Series analog1 = new Series();
            string data;
            int i;

            chart1.ChartAreas.Clear();
            chart1.Series.Clear();

            // 「chartArea1」,「chartArea2」という名前の2つのエリアを生成
            ChartArea chartArea1 = new ChartArea("chartArea1");
            ChartArea chartArea2 = new ChartArea("chartArea2");
            // 生成したエリアをChartコントロールに追加
            chart1.ChartAreas.Add(chartArea1);
            chart1.ChartAreas.Add(chartArea2);

            for (i = 0; i < n; i++)
            {
                analog0.Points.AddXY(i*dt, fa0[i]);
                analog1.Points.AddXY(i*dt, fa1[i]);
            }
            Thread.Sleep(50);//少し待たないとchart1の準備ができないようだ。

            chart1.Series.Add(analog0);
            chart1.Series.Add(analog1);

            analog0.ChartArea = "chartArea1";
            analog1.ChartArea = "chartArea2";

            //描画範囲設定
            chart1.ChartAreas["chartArea1"].AxisX.Minimum = 0;
            chart1.ChartAreas["chartArea1"].AxisX.Maximum = n*dt-1;
            chart1.ChartAreas["chartArea1"].AxisY.Minimum = 0;
            chart1.ChartAreas["chartArea1"].AxisY.Maximum = 5.0;
            chart1.ChartAreas["chartArea2"].AxisX.Minimum = 0;
            chart1.ChartAreas["chartArea2"].AxisX.Maximum = n*dt-1;
            chart1.ChartAreas["chartArea2"].AxisY.Minimum = 0;
            chart1.ChartAreas["chartArea2"].AxisY.Maximum = 5.0;

            analog0.ChartType = SeriesChartType.Line;
            analog1.ChartType = SeriesChartType.Line;
            chart1.Series[0].IsVisibleInLegend = false;
            chart1.Series[1].IsVisibleInLegend = false;
        }


        private string arduino_send_recv(string send_message)
        {
            string recv_message;

            serialPort1.ReadExisting(); //バッファを空に

            serialPort1.Write(send_message); //メッセージ送信
            while (serialPort1.BytesToRead == 0)
            {
                // バッファにデータが溜まるまでまつ（エコーバック用）
            }
            recv_message = serialPort1.ReadLine();//メッセージ受信

            return recv_message;        
        }
    }
}
