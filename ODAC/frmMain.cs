using Dapper;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows.Forms;

namespace ODAC
{
    public partial class frmMain : Form
    {
        /// <summary>
        /// 
        /// </summary>
        int _length;

        /// <summary>
        /// 
        /// </summary>
        object[] _f1Array;

        /// <summary>
        /// 
        /// </summary>
        object[] _f2Array;

        public frmMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 執行查詢命令
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {

                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT SYSDATE FROM DUAL";
                Object obj = cmd.ExecuteScalar();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行新增命令
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {

                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO T1 (F1,F2) VALUES('SFIS','Jet')";
                int count = cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行有參數新增命令
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {

                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO T1 (F1,F2) VALUES(:F1,:F2)";
                cmd.Parameters.Add("F1", OracleDbType.Varchar2).Value = "SFIS";
                cmd.Parameters.Add("F2", OracleDbType.Varchar2).Value = "Eric";
                int count = cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        ///  執行更新命令
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE T1 SET F1 = 'SDD1' WHERE F2='Jet'";
                int count = cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行並取得第一列第一欄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {

                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM T1";
                Object obj = cmd.ExecuteScalar();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行並取得讀取器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM T1";
                OracleDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    Debug.WriteLine(dr.GetString(0) + dr.GetString(1));
                }
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行並取得讀取器(名稱取值)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT F2, F1 FROM T1";
                OracleDataReader dr = cmd.ExecuteReader();
                int idxF1 = dr.GetOrdinal("F1");
                int idxF2 = dr.GetOrdinal("F2");
                while (dr.Read())
                {
                    Debug.WriteLine(dr.GetString(idxF1) + dr.GetString(idxF2));
                }
                dr.Close();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行並將結果填至DataSet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM T1";
                OracleDataAdapter da = new OracleDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行並將結果填至DataTable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM T1";
                OracleDataAdapter da = new OracleDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 執行 Procedure (預儲程序)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";

            string ret = "";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "SDD1TEST.P1";

                cmd.Parameters.Add("P_1", OracleDbType.Varchar2, ParameterDirection.Input).Value = "SFIS";
                cmd.Parameters.Add("P_2", OracleDbType.Varchar2, ParameterDirection.Input).Value = "David";
                OracleParameter p3 = cmd.Parameters.Add("P_3", OracleDbType.Varchar2, ParameterDirection.Output);
                p3.Size = 10;

                int count = cmd.ExecuteNonQuery();
                ret = cmd.Parameters["P_3"].Value.ToString();

                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show(ret);
        }

        /// <summary>
        /// 集區連線要求逾時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            int count = 0;
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=1;Max Pool Size=3;Connection Timeout=1";
            try
            {
                while (count < 10)
                {
                    count++;
                    OracleConnection cn = new OracleConnection(connectionString);
                    cn.Open();
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT SYSDATE FROM DUAL";
                    Object obj = cmd.ExecuteScalar();
                }


            }
            catch (OracleException ex)
            {
                MessageBox.Show(string.Format("({0}){1}", count, ex.Message));
            }
        }

        /// <summary>
        /// 使用 using 確保關閉連線 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {
            int count = 0;
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=1;Max Pool Size=3;Connection Timeout=10";
            try
            {
                while (count < 10)
                {
                    count++;
                    using (OracleConnection cn = new OracleConnection(connectionString))
                    {
                        cn.Open();
                        OracleCommand cmd = new OracleCommand();
                        cmd.Connection = cn;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "SELECT SYSDATE FROM DUAL";
                        Object obj = cmd.ExecuteScalar();
                    }
                }
            }
            catch (OracleException ex)
            {
                MessageBox.Show(string.Format("({0}){1}", count, ex.Message));
            }
            MessageBox.Show("重複開啟連線完成");
        }

        /// <summary>
        /// 關閉連線時順便關閉讀取器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button13_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                using (OracleConnection cn = new OracleConnection(connectionString))
                {
                    cn.Open();
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM T1";
                    OracleDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 本機式交易
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleTransaction transaction = cn.BeginTransaction();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO T1 (F1,F2) VALUES('SDD1','Jetaa')";
                Object obj = cmd.ExecuteNonQuery();
                //transaction.Rollback(); // 還原
                transaction.Commit(); // 確認
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 明確式交易
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button15_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                CommittableTransaction transaction = new CommittableTransaction(); //必須參考 System.Transactions
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                cn.EnlistTransaction(transaction); // 從這裡就可以看出，可以跨連線是他的優點
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO T1 (F1,F2) VALUES('SDD1','Jetq')";
                int count = cmd.ExecuteNonQuery();

                transaction.Rollback(); // 還原
                //transaction.Commit(); // 確認

                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Executing SQL Scripts 執行一段PL SQL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button16_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"BEGIN
                                        INSERT INTO T1 (F1,F2) VALUES('SDD1','Jet');
                                        INSERT INTO T1 (F1,F2) VALUES('SFIS','David');
                                    END;";

                // ROLLBACK;
                int count = cmd.ExecuteNonQuery();


                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 大量存取批次執行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button17_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;

                object[] f1Array = { "SDD1", "SFIS", "SFIS" };
                object[] f2Array = { "Jet", "Eric", "David" };
                cmd.ArrayBindCount = 3; // 取得參數筆數

                cmd.CommandText = "INSERT INTO T1 (F1,F2) VALUES(:F1,:F2)";
                cmd.Parameters.Add("F1", OracleDbType.Varchar2).Value = f1Array;
                cmd.Parameters.Add("F2", OracleDbType.Varchar2).Value = f2Array;
                int count = cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Dapper (ORM套件)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button18_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                List<T1> ret = cn.Query<T1>("SELECT * FROM T1", null, commandType: CommandType.Text) as List<T1>;
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 準備資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button19_Click(object sender, EventArgs e)
        {
            _length = 10000;
            _f1Array = new object[_length];
            _f2Array = new object[_length];

            for (int i = 0; i < _length; i++)
            {
                _f1Array[i] = "SDD1";
                _f2Array[i] = string.Format("Jet{0}", i);
            }

            MessageBox.Show(_length + " 筆資料，準備完成");
        }

        /// <summary>
        /// 大量新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button20_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            string connectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.18.196.18)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=PTWSFDEV)));Persist Security Info=False;User ID=sdd1test;Password=PEGATRON;Min Pool Size=2;Max Pool Size=5;Connection Timeout=10";
            try
            {
                OracleConnection cn = new OracleConnection(connectionString);
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;

                cmd.ArrayBindCount = _length; // 取得參數筆數
                cmd.CommandTimeout = 100;
                cmd.CommandText = "INSERT INTO T1 (F1,F2) VALUES(:F1,:F2)";
                cmd.Parameters.Add("F1", OracleDbType.Varchar2).Value = _f1Array;
                cmd.Parameters.Add("F2", OracleDbType.Varchar2).Value = _f2Array;
                int count = cmd.ExecuteNonQuery();
                cn.Close();
            }
            catch (OracleException ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show(stopwatch.Elapsed.ToString());
        }
    }
}
