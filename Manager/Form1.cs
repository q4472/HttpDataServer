using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Manager
{
    public partial class Form1 : Form
    {
        private static DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Загрузить в память ПереводыДляОбеспеченияКонтрактов из общей базы
            // Загрузка с SQL сервера в память
            StringBuilder sb = new StringBuilder();
            try
            {
                dt = SqlServer.DownloadPaymentList();
                /*
                foreach (DataRow dr in dt.Rows)
                {
                    sb.AppendFormat("{0:yyyy-MM-dd} ", dr["Дата"]);
                    sb.AppendFormat("{0} ", dr["Номер"]);
                    sb.AppendFormat("{0} ", dr["Сумма"]);
                    sb.Append("\n");
                }
                */
            }
            catch (Exception ex) { richTextBox1.Text += ex.ToString(); return; }
            if (dt != null)
            {
                sb.AppendFormat("Загружено в память {0} строк.", dt.Rows.Count);
            }
            else
            {
                sb.Append("Ничего не загружено в память.");
            }
            WriteLine(sb.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Выгрузить в 1с77 ПереводыДляОбеспеченияКонтрактов из памяти
            StringBuilder sb = new StringBuilder();
            if (dt == null)
            {
                sb.Append("Ничего не загружено в память.");
            }
            else
            {
                String msg = OcStoredProcedure.Exec0(dt, true);
                sb.AppendFormat("{0}", msg);
            }
            WriteLine(sb.ToString());
        }
        private void WriteLine(String msg)
        {
            richTextBox1.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + msg + "\n" + richTextBox1.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Загрузить в общую базу Переводы из памяти
            if ((dt != null) && (dt.Rows.Count > 0))
            {
                SqlServer.UploadPaymentList(dt);
                WriteLine(dt.Rows.Count.ToString());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Загрузить в память Переводы из 1с83
            Dictionary<String, Object> pars = new Dictionary<String, Object>();
            pars.Add("НачалоПериода", new DateTime(2018, 7, 1, 0, 0, 0));
            pars.Add("КонецПериода", new DateTime(2018, 12, 31, 23, 59, 59));

            DataSet ds = OcStoredProcedures.Exec0("GetPaymentList", pars);
            
            if ((ds != null) && (ds.Tables.Count > 0))
            {
                dt = ds.Tables[0];
                WriteLine(dt.Rows.Count.ToString());
            }
        }
    }
}
