using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace AppEditDB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable data = new DataTable();
        List<string> user_file = new List<string>();
        private void Form1_Load(object sender, EventArgs e)
        {
            String name = "Sheet1";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "temp.xlsx" +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);

            sda.Fill(data);

            data.Columns.Add("Username");

            for(int i = 0; i< data.Rows.Count; i++)
            {

                data.Rows[i]["Username"] = data.Rows[i]["Email"].ToString().Replace("@cmc.com.vn", "");

                user_file.Add(data.Rows[i]["Username"].ToString());

                string temp = data.Rows[i]["Ngày kết thúc hợp đồng"].ToString();
                List<string> dat = temp.Split('/').ToList();

                data.Rows[i]["Ngày kết thúc hợp đồng"] = Convert.ToDateTime(dat[1] + "/" + dat[0] + "/" + dat[2]);
            }

            
            dataGridView1.DataSource = data;
        }

        DataProvider dataProvider = new DataProvider();
        private void button1_Click(object sender, EventArgs e)
        {
            string user_nghi_sau_31_3 = string.Join("', '", user_file);

            DataTable dt = dataProvider.getDatatable("select * from AppUsers where Status = 0 and Username not in ( '"+ user_nghi_sau_31_3 + "') ");

            listBox1.Items.Add("Có " + dt.Rows.Count + " user nghỉ trước 31/3.");

            int tem = dataProvider.ExecuteNonQuery("update [UserGroup] set [StartOff] = '" + new DateTime(2019, 1, 1) + "', [EndOff] = '" + new DateTime(2019, 3, 31) + "' where [IdUser] in (select ID from AppUsers where Status = 0 and Username not in ( '" + user_nghi_sau_31_3 + "') )");

            listBox1.Items.Add("updated " + tem + " rows");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string user_nghi_sau_31_3 = string.Join("', '", user_file);

            listBox1.Items.Add("Có " + data.Rows.Count + " user nghỉ sau 31/3.");
            int count = 0;
            for (int i = 0; i< data.Rows.Count; i++)
            {
                dataProvider.ExecuteNonQuery("update appusers set status = 0 where Username in ('" + user_nghi_sau_31_3 + "'); ");
                int tem = dataProvider.ExecuteNonQuery("update [UserGroup] set [StartOff] = '" + new DateTime(2019, 4, 1) + "', [EndOff] = '" + data.Rows[i]["Ngày kết thúc hợp đồng"] + "' where [IdUser] in (select ID from AppUsers where Username = '"+ data.Rows[i]["Username"] + "')");

                if(tem > 0)
                {
                    listBox1.Items.Add((i + 1) + ": updated " + data.Rows[i]["Username"] + " endDate là " + data.Rows[i]["Ngày kết thúc hợp đồng"]);
                    count++;
                }

            }
            listBox1.Items.Add("updated " + count + " user.");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int tem = dataProvider.ExecuteNonQuery("update [UserGroup] set [StartOff] = '" + new DateTime(2019, 4, 1) + "', [EndOff] = NULL where [IdUser] in (select ID from AppUsers where Status = 1)");

            listBox1.Items.Add("Có " + tem + " user đang làm việc.");
        }
    }
}
