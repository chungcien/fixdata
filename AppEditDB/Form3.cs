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

namespace AppEditDB
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }


        DataTable data = new DataTable();
        List<string> user_file = new List<string>();

        Dictionary<string, string> User_Active = new Dictionary<string, string>();
        Dictionary<string, string> _User_DU = new Dictionary<string, string>();
        DataProvider dataProvider = new DataProvider();

        private void Form3_Load(object sender, EventArgs e)
        {
            String name = "Sheet1";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "DU.xlsx" +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);

            sda.Fill(data);

            data.Columns.Add("Username");
            data.Columns.Add("DU in DB");

            //get DU from DB
            DataTable User_DU = dataProvider.getDatatable("  select UserName, Department.Name, DepartmentID from AppUsers, Department, Position where AppUsers.PositionID = Position.ID and Position.DepartmentID = Department.ID");

            int _count = 0;

            for (int i = 0; i < data.Rows.Count; i++)
            {
                //parse Username
                data.Rows[i]["Username"] = data.Rows[i]["Email"].ToString().Replace("@cmc.com.vn", "");

                string temp = data.Rows[i]["Username"].ToString();

                user_file.Add(temp);

                for (int j = 0; j < User_DU.Rows.Count; j++)
                {
                    //_User_DU.Add(User_DU.Rows[i]["UserName"].ToString(), User_DU.Rows[i]["Name"].ToString());

                    if (User_DU.Rows[j]["UserName"].ToString() == temp)
                    {
                        data.Rows[i]["DU in DB"] = User_DU.Rows[j]["Name"].ToString();

                        break;
                    }
                }

                if (data.Rows[i]["DU in DB"].ToString().ToLower() != data.Rows[i]["Bộ phận chuyển đến"].ToString().ToLower())
                {
                    listBox1.Items.Add("Bộ phận sai: " + temp + ": " + data.Rows[i]["DU in DB"].ToString() + " -> " + data.Rows[i]["Bộ phận chuyển đến"].ToString());
                    _count++;
                }

                try
                {
                    User_Active.Add(temp, temp);
                }
                catch (Exception op)
                {
                    listBox1.Items.Add("trùng " + temp);
                }
            }

            listBox1.Items.Add("Bộ phận sai: " + _count);
            dataGridView1.DataSource = data;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int count = 0;

            for (int i = 0; i < data.Rows.Count; i++)
            {

                int tem = dataProvider.ExecuteNonQuery("update [UserGroup] set [EndOff] = " + data.Rows[i]["Điều chuyển từ"].ToString() + " where [IdUser] = (select ID from AppUsers where Username = '" + data.Rows[i]["UserName"].ToString() + "')");
                count = count + tem;

            }

            listBox1.Items.Add("Có " + count + " user đang làm việc.");
        }
    }
}
