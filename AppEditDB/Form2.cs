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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        DataTable data = new DataTable();
        List<string> user_file = new List<string>();

        Dictionary<string, string> User_Active = new Dictionary<string, string>();
        Dictionary<string, string> _User_DU = new Dictionary<string, string>();


        private void Form2_Load(object sender, EventArgs e)
        {
            String name = "Sheet1";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            "CBCNV.xlsx" +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);

            sda.Fill(data);

            data.Columns.Add("Username");
            data.Columns.Add("Ngày bắt đầu hợp đồng");
            data.Columns.Add("Ngày kết thúc hợp đồng");
            data.Columns.Add("DU in DB");

            //get DU from DB
            DataTable User_DU = dataProvider.getDatatable("  select UserName, Department.Name from AppUsers, Department, Position where AppUsers.PositionID = Position.ID and Position.DepartmentID = Department.ID");

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

                    if(User_DU.Rows[j]["UserName"].ToString() == temp)
                    {
                        data.Rows[i]["DU in DB"] = User_DU.Rows[j]["Name"].ToString();

                        break;
                    }
                }

                if(data.Rows[i]["DU in DB"].ToString().ToLower() != data.Rows[i]["Bộ phận"].ToString().ToLower())
                {
                    listBox1.Items.Add("Bộ phận sai: " + temp + ": " + data.Rows[i]["DU in DB"].ToString() + " -> " + data.Rows[i]["Bộ phận"].ToString());
                    _count++;
                }

                try
                {
                    User_Active.Add(temp, temp);
                }
                catch(Exception op)
                {
                    listBox1.Items.Add("trùng " + temp);
                }


                string date_start = data.Rows[i]["Ngày bắt đầu"].ToString();

                string date_end = data.Rows[i]["Ngày chấm dứt"].ToString();

                if(!string.IsNullOrEmpty( date_end))
                {
                    List<string> lis_date_end = date_end.Split('/').ToList();

                    data.Rows[i]["Ngày kết thúc hợp đồng"] = "'" + Convert.ToDateTime(lis_date_end[1] + "/" + lis_date_end[0] + "/" + lis_date_end[2])+"'";
                }
                else
                {
                    data.Rows[i]["Ngày kết thúc hợp đồng"] = "NULL";
                }

                if (!string.IsNullOrEmpty(date_start))
                {
                    List<string> lis_date_start = date_start.Split('/').ToList();

                    data.Rows[i]["Ngày bắt đầu hợp đồng"] = Convert.ToDateTime(lis_date_start[1] + "/" + lis_date_start[0] + "/" + lis_date_start[2]);

                    if((lis_date_start[2] + "/" + lis_date_start[1] + "/" + lis_date_start[0]).CompareTo("2019/04/01") <= 0)
                    {
                        data.Rows[i]["Ngày bắt đầu hợp đồng"] = Convert.ToDateTime("04/01/2019");
                    }
                }
                else
                {
                    data.Rows[i]["Ngày bắt đầu hợp đồng"] = Convert.ToDateTime("04/01/2019");
                }


            }

            listBox1.Items.Add("Bộ phận sai: " + _count);
            dataGridView1.DataSource = data;
        }
        DataProvider dataProvider = new DataProvider();
        private void button1_Click(object sender, EventArgs e)
        {
            string user_Active = string.Join("', '", user_file);

            DataTable dt = dataProvider.getDatatable("select * from AppUsers where Status = 1 and Username not in ( '" + user_Active + "') ");

            listBox1.Items.Add("Có " + dt.Rows.Count + " user đã nghỉ nhưng chưa inactive");

            int tem = dataProvider.ExecuteNonQuery("update [AppUsers] set [Status] = 0 where Username not in ( '" + user_Active + "')");

            listBox1.Items.Add("updated " + tem + " User inActive");

            tem = dataProvider.ExecuteNonQuery("update [AppUsers] set [Status] = 1 where Username in ( '" + user_Active + "')");

            listBox1.Items.Add("updated " + tem + " User Active ");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int count = 0;

            for(int i = 0; i< data.Rows.Count; i++)
            {

                int tem = dataProvider.ExecuteNonQuery("update [UserGroup] set [StartOff] = '" + data.Rows[i]["Ngày bắt đầu hợp đồng"].ToString() + "', [EndOff] = " + data.Rows[i]["Ngày kết thúc hợp đồng"].ToString()+" where [IdUser] = (select ID from AppUsers where Username = '"+ data.Rows[i]["UserName"].ToString() + "')");
                count = count + tem;

            }

            listBox1.Items.Add("Có " + count + " user đang làm việc.");

        }


    }
}
