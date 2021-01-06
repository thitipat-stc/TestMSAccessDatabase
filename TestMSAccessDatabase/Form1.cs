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

namespace TestMSAccessDatabase
{
    public partial class Form1 : Form
    {
        OleDbConnection con;
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;

        private string ConnectionString = @"Provider = Microsoft.ACE.Oledb.12.0; Data Source = C:\Users\Pattithi\source\repos\TestMSAccessDatabase\Database.accdb";
        private AccessDbms Dbms { get; set; } = new AccessDbms();

        public Form1()
        {
            InitializeComponent();
            Dbms.ConnectionString = this.ConnectionString;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            GetStudent();
        }

        private void GetStudent()
        {
            /*con = new OleDbConnection(@"Provider=Microsoft.ACE.Oledb.12.0;Data Source=C:\Users\Pattithi\source\repos\TestMSAccessDatabase\Database.accdb");
            da = new OleDbDataAdapter("SELECT * FROM Tbl_User", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "Tbl_User");
            dataGridView1.DataSource = ds.Tables["Tbl_User"];
            con.Close();*/

            DataTable dt = Dbms.GetDataTable("SELECT * FROM Tbl_User");
            dataGridView1.DataSource = dt;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            /*string query = "Insert into Tbl_User (Name,Age) values (@Name,@Age)";
            cmd = new OleDbCommand(query, con);
            cmd.Parameters.AddWithValue("@Name", tbName.Text);
            cmd.Parameters.AddWithValue("@Age", Convert.ToInt32(tbAge.Text));
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();*/

            List<OleDbParameter> parameters = new List<OleDbParameter>();
            parameters.Add(new OleDbParameter("@Name", tbName.Text));
            parameters.Add(new OleDbParameter("@Age", Convert.ToInt32(tbAge.Text)));
            Dbms.Execute("Insert into Tbl_User (Name,Age) values (@Name,@Age)", parameters);

            GetStudent();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            /*string query = "Update Tbl_User Set Name=@Name,Age=@Age Where ID=@ID";
            cmd = new OleDbCommand(query, con);
            cmd.Parameters.AddWithValue("@Name", tbName.Text);
            cmd.Parameters.AddWithValue("@Age", Convert.ToInt32(tbAge.Text));
            cmd.Parameters.AddWithValue("@ID", Convert.ToInt32(tbID.Text));
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            GetStudent();*/

            List<OleDbParameter> parameters = new List<OleDbParameter>();
            parameters.Add(new OleDbParameter("@Age", Convert.ToInt32(tbAge.Text)));
            parameters.Add(new OleDbParameter("@Name", tbName.Text));
            parameters.Add(new OleDbParameter("@ID", Convert.ToInt32(tbID.Text)));
            Dbms.Execute("Update Tbl_User Set Name=@Name, Age=@Age Where ID=@ID", parameters);

            GetStudent();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            /*string query = "Delete From Tbl_User Where Id = @ID";
            cmd = new OleDbCommand(query, con);
            cmd.Parameters.AddWithValue("@ID", dataGridView1.CurrentRow.Cells[0].Value);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();*/

            List<OleDbParameter> parameters = new List<OleDbParameter>(); 
            parameters.Add(new OleDbParameter("@ID", Convert.ToInt32(tbID.Text)));
            Dbms.Execute("Delete From Tbl_User Where Id = @ID", parameters);

            GetStudent();
        }

        private void gridView1_Click(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                tbID.Text = row.Cells[0].Value.ToString();
                tbName.Text = row.Cells[1].Value.ToString();
                tbAge.Text = row.Cells[2].Value.ToString();
            }
        } 
    }
}
