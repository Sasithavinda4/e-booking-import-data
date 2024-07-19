using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace e_booking_import_data
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Load_Data();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtPath.Text = openFile.FileName;
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            string PathCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtPath.Text + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            OleDbConnection conn = new OleDbConnection(PathCon);
            OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + txtSname.Text + "$", conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void Load_Data()
        {
            SqlDataAdapter da = new SqlDataAdapter("Select * from tblStudents", ConClass.con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string PathCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtPath.Text + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            OleDbConnection conn = new OleDbConnection(PathCon);
            OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + txtSname.Text + "$", conn);
            DataTable dt = new DataTable();
            da.Fill(dt);

            for (int i = 0; i < Convert.ToInt32(dt.Rows.Count.ToString()); i++)
            {
                string SaveStr = "Insert into tblStudents (StudentID, StudentName, StudentCourse) values (@ID, @NAME, @COURSE)";
                SqlCommand SaveCmd = new SqlCommand(SaveStr, ConClass.con);
                SaveCmd.Parameters.AddWithValue("@ID", dt.Rows[i][0].ToString());
                SaveCmd.Parameters.AddWithValue("@NAME", dt.Rows[i][1].ToString());
                SaveCmd.Parameters.AddWithValue("@COURSE", dt.Rows[i][2].ToString());
                ConClass.con.Open();
                SaveCmd.ExecuteNonQuery();
                ConClass.con.Close();

            }
            Load_Data();
        }
    }
}
