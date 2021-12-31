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
using System.Data.OleDb;




namespace WindowsFormsApplication5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OleDbConnection myConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:/cess2007file1.accdb;Persist Security Info=False;");
            try
            {
                myConn.Close();
                MessageBox.Show("ok");
                myConn.Close();
            }
            catch
            {
                MessageBox.Show("eroor");
            }

            /*
             * 
             *  OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\kenny\Documents\Visual Studio 2010\Projects\Copy Cegees\Cegees\Cegees\Login.accdb";

            String Username = TEXTNewUser.Text;
            String Password = TEXTNewPass.Text;

            OleDbCommand cmd = new OleDbCommand("INSERT into Login (Username, Password) Values(@Username, @Password)");
            cmd.Connection = conn;

            conn.Open();

            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Username", OleDbType.VarChar).Value = Username;
                cmd.Parameters.Add("@Password", OleDbType.VarChar).Value = Password;

                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Data Added");
                    conn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Source);
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Connection Failed");
            }
            */

            /*
             *  myConn.Open();
            OleDbCommand myQuery = new OleDbCommand("select CustID from Customers WHERE CustID = 1;", myConn);
            OleDbDataReader myReader = myQuery.ExecuteReader();
            if (myReader.HasRows)
            {
                myReader.Read();
                label1.Text = myReader.ToString();
            }
            */

            // CreateNewAccessDatabase(@"D:\cess2007file1.accdb");
            /*  ADOX.Catalog cat = new ADOX.Catalog();

              cat.Create(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\myAccess2007file.accdb;");

              Console.WriteLine("Database Created Successfully");

              cat = null;*/
            /*  OleDbConnection myConnection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source= D:\myAccess2007file.accdb \\ConfigStructure.mdb");

              myConnection.Open();
              string strTemp = " [KEY] Text, [VALUE] Text ";
              OleDbCommand myCommand = new OleDbCommand();
              myCommand.Connection = myConnection;
              myCommand.CommandText = "CREATE TABLE table1(" + strTemp + ")";
              myCommand.ExecuteNonQuery();
              myCommand.Connection.Close();*/




        }
        public void CreateNewAccessDatabase(string fileName)
        {
            bool result = false;

            ADOX.Catalog cat = new ADOX.Catalog();
            ADOX.Table table = new ADOX.Table();
            ADOX.Table table1 = new ADOX.Table();
            ADOX.Table table2 = new ADOX.Table();
            ADOX.Table table3 = new ADOX.Table();
            ADOX.Table table4 = new ADOX.Table();
            ADOX.Table table5 = new ADOX.Table();

            //Create the table and it's fields. 
            table.Name = "table_class_student";
            table.Columns.Append("id");
            table.Columns.Append("name_class");
            table.Columns.Append("student_id");
            
            table1.Name = "table_classes";
            table1.Columns.Append("id_classes");
            table1.Columns.Append("name_classes");

            table2.Name = "table_courese_semester";
            table2.Columns.Append("id_courese_student");
            table2.Columns.Append("semester_id");
            table2.Columns.Append("student_id");
            table2.Columns.Append("courses_id");
            table2.Columns.Append("d1");
            table2.Columns.Append("d2");
            table2.Columns.Append("d_total");
            table2.Columns.Append("success_t_f");
            table2.Columns.Append("created_at");
            table2.Columns.Append("update_at");

            table3.Name = "table_courese_semester_old";
            table3.Columns.Append("id_courese_old_student");
            table3.Columns.Append("semester_old_id");
            table3.Columns.Append("student_id");
            table3.Columns.Append("courses_id");
            table3.Columns.Append("d1");
            table3.Columns.Append("d2");
            table3.Columns.Append("d_total");
            table3.Columns.Append("success_t_f");
            table3.Columns.Append("number_type");
            table3.Columns.Append("created_at");
            table3.Columns.Append("update_at");

            table4.Name = "table_coureses";
            table4.Columns.Append("id_courese");
            table4.Columns.Append("name_courese");
            table4.Columns.Append("code_courese");
            table4.Columns.Append("courese_prer");
            table4.Columns.Append("created_at");
            table4.Columns.Append("update_at");
            table4.Columns.Append("unty_courese");
            table4.Columns.Append("type_courese");
            table4.Columns.Append("classess_id");

            table5.Name = "Table_coureses_prer";
            table5.Columns.Append("id_prer");
            table5.Columns.Append("name_courese_prer");
            table5.Columns.Append("code_couresr_prer");
            table5.Columns.Append("created_at");
            table5.Columns.Append("update_at");
            table5.Columns.Append("courese_id");
            
                cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");
                //cat.Tables.Append(table);
                cat.Tables.Append(table1);
                cat.Tables.Append(table2);
                cat.Tables.Append(table3);
                cat.Tables.Append(table4);
                cat.Tables.Append(table5);

                //Now Close the database
                ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;
                if (con != null)
                    con.Close();

              
            
           


        }
    }
}
