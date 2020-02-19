using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.Drawing.Printing;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        SqlConnection SqlConnection;

       // MySqlConnection mySqlConnection;
        MySqlConnection SqlConnection2;



        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "saloonDataSet4.service". При необходимости она может быть перемещена или удалена.
            this.serviceTableAdapter1.Fill(this.saloonDataSet4.service);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "saloonDataSet3.master". При необходимости она может быть перемещена или удалена.
            this.masterTableAdapter2.Fill(this.saloonDataSet3.master);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "saloonDataSet2.service". При необходимости она может быть перемещена или удалена.
            //       this.serviceTableAdapter.Fill(this.saloonDataSet2.service);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "saloonDataSet1.master". При необходимости она может быть перемещена или удалена.
            //       this.masterTableAdapter1.Fill(this.saloonDataSet1.master);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "saloonDataSet.master". При необходимости она может быть перемещена или удалена.
            //     this.masterTableAdapter.Fill(this.saloonDataSet.master);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataDataSet4.master". При необходимости она может быть перемещена или удалена.
            //   this.masterTableAdapter1.Fill(this.dataDataSet4.master);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataDataSet3.service". При необходимости она может быть перемещена или удалена.
            //    this.serviceTableAdapter1.Fill(this.dataDataSet3.service);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataDataSet2.master". При необходимости она может быть перемещена или удалена.
            //     this.masterTableAdapter.Fill(this.dataDataSet2.master);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataDataSet1.service". При необходимости она может быть перемещена или удалена.
            //      this.serviceTableAdapter.Fill(this.dataDataSet1.service);


            string connectionString = @"Data Source=DESKTOP-FITMBD4;Initial Catalog=saloon;" + "Integrated Security=true;";   //Подключение к Sql
            SqlConnection = new SqlConnection(connectionString);
            SqlConnection.Open();

            //  Начальный вывод таблицы

            string query = "SELECT salon.ID,NCust, service.Nservice, master.FName,service.Price FROM salon inner join service on salon.ID_service = service.ID_service inner join master on salon.ID_master = master.ID_master order by ID ";

            SqlCommand command = new SqlCommand(query, SqlConnection);

            SqlDataReader reader = command.ExecuteReader();

            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            foreach (string[] s in data)
                dataGridView1.Rows.Add(s);

        }


        private async void button3_Click(object sender, EventArgs e)
        {
            if (label9.Visible)
                label9.Visible = false;

            if (!string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox8.Text))
            {
                SqlCommand command = new SqlCommand("DELETE FROM [salon] where [ID]=@Shop", SqlConnection); //Удаление

                command.Parameters.AddWithValue("Shop ", textBox8.Text);

                await command.ExecuteNonQueryAsync();


            }
            else
            {
                label9.Visible = true;

                label9.Text = "Заполните ID!"; //Проверка

            }
        }
            
        private void RefToolStripMenuItem_Click(object sender, EventArgs e) //Обновление при смене вкладок
        {

            dataGridView1.Rows.Clear();

            string query = "SELECT salon.ID,NCust, service.Nservice, master.FName,service.Price FROM salon inner join service on salon.ID_service = service.ID_service inner join master on salon.ID_master = master.ID_master order by ID ";

            SqlCommand command = new SqlCommand(query, SqlConnection);

            SqlDataReader reader = command.ExecuteReader();

            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            foreach (string[] s in data)
                dataGridView1.Rows.Add(s);
        }



        private async void button1_Click(object sender, EventArgs e) //Добавление
        {
            SqlCommand command = new SqlCommand("INSERT INTO salon (NCust, ID_master, ID_service) VALUES (@NCust, @ID_master, @Nservice)", SqlConnection);

            if (label11.Visible)
                label11.Visible = false;

            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text) &&
               (!string.IsNullOrEmpty(comboBox2.Text) && !string.IsNullOrWhiteSpace(comboBox2.Text)) &&
               (!string.IsNullOrEmpty(comboBox1.Text) && !string.IsNullOrWhiteSpace(comboBox1.Text)))
            {
                command.Parameters.AddWithValue("NCust", textBox1.Text);
                command.Parameters.AddWithValue("ID_master", comboBox2.SelectedValue);
                command.Parameters.AddWithValue("Nservice", comboBox1.SelectedValue);

                await command.ExecuteNonQueryAsync();

            }

            else
            {
                label11.Visible = true;

                label11.Text = "Что то не то !";

            }
        }

        private void аЯToolStripMenuItem_Click(object sender, EventArgs e)//Сортировка А-Я
        {

            dataGridView1.Rows.Clear();
            string query = "SELECT salon.ID,NCust, service.Nservice, master.FName,service.Price FROM salon inner join service on salon.ID_service = service.ID_service inner join master on salon.ID_master = master.ID_master order by ID ";

            SqlCommand command = new SqlCommand(query, SqlConnection);

            SqlDataReader reader = command.ExecuteReader();

            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            foreach (string[] s in data)
                dataGridView1.Rows.Add(s);
        }


        private void яАToolStripMenuItem_Click(object sender, EventArgs e)//Сортировка Я-А
        {

            dataGridView1.Rows.Clear();

            string query = "SELECT salon.ID,NCust, service.Nservice, master.FName,service.Price FROM salon inner join service on salon.ID_service = service.ID_service inner join master on salon.ID_master = master.ID_master order by ID desc";

            SqlCommand command = new SqlCommand(query, SqlConnection);

            SqlDataReader reader = command.ExecuteReader();

            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            foreach (string[] s in data)
                dataGridView1.Rows.Add(s);
        }

        private async void button2_Click(object sender, EventArgs e)//Редактировние
        {
            if (label12.Visible)
                label12.Visible = false;


            //ПРОВЕРКА ЧТО ПОЛЯ ЗАПОЛНЕНЫ
            if (!string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrWhiteSpace(textBox7.Text) &&
                !string.IsNullOrEmpty(comboBox4.Text) && !string.IsNullOrWhiteSpace(comboBox4.Text) &&
                !string.IsNullOrEmpty(comboBox5.Text) && !string.IsNullOrWhiteSpace(comboBox5.Text) &&
                 !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
            {
                SqlCommand command = new SqlCommand("UPDATE  salon set  [ID_master]=@ID_master ,[ID_service]=@ID_service,  [NCust]=@NCust where [ID]=@ID ", SqlConnection);

                command.Parameters.AddWithValue("ID", textBox7.Text);
                command.Parameters.AddWithValue("ID_master", comboBox4.SelectedValue);
                command.Parameters.AddWithValue("ID_service", comboBox5.SelectedValue);
                command.Parameters.AddWithValue("NCust", textBox2.Text);

                await command.ExecuteNonQueryAsync();

            }
            else if (!string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrWhiteSpace(textBox7.Text))
            {
                label12.Visible = true;

                label12.Text = "ID не заполнен!";//Проверка на существание ID

            }
            else
            {

                label12.Visible = true;

                label12.Text = "Поля '1','2' и '3' не заполнены!";//Проверка на заполненость строка радактирования

            }

        }

        private void button4_Click(object sender, EventArgs e) //Проверка и заполнение выпадающих списков
        {

            SqlCommand cmd = new SqlCommand("SELECT count(*) FROM [salon] WHERE ID = @ID ", SqlConnection);//Количество строк соответствует where
            cmd.Parameters.AddWithValue("@id", textBox7.Text);
            int countRows = (int)cmd.ExecuteScalar();
            if (countRows != 0) //Если ноль тогда мимо
            {
                SqlCommand command = new SqlCommand("SELECT * FROM salon inner join service on salon.ID_service = service.ID_service inner join master on salon.ID_master = master.ID_master  WHERE ID = @ID ", SqlConnection);


                command.Parameters.AddWithValue("ID", textBox7.Text);


                SqlDataReader reader = command.ExecuteReader();

                reader.Read();
                comboBox4.Text = (Convert.ToString(reader["ID_master"]));

                comboBox5.Text = (Convert.ToString(reader["ID_service"]));



                MessageBox.Show("Такая запись есть");
                comboBox4.Text = (Convert.ToString(reader["FName"]));
                comboBox5.Text = (Convert.ToString(reader["Nservice"]));

                textBox2.Text = (Convert.ToString(reader["NCust"]));


                reader.Close();
            }
            else
            {
                MessageBox.Show("Такой записи нет");
            }

        }

        private void button7_Click(object sender, EventArgs e) //Подключение Sql MySql
        {

            string connectionString = @"Data Source=DESKTOP-FITMBD4;Initial Catalog=saloon;" + "Integrated Security=true;";//Sql
            SqlConnection = new SqlConnection(connectionString);
            SqlConnection.Open();

            string connectionString2 = "server=localhost; user=root; database=saloon; password='';";//Mysql
            SqlConnection2 = new MySqlConnection(connectionString2);
            SqlConnection2.Open();




            // Проверка на наличение таблицы master если есть удаляем
            MySqlCommand check2 = new MySqlCommand("SELECT count(*) as Exist from INFORMATION_SCHEMA.TABLES WHERE table_name = 'master'", SqlConnection2);
            if ((long)check2.ExecuteScalar() == 1)
            {


                MySqlCommand drop2 = new MySqlCommand("DROP TABLE master", SqlConnection2);//Удаление
                MySqlDataReader reader2 = drop2.ExecuteReader();
                reader2.Close();
                h_master();
            }
            else
            {

                h_master();

            }

            // Проверка на наличение таблицы salon если есть удаляем
            MySqlCommand check1 = new MySqlCommand("SELECT count(*) as Exist from INFORMATION_SCHEMA.TABLES WHERE table_name = 'salon'", SqlConnection2);
            if ((long)check1.ExecuteScalar() == 1)
            {
                MySqlCommand drop1 = new MySqlCommand("DROP TABLE salon", SqlConnection2);//Удаление
                MySqlDataReader reader1 = drop1.ExecuteReader();
                reader1.Close();
                h_salon();
            }
            else
            {

                h_salon();
            }

            // Проверка на наличение таблицы service если есть удаляем
            MySqlCommand check3 = new MySqlCommand("SELECT count(*) as Exist from INFORMATION_SCHEMA.TABLES WHERE table_name = 'service'", SqlConnection2);
            if ((long)check3.ExecuteScalar() == 1)
            {

                MySqlCommand drop3 = new MySqlCommand("DROP TABLE service", SqlConnection2);//Удаление
                MySqlDataReader reader3 = drop3.ExecuteReader();
                reader3.Close();
                h_service();
            }
            else
            {

                h_service();
            }


        }

        public int Flag = 0;

        public object Response { get; private set; }

        public void h_master() //Создание таблицы и перенос данных
        {
            MySqlCommand Create_table = new MySqlCommand("CREATE TABLE master (ID_master INT NOT NULL , FName VARCHAR(30) NOT NULL, MName VARCHAR(30) NOT NULL,  LName VARCHAR(30) NOT NULL )", SqlConnection2);
            Create_table.ExecuteNonQuery();
            SqlDataReader sqlReader = null;
            SqlCommand command_0 = new SqlCommand("SELECT count(*) FROM [master]", SqlConnection);
            Flag = (int)command_0.ExecuteScalar(); // количество строк в таблице
            while (Flag != 0)
            {
                SqlCommand command_1 = new SqlCommand("SELECT * FROM (SELECT [FName], [MName], [LName], ROW_NUMBER() OVER(ORDER BY ID_master DESC) AS ROW FROM [master]) AS TMP WHERE ROW =" + Flag + "", SqlConnection);
                sqlReader = command_1.ExecuteReader(); // присваивание возвращаемого значения
                sqlReader.Read();
                string FName = Convert.ToString(sqlReader["FName"]);
                string MName = Convert.ToString(sqlReader["MName"]);
                string LName = Convert.ToString(sqlReader["LName"]);



                MySqlCommand command_2 = new MySqlCommand("INSERT INTO master (FName, MName, LName) VALUES (@FName, @MName, @LName)", SqlConnection2);
                command_2.Parameters.AddWithValue("FName", FName);
                command_2.Parameters.AddWithValue("MName", MName);
                command_2.Parameters.AddWithValue("LName", LName);

                command_2.ExecuteNonQuery();
                sqlReader.Close();
                Flag -= 1;
            }
            MessageBox.Show("Таблица master создана и заполнена!");
        }

        public void h_salon()
        {
            MySqlCommand Create_table = new MySqlCommand("CREATE TABLE salon (ID INT NOT NULL AUTO_INCREMENT, ID_master INT NOT NULL, NCust VARCHAR(50) NOT NULL,  ID_service int NOT NULL, PRIMARY KEY (ID))", SqlConnection2);
            Create_table.ExecuteNonQuery();
            SqlDataReader sqlReader = null;
            SqlCommand command_0 = new SqlCommand("SELECT count(*) FROM [salon]", SqlConnection);
            Flag = (int)command_0.ExecuteScalar(); // количество строк в таблице
            while (Flag != 0)
            {
                SqlCommand command_1 = new SqlCommand("SELECT * FROM (SELECT [ID_master], [NCust], [ID_service], ROW_NUMBER() OVER(ORDER BY ID DESC) AS ROW FROM [salon]) AS TMP WHERE ROW =" + Flag + "", SqlConnection);
                sqlReader = command_1.ExecuteReader(); // присваивание возвращаемого значения
                sqlReader.Read();
                string ID_master = Convert.ToString(sqlReader["ID_master"]);
                string NCust = Convert.ToString(sqlReader["NCust"]);
                string ID_service = Convert.ToString(sqlReader["ID_service"]);


                MySqlCommand command_2 = new MySqlCommand("INSERT INTO salon (ID_master, NCust, ID_service) VALUES (@ID_master, @NCust, @ID_service)", SqlConnection2);
                command_2.Parameters.AddWithValue("ID_master", ID_master);
                command_2.Parameters.AddWithValue("NCust", NCust);
                command_2.Parameters.AddWithValue("ID_service", ID_service);
                command_2.ExecuteNonQuery();
                sqlReader.Close();
                Flag -= 1;
            }
            MessageBox.Show("Таблица salon создана и заполнена!");
        }

        public void h_service()
        {
            MySqlCommand Create_table = new MySqlCommand("CREATE TABLE service (ID_service INT NOT NULL AUTO_INCREMENT, Nservice VARCHAR(50) NOT NULL, PRIMARY KEY (ID_service))", SqlConnection2);
            Create_table.ExecuteNonQuery();
            SqlDataReader sqlReader = null;
            SqlCommand command_0 = new SqlCommand("SELECT count(*) FROM [service]", SqlConnection);
            Flag = (int)command_0.ExecuteScalar(); // количество строк в таблице
            while (Flag != 0)
            {
                SqlCommand command_1 = new SqlCommand("SELECT * FROM (SELECT [ID_service], [Nservice], ROW_NUMBER() OVER(ORDER BY ID_service DESC) AS ROW FROM [service]) AS TMP WHERE ROW =" + Flag + "", SqlConnection);
                sqlReader = command_1.ExecuteReader(); // присваивание возвращаемого значения
                sqlReader.Read();
                string ID_service = Convert.ToString(sqlReader["ID_service"]);
                string Nservice = Convert.ToString(sqlReader["Nservice"]);


                MySqlCommand command_2 = new MySqlCommand("INSERT INTO service (ID_service, Nservice) VALUES (@ID_service, @Nservice)", SqlConnection2);
                command_2.Parameters.AddWithValue("ID_service", ID_service);
                command_2.Parameters.AddWithValue("Nservice", Nservice);
                command_2.ExecuteNonQuery();
                sqlReader.Close();
                Flag -= 1;
            }
            MessageBox.Show("Таблица service создана и заполнена!");
        }


        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            int x = 0;
            int y = 0;
            int cell_height = 0;

            int colCount = dataGridView1.ColumnCount;
            int rowCount = dataGridView1.RowCount - 1;

            Font font = new Font("Tahoma", 9, FontStyle.Bold, GraphicsUnit.Point);

            int widthC = 0;

            int current_col = 0;
            int current_row = 10;

            while (current_row < rowCount)
            {
                while (current_col < colCount)
                {
                    if (g.MeasureString(dataGridView1[current_col, current_row].Value.ToString(), font).Width < widthC)
                    {
                        widthC = (int)g.MeasureString(dataGridView1[current_col, current_row].Value.ToString(), font).Width;
                    }
                    current_col++;
                }
                current_col = 0;
                current_row++;
            }

            current_col = 0;
            current_row = 0;

            string value = "";

            int width = widthC;
            int height = dataGridView1[current_col, current_row].Size.Height;

            Rectangle cell_border;
            SolidBrush brush = new SolidBrush(Color.Black);


            while (current_col < colCount)
            {

                cell_height = dataGridView1[current_col, current_row].Size.Height;
                cell_border = new Rectangle(x, y, width, height);
                value = dataGridView1.Columns[current_col].HeaderText.ToString();
                g.DrawRectangle(new Pen(Color.Black), cell_border);
                g.DrawString(value, font, brush, x, y);
                x += widthC;
                current_col++;
            }
            while (current_row < rowCount)
            {
                while (current_col < colCount)
                {

                    cell_height = dataGridView1[current_col, current_row].Size.Height;
                    cell_border = new Rectangle(x, y, width, height);
                    value = dataGridView1[current_col, current_row].Value.ToString();
                    g.DrawRectangle(new Pen(Color.Black), cell_border);
                    g.DrawString(value, font, brush, x, y);
                    current_col++;
                    x += widthC;
                }
                current_col = 0;
                current_row++;
                x = 0;
                y += cell_height;
            }
        }

        private void button81_Click(object sender, EventArgs e)
        {
            

            ClsPrint _ClsPrint = new ClsPrint(dataGridView1, "Салон");
            _ClsPrint.PrintForm();
        }


        private void button8_Click(object sender, EventArgs e)
        {
            PrintDocument Document = new PrintDocument();
            Document.PrintPage += new PrintPageEventHandler(printDocument12_PrintPage);
            PrintPreviewDialog dlg = new PrintPreviewDialog();
            dlg.Document = Document;
            dlg.ShowDialog();

            
        }







        private void ExportToExcel()
        {
            // Creating a Excel object.
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "ExportedFromDatGrid";

                int cellRowIndex = 2;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column.
                for (int i = 0; i < dataGridView1.Rows.Count  ; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user.
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Excel!");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

      


        private void printDocument12_PrintPage(object sender, EventArgs e)
        {
            string connectionString = @"Data Source=DESKTOP-FITMBD4;Initial Catalog=saloon;" + "Integrated Security=true;";//Sql
            

            Cursor.Current = Cursors.WaitCursor;
            SqlConnection sqlConnection = null;
            SqlCommand sqlCommand = null;
            SqlDataReader sqlReader = null;

            try
            {
                string strQuery = "SELECT salon.ID,NCust, service.Nservice, master.FName,service.Price FROM salon inner join service on salon.ID_service = service.ID_service inner join master on salon.ID_master = master.ID_master order by ID ";
                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();
                sqlCommand = new SqlCommand(strQuery, SqlConnection);
                sqlReader = sqlCommand.ExecuteReader();
                while (sqlReader.Read())
                {
                    object[] row = { sqlReader[0], sqlReader[1], sqlReader[2], sqlReader[3],sqlReader[4] };
                    dataGridView1.Rows.Add(row);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                sqlConnection.Close();
                if (sqlReader != null)
                {
                    sqlReader.Dispose();
                    sqlReader = null;
                }
                if (sqlCommand != null)
                {
                    sqlCommand.Dispose();
                    sqlCommand = null;
                }
            }
        }


        class ClsPrint
        {
            #region Variables

            int iCellHeight = 0; //Used to get/set the datagridview cell height
            int iTotalWidth = 0; //
            int iRow = 0;//Used as counter
            bool bFirstPage = false; //Used to check whether we are printing first page
            bool bNewPage = false;// Used to check whether we are printing a new page
            int iHeaderHeight = 0; //Used for the header height
            StringFormat strFormat; //Used to format the grid rows.
            ArrayList arrColumnLefts = new ArrayList();//Used to save left coordinates of columns
            ArrayList arrColumnWidths = new ArrayList();//Used to save column widths
            private PrintDocument _printDocument = new PrintDocument();
            private DataGridView gw = new DataGridView();
            private string _ReportHeader;

            #endregion

            public ClsPrint(DataGridView gridview, string ReportHeader)
            {
                _printDocument.PrintPage += new PrintPageEventHandler(_printDocument_PrintPage);
                _printDocument.BeginPrint += new PrintEventHandler(_printDocument_BeginPrint);
                gw = gridview;
                _ReportHeader = ReportHeader;
            }

            public void PrintForm()
            {
                ////Open the print dialog
                //PrintDialog printDialog = new PrintDialog();
                //printDialog.Document = _printDocument;
                //printDialog.UseEXDialog = true;

                ////Get the document
                //if (DialogResult.OK == printDialog.ShowDialog())
                //{
                //    _printDocument.DocumentName = "Test Page Print";
                //    _printDocument.Print();
                //}

                //Open the print preview dialog
                PrintPreviewDialog objPPdialog = new PrintPreviewDialog();
                objPPdialog.Document = _printDocument;
                objPPdialog.ShowDialog();
            }

            private void _printDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
            {
                //try
                //{
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in gw.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                            (double)iTotalWidth * (double)iTotalWidth *
                            ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                            GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headers
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= gw.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = gw.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allows more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {

                        if (bNewPage)
                        {
                            //Draw Header
                            e.Graphics.DrawString(_ReportHeader,
                                new Font(gw.Font, FontStyle.Bold),
                                Brushes.Black, e.MarginBounds.Left,
                                e.MarginBounds.Top - e.Graphics.MeasureString(_ReportHeader,
                                new Font(gw.Font, FontStyle.Bold),
                                e.MarginBounds.Width).Height - 13);

                            String strDate = "";
                            //Draw Date
                            e.Graphics.DrawString(strDate,
                                new Font(gw.Font, FontStyle.Bold), Brushes.Black,
                                e.MarginBounds.Left +
                                (e.MarginBounds.Width - e.Graphics.MeasureString(strDate,
                                new Font(gw.Font, FontStyle.Bold),
                                e.MarginBounds.Width).Width),
                                e.MarginBounds.Top - e.Graphics.MeasureString(_ReportHeader,
                                new Font(new Font(gw.Font, FontStyle.Bold),
                                FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            DataGridViewColumn[] _GridCol = new DataGridViewColumn[gw.Columns.Count];
                            int colcount = gw.Columns.Count - 1;
                            //Convert ltr to rtl
                            foreach (DataGridViewColumn GridCol in gw.Columns)
                            {
                                _GridCol[colcount--] = GridCol;
                            }
                            for (int i = (_GridCol.Count() - 1); i >= 0; i--)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(_GridCol[i].HeaderText,
                                    _GridCol[i].InheritedStyle.Font,
                                    new SolidBrush(_GridCol[i].InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        DataGridViewCell[] _GridCell = new DataGridViewCell[GridRow.Cells.Count];
                        int cellcount = GridRow.Cells.Count - 1;
                        //Convert ltr to rtl
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            _GridCell[cellcount--] = Cel;
                        }
                        //Draw Columns Contents                
                        for (int i = (_GridCell.Count() - 1); i >= 0; i--)
                        {
                            if (_GridCell[i].Value != null)
                            {
                                e.Graphics.DrawString(_GridCell[i].FormattedValue.ToString(),
                                    _GridCell[i].InheritedStyle.Font,
                                    new SolidBrush(_GridCell[i].InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount],
                                    (float)iTopMargin,
                                    (int)arrColumnWidths[iCount], (float)iCellHeight),
                                    strFormat);
                            }
                            //Drawing Cells Borders 
                            e.Graphics.DrawRectangle(Pens.Black,
                                new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                (int)arrColumnWidths[iCount], iCellHeight));
                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }
                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
                //}
                //catch (Exception exc)
                //{
                //    MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK,
                //       MessageBoxIcon.Error);
                //}
            }

            private void _printDocument_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
            {
                try
                {
                    strFormat = new StringFormat();
                    strFormat.Alignment = StringAlignment.Center;
                    strFormat.LineAlignment = StringAlignment.Center;
                    strFormat.Trimming = StringTrimming.EllipsisCharacter;

                    arrColumnLefts.Clear();
                    arrColumnWidths.Clear();
                    iCellHeight = 0;
                    iRow = 0;
                    bFirstPage = true;
                    bNewPage = true;

                    // Calculating Total Widths
                    iTotalWidth = 0;
                    foreach (DataGridViewColumn dgvGridCol in gw.Columns)
                    {
                        iTotalWidth += dgvGridCol.Width;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

    }
}



