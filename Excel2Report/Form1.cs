using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iText.Layout;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using iText.Layout.Properties;
using QRCoder;
using System.IO;

namespace Report2Pdf
{
    public partial class Form1 : Form
    {

        List<Order> lstOrders;
        public Form1()
        {
            InitializeComponent();           
        }
        const string DEST = "hello.pdf";
        string DEST_EXCEL = "excel.xls";
        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            button2.Enabled = false;
            string path = textBox1.Text ?? Application.StartupPath;
            backgroundWorker2.RunWorkerAsync(path + "?"+ (chkExportAll.Checked ? -1 : listBox1.SelectedIndex).ToString());
            while (backgroundWorker2.IsBusy)
            {
                progressBar1.Increment(1);
                // Keep UI messages moving, so the form remains
                // responsive during the asynchronous operation.

                // Wait 100 milliseconds.
                System.Threading.Thread.Sleep(50);


                Application.DoEvents();
            }
            
        }

        private void ExportExcel(string param)
        {
            string[] ps = param.Split('?');
            string path = ps[0];
            int index = int.Parse(ps[1]);

            ExcelUtil util = new ExcelUtil("", path);
            util.WriteOrders(lstOrders, index);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDiglog = new OpenFileDialog();
            openDiglog.InitialDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "Downloads");
            openDiglog.Filter = "Excel Files|*.xls";
            if (openDiglog.ShowDialog() == DialogResult.OK)
            {
                progressBar1.Value = 0;
                button2.Enabled = false;
                backgroundWorker1.RunWorkerAsync(openDiglog.FileName);
                while (backgroundWorker1.IsBusy)
                {
                    progressBar1.Increment(1);
                    // Keep UI messages moving, so the form remains
                    // responsive during the asynchronous operation.

                        // Wait 100 milliseconds.
                        System.Threading.Thread.Sleep(100);

                   
                    Application.DoEvents();
                }
            }
        }


        private void createQR(string id)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(id, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            pictureBox1.Image = qrCodeImage;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int idx = listBox1.SelectedIndex;
            FillOrder(lstOrders[idx]);
        }

        private void FillOrder(Order ord)
        {
            if (ord != null)
            {
                createQR(ord.OrderId);

                lblID.Text = ord.OrderId;
                lblShop.Text = lblName.Text = ord.UserName;
                lblPhone.Text = ord.Phone;
                lblAddress.Text = ord.GetFullAddress();

                FillProducts(ord.Products);
            }
        }

        private void FillProducts(List<Product> products)
        {
            // grvProducts.ColumnCount = 6;
            grvProducts.Rows.Clear();
            int rowcount = products.Count + 2;
            grvProducts.RowCount = rowcount;

            double sum = 0;

            for (int r=0; r< products.Count; r++)
            {
                Product p = products[r];
                grvProducts.Rows[r].Cells[0].Value =  (r+1).ToString();
                grvProducts.Rows[r].Cells[1].Value = p.SKU;
                grvProducts.Rows[r].Cells[2].Value = p.Name;
                grvProducts.Rows[r].Cells[3].Value = p.Count;
                grvProducts.Rows[r].Cells[4].Value = p.Price;
                grvProducts.Rows[r].Cells[5].Value = p.Count * p.Price;
                sum += p.Count * p.Price;
            }

            grvProducts.Rows[rowcount - 2].Height = 5;
            grvProducts.Rows[rowcount - 2].DefaultCellStyle.BackColor = Color.Black;
            grvProducts.Rows[rowcount - 1].Cells[2].Value = "TỔNG";
            grvProducts.Rows[rowcount - 1].Cells[5].Value = sum;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lstOrders = new List<Order>();
            textBox1.Text = Application.StartupPath;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)e.Argument;
            ExcelUtil util = new ExcelUtil(fileName);
            lstOrders = util.ReadAllOrders();
            e.Result = lstOrders.Count;

        }


        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            listBox1.DataSource = lstOrders;
            listBox1.DisplayMember = "OrderId";

            FillOrder(lstOrders[0]);

            progressBar1.Value = progressBar1.Maximum;
            button1.Enabled = true;
            button2.Enabled = true;
        }

        private void btnBrowser_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbDialog = new FolderBrowserDialog();
            fbDialog.Description = "Chọn nơi lưu report files.";
            fbDialog.RootFolder = Environment.SpecialFolder.MyDocuments;
            if (fbDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fbDialog.SelectedPath;
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            ExportExcel((string)e.Argument);
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = progressBar1.Maximum;
            button2.Enabled = true;
        }
    }
}
