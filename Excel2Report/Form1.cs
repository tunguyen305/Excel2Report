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

            lstOrders = new List<Order>();
        }
        const string DEST = "hello.pdf";
        string DEST_EXCEL = "excel.xls";
        private void button1_Click(object sender, EventArgs e)
        {
            //ExportPdf();
            ExportExcel();
        }

        private void ExportExcel()
        {
            string fullFileName = System.IO.Path.Combine(Application.StartupPath, DEST_EXCEL);

            ExcelUtil util = new ExcelUtil("");
            util.WriteOrders(lstOrders, chkExportAll.Checked? -1: listBox1.SelectedIndex);
        }

        private  void ExportPdf()
        {
            string fullFileName = System.IO.Path.Combine(Application.StartupPath, DEST);
            PdfDocument pdf = new PdfDocument(new PdfWriter(fullFileName));
            Document document = new Document(pdf);
            string line = "BIÊN NHẬN ĐÓNG GÓI";
            Paragraph pTitle = new Paragraph(line);
            Style pStyle = new Style();
            pStyle.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
            pStyle.SetFontSize(16);
            pTitle.AddStyle(pStyle);

            document.Add(pTitle);
            document.Add(new Paragraph());
            document.Add(new Paragraph());

            string txtTime = DateTime.Now.ToLongDateString();
            Paragraph pTime = new Paragraph(txtTime);
            Style sTime = new Style();
            sTime.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT);
            //sTime.SetWidth(new UnitValue(UnitValue.PERCENT, 30));
            sTime.SetTextAlignment(TextAlignment.JUSTIFIED);
            sTime.SetVerticalAlignment(VerticalAlignment.TOP);
            sTime.SetFontSize(10);
            pTime.AddStyle(sTime);



            line = "SHOP Minh Tuyết Hotline Sđt/Zalo: 0915.332.489 - 0365.189.935";
            Paragraph pAddress = new Paragraph(line);
            Style sAddress = new Style();
            sAddress.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
            sAddress.SetWidth(new UnitValue(UnitValue.PERCENT, 30));
            sAddress.SetTextAlignment(TextAlignment.JUSTIFIED);
            sAddress.SetFontSize(10);
            pAddress.AddStyle(sAddress);
            pAddress.SetBold();
            //document.Add(pAddress);
            pTime.Add(pAddress);

            string txtID = "20221003093822760";
            Paragraph pID = new Paragraph(txtID);
            Style sID = new Style();
            sID.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.RIGHT);
            sID.SetWidth(new UnitValue(UnitValue.PERCENT, 30));
            sID.SetTextAlignment(TextAlignment.JUSTIFIED);
            sID.SetVerticalAlignment(VerticalAlignment.TOP);
            sID.SetFontSize(10);
            pID.AddStyle(sID);
            //document.Add(pID);
            pTime.Add(pID);
            document.Add(pTime);
            document.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDiglog = new OpenFileDialog();
            openDiglog.InitialDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "Downloads");
            openDiglog.Filter = "Excel Files|*.xls";
            if (openDiglog.ShowDialog() == DialogResult.OK)
            {
                progressBar1.Value = 0;
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
            textBox1.Text = Environment.SpecialFolder.MyDocuments.ToString();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)e.Argument;
            ExcelUtil util = new ExcelUtil(fileName);
            lstOrders = util.ReadAllOrders();
            e.Result = lstOrders.Count;

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            listBox1.DataSource = lstOrders;
            listBox1.DisplayMember = "OrderId";

            FillOrder(lstOrders[0]);

            progressBar1.Value = progressBar1.Maximum;
            button1.Enabled = true;
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
    }
}
