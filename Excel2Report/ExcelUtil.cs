using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using QRCoder;
using System.Drawing;

namespace Report2Pdf
{
    

    public class ExcelUtil
    {
        FileInfo finfo;
        

        string fileName = @"F:\Work\Report2Pdf\Report2Pdf\shop.xls";
        
        string fileTemplateName = @"C:\Git\Projects\Excel2Report\Excel2Report\rptTpl.xlt";

        public ExcelUtil(string fileName)
        {
            this.fileName = fileName;
        }

        public void Open()
        {
            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            ExcelObj.Visible = false;

            Workbook theWorkbook;
            Worksheet worksheet;
            finfo = new FileInfo(fileName);
            if (finfo.Extension == ".xls" || finfo.Extension == ".xlsx" || finfo.Extension == ".xlt" || finfo.Extension == ".xlsm" || finfo.Extension == ".csv")
            {
                theWorkbook = ExcelObj.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, false, false);
                string s = "";
                for (int count = 1; count <= theWorkbook.Sheets.Count; count++)
                {
                    worksheet = (Worksheet)theWorkbook.Worksheets.get_Item(count);
                    worksheet.Activate();
                   // worksheet.Visible = XlSheetVisibility.xlSheetHidden;
                    worksheet.UsedRange.Cells.Select();



                    s += worksheet.Range["A2", "A2"].Text;
                }
                MessageBox.Show(s);


            }
        }

        #region Read
        public List<Order> ReadAllOrders()
        {
            List<Order> lstOrder = new List<Order>();

            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            ExcelObj.Visible = false;

            Workbook theWorkbook;
            Worksheet worksheet;
            finfo = new FileInfo(fileName);
            if (finfo.Extension == ".xls" || finfo.Extension == ".xlsx")
            {
                theWorkbook = ExcelObj.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, false, false);
                string s = "";
                if (theWorkbook.Sheets.Count >= 1)
                {
                    worksheet = (Worksheet)theWorkbook.Worksheets.get_Item(1);
                    int rowcount = worksheet.Rows.Count;
                    worksheet.Activate();
                    int r = 2;
                    while (r < rowcount)
                    {
                        Order o = ReadOrder(worksheet, r++);
                        if (o == null) break;

                        MergeOrder(lstOrder, o);
                    }

                    Marshal.ReleaseComObject(worksheet);
                }
                
                theWorkbook.Close();
                ExcelObj.Quit();
                releaseObject(theWorkbook);
            }
            return lstOrder;
        }

        private void MergeOrder(List<Order> current, Order newOne)
        {
            if (current.Count > 0)
            {
                Order order = current.Find(o => { return o.OrderId == newOne.OrderId; });
                if (order != null)
                {
                    order.Products.AddRange(newOne.Products);
                    return;
                }
            }
            current.Add(newOne);
        }

        public Order ReadOrder(Worksheet sheet, int lineNumber)
        {
            string range = "A" + lineNumber.ToString();
            string id = sheet.Range[range, range].Text;
            if (id != "")
            {
                Order o = new Order(id);
                range = "B" + lineNumber.ToString();
                o.UserName = sheet.Range[range, range].Text;

                range = "C" + lineNumber.ToString();
                o.Phone = sheet.Range[range, range].Text;

                range = "F" + lineNumber.ToString();
                o.AddrDetail = sheet.Range[range, range].Text;

                range = "G" + lineNumber.ToString();
                o.Province = sheet.Range[range, range].Text;

                range = "H" + lineNumber.ToString();
                o.District = sheet.Range[range, range].Text;

                range = "I" + lineNumber.ToString();
                o.Wards = sheet.Range[range, range].Text;

                range = "J" + lineNumber.ToString();
                o.Country = sheet.Range[range, range].Text;

                range = "L" + lineNumber.ToString();
                o.Notes = sheet.Range[range, range].Text;

                //product
                Product p = new Product();
                range = "M" + lineNumber.ToString();
                p.SKU = sheet.Range[range, range].Text;

                range = "N" + lineNumber.ToString();
                p.Name = sheet.Range[range, range].Text;

                range = "O" + lineNumber.ToString();
                p.Polyphym = sheet.Range[range, range].Text;

                range = "Q" + lineNumber.ToString();
                p.Price = double.Parse(sheet.Range[range, range].Text);

                range = "T" + lineNumber.ToString();
                p.Count = int.Parse(sheet.Range[range, range].Text);

                o.Products.Add(p);
                return o;
            }
            return null;
        }

        #endregion

        #region Write

        /// <summary>
        /// Write the list into report excel file
        /// </summary>
        /// <param name="lst"></param>
        /// <param name="index">-1/skip if want to export all</param>
        public void WriteOrders(List<Order> lst, int index = -1)
        {
            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            ExcelObj.Visible = false;
            int MAX_LINE_COUNT = 48;
            Workbook theWorkbook;
            Worksheet worksheet;
            finfo = new FileInfo(fileTemplateName);
            if (finfo.Extension == ".xls" || finfo.Extension == ".xlsx" || finfo.Extension == ".xlt" || finfo.Extension == ".xltx" || finfo.Extension == ".xlsm" || finfo.Extension == ".csv")
            {
                theWorkbook = ExcelObj.Workbooks.Open(fileTemplateName, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", true, false, 0, true, false, false);
                string s = "";
                for (int count = 1; count <= theWorkbook.Sheets.Count; count++)
                {
                    worksheet = (Worksheet)theWorkbook.Worksheets.get_Item(count);
                    worksheet.Activate();
                    // worksheet.Visible = XlSheetVisibility.xlSheetHidden;
                    worksheet.UsedRange.Cells.Select();
                    if (index != -1)
                    {
                        WriteAnOrder(worksheet, lst[Math.Min(index, lst.Count-1)]);
                    }
                    else
                    {
                        for (int i=0; i< 3 /* lst.Count*/; i++)
                        {
                            WriteAnOrder(worksheet, lst[i], i* MAX_LINE_COUNT );
                        }
                    }
                }
                theWorkbook.SaveAs(@"C:\Git\Projects\Excel2Report\out.xls", XlFileFormat.xlWorkbookNormal);
                theWorkbook.Close();

                ExcelObj.Quit();

                releaseObject(theWorkbook);
                MessageBox.Show("OK", "Exported", MessageBoxButtons.OK);
            }
            else
            {
                MessageBox.Show("wrong format file!");
            }

        }

        /// <summary>
        /// write an order to excel file
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="ord"></param>
        /// <param name="startRowIndex">row index that start to write</param>
        internal void WriteAnOrder(Worksheet sheet, Order ord, int startRowIndex = 0)
        {
            string line = (2 + startRowIndex).ToString();
            sheet.Range["D" + line].Value = "BIÊN NHẬN ĐÓNG GÓI";
            sheet.Range["D" + line].Font.Size = 18;


            line = (4 + startRowIndex).ToString();
            //time
            sheet.Range["A" + line, "C" + line].Merge();
            sheet.Range["A" + line].Font.Size = 11;
            sheet.Range["A" + line].Value = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss");

            //NAME            
            sheet.Range["D" + line, "D" + (5 + startRowIndex).ToString()].Merge();
            sheet.Range["D" + line].Font.Size = 14;
            sheet.Range["D" + line].Value = ord.UserName.ToUpper();

            //ID
            sheet.Range["F" + line, "G" + line].Merge();
            sheet.Range["F" + line].NumberFormat = "@";
            sheet.Range["F" + line].Font.Size = 11;
            sheet.Range["F" + line].Value = ord.OrderId;

            //name
            sheet.Range["A" + (7 + startRowIndex).ToString()].Value = "Tên người nhận:";
            sheet.Range["D" + (7 + startRowIndex).ToString()].Value = ord.UserName;
            //phone
            sheet.Range["A" + (8 + startRowIndex).ToString()].Value = "Số điện thoại:";
            sheet.Range["D" + (8 + startRowIndex).ToString()].NumberFormat = "@";
            sheet.Range["D" + (8 + startRowIndex).ToString()].Value = ord.Phone;

            //
            line = (9 + startRowIndex).ToString();
            sheet.Range["A" + line, "C" + line].Merge();
            sheet.Range["A" + line].WrapText = true;
            sheet.Range["A" + line].Value = "Phương thức giao hàng:";
            sheet.Range["D" + line].Value = "Giao hàng toàn quốc";
            sheet.Range["A" + line, "D" + line].RowHeight = 26;

            //store
            sheet.Range["A" + (10 + startRowIndex).ToString()].Value = "Cửa hàng:";
            sheet.Range["D" + (10 + startRowIndex).ToString()].Value = ord.AddrDetail;
            //full address
            sheet.Range["A" + (11 + startRowIndex).ToString()].Value = "Địa chỉ:";
            sheet.Range["D" + (11 + startRowIndex).ToString()].Value = ord.GetFullAddress();

            //products list
            int prdCount = ord.Products.Count;
            if (startRowIndex > 0)
            {
                var selectedRow = sheet.Range["A14:G18"];
                var destinationRow = sheet.Range[string.Format("A{0}:G{1}", 14 + startRowIndex, 18 + startRowIndex)];

                selectedRow.Copy(destinationRow);
            }
            //sheet.Copy(sheet.Range["A14:G18"], sheet.Range[string.Format("A{0}:G{1}", 14 + startRowIndex, 18 + startRowIndex)]);

            while (prdCount-- > 1)
            {
                sheet.Rows[16 + startRowIndex].Insert();
            }

            for (int r=0; r< ord.Products.Count; r++)
            {
                WriteAProduct(sheet, ord.Products[r], r + startRowIndex);
            }

            line = (startRowIndex + ord.Products.Count + 15).ToString();
            string lineEnd = (startRowIndex + ord.Products.Count + 14).ToString();
            //sum
            sheet.Range["B" + line].Value = "Tổng cộng";
            sheet.Range["E" + line].Value = string.Format("=sum(E{0}:E{1})", 15 + startRowIndex, lineEnd);
            sheet.Range["G" + line].Value = string.Format("=sum(G{0}:G{1})", 15 + startRowIndex, lineEnd);

            sheet.Range["G" + (ord.Products.Count + 17 + startRowIndex).ToString()].Value = string.Format("=G{0} + G{1}", ord.Products.Count + 15 + startRowIndex, ord.Products.Count + 16 + startRowIndex);

            sheet.PageSetup.LeftMargin = 1.2;
            sheet.PageSetup.RightMargin = 0.1;

            //insert QR code image
            //sheet.Shapes.AddShape(ShapeType)
            try
            {
                Bitmap pic = createQR(ord.OrderId);
                Clipboard.SetImage(pic);
                Range position = sheet.Range["F" +(5+ startRowIndex).ToString()];
                sheet.Paste(position); //copy the clipboard to the given position
            }
            catch (Exception)
            {
                MessageBox.Show("Tạo mã QR bị lỗi.");
            }
            

        }

        internal void WriteAProduct(Worksheet sheet, Product prod, int startIdx)
        {
            string line = (startIdx + 15).ToString();
            sheet.Range["A" + line].Value = startIdx + 1;
            sheet.Range["B" + line].Value = prod.SKU;
            string sp = prod.Name;
            if (prod.Polyphym != "")
            {
                sp += " / " + prod.Polyphym;
            }
            sheet.Range["D" + line].Value = sp;
            sheet.Range["E" + line].Value = prod.Count;
            sheet.Range["F" + line].Value = prod.Price;
            sheet.Range["G" + line].Value = string.Format("=E{0}*F{0}", line);

        }
        #endregion

        #region utils
        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private Bitmap createQR(string id)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(id, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(3);
            return qrCodeImage;
        }
        #endregion
    }
}
