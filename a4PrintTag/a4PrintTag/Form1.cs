
using Excel_using_OleDB.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace a4PrintTag
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public Image turnImage;
        private void button1_Click(object sender, EventArgs e)
        {
            PrintDialog MyPrintDg = new PrintDialog();

            this.printDocument1.DefaultPageSettings.Landscape = false;

            MyPrintDg.Document = printDocument1;
            if (MyPrintDg.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
        
        }

        Brush myBrush = new SolidBrush(Color.FromArgb(0, 82, 73));//0x005249
        Brush brush = new SolidBrush(Color.White);
        Font font = new Font("微软雅黑", 11, GraphicsUnit.Pixel);
        Font font2 = new Font("微软雅黑", 14.6f, FontStyle.Bold, GraphicsUnit.Pixel);
        Pen penGreen = new Pen(new SolidBrush(Color.FromArgb(0, 82, 73)));

        public void DrawGraphic(PrintInfo pi, System.Drawing.Printing.PrintPageEventArgs e,int x,int y)
        {              
            EncodeEan13(pi.ERPCode);
            string enStr = TurnString(pi.ERPCode);
            e.Graphics.FillRectangle(myBrush, 37 + x, 20 + y, MillimetersToPixel(86, fDpiX), MillimetersToPixel(10, fDpiY));
            e.Graphics.DrawRectangle(new Pen(Color.Black, 1), 37 + x + MillimetersToPixel(2, fDpiX), 20 + y + MillimetersToPixel(10, fDpiY), MillimetersToPixel(82, fDpiX), MillimetersToPixel(15, fDpiY));
           
            e.Graphics.DrawString("国家电网", font, brush, 85+x, 26+y);
            e.Graphics.DrawString("S", new Font("cambria", 13,FontStyle.Bold, GraphicsUnit.Pixel), brush, 85+x, 39+y);
            e.Graphics.DrawString("TATE", new Font("cambria", 11, FontStyle.Bold, GraphicsUnit.Pixel), brush, 93 + x, 41 + y);
            e.Graphics.DrawString("G", new Font("cambria", 13, FontStyle.Bold, GraphicsUnit.Pixel), brush, 124 + x, 39 + y);
            e.Graphics.DrawString("RID", new Font("cambria", 11, FontStyle.Bold, GraphicsUnit.Pixel), brush, 132 + x, 41 + y);
            e.Graphics.DrawString("天荒坪智慧资产", new Font("华文行楷", 21.5f, GraphicsUnit.Pixel), brush, 200+x, 32+y);
            e.Graphics.DrawImage(turnImage, new RectangleF(55+x, 28+y, 26, 26)); //logo

            e.Graphics.DrawLine(Pens.Black, 37 + x + MillimetersToPixel(2, fDpiX), 85+y, 220+x, 85+y);//中间横线
            e.Graphics.DrawLine(Pens.Black, 220+x, 58+y, 220+x, 115+y); //中间竖线
            e.Graphics.DrawLine(Pens.Black, -10, 118 + y, MillimetersToPixel(230, fDpiX), 118 + y);//下侧黑线            
            e.Graphics.DrawLine(penGreen, -10, 22+y, MillimetersToPixel(230, fDpiX), 22+y); //上侧绿线
     
            e.Graphics.DrawString("资产名称："+pi.assetName, new Font("微软雅黑", 12, FontStyle.Bold,GraphicsUnit.Pixel), new SolidBrush(Color.Black), 62+x, 63+y);
            e.Graphics.DrawString("启用日期："+pi.capitalizationDate, new Font("微软雅黑", 12,FontStyle.Bold, GraphicsUnit.Pixel), new SolidBrush(Color.Black), 62+x, 91+y);
            e.Graphics.DrawImage(imgBar,new Rectangle(227+x,67+y,120,25));//barcode

            e.Graphics.DrawString(enStr, font2, new SolidBrush(Color.Black), 225+x, 94+y);
        }


        public int count = 0;
        public int page = 0;
        public int pageCount = 0;
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int totalData = list.Count;
            page = totalData / 18  ;

            PrintInfo l = list[count];

            e.HasMorePages = true;

            /*e.Graphics.DrawLine(penGreen, 40 + MillimetersToPixel(84, fDpiX) , 0, 40 + MillimetersToPixel(84, fDpiX) , MillimetersToPixel(320, fDpiY));//右侧红线
            e.Graphics.DrawLine(penGreen, 34 + MillimetersToPixel(2, fDpiX) , 0, 34 + MillimetersToPixel(2, fDpiX) , MillimetersToPixel(320, fDpiY)); //左侧红线
            e.Graphics.DrawLine(penGreen, 40 + MillimetersToPixel(84, fDpiX) + 400, 0, 40 + MillimetersToPixel(84, fDpiX) + 400, MillimetersToPixel(320, fDpiY));//右侧红线
            e.Graphics.DrawLine(penGreen, 34 + MillimetersToPixel(2, fDpiX) + 400, 0, 34 + MillimetersToPixel(2, fDpiX) + 400, MillimetersToPixel(320, fDpiY)); //左侧红线*/


            e.Graphics.DrawLine(penGreen, 39 + MillimetersToPixel(84, fDpiX), 0, 39 + MillimetersToPixel(84, fDpiX), MillimetersToPixel(320, fDpiY));//右侧红线
            e.Graphics.DrawLine(penGreen, 35 + MillimetersToPixel(2, fDpiX), 0, 35 + MillimetersToPixel(2, fDpiX), MillimetersToPixel(320, fDpiY)); //左侧红线
            e.Graphics.DrawLine(penGreen, 39 + MillimetersToPixel(84, fDpiX) + 400, 0, 39 + MillimetersToPixel(84, fDpiX) + 400, MillimetersToPixel(320, fDpiY));//右侧红线
            e.Graphics.DrawLine(penGreen, 35 + MillimetersToPixel(2, fDpiX) + 400, 0, 35 + MillimetersToPixel(2, fDpiX) + 400, MillimetersToPixel(320, fDpiY)); //左侧红线

            for (int i = 0 + pageCount * 18; i < (pageCount == page ? (totalData ) : (pageCount + 1) * 18); i++) 
            {
                
                if (i == totalData)
                    break;

                if (i % 2 == 0)
                    DrawGraphic(list[i], e, 0, 120 * ((i % 18) / 2));
                else
                    DrawGraphic(list[i], e, 400, 120 * ((i % 18) / 2));
               
            }
            ++pageCount;

            e.Graphics.Dispose();
            if(pageCount>page)
                e.HasMorePages = false;           
        }



        public static byte[] ImageToBytes(Image image)
        {
            ImageFormat format = image.RawFormat;
            using (MemoryStream ms = new MemoryStream())
            {
                if (format.Equals(ImageFormat.Jpeg))
                {
                    image.Save(ms, ImageFormat.Jpeg);
                }
                else if (format.Equals(ImageFormat.Png))
                {
                    image.Save(ms, ImageFormat.Png);
                }
                else if (format.Equals(ImageFormat.Bmp))
                {
                    image.Save(ms, ImageFormat.Bmp);
                }
                else if (format.Equals(ImageFormat.Gif))
                {
                    image.Save(ms, ImageFormat.Gif);
                }
                else if (format.Equals(ImageFormat.Icon))
                {
                    image.Save(ms, ImageFormat.Icon);
                }
                else if (format.Equals(ImageFormat.Emf))
                {
                    image.Save(ms, ImageFormat.Emf);
                }
                byte[] buffer = new byte[ms.Length];
                //Image.Save()会改变MemoryStream的Position，需要重新Seek到Begin
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(buffer, 0, buffer.Length);
                return buffer;
            }
        }

        /// <summary>
        /// Convert Byte[] to Image
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public static Image BytesToImage(byte[] buffer)
        {
            MemoryStream ms = new MemoryStream(buffer);
            Image image = System.Drawing.Image.FromStream(ms);
            return image;
        }

        /// <summary>
        /// Convert Byte[] to a picture and Store it in file
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public static string CreateImageFromBytes(string fileName, byte[] buffer)
        {
            string file = fileName;
            Image image = BytesToImage(buffer);
            ImageFormat format = image.RawFormat;
            if (format.Equals(ImageFormat.Jpeg))
            {
                file += ".jpeg";
            }
            else if (format.Equals(ImageFormat.Png))
            {
                file += ".png";
            }
            else if (format.Equals(ImageFormat.Bmp))
            {
                file += ".bmp";
            }
            else if (format.Equals(ImageFormat.Gif))
            {
                file += ".gif";
            }
            else if (format.Equals(ImageFormat.Icon))
            {
                file += ".icon";
            }
            System.IO.FileInfo info = new System.IO.FileInfo(file);
            System.IO.Directory.CreateDirectory(info.Directory.FullName);
            File.WriteAllBytes(file, buffer);
            return file;
        }


        public Image imgLogo;
        public float fDpiX, fDpiY;
        private void Form1_Load(object sender, EventArgs e)
        {
           /* path = @"E:\ProjectDoc\资产\报废标签打印.xls";          
            connectionString = collect_Con(System.IO.Path.GetExtension(path), path);*/
            fDpiX = CreateGraphics().DpiX;
            fDpiY = CreateGraphics().DpiY;
            //button2.Enabled = false;
            //button1.Enabled = false;

           
        }


        private float MillimetersToPixel(float fValue, float fDPI)
        {
            return (fValue / 25.4f) * fDPI;
        }

        BarcodeLib.Barcode b = new BarcodeLib.Barcode();
        public Image imgBar;
        public void EncodeEan13(string erpCode)
        {
            
            BarcodeLib.TYPE type = BarcodeLib.TYPE.EAN13;
            
            try
            {
                if (type != BarcodeLib.TYPE.UNSPECIFIED)
                {
                    b.IncludeLabel = false;
                    imgBar = b.Encode(type, erpCode, Color.Black, Color.White, 200, 32);
                    

                }//if
            }//try
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }//catch
        }


        public string TurnString(string str)
        {
            if (str.Length == 12)
            {
                string s1 = str.Substring(0, 4);
                string s2 = str.Substring(4, 4);
                string s3 = str.Substring(8, 4);
                str = s1 + " " + s2 + " " + s3;
            }
            return str;
        }

        string str;
        public void Read()
        {
            StreamReader sr = new StreamReader(System.AppDomain.CurrentDomain.BaseDirectory + "barcode.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                str = str + " " + line.ToString();
            }

        }


        List<PrintInfo> list = new List<PrintInfo>();
        string path;

        /// <summary>
        /// 读取Excel 和 logo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            #region MyRegion
            /*//DataSet objDS = new DataSet();

            Image imgLogo = Image.FromFile(System.AppDomain.CurrentDomain.BaseDirectory + "logo4.png");
            
            byte[] imgByte = ImageToBytes(imgLogo);
            turnImage = BytesToImage(imgByte);

            string szConStr = collect_Con(System.IO.Path.GetExtension(path), path);
            try
            {
                objCon = new OleDbConnection(szConStr);
                objCon.Open();
                objDA = new System.Data.OleDb.OleDbDataAdapter("select [ERP资产编号] as [ERPCode],[描述一（资产描述）] as [assetName],[资本化日期]as [capitalizationDate],[资产座落地点] as [location] from [sheet1$] where [ERP资产编号] is not null ", objCon);
                DataTable objDTExcel = new DataTable();
                objDA.Fill(objDTExcel);
                //dataGridView1.DataSource = objDTExcel;

                foreach (DataRow row in objDTExcel.Rows)
                {
                    list.Add(new PrintInfo()
                    {                        
                        ERPCode = row["ERPCode"].ToString(),
                        assetName = row["assetName"].ToString(),                        
                        capitalizationDate = row["capitalizationDate"].ToString(),                                                
                        //depreciationlife = row["depreciationlife"].ToString(),
                        
                    });
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (objDA != null)
                {
                    objDA.Dispose();
                    objDA = null;
                }

                if (objCon != null)
                {
                    objCon.Close();
                    objCon.Dispose();
                    objCon = null;
                }
            }

            button2.Enabled = false;*/
            
            #endregion

            Image imgLogo = Image.FromFile(System.AppDomain.CurrentDomain.BaseDirectory + "logo4.png");

            byte[] imgByte = ImageToBytes(imgLogo);
            turnImage = BytesToImage(imgByte);
            Read();
            list = JsonTransform.JsonToAnything<List<PrintInfo>>(str);
            label2.Text = list.Count.ToString();
        }




       
    }
}
