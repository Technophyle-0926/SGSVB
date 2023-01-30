using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.IO;
using System.Configuration;
using System.Drawing.Printing;
using System.Reflection;

namespace Payment_Application
{
    public partial class Form1 : Form
    {
        int i = 1;
        int ReceiptNo = 0;
        string MembershipNo = string.Empty;
        string FullName = string.Empty;
        string Address = string.Empty;
        decimal Amount;
        string Counter = string.Empty;
        string Last_Name = string.Empty;
        string First_Name = string.Empty;
        string Middle_Name = string.Empty;
        public Form1()
        {
            InitializeComponent();

        }

        private void Submit_Click(object sender, EventArgs e)
        {
            MembershipNo = Membershiptxt.Text;
            FullName = FullNametxt.Text;
            Address = Addresstxt.Text;
            Amount = Convert.ToDecimal(Amounttxt.Text);
            Counter = Countertxt.Text;
            /**************************************Insert the data into the databasae*************************************************************/
            string query = "Insert into ams(Membership_No,FullName,Address,Amount,Counter) values (@MembershipNo,@FullName,@Address,@Amount,@Counter)";
            string conection = ConfigurationManager.ConnectionStrings["sgsvbdataConnectionString"].ToString();
            MySqlConnection conn = new MySqlConnection(conection);
            MySqlCommand command = conn.CreateCommand();
            MySqlCommand cmd = new MySqlCommand(query, conn);
            conn.Open();
            cmd.Parameters.AddWithValue("@MembershipNo", MembershipNo);
            cmd.Parameters.AddWithValue("@FullName", FullName);
            cmd.Parameters.AddWithValue("@Address", Address);
            cmd.Parameters.AddWithValue("@Amount", Amount);
            cmd.Parameters.AddWithValue("@Counter", Counter);


            int receipt_no;


            if ((MembershipNo == "") || (FullName == "") || (Address == "") || (Amount == null) || (Counter == ""))
            {
                cmd.ExecuteNonQuery();
                receipt_no = (int)cmd.LastInsertedId;
                //MessageBox.Show("Data Inserted Successfully");
            }
            else
            {
                cmd.ExecuteNonQuery();
                receipt_no = (int)cmd.LastInsertedId;
            }
            MessageBox.Show("Data Inserted Successfully");
            conn.Close();
            /*************************************************Take the receipt number***************************************************************/
           //// //string totaldata = Last_Name + " " + First_Name + " " + Middle_Name;
           //// //string query1 = "Select count(Receipt_No) from bhet";
           //// //string query1 = "SELECT Receipt_No FROM bhet where Membership_No='" + MembershipNo + "'and FullName='" + FullName + "'";
           //// //string query1 = "select * from bhet";
           //// string conection1 = ConfigurationManager.ConnectionStrings["sgsvbdataConnectionString"].ToString();
           //// MySqlConnection conn1 = new MySqlConnection(conection1);
           //// conn1.Open();
           //// MySqlCommand command1 = conn1.CreateCommand();
           //// command1.CommandText = "SELECT Receipt_No FROM bhet where Membership_No='" + MembershipNo + "'and FullName='" + FullName + "'";
           ////// MySqlCommand cmd1 = new MySqlCommand(query1, conn1);
           
           //// //ReceiptNo = Convert.ToInt32(cmd1.ExecuteScalar());
           //// //string data = MembershipNo;
           //// //string full = FullName;
           //// //ReceiptNo = Convert.ToInt32(cmd1.ExecuteScalar());
           //// MySqlDataReader dr = command1.ExecuteReader();

           //// while (dr.Read())
           //// {
           ////     string val = dr[0].ToString();
           ////   // ReceiptNo = this.GetDBString("Receipt_No", dr);               

           //// }
            ExcelHelper ex = new ExcelHelper();
            ex.GenerateExcel(i, FullName, Address, Amount, receipt_no);
            i++;
           // conn1.Close();

            Membershiptxt.Clear();
            FullNametxt.Clear();
            Addresstxt.Clear();
            //Amounttxt.Clear();
            Countertxt.Clear();

        }

        private string GetDBString(string SqlFieldName, MySqlDataReader Reader)
        {
            return Reader[SqlFieldName].Equals(DBNull.Value) ? String.Empty : Reader.GetString(SqlFieldName);
        }



        public class ExcelHelper
        {
            public string GenerateExcel(int id, string FullName, string Address, decimal Amount, int ReceiptNumber)
            {
                string targetExcelPath = string.Empty;
                try
                {
                    // string connectionString = ConfigurationManager.ConnectionStrings["BillConnectionString"].ConnectionString;
                    // BillRecord billRecord = new BillService(connectionString).GetBillDetail(id);
                    string folder = ConfigurationManager.AppSettings["BillStorePath"];
                    string path = string.Empty;
                    path = "Bills_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year;
                    System.IO.DirectoryInfo dir = new DirectoryInfo(folder);
                    if (dir.Exists)
                    {
                        //Check is Bill Folder for today's date exists
                        System.IO.DirectoryInfo billFolder = new DirectoryInfo(folder + "\\" + path);
                        if (!billFolder.Exists)
                            billFolder.Create();
                    }

                    //Check is Excel Exists?
                    string templateFilePath = ConfigurationManager.AppSettings["TemplatePath"]; ;
                    FileInfo file = new FileInfo(templateFilePath);
                    if (file.Exists)
                    {
                        //Create Blank Excel File with BillId and time
                        Excel excel = new Excel();
                        string targetExcelName = string.Empty;

                        if (!string.IsNullOrEmpty(FullName))
                            targetExcelName = FullName + "_" + DateTime.Now.ToShortTimeString();
                        else
                            targetExcelName = "Bill" + FullName + "_" + DateTime.Now;
                        targetExcelName = targetExcelName.Replace(":", "-");
                        targetExcelPath = folder + "\\" + path + "\\" + targetExcelName + ".xls";
                        File.Copy(templateFilePath, targetExcelPath);

                        excel.UpdateExcel(targetExcelPath, id, FullName, Address, Amount, ReceiptNumber);
                    }
                }
                catch (Exception ex)
                {

                }
                return targetExcelPath;
            }
        }


        public class Excel
        {

            private Microsoft.Office.Interop.Excel.Application _excelApplication = null;
            private Microsoft.Office.Interop.Excel.Workbooks _workBooks = null;
            private Microsoft.Office.Interop.Excel._Workbook _workBook = null;
            private object _value = Missing.Value;
            private Microsoft.Office.Interop.Excel.Sheets _excelSheets = null;
            private Microsoft.Office.Interop.Excel._Worksheet _excelSheet = null;


            public void ActivateExcel()
            {
                _excelApplication = new Microsoft.Office.Interop.Excel.Application();
                _workBooks = (Microsoft.Office.Interop.Excel.Workbooks)_excelApplication.Workbooks;
                _workBook = (Microsoft.Office.Interop.Excel._Workbook)(_workBooks.Add(_value));
                _excelSheets = (Microsoft.Office.Interop.Excel.Sheets)_workBook.Worksheets;
                _excelSheet = (Microsoft.Office.Interop.Excel._Worksheet)(_excelSheets.get_Item(1));


            }

            public void SaveExcel(string fileName)
            {
                _workBook.SaveAs(fileName, _value, _value,
                    _value, _value, _value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    _value, _value, _value, _value, null);
                _workBook.Close(false, _value, _value);
                _excelApplication.Quit();
            }

            public string CreateExcelFile(string filePath)
            {
                this.ActivateExcel();
                this.SaveExcel(filePath);
                return filePath;

            }

            public void UpdateExcel(string filePath, int billId, string FullName, string Address, decimal Amount, int ReceiptNumber)
            {

                //Open Excel
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook Workbook = app.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                    , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Excel.Sheets wrksheets = Workbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)wrksheets.get_Item(1);

                //Enter Values in worksheet
                worksheet.Cells[11, 3] = ReceiptNumber;
                worksheet.Cells[13, 4] = FullName;
                worksheet.Cells[15, 2] = Address;
                worksheet.Cells[17, 2] = Amount;
                //string address1, address2;
                //address1 = address2 = string.Empty;
                //this.DivideAddressInTwoLine(billRecord.BillClient.ClientAddress, ref address1, ref address2);
                //worksheet.Cells[11, 2] = address1;
                //worksheet.Cells[12, 2] = address2;

                 //WordHelper wordHelper = new WordHelper();
                //worksheet.Cells[39, 1] = "RS.";
               // worksheet.Cells[19, 2] = wordHelper.IntergerToWord(Amount);
                Workbook.Save();
                Workbook.Close(false, Type.Missing, Type.Missing);
                app.Quit();

            }

        }

        private void Search_Click(object sender, EventArgs e)
        {
            string MembershipNo = Membershiptxt.Text;
            string query1 = "Select * from address_book where Membership_No='" + MembershipNo + "'";
            string conection = ConfigurationManager.ConnectionStrings["sgsvbdataConnectionString"].ToString();
            MySqlConnection conn = new MySqlConnection(conection);
            MySqlCommand command = conn.CreateCommand();
            MySqlCommand cmd = new MySqlCommand(query1, conn);

            conn.Open();
            MySqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {

                Last_Name = dr[1].ToString();
                First_Name = dr[2].ToString();
                Middle_Name = dr[3].ToString();
                string FullName = Last_Name + " " + First_Name + " " + Middle_Name;
                FullNametxt.Text = FullName;
            }

            conn.Close();

        }









        //public class WordHelper
        //{
        //    string[] ones = {
        //                    "",
        //                    "one",
        //                    "two",
        //                    "three",
        //                    "four",
        //                    "five",
        //                    "six",
        //                    "seven",
        //                    "eight",
        //                    "nine",
        //                    "ten",
        //                    "eleven",
        //                    "twelve",
        //                    "thirteen",
        //                    "fourteen",
        //                    "fifteen",
        //                    "sixteen",
        //                    "seventeen",
        //                    "eighteen",
        //                    "nineteen"
        //                };
        //    string[] tens =
        //    {
        //         "",
        //        "ten",
        //        "twenty",
        //        "thirty",
        //        "forty",
        //        "fifty",
        //        "sixty",
        //        "seventy",
        //        "eighty",
        //        "ninety"
        //    };

        //    string[] Location =
        //    {
        //        "",
        //        "hundred",
        //        "thousand",
        //        "lakh",
        //        "crore"
              
        //    };



        //    public string IntergerToWord(decimal inputnum)
        //    {
        //        string returnValue = string.Empty;
        //        string value = inputnum.ToString();
        //        string numInString = value.PadLeft(8, '0');
        //        //Add underscore to make pattern as Crore_Lakh_Thousand_hundred_digit(0_00_00_0_00)
        //        numInString = numInString.Insert(1, "_");
        //        numInString = numInString.Insert(4, "_");
        //        numInString = numInString.Insert(7, "_");
        //        numInString = numInString.Insert(9, "_");

        //        string[] numbers = numInString.Split('_');
        //        List<string> arrayInNumbers=new List<string>();
        //        for (int j = numbers.Length - 1; j >= 0; j--)
        //        {
        //            arrayInNumbers.Add(numbers[j]);
        //        }
        //      //  List<string> arrayInNumbers = numbers.Reverse.ToString();
        //        for (int i = 0; i < arrayInNumbers.Count; i++)
        //        {
        //            int num = int.Parse(arrayInNumbers[i]);
        //            int digit1, digit2;
        //            digit1 = int.Parse(arrayInNumbers[i].Substring(0, 1));
        //            if (i == 1 || i == 4)
        //            {
        //                if (digit1 > 0)
        //                    returnValue = ones[digit1] + " " + Location[i] + " " + returnValue;
        //            }
        //            else
        //            {
        //                digit2 = int.Parse(arrayInNumbers[i].Substring(1, 1));
        //                if (num > 0)
        //                {
        //                    if (num < 20) // if less than 20, use "ones" only
        //                        returnValue = ones[num] + " " + Location[i] + " " + returnValue;
        //                    else // otherwise, use both "tens" and "ones" array
        //                    {
        //                        string digit1InString = tens[digit1];
        //                        string digit2InString = (digit2 > 0) ? ones[digit2] + " " : string.Empty;
        //                        returnValue = tens[digit1] + " " + digit2InString + Location[i] + " " + returnValue;
        //                    }
        //                }
        //            }
        //        }
        //        if (!string.IsNullOrEmpty(returnValue))
        //        {
        //            returnValue = char.ToUpper(returnValue[0]) + returnValue.Substring(1);
        //            returnValue += "only";

        //        }
        //        return returnValue;

        //    }

        //}

        public class SplitAmount
        {
            public string RsString;
            public string PsString;

        }
    }
}

 
