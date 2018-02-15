using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Project1
{
    class ExcelToXml
    {

        static private void BuildXml(Dictionary<string, string> map, string outputFile, int row)
        {
            string xmlData =
              "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
              "<DOC>" +
                "<HEADER> " +
                 "<MESSAGETYPEID>{0}</MESSAGETYPEID> " +
                  "<DESTINATIONNAME>{1}</DESTINATIONNAME>" +
                  "<MESSAGEQUEUEID>{2}</MESSAGEQUEUEID>" +
                  "<MESSAGELOGTIME>{3}</MESSAGELOGTIME>" +
                "</HEADER>" +
                "<ROW>" +
                  "<FLAG>{4}</FLAG>" +
                  "<ROWID>{5}</ROWID>" +
                  "<COLUMN>" +
                    "<EMP_ID>{6}</EMP_ID>" +
                    "<LASTNAME>{7}</LASTNAME>" +
                    "<FIRSTNAME>{8}</FIRSTNAME>" +
                    "<MI>{9}</MI>" +
                    "<HIERARCHY>{10}</HIERARCHY>" +
                    "<SSN />" +
                    "<ADDRESS1>{11}</ADDRESS1>" +
                    "<ADDRESS2 />" +
                    "<CITY>{12}</CITY>" +
                    "<STATE>{13}</STATE>" +
                    "<ZIPCODE>{14}</ZIPCODE>" +
                    "<EM_CONTACT>{15}</EM_CONTACT>" +
                    "<EM_RELATN>{16}</EM_RELATN>" +
                    "<EM_PHONE>{17}</EM_PHONE>" +
                    "<EM_ADDRESS>{18}</EM_ADDRESS>" +
                    "<EM_CITY>{19}</EM_CITY>" +
                    "<EM_STATE>{20}</EM_STATE>" +
                    "<EM_ZIP>{21}</EM_ZIP>" +
                    "<GENDER>{22}</GENDER>" +
                    "<BIRTH_DATE>{23}</BIRTH_DATE>" +
                    "<PSLSTADATE>{24}</PSLSTADATE>" +
                    "<PSLENDDATE />" +
                    "<RANK>{25}</RANK>" +
                    "<VENDOR />" +
                    "<PIN />" +
                    "<OSN />" +
                    "<USERNAME>{26}</USERNAME>" +
                    "<PASSWORD>{27}</PASSWORD>" +
                  "</COLUMN>" +
                "</ROW>" +
              "</DOC>";

            string xmlString = string.Empty;
            string messageTypeId = map.ContainsKey("MESSAGETYPEID") ? map["MESSAGETYPEID"] : string.Empty;
            string destinationName = map.ContainsKey("DESTINATIONNAME") ? map["DESTINATIONNAME"] : string.Empty;
            string messsageQueueId = map.ContainsKey("MESSAGEQUEUEID") ? map["MESSAGEQUEUEID"] : string.Empty;
            string messageLogTime = map.ContainsKey("MESSAGELOGTIME") ? map["MESSAGELOGTIME"] : string.Empty;
            string flag = map.ContainsKey("FLAG") ? map["FLAG"] : string.Empty;
            string rowId = map.ContainsKey("ROWID") ? map["ROWID"] : string.Empty;
            string empId = map.ContainsKey("EMP_ID") ? map["EMP_ID"] : string.Empty;
            string lastName = map.ContainsKey("LASTNAME") ? map["LASTNAME"] : string.Empty;
            string firstName = map.ContainsKey("FIRSTNAME") ? map["FIRSTNAME"] : string.Empty;
            string middleInitial = map.ContainsKey("MI") ? map["MI"] : string.Empty;
            string hierarchy = map.ContainsKey("HIERARCHY") ? map["HIERARCHY"] : string.Empty;
            string rank = map.ContainsKey("RANK") ? map["RANK"] : string.Empty;
            string userName = map.ContainsKey("USERNAME") ? map["USERNAME"] : string.Empty;
            string password = map.ContainsKey("PASSWORD") ? map["PASSWORD"] : string.Empty;

            string data = string.Format(xmlData,
                messageTypeId,
                destinationName,
                messsageQueueId,
                messageLogTime,
                flag,
                rowId,
                empId,
                lastName,
                firstName,
                middleInitial,
                hierarchy,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                userName,
                string.Empty);

            //String outPutFileName = outputFile + lastName + row + ".xml";
            String outPutFileName = "Personnel" + row + ".xml";
            Console.WriteLine("writing file "+ outPutFileName +" for user: " + lastName +", "+firstName);
            var stream = System.IO.File.Create( outPutFileName ); ;
            stream.Write(ASCIIEncoding.ASCII.GetBytes(data.ToCharArray()), 0, data.Length);
            stream.Close();
        }

        public static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;
            try
            {
                if (args.Length != 2)
                {
                    Console.WriteLine("usage ExcelToXml <path to xslx> <output path>");
                    return;
                }

                if ( !System.IO.File.Exists(args[0]))
                {
                    Console.WriteLine("File " + args[0] + " does not exist.");
                    return;
                }

                if ( !System.IO.Directory.Exists(args[1]))
                {
                    Console.WriteLine("The directory " + args[1] + " does not exist.");
                    return;
                }
 
                xlWorkbook = xlApp.Workbooks.Open(args[0]); ;

                Console.WriteLine("Opening file " + args[0] + " for input...");

                int x = xlWorkbook.Worksheets.Count;

                Excel.Worksheet activeSheet = xlWorkbook.ActiveSheet;

                int count = activeSheet.Rows.Count;

                // Get headers for from the spreadsheet.
                var headers = GetHeaders(activeSheet);
                bool moreRows = false;

                for (int i = 2; i < count; i++)
                {
                    var map = CreateRowMapping(activeSheet, headers, i, out moreRows);
                    if (!moreRows)
                    {
                        break;
                    }
                    BuildXml(map, args[1], i);
                }
            }
            finally
            {
                Console.WriteLine("Preparing to shutdown Excel");
                xlWorkbook.Close();
                xlApp.Quit();
                Console.WriteLine("Excel should be shut down");
            }

        }

        static private List<string> GetHeaders(Excel.Worksheet sheet)
        {
            List<string> headers = new List<string>();
            int headerRow = 1;
            for (int i = 1; i < 15; i++)
            {
                Excel.Range range = sheet.Cells[headerRow, i];
                headers.Add(range.Text);
            }

            return headers;
        }

        static Dictionary<string, string> CreateRowMapping(Excel.Worksheet sheet, List<string> headers, int dataRow, out bool moreRows)
        {
            moreRows = true;
            var map = new Dictionary<string, string>();
            Excel.Range moreRowsRange = sheet.Cells[dataRow, 1];
            if (String.IsNullOrEmpty(moreRowsRange.Text))
            {
                moreRows = false;
                return map;
            }
            
            for (int i =1; i <= headers.Count; i ++)
            {
                Excel.Range range = sheet.Cells[dataRow, i];
                if (range.NumberFormat == "m/d/yyyy h:mm")
                {
                    try
                    {
                        DateTime dt = DateTime.Parse(ConvertToDateTime(range.Value2.ToString()));
                        map.Add(headers[i - 1], dt.ToLongDateString());
                    }
                    catch( Exception )
                    {
                        throw;
                    }
                }
                else
                {
                    map.Add(headers[i - 1], range.Text);
                }
                
            }

            return map;
        }


        public static string ConvertToDateTime(string strExcelDate)
        {
            double excelDate;
            try
            {
                excelDate = Convert.ToDouble(strExcelDate);
            }
            catch
            {
                return strExcelDate;
            }
            if (excelDate < 1)
            {
                throw new ArgumentException("Excel dates cannot be smaller than 0.");
            }
            DateTime dateOfReference = new DateTime(1900, 1, 1);
            if (excelDate > 60d)
            {
                excelDate = excelDate - 2;
            }
            else
            {
                excelDate = excelDate - 1;
            }
            return dateOfReference.AddDays(excelDate).ToShortDateString();
        }
    }
}
