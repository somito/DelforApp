using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace DelforApp
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            FolderBrowserDialog dialog = new FolderBrowserDialog();

            string dirpath = "";
            string filepath = "";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                dirpath = dialog.SelectedPath;


                string[] fileEntries = Directory.GetFiles(dirpath);

                //Initialize an excel file and a line number to be passed to the Output method
                int rowNumber = 1;

                CreateExcel(dirpath);

                bool save = false;

                for (int i = 0; i < fileEntries.Length; i++)
                {
                    filepath = fileEntries[i];

                    //Check if the inspected file is a .rcv or .xml or .env, if not, continue to next file
                    string extension = Path.GetExtension(filepath);

                    if (extension != ".rcv" && extension != ".xml" && extension != ".env")
                    {
                        continue;
                    }

                    Delfor Message = new Delfor(filepath);

                    rowNumber++;

                    //Flag for saving only at last entry to the excel
                    if (i == fileEntries.Length - 1)
                    {
                        save = true;
                    }

                    OutputToExcel(Message, rowNumber, dirpath + "\\Delfor_Analysis.xlsx", save);
                }

            }
        }

        private static void CreateExcel(string dirpath)
        {
            //Check if the result excel already exists. If yes, delete it to overwrite

            try
            {
                if (File.Exists(dirpath + "\\Delfor_Analysis.xlsx"))
                {
                    File.Delete(dirpath + "\\Delfor_Analysis.xlsx");
                }
            }

            catch (IOException)
            {
                MessageBox.Show("The result excel file exists and is in use.");
                Environment.Exit(0);
            }

            //Create the result excel

            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            Directory.SetCurrentDirectory(dirpath);
            Excel.Workbook wb = excel.Workbooks.Add();
            Excel.Worksheet sh = wb.Sheets.Add();
            sh.Name = "Delfor";
            sh.Cells[1, "A"].Value2 = "FileName";
            sh.Cells[1, "B"].Value2 = "Envelope:From";
            sh.Cells[1, "C"].Value2 = "Envelope:To";
            sh.Cells[1, "D"].Value2 = "Item Alias";
            sh.Cells[1, "E"].Value2 = "BuyerID";
            sh.Cells[1, "F"].Value2 = "SellerID";
            sh.Cells[1, "G"].Value2 = "Loading Place";
            sh.Cells[1, "H"].Value2 = "Ship-To Place";
            sh.Cells[1, "I"].Value2 = "Plant Code";
            sh.Cells[1, "J"].Value2 = "Place of Discharge";
            sh.Cells[1, "K"].Value2 = "Place of Delivery";
            sh.Cells[1, "L"].Value2 = "VW Partner Type 10";

            Excel.Range formatRange;
            formatRange = sh.get_Range("A1");
            formatRange.EntireRow.Font.Bold = true;
            formatRange = sh.get_Range("A1, B1, C1, D1, E1, F1, G1, H1, I1, J1, K1, L1");
            formatRange.EntireColumn.NumberFormat = "@";

            wb.SaveAs(dirpath + "\\Delfor_Analysis.xlsx");
            wb.Close(true);
            excel.Quit();
        }

        private static void OutputToExcel(Delfor Message, int rowNumber, string path, bool save)
        {
            Message.MessageFrom = Message.GetMessageFrom(Message.DelforMessageXml);
            Message.MessageTo = Message.GetMessageTo(Message.DelforMessageXml);
            Message.AliasNumber = Message.GetAliasNumber(Message.DelforMessageXml);
            Message.BuyerID = Message.GetBuyerID(Message.DelforMessageXml);
            Message.SellerID = Message.GetSellerID(Message.DelforMessageXml);
            Message.LoadingPlace = Message.GetLoadingPlace(Message.DelforMessageXml);
            Message.ShipToPlace = Message.GetShipToPlace(Message.DelforMessageXml);
            Message.PlantCode = Message.GetPlantCode(Message.DelforMessageXml);
            Message.Loc1 = Message.GetLoc1(Message.DelforMessageXml);
            Message.Loc2 = Message.GetLoc2(Message.DelforMessageXml);
            Message.VWType10 = Message.SellerID + Message.PlantCode;

            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            Excel.Workbook wb = excel.Workbooks.Open(path);
            Excel.Worksheet sh = wb.ActiveSheet;

            sh.Cells[rowNumber, "A"] = Message.FileName;
            sh.Cells[rowNumber, "B"] = Message.MessageFrom;
            sh.Cells[rowNumber, "C"] = Message.MessageTo;
            sh.Cells[rowNumber, "D"] = Message.AliasNumber;
            sh.Cells[rowNumber, "E"] = Message.BuyerID;
            sh.Cells[rowNumber, "F"] = Message.SellerID;
            sh.Cells[rowNumber, "G"] = Message.LoadingPlace;
            sh.Cells[rowNumber, "H"] = Message.ShipToPlace;
            sh.Cells[rowNumber, "I"] = Message.PlantCode;
            sh.Cells[rowNumber, "J"] = Message.Loc1;
            sh.Cells[rowNumber, "K"] = Message.Loc2;
            sh.Cells[rowNumber, "L"] = Message.VWType10;

            if (save)
            {
                wb.Save();
            }

            wb.Close(true);
            excel.Quit();
        }
    }
}
