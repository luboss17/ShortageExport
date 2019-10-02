using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Win32;
using System.Windows.Forms;
using ShortageExport;

namespace ShortageExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"\\";
        string globalFileName=DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xls";
        string globalPath = @"Z:\Shortage\";
        string localPath = "";
        string outputPath;
        const string templatePath = @"Z:\Shortage\shortage_template.xlsx";
        string shortageCSVPath="";
        private const string userRoot = "HKEY_CURRENT_USER";
        //Keys for Registry
        private const string defaultSaveLoc_keyName = userRoot + "\\" + "ShortageLocations";
        //Values for Registry
        private const string saveLoc_valueName = "shortageLocation";

        public MainWindow()
        {
            InitializeComponent();
            local_radio.IsChecked = true;
            
            localPath = getRegistryValue(defaultSaveLoc_keyName, saveLoc_valueName);
            if (localPath==null)
            {
                localPath = DesktopPath;
                setRegistryValue(defaultSaveLoc_keyName, saveLoc_valueName, localPath);
            }
            updateSavePath(localPath+ updateFileName());
        }
        //Get value for specified Registry
        private string getRegistryValue(string keyName, string valueName)
        {
            string registryValue = (string)Registry.GetValue(keyName, valueName, "");
            return registryValue;
        }

        //Set value to Registry
        private void setRegistryValue(string keyName, string valueName, string value)
        {
            Registry.SetValue(keyName, valueName, value);
        }


        private void dropPanel_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            e.Effects = System.Windows.DragDropEffects.All;
            e.Handled = true;
        }
        
        List<string> CSVContent = new List<string>();
        private List<string> readExcel(string path)
        {
            List<string> fileStream = new List<string>();
            using (var reader = new StreamReader(path))
            {
                List<string> listA = new List<string>();
                
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    //var values = line.Split(',');
                    fileStream.Add(line+'\n');
                }
            }
            return fileStream;
        }
        private string cutBeginEndQuote(string str)
        {
            string returnStr = str;
            if ((str.StartsWith("\"")) && (str.EndsWith("\"")))
                returnStr = str.Substring(1, str.Length - 2);
            return returnStr;
        }
        //extract all line items to 2D arrays with right format to write to excel
        private string[,] parseAllItemsToArray(List<string> content)
        {
            string bin = "", customer = "", orderDate = "", dueDate = "", priority = "", qty="",shortage="",type="",topItem="";
            const int colCount = 9;
            string[,] arrayReturn = new string[content.Count(), colCount];
            int rowCounter = 0;
            foreach (string line in content)
            {
                string[] cells = line.Split(',');

                if ((cells[0] != "") && (cells[0] != ">") && (cells[0] != "\n"))
                {//begin of the job, set bin, customer, orderDate, dueDate and priority for following shortage line
                    bin = cells[0].ToString();
                    customer = cells[5].ToString();
                    orderDate = cells[10].ToString();
                    dueDate = cells[11].ToString();
                    priority = cells[12].ToString();
                }
                else if (cells[0] == ">")
                {//this line is shortage
                    qty = cells[2].ToString();
                    shortage =cutBeginEndQuote(cells[4].ToString());
                    type = cells[6].ToString();

                    //set value of this line item to array with following order
                    //customer, bin, orderDate, dueDate, qty, description, type, priority;
                    arrayReturn[rowCounter, 0] = customer;
                    arrayReturn[rowCounter, 1] = bin;
                    arrayReturn[rowCounter, 2] = orderDate;
                    arrayReturn[rowCounter, 3] = dueDate;
                    arrayReturn[rowCounter, 4] = qty;
                    arrayReturn[rowCounter, 6] = shortage;
                    arrayReturn[rowCounter, 7] = type;
                    arrayReturn[rowCounter, 8] = priority;
                    rowCounter++;
                }
                else if ((cells[0]=="")&&(cells.Length>1))
                {//this line is top line item
                    qty = cells[1].ToString();
                    topItem = cutBeginEndQuote(cells[3].ToString());
                    
                    //set value of this line item to array with following order
                    //customer, bin, orderDate, dueDate, qty, description, type, priority;
                    arrayReturn[rowCounter, 0] = customer;
                    arrayReturn[rowCounter, 1] = bin;
                    arrayReturn[rowCounter, 2] = orderDate;
                    arrayReturn[rowCounter, 3] = dueDate;
                    arrayReturn[rowCounter, 4] = qty;
                    arrayReturn[rowCounter, 5] = topItem;
                    arrayReturn[rowCounter, 8] = priority;
                    rowCounter++;
                }
            }
            return arrayReturn;
        }
        //extract all line items to 2D arrays with right format to write to excel
        private string[,] parseShortageTopToArray(List<string> content)
        {
            string bin = "", customer = "", orderDate = "", dueDate = "", priority = "", qty = "", shortage = "", type = "", topItem = "";
            const int colCount = 9;
            string[,] arrayReturn = new string[content.Count(), colCount];
            int rowCounter = 0;
            foreach (string line in content)
            {
                string[] cells = line.Split(',');

                if ((cells[0] != "") && (cells[0] != ">") && (cells[0] != "\n"))
                {//begin of the job, set bin, customer, orderDate, dueDate and priority for following shortage line
                    bin = cells[0].ToString();
                    customer = cells[5].ToString();
                    orderDate = cells[10].ToString();
                    dueDate = cells[11].ToString();
                    priority = cells[12].ToString();
                }
                else if (cells[0] == ">")
                {//this line is shortage
                    qty = cells[2].ToString();
                    shortage = cutBeginEndQuote(cells[4].ToString());
                    type = cells[6].ToString();

                    //set value of this line item to array with following order
                    //customer, bin, orderDate, dueDate, qty, top, shortage, type, priority;
                    arrayReturn[rowCounter, 0] = customer;
                    arrayReturn[rowCounter, 1] = bin;
                    arrayReturn[rowCounter, 2] = orderDate;
                    arrayReturn[rowCounter, 3] = dueDate;
                    arrayReturn[rowCounter, 4] = qty;
                    arrayReturn[rowCounter, 5] = topItem;
                    arrayReturn[rowCounter, 6] = shortage;
                    arrayReturn[rowCounter, 7] = type;
                    arrayReturn[rowCounter, 8] = priority;
                    rowCounter++;
                }
                else if ((cells[0] == "") && (cells.Length > 1))
                {//this line is top line item
                    //qty = cells[1].ToString();
                    topItem = cutBeginEndQuote(cells[3].ToString());
                }
            }
            return arrayReturn;
        }
        private static string[] RemoveRangeArray<String>(string[] source, int startIndex, int endIndex)
        {
            string[] dest = new string[source.Length - (endIndex-startIndex)];
            if (startIndex > 0)
                Array.Copy(source, 0, dest, 0, startIndex);

            if (endIndex < source.Length - 1)
                Array.Copy(source, startIndex + 1, dest, startIndex, source.Length - startIndex - 1);

            return dest;
        }
        //check if the price is in spread out, then combine them and return array
        private string[] combinePriceInArray(string[] cells)
        {
            int priceEndPos = 9,priceStartPos=9;
            if (cells[priceStartPos].ToString().StartsWith("\""))
            {
                int runningPos = priceStartPos+1;
                string price = cells[priceStartPos].ToString();
                while(runningPos<cells.Count())
                {
                    price += cells[runningPos].ToString();
                    if (cells[runningPos].ToString().Contains("\""))
                    {
                        priceEndPos = runningPos;
                        break;
                    }
                    runningPos++;
                    
                }
                //sucess combine multiple cell for price
                if (priceEndPos>priceStartPos)
                {
                    /*var tempCells = new List<String>(cells);
                    tempCells.RemoveRange(priceStartPos + 1, priceEndPos - priceStartPos);
                    tempCells[priceStartPos] = price;
                    cells = tempCells.ToArray();*/
                    cells = RemoveRangeArray(cells,;
                }
            }
        }
        //extract shortage to 2D arrays with right format to write to excel
        private string[,] parseShortageToArray(List<string> content)
        {
            string bin = "", customer = "", orderDate = "", dueDate = "", priority = "";
            const int colCount = 8;
            string[,] arrayReturn = new string[content.Count(), colCount];
            int rowCounter = 0;
            foreach (string line in content)
            {
                string[] cells = line.Split(',');
                
                if ((cells[0] != "") && (cells[0] != ">") && (cells[0] != "\n"))
                {//begin of the job, set bin, customer, orderDate, dueDate and priority for following shortage line
                    bin = cells[0].ToString();
                    customer = cells[5].ToString();
                    //check if the price is in spread out, then combine them and return array
                    orderDate = cells[10].ToString();
                    dueDate = cells[11].ToString();
                    priority = cells[12].ToString();
                }
                else if (cells[0] == ">")
                {//this line is shortage
                    string qty = cells[2].ToString();
                    string description = cutBeginEndQuote(cells[4].ToString());
                    string type = cells[6].ToString();

                    //set value of this line item to array with following order
                    //customer, bin, orderDate, dueDate, qty, description, type, priority;
                    arrayReturn[rowCounter, 0] = customer;
                    arrayReturn[rowCounter, 1] = bin;
                    arrayReturn[rowCounter, 2] = orderDate;
                    arrayReturn[rowCounter, 3] = dueDate;
                    arrayReturn[rowCounter, 4] = qty;
                    arrayReturn[rowCounter, 5] = description;
                    arrayReturn[rowCounter, 6] = type;
                    arrayReturn[rowCounter, 7] = priority;
                    rowCounter++;
                }
            }
            return arrayReturn;
        }

        //write 2d array of shortage array to wsheet
        private Excel.Worksheet write2DArraysToWsheet(Excel.Worksheet wsheet, string[,] datas)
        {
            int startRow = 3;
            int startCol = 2;
            Excel.Range datarange = (Excel.Range)wsheet.Cells[startRow,startCol];//Set initial cell of datarange
            datarange = datarange.get_Resize(datas.GetLength(0),datas.GetLength(1));//Set the size of ddddatarange
            datarange.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, datas);
            return wsheet;
        }

        //main function to write To Excel
        private void writeShortageToExcel()
        {
            //get rawCSV content
            CSVContent = readExcel(shortageCSVPath);
            //Get shortage data from CSV file
            string[,] shortage_data = parseShortageToArray(CSVContent);
            string[,] shortageTop_data = parseShortageTopToArray(CSVContent);
            string[,] allLineItems_data = parseAllItemsToArray(CSVContent);
            
            //open template File
            ExcelFile xcelFile = new ExcelFile();
            xcelFile.ExcelFilePath = templatePath;
            xcelFile.openExcel();

            //get and write shortage data to new shortageSheet
            Excel.Worksheet shortage_sheet = write2DArraysToWsheet(xcelFile.ShortageSheet, shortage_data);
            shortage_sheet = autofitWSheet(shortage_sheet);
            xcelFile.ShortageSheet = shortage_sheet;

            //get and write Shortage with Top Items data to new ShortTop wSheet
            Excel.Worksheet shortageTop_sheet = write2DArraysToWsheet(xcelFile.ShortageTopSheet, shortageTop_data);
            shortageTop_sheet = autofitWSheet(shortageTop_sheet);
            xcelFile.ShortageTopSheet = shortageTop_sheet;

            //get and write all Items data to new All Item Sheet
            Excel.Worksheet allItems_sheet = write2DArraysToWsheet(xcelFile.AllItemsSheet, allLineItems_data);
            allItems_sheet = autofitWSheet(allItems_sheet);
            xcelFile.AllItemsSheet = allItems_sheet;

            //save excel to new path
            xcelFile.ExcelFilePath = outputPath;
            xcelFile.saveExcel();
        }

        private Excel.Worksheet autofitWSheet(Excel.Worksheet wsheet)
        {
            if (wsheet != null)
            {
                Excel.Range excelRange = wsheet.get_Range("B:T", System.Type.Missing);
                excelRange.Columns.AutoFit();
            }
            return wsheet;
        }
        

        private void dropBox_PreviewDrop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                //get file path
                string[] Paths = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                shortageCSVPath = Paths[0];
                if (local_radio.IsChecked==true)
                    updateSavePath(localPath + updateFileName());
                //write to Excel
                writeShortageToExcel();

                /*  For future multiple files drop support
                 * string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string filePath in files)
                {
                    panel1.Text += filePath + "\n";
                }*/
            }
            e.Handled = true;
        }

        private void dropBox_PreviewDragOver(object sender, System.Windows.DragEventArgs e)
        {
            e.Effects = System.Windows.DragDropEffects.All;
            e.Handled = true;
        }
        private void updateSavePath(string path)
        {
            file_loca_txtBox.Content = extractPathFromExcelPathName(path);
            outputPath = path;
        }
        private void global_radio_Checked(object sender, RoutedEventArgs e)
        {
            changeSaveLoc_btn.IsEnabled = false;
            updateSavePath(globalPath+updateFileName());
        }
        private string updateFileName()
        {
            string fileName;
            if (local_radio.IsChecked == true)
            {
                fileName = extractFileNameFromExcelPathName(shortageCSVPath);
            }
            else
                fileName=globalFileName;
            return fileName;
        }
        private void local_radio_Checked(object sender, RoutedEventArgs e)
        {
            changeSaveLoc_btn.IsEnabled = true;
            updateSavePath(localPath+ updateFileName());
        }
        //Return just the folder location of a Full Path File Name
        private static string extractPathFromExcelPathName(string pathName)
        {
            int index = pathName.Length - 1;
            if ((pathName.Contains(".xls"))||(pathName.Contains(".csv")))
            {
                while ((pathName[index] != '\\') && (index > 0))
                {
                    pathName = pathName.Substring(0, index);
                    index--;
                }
            }
            return pathName;
        }

        //return just the file name of full path
        private static string extractFileNameFromExcelPathName(string pathName)
        {
            int index = pathName.Length - 1;
            int csvLength = 3;
            if (pathName.Contains(".csv"))
            {
                while ((pathName[index] != '\\') && (index > 0))
                {
                    index--;
                }
                index++;//offset the "\\" at beginning of file name
                pathName = pathName.Substring(index,pathName.Length-index-csvLength)+"xls";
                
            }
            return pathName;
        }
        private string showFolderBrowserDialog(string lastPath)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = lastPath;
            DialogResult result = fbd.ShowDialog();
            string path = "";
            path = fbd.SelectedPath;
            return path;
        }

        //Show save dialog box and return path name
        private string getSavePathName(string lastPath)
        {

            System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
            if (lastPath!="")
                saveDlg.InitialDirectory = lastPath;
            saveDlg.Filter = "Excel|*.xls";
            saveDlg.ShowDialog();
            string path = "";
            path = saveDlg.FileName;
            //if (path != "")
            //    lastPathName = extractPathFromExcelPathName(path);
            return path;
        }
        private void changeSaveLoc_btn_Click(object sender, RoutedEventArgs e)
        {
            localPath = showFolderBrowserDialog(localPath)+"\\";//getSavePathName(localPath);
            if (localPath.EndsWith(@"\\"))
            {
                localPath = localPath.Substring(0, localPath.Length - 1);
            }
            setRegistryValue(defaultSaveLoc_keyName, saveLoc_valueName, localPath);
            updateSavePath(localPath+ updateFileName());
        }
    }
}
