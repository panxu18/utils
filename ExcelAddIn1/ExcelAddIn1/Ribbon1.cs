using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        /**
         * 将多个excel合并到一个工作簿中
         * 1. 选择待合并的目录
         * 2. 合并excel
         * 3. 选择保存目录
         */
        private void mergeWorkbooks(object sender, RibbonControlEventArgs e)
        {

            List<String> files = chooseFiles();
            if (files.Count > 0)
            {
                Excel.Application Application = Globals.ThisAddIn.Application;
                Excel.Workbooks Workbooks = Application.Workbooks;
                Application.ScreenUpdating = false;
                Excel.Workbook desWb = Workbooks.Add();
                Excel.Worksheet activeSheet = desWb.ActiveSheet;
                /*
                foreach (FileInfo f in Dir.GetFiles("*.xls", SearchOption.TopDirectoryOnly))
                {
                    list.Add(f.FullName);
                }
                */
                merge(Workbooks, activeSheet, files);

                Application.ScreenUpdating = true;
                string result = "合并文以下件 " + files.Count + " 个：\n" + string.Join("\n", files.ToArray());
                MessageBox.Show(result, "合并成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }    
        }

        private void merge(Excel.Workbooks Workbooks, Excel.Worksheet desWorksheet, List<String> files)
        {
            for (int i = 0; i < files.Count; i++)
            {
                Excel.Workbook wb = Workbooks.Open(files[i]);
                if (i == 0)
                {
                    wb.Sheets[1].Range["a1"].CurrentRegion.Copy(desWorksheet.Cells[1, 1]);
                }
                else
                {
                    int row = desWorksheet.Range["a1048576"].End[Excel.XlDirection.xlUp].Row;
                    wb.Sheets[1].Range["a1"].CurrentRegion.Offset(1).Copy(desWorksheet.Cells[row + 1, 1]);
                }
                wb.Close();
            }
        }

        private void saveFile(Excel.Workbook wb)
        {
            
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "All files(*.*)|*.*";
            saveFileDialog.FileName = "mergeResult";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.AddExtension = false;
            saveFileDialog.ShowDialog();
            
            string fileNameExt = saveFileDialog.FileName.ToString();
            int index = fileNameExt.LastIndexOf(".");
            string fileName = fileNameExt;
            if (index >= 0)
            {
                fileName = fileNameExt.Substring(0, fileNameExt.LastIndexOf("."));
            }
            wb.SaveAs(fileName);
        }
        
        /**
         * 选择多个文件
         * */
        private List<String> chooseFiles()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Title = "选择需要合并的Excel文件";
            dlg.Filter = "All files|*.*|xlsx|*.xlsx|xls|*.xls";
            List<String> files = new List<string>();
            if (dlg.ShowDialog() == DialogResult.OK)
            {

                foreach (string file in dlg.FileNames)
                {
                    files.Add(file);
                    MessageBox.Show(file);
                }

            }
            return files;

        }

        /**
         * 选择目录
         * */
        private string choosepath()
        {
            FolderBrowserDialog dlgOpenPath = new FolderBrowserDialog();
            dlgOpenPath.Description = "选择待合并文件的目录";
            DialogResult dr = dlgOpenPath.ShowDialog();
            return dlgOpenPath.SelectedPath;
        }

        private string makeUrl (string urlHeader, string keyword)
        {
            return WebUtility.UrlEncode(urlHeader + keyword);
        }
    }
}
