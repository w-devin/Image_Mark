using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Image_Mark
{
    class excelEdit
    {
        public String excFilePath;
        public Microsoft.Office.Interop.Excel.Application App;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;

        //struct fun
        public excelEdit()
        {

        }

        //Create a Microsoft.Office.Interop.Excel object
        public void Create()
        {
            App = new Microsoft.Office.Interop.Excel.Application();
            wbs = App.Workbooks;
            wb = wbs.Add(true);
        }

        //Open a excel file
        public void Open(String filePath)
        {
            App = new Microsoft.Office.Interop.Excel.Application();
            wbs = App.Workbooks;
            wb = wbs.Add(filePath);
            excFilePath = filePath;
        }

        //get a worksheet
        public Microsoft.Office.Interop.Excel.Worksheet getSheet(String sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet) wb.Worksheets[sheetName];

            return s;
        }

        //add a worksheet
        public Microsoft.Office.Interop.Excel.Worksheet addSheet(String sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = sheetName;

            return s;
        }

        //delete a worksheet
        public void delSheet(String sheetName)
        {
            ((Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[sheetName]).Delete();
        }

        //rename a sheet with old name and new name
        public Microsoft.Office.Interop.Excel.Worksheet renameSheet(String oldName, String newName)
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[oldName];
            s.Name = newName;
            return s;
        }
       
        //rename a sheet with new name 
        public Microsoft.Office.Interop.Excel.Worksheet renameSheet(Microsoft.Office.Interop.Excel.Worksheet s, String newName)
        {
            s.Name = newName;
            return s;
        }

        //set cell value
        public void setCellValue(Microsoft.Office.Interop.Excel.Worksheet ws, int x, int y, object value)
        {
            ws.Cells[x, y] = value;
        }

        //set cell value
        public void setCellValue(string ws, int x, int y, object value)
        {
            getSheet(ws).Cells[x, y] = value;
        }

        //insert talbe with System.Data.DataTable
        public void insertTable(System.Data.DataTable dt, String ws, int startX, int startY)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
                for (int j = 0; j < dt.Columns.Count; j++)
                    getSheet(ws).Cells[startX + i, startY + j] = dt.Rows[i][j].ToString();
        }

        //insert table with 
        public void insertTable(System.Data.DataTable dt, Microsoft.Office.Interop.Excel.Worksheet ws, int startX, int startY)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
                for (int j = 0; j < dt.Columns.Count; j++)
                    ws.Cells[startX + i, startY + j] = dt.Rows[i][j].ToString();
        }

        //save 
        public bool Save()
        {
            if (excFilePath == "")
                return false;

            try
            {
                wb.Save();
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        //Save as..
        public bool SaveAs(object fileName)
        {
            try
            {
                wb.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        } 

        //close
        public void Close()
        {
            wb.Save();
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            App.Quit();
            wb = null;
            wbs = null;
            App = null;
            GC.Collect();
        }




    }
}
