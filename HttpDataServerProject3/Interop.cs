using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Diagnostics;

namespace HttpDataServerProject3
{
    public class NskdInterop
    {
        public class WordTablesReader
        {
            public static DataSet FromDocFile(String fileName)
            {
                DataSet ds = new DataSet();
                try
                {
                    Word.Application word = new Word.Application();
                    word.Visible = false;

                    Word.Document doc = word.Documents.Open(FileName: fileName, ReadOnly: true);
                    try
                    {
                        foreach (Word.Table table in doc.Tables)
                        {
                            System.Data.DataTable dt = new System.Data.DataTable();
                            ds.Tables.Add(dt);

                            for (int ci = 0; ci < table.Columns.Count; ci++)
                            {
                                DataColumn dc = new DataColumn("Column" + ci.ToString(), typeof(String));
                                dt.Columns.Add(dc);
                            }
                            for (int ri = 0; ri < table.Rows.Count; ri++)
                            {
                                DataRow dr = dt.NewRow();
                                dt.Rows.Add(dr);
                            }
                            foreach (Word.Cell cell in table.Range.Cells)
                            {
                                String text = cell.Range.Text;
                                text = (new Regex(@"[\x00-\x1f]")).Replace(text, " ");
                                text = text.Trim();
                                dt.Rows[cell.RowIndex - 1][cell.ColumnIndex - 1] = text;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        ds = null; throw new Exception("Ошибка при разборе документа doc.", e);
                    }
                    finally
                    {
                        doc.Close(SaveChanges: false);
                        word.Quit(SaveChanges: false);
                    }
                }
                catch (Exception e) { ds = null; throw new Exception("Ошибка при разборе файла doc.", e); }
                return ds;
            }
            /*
private void button1_Click(object sender, System.EventArgs e)
{

	//Excel Application Object
	Excel.Application oExcelApp;

	this.Activate();

	//Get reference to Excel.Application from the ROT.
	oExcelApp =  (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

	//Display the name of the object.
	MessageBox.Show(oExcelApp.ActiveWorkbook.Name);

	//Release the reference.
	oExcelApp = null;
}

private void button2_Click(object sender, System.EventArgs e)
{			Excel.Workbook  xlwkbook;
	Excel.Worksheet xlsheet;

	//Get a reference to the Workbook object by using a file moniker.
	//The xls was saved earlier with this file name.
	xlwkbook = (Excel.Workbook)  System.Runtime.InteropServices.Marshal.BindToMoniker(textBox1.Text); 

	string sFile = textBox1.Text.Substring(textBox1.Text.LastIndexOf("\\")+1);
	xlwkbook.Application.Windows[sFile].Visible = true;
	xlwkbook.Application.Visible = true;
	xlsheet = (Excel.Worksheet) xlwkbook.ActiveSheet;
	xlsheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
	xlsheet.Cells[1,1] = 100;

	//Release the reference.
	xlwkbook = null;
	xlsheet = null;		}

private void button3_Click(object sender, System.EventArgs e)
{
	Word.Application wdapp;

	//Shell Word
	System.Diagnostics.Process.Start("<Path to WINWORD.EXE>");

	this.Activate();

	//Word and other Office applications register themselves in 
	//ROT when their top-level window loses focus. Having a MessageBox 
	//forces Word to lose focus and then register itself in the ROT.

	MessageBox.Show("Launched Word");

	//Get the reference to Word.Application from the ROT.
	wdapp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");

	//Display the name.
	MessageBox.Show(wdapp.Name);

	//Release the reference.
	wdapp = null;
}
		             * */
        }
    }
}