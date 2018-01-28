using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordTemplate1
{
    public class Angajat
    {
        private string nume;
        private string sex;
        private int oreLucrate;
        private int salariuOrar;

        public Angajat(string nume, string sex, int oreLucrate, int salariuOrar)
        {
            this.nume = nume;
            this.sex = sex;
            this.oreLucrate = oreLucrate;
            this.salariuOrar = salariuOrar;
        }

        public string Nume { get => nume; set => nume = value; }
        public string Sex { get => sex; set => sex = value; }
        public int OreLucrate { get => oreLucrate; set => oreLucrate = value; }
        public int SalariuOrar { get => salariuOrar; set => salariuOrar = value; }

        public int calculSalariu ()
        {
            return OreLucrate * SalariuOrar;
        }
    }


    public partial class ThisDocument
    {
        Button load, closeAndCalculate;
        Excel.Application eapl;
        Excel.Workbook wb;
        Excel.Worksheet sheet;
        List<Angajat> lAng = new List<Angajat>();
        int row = 1;
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            eapl = new Excel.Application();
            wb = eapl.Workbooks.Add();
            sheet = wb.Sheets.Add();

            eapl.Visible = true;
            load = new Button();
            closeAndCalculate = new Button();
            load.Text = "Incarca date in sheet";
            closeAndCalculate.Text = "Inchide document si calculeaza fondul de salarii";

            Globals.ThisDocument.ActionsPane.Controls.Add(load);
            Globals.ThisDocument.ActionsPane.Controls.Add(closeAndCalculate);

            load.Click += new EventHandler(load_Click);
            closeAndCalculate.Click += new EventHandler(closeAndCalculate_Click);


        }

        private void load_Click(object sender, EventArgs e)
        {
            Angajat temp = new Angajat(
                nameTextControl.Text, 
                dropDownControl.Text, 
                Int32.Parse(oreTextControl.Text), 
                Int32.Parse(salariuTextControl.Text));

            lAng.Add(temp);

            ((Excel.Range)sheet.Cells[row, 1]).Value = temp.Nume;
            ((Excel.Range)sheet.Cells[row, 2]).Value = temp.Sex;
            ((Excel.Range)sheet.Cells[row, 3]).Value = temp.OreLucrate;
            ((Excel.Range)sheet.Cells[row, 4]).Value = temp.SalariuOrar;

            row++;
        }

        private void closeAndCalculate_Click(object sender, EventArgs e)
        {
            int sum = 0;
            foreach(Angajat temp in lAng)
            {
                sum += temp.calculSalariu();
            }
            ((Excel.Range)sheet.Cells[row, 1]).Value = "Fond salarial total";
            ((Excel.Range)sheet.Cells[row, 2]).Value = sum.ToString();

            Globals.ThisDocument.Application.Quit();
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
