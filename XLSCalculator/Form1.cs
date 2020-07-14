using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace XLSCalculator
{
    public partial class Form1 : Form
    {
        public string ieskomas_nr = "";
        public int bendros_pajamos = 0;
        public int bendros_islaidos = 0;
        public int bendras_uzdarbis = 0; //pajamos-islaidos

        //Sheetu strukturos irasams saugoti
        public struct darbasRow
        {
            public string nr { get; set; }
            public string aprasas { get; set; }
            public string pirk_moketojo_nr { get; set; }
            public DateTime sukurimo_data { get; set; }
            public DateTime pradzios_data { get; set; }
            public DateTime pabaigos_data { get; set; }
            public string busena { get; set; }
            public string atsakingas_asmuo { get; set; }
            public string pirmas_glob_dimensijos_kodas { get; set; }
            public string antras_glob_dimensijos_kodas { get; set; }

            public override string ToString()
            {
                if(nr == "")
                {
                    nr = "   ......   ";
                }
                if (aprasas == "")
                {
                    aprasas = "   ......   ";
                }
                if (nr == "")
                {
                    pirk_moketojo_nr = "   ......   ";
                }
                if (busena == "")
                {
                    busena = "   ......   ";
                }
                if (atsakingas_asmuo == "")
                {
                    atsakingas_asmuo = "   ......   ";
                }
                if (pirmas_glob_dimensijos_kodas == "")
                {
                    pirmas_glob_dimensijos_kodas = "   ......   ";
                }
                if (antras_glob_dimensijos_kodas == "")
                {
                    antras_glob_dimensijos_kodas = "   ......   ";
                }

                return String.Format("{0,-10} | {1,-10} | {2,-10} | {3,-10} | {4,-10} | {5,5} | {6,5} | {7,10} | {8,15} " +
                "| {9,5}", nr,aprasas,pirk_moketojo_nr,sukurimo_data,pradzios_data,
                pabaigos_data,busena,atsakingas_asmuo,pirmas_glob_dimensijos_kodas,antras_glob_dimensijos_kodas);

            }

        }
        List<darbasRow> darbasTable= new List<darbasRow>();

        public struct darboUzdRow
        {
            public string darbo_nr { get; set; }
            public string darbo_uzd_nr { get; set; }
            public string aprasas { get; set; }
            public string darbo_uzd_tipas { get; set; }
            public string pirmas_glob_dimensijos_kodas { get; set; }
            public string antras_glob_dimensijos_kodas { get; set; }

            public override string ToString()
            {
                if (darbo_nr == "")
                {
                    darbo_nr = "   ......   ";
                }
                if (darbo_uzd_nr == "")
                {
                    darbo_uzd_nr = "   ......   ";
                }
                if (aprasas == "")
                {
                    aprasas = "   ......   ";
                }
                if (darbo_uzd_tipas == "")
                {
                    darbo_uzd_tipas = "   ......   ";
                }
                if (pirmas_glob_dimensijos_kodas == "")
                {
                    pirmas_glob_dimensijos_kodas = "   ......   ";
                }
                if (antras_glob_dimensijos_kodas == "")
                {
                    antras_glob_dimensijos_kodas = "   ......   ";
                }

                return String.Format("{0,-10} | {1,-10} | {2,-10} |  {3,-10} | {4,20} | {5,5} ",
                darbo_nr,darbo_uzd_nr,aprasas,darbo_uzd_tipas,pirmas_glob_dimensijos_kodas,antras_glob_dimensijos_kodas);
            }

        }
        List<darboUzdRow> darboUzdTable= new List<darboUzdRow>();

        public struct darboUzdDimRow
        {
            public string darbo_nr { get; set; }
            public string darbo_uzd_nr { get; set; }
            public string dimensijos_kodas { get; set; }
            public string dimensijos_vertes_kodas { get; set; }

            public override string ToString()
            {
                if (darbo_nr == "")
                {
                    darbo_nr = "   ......   ";
                }
                if (darbo_uzd_nr == "")
                {
                    darbo_uzd_nr = "   ......   ";
                }
                if (dimensijos_kodas == "")
                {
                    dimensijos_kodas = "   ......   ";
                }
                if (dimensijos_vertes_kodas == "")
                {
                    dimensijos_vertes_kodas = "   ......   ";
                }

                return String.Format("{0,-10} | {1,-10} | {2,10} | {3,15} ", darbo_nr,
                    darbo_uzd_nr,dimensijos_kodas,dimensijos_vertes_kodas);

            }
        }
        List<darboUzdDimRow> darboUzdDimTable= new List<darboUzdDimRow>();

        public struct darbuPlanEilRow
        {
            public int eil_nr { get; set; }
            public string darbo_nr { get; set; }
            public string darbo_uzd_nr { get; set; }
            public DateTime planavimo_data { get; set; }
            public string dokumento_nr { get; set; }
            public string tipas { get; set; }
            public string nr { get; set; }
            public string aprasas { get; set; }
            public int kiekis { get; set; }
            public int vieneto_savikaina { get; set; }
            public int vieneto_kaina { get; set; }
            public string mat_vnt_kodas { get; set; }

            public override string ToString()
            {
                if (darbo_nr == "")
                {
                    darbo_nr = "   ......   ";
                }
                if (darbo_uzd_nr == "")
                {
                    darbo_uzd_nr = "   ......   ";
                }
                if (dokumento_nr == "")
                {
                    dokumento_nr = "   ......   ";
                }
                if (tipas == "")
                {
                    tipas = "   ......   ";
                }
                if (nr == "")
                {
                    nr = "   ......   ";
                }
                if (aprasas == "")
                {
                    aprasas = "   ......   ";
                }
                if (mat_vnt_kodas == "")
                {
                    mat_vnt_kodas = "   ......   ";
                }

                return String.Format("{0,-10} | {1,-10} | {2,-10} | {3,-10} | {4,-10} | {5,5} | {6,5} | {7,10} | {8,20} " +
            "| {9,20} | {10,20} | {11,20}", eil_nr,darbo_nr,darbo_uzd_nr,planavimo_data,dokumento_nr,
                    tipas,nr,aprasas,kiekis,vieneto_savikaina,vieneto_kaina,mat_vnt_kodas);
            }
        }
        List<darbuPlanEilRow> darbuPlanEilTable= new List<darbuPlanEilRow>();

        public Form1()
        {
            InitializeComponent();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {  }

        //is raw string tipo formuojami sheet objektai su skirtingo tipo reiksmem
        private void formDarbasObject(string dataStr)
        {
            string[] separatingStrings = { "\r\n", ",", "\r", "\n", "\t" };

            string[] values = dataStr.Split(separatingStrings, System.StringSplitOptions.None);

            for (int i=0; i<values.Length; i++)
            {
                if (i+9<=values.Length)
                {
                    var row = new darbasRow();
                    row.nr = values[i];
                    row.aprasas = values[++i];
                    row.pirk_moketojo_nr = values[++i];
                    row.sukurimo_data = DateTime.ParseExact(values[++i], "yyyy-MM-dd",
                                            System.Globalization.CultureInfo.InvariantCulture);
                    row.pradzios_data = DateTime.ParseExact(values[++i], "yyyy-MM-dd",
                                            System.Globalization.CultureInfo.InvariantCulture);
                    row.pabaigos_data = DateTime.ParseExact(values[++i], "yyyy-MM-dd",
                                            System.Globalization.CultureInfo.InvariantCulture);
                    row.busena = values[++i];
                    row.atsakingas_asmuo = values[++i];
                    row.pirmas_glob_dimensijos_kodas = values[++i];
                    row.antras_glob_dimensijos_kodas = values[++i];

                    addToDarbasTable(row);
                }
            }
        }
        private void formDarboUzdObject(string dataStr)
        {
            string[] separatingStrings = { "\r\n", ",", "\r", "\n", "\t" };

            string[] values = dataStr.Split(separatingStrings, System.StringSplitOptions.None);

            for (int i = 0; i < values.Length; i++)
            {
                if (i + 5 <= values.Length)
                {
                    var row = new darboUzdRow();
                    row.darbo_nr = values[i];
                    row.darbo_uzd_nr = values[++i];
                    row.aprasas = values[++i];
                    row.darbo_uzd_tipas = values[++i];
                    row.pirmas_glob_dimensijos_kodas = values[++i];
                    row.antras_glob_dimensijos_kodas = values[++i];

                    addToDarboUzdTable(row);
                }
            }
        }
        private void formDarboUzdDimObject(string dataStr)
        {
            string[] separatingStrings = { "\r\n", ",", "\r", "\n", "\t" };

            string[] values = dataStr.Split(separatingStrings, System.StringSplitOptions.None);

            for (int i = 0; i < values.Length; i++)
            {
                if (i + 3 <= values.Length)
                {
                    var row = new darboUzdDimRow();
                    row.darbo_nr = values[i];
                    row.darbo_uzd_nr = values[++i];
                    row.dimensijos_kodas = values[++i];
                    row.dimensijos_vertes_kodas = values[++i];

                    addToDarboUzdDimTable(row);
                }
            }
        }
        private void formDarbuPlanEilObject(string dataStr)
        {
            string[] separatingStrings = { "\r\n", ",", "\r", "\n", "\t" };

            string[] values = dataStr.Split(separatingStrings, System.StringSplitOptions.None);

            for (int i = 0; i < values.Length; i++)
            {
                if (i + 10 <= values.Length)
                {
                    var row = new darbuPlanEilRow();
                    row.eil_nr = Int32.Parse(values[i]);
                    row.darbo_nr = values[++i];
                    row.darbo_uzd_nr = values[++i];
                    row.planavimo_data = DateTime.ParseExact(values[++i], "yyyy-MM-dd",
                                            System.Globalization.CultureInfo.InvariantCulture);
                    row.dokumento_nr = values[++i];
                    row.tipas = values[++i];
                    row.nr = values[++i];
                    row.aprasas = values[++i];
                    row.kiekis = Int32.Parse(values[++i]);
                    row.vieneto_savikaina = Int32.Parse(values[++i]);
                    row.vieneto_kaina = Int32.Parse(values[++i]);
                    row.mat_vnt_kodas = values[++i];

                    addToDarbuPlanEilTable(row);
                }
            }
        }

        //objektai dedami i struct tipo listus
        public void addToDarbasTable(darbasRow row)
        {
            darbasTable.Add(row);
        }
        public void addToDarboUzdTable(darboUzdRow row)
        {
            darboUzdTable.Add(row);
        }
        public void addToDarboUzdDimTable(darboUzdDimRow row)
        {
            darboUzdDimTable.Add(row);
        }
        public void addToDarbuPlanEilTable(darbuPlanEilRow row)
        {
            darbuPlanEilTable.Add(row);
        }

        //logika
        private void button1_Click(object sender, EventArgs e)
        {
            label6.Hide();

            string darbasTableStr = "";
            string darboUzdTableStr = "";
            string darboUzdDimTableStr = "";
            string darbuPlanavimoEilTableStr = "";

            //failo destination gauti
            string file = textBox5.Text;
            //check if file exists

            if (File.Exists(file))
            {
                //using OfficeOpenXml;
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@file)))
                {
                    var sheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                    var totalRows = sheet.Dimension.End.Row;
                    var totalColumns = sheet.Dimension.End.Column;

                    var sb = new StringBuilder(); //this is your data
                    for (int rowNum = 4; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = sheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        sb.AppendLine(string.Join(",", row));
                    }
                    darbasTableStr = sb.ToString();
                    formDarbasObject(darbasTableStr);


                    sheet = xlPackage.Workbook.Worksheets.ElementAt(1); //select sheet here
                    totalRows = sheet.Dimension.End.Row;
                    totalColumns = sheet.Dimension.End.Column;

                    sb = new StringBuilder(); //this is your data
                    for (int rowNum = 4; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = sheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        sb.AppendLine(string.Join(",", row));
                    }
                    darboUzdTableStr = sb.ToString();
                    formDarboUzdObject(darboUzdTableStr);


                    sheet = xlPackage.Workbook.Worksheets.ElementAt(2); //select sheet here
                    totalRows = sheet.Dimension.End.Row;
                    totalColumns = sheet.Dimension.End.Column;

                    sb = new StringBuilder(); //this is your data
                    for (int rowNum = 4; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = sheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        sb.AppendLine(string.Join(",", row));
                    }
                    darboUzdDimTableStr = sb.ToString();
                    formDarboUzdDimObject(darboUzdDimTableStr);


                    sheet = xlPackage.Workbook.Worksheets.ElementAt(3); //select sheet here
                    totalRows = sheet.Dimension.End.Row;
                    totalColumns = sheet.Dimension.End.Column;

                    sb = new StringBuilder(); //this is your data
                    for (int rowNum = 4; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = sheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        sb.AppendLine(string.Join(",", row));
                    }
                    darbuPlanavimoEilTableStr = sb.ToString();
                    formDarbuPlanEilObject(darbuPlanavimoEilTableStr);
                }
            }
            else {
                label6.Show();
            }

            richTextBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();

            List<darbasRow> darbasTableResult = new List<darbasRow>();
            List<darboUzdRow> darboUzdTableResult = new List<darboUzdRow>();
            List<darboUzdDimRow> darboUzdDimTableResult = new List<darboUzdDimRow>();
            List<darbuPlanEilRow> darbuPlanEilTableResult = new List<darbuPlanEilRow>();

            bendros_pajamos = 0;
            bendros_islaidos = 0;
            bendras_uzdarbis = 0; //pajamos-islaidos

            ieskomas_nr = textBox1.Text;
            
            richTextBox1.AppendText("DARBAS");
            richTextBox1.AppendText(Environment.NewLine);

            richTextBox1.AppendText(String.Format("{0,-21}  {1,-11}                {2,-28}  {3,-10}      {4,-10}  {5,30} " +
                "   {6,20}  {7,10}       {8,5} " +
            " {9,5}", "Nr.","Aprašas","Pirk.-mok nr.","Sukūrimo data","Pradžios data","Pabaigos data","Būsena",
            "Ats.asm.","1gl.dim.kodas","2gl.dim.kodas"));

            richTextBox1.AppendText(Environment.NewLine);
            foreach (var row in darbasTable)
            {
                if(row.nr==ieskomas_nr)
                {
                    richTextBox1.AppendText(row.ToString());
                    richTextBox1.AppendText(Environment.NewLine);
                }
            }
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("DARBO UŽDUOTIS");
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText(String.Format("{0,-21}  {1,-11}                {2,-10}  {3,-5} " +
                "   {4,5}  {5,5} ", "Darbo nr.","Darbo užd.nr.","Aprašas","Darbo užd.tipas",
                "1gl.dim.kodas",  "2gl.dim.kodas"));
            richTextBox1.AppendText(Environment.NewLine);
            foreach (var row in darboUzdTable)
            {
                if(row.darbo_nr==ieskomas_nr)
                {
                    richTextBox1.AppendText(row.ToString());
                    richTextBox1.AppendText(Environment.NewLine);
                }
            }
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("DARBO UŽDUOTIES DIMENSIJA");
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText(String.Format("{0,-21}  {1,-11}    " +
                "   {2,10}       {3,5} ", "Darbo nr.","Darbo užduoties nr.","Dim.kodas","Dim.vertės kodas"));
            richTextBox1.AppendText(Environment.NewLine);
            foreach (var row in darboUzdDimTable)
            {
                if (row.darbo_nr == ieskomas_nr)
                {
                    richTextBox1.AppendText(row.ToString());
                    richTextBox1.AppendText(Environment.NewLine);
                }
            }
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText("DARBŲ PLANAVIMO EILUTĖ");
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.AppendText(String.Format("{0,-16}  {1,-11}    {2,-28}  {3,-30}   {4,-10}  {5,3} " +
                " {6,12}  {7,5}  {8,18}  {9,10}   {10,15} " +
            " {11,25}", "Eilutės nr.","Darbo nr.","Darbo užduoties nr.","Planavimo data","Dok nr.","Tipas","Nr.",
            "Aprašas","Kiekis","Vnt.savikaina",	"Vnt.kaina","Mat.vnt.kodas"));
            richTextBox1.AppendText(Environment.NewLine);
            foreach (var row in darbuPlanEilTable)
            {
                if (row.darbo_nr == ieskomas_nr)
                {
                    richTextBox1.AppendText(row.ToString());
                    richTextBox1.AppendText(Environment.NewLine);

                    //pridedamos islaidos prie islaidu sumos
                    bendros_islaidos += row.vieneto_savikaina;
                    //pridedamos pajamos prie pajamu sumos
                    bendros_pajamos += row.vieneto_kaina;
                }
            }
            //apskaiciuojamas uzdarbis: pajamos - islaidos
            bendras_uzdarbis = bendros_pajamos - bendros_islaidos;

            textBox2.AppendText(bendros_islaidos.ToString());
            textBox3.AppendText(bendros_pajamos.ToString());
            textBox4.AppendText(bendras_uzdarbis.ToString());
        }
    }
}
