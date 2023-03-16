using Aspose.Cells;
using Aspose.Words;
using Aspose.Words.Replacing;
using Microsoft.VisualBasic;
using System;
using System.Windows.Forms;
using System.Reflection.Metadata;
using System.IO;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using Document = Aspose.Words.Document;
using System.Diagnostics;
using Aspose.Cells.Drawing;

namespace DocsRead
{
    public partial class Form1 : Form
    {
        int ij = 0;
        string text = "";
        string text1 = "";
        string text2 = "",text3="",text4="",text5="",text6="",text7="",text8="";

        public Form1()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox11.Text = fbd.SelectedPath;
                    string[] files = Directory.GetFiles(@fbd.SelectedPath);
                    for (int i =0; i<files.Length;i++) { 
                        FileInfo file = new FileInfo(files[i]);
                        comboBox3.Items.Add(file.Name);
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox3.SelectedIndex = 0;
            string pat = textBox11.Text+"\\"+comboBox3.SelectedItem;
            /*if (comboBox3.SelectedItem == "Bonafide Letter")
            {
                pat = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\DEM.docx";
            }
            if (comboBox3.SelectedItem == "Demand Letter")
            {
                pat = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\DEM1.docx";
            }
            if (comboBox3.SelectedItem == "Letter Head")
            {
                pat = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\DEM2.docx";
            }*/
            var doc = new Document(pat);
            string allText = doc.ToString(Aspose.Words.SaveFormat.Text);
            label2.Text = allText;
            comboBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox2.AutoCompleteMode = AutoCompleteMode.Suggest;
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox3.AutoCompleteMode = AutoCompleteMode.Suggest;
            comboBox3.AutoCompleteSource = AutoCompleteSource.ListItems;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            ij += 1;
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox12.Text = fbd.SelectedPath;
                }
            }

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox13.Text = fbd.SelectedPath;
                }
            }
            readword();
        }
        private void readword(){
            Aspose.Words.License license = new Aspose.Words.License();
            license.SetLicense("C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\bin\\Debug\\net7.0-windows\\Aspose.Total.Product.Family.lic");
            var dateTime = DateTime.Now;
            text = textBox1.Text;
            text1 = textBox2.Text;
            text2 = textBox3.Text;
            text3 = textBox4.Text;
            text4 = textBox5.Text;
            text5 = textBox6.Text;
            text6 = textBox7.Text;
            text7 = textBox8.Text;
            text8 = textBox9.Text;
            int st = 0, gt=0, gef=0, gtf=0, ghf=0, st1=0;
            int y = 0;
            var i1 = "";
            var i2 = "";
            string i = "", it="";
            string pat = textBox11.Text + "\\" + comboBox3.SelectedItem;
            string pa1 = comboBox3.Text.Remove(comboBox3.Text.Length-5,5);
            string pat1 = textBox12.Text + "\\";
            string pat2 = textBox13.Text + "\\";
            string selected = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
            string lll = this.comboBox3.GetItemText(this.comboBox3.SelectedItem);
            var longDateValue = dateTime.ToLongDateString();
            var longDateValue1 = dateTime.ToString("dd-MM-yyyy");
            var dateValue1 = dateTime.ToString("yyyy");
            /*if (comboBox3.SelectedItem == "Bonafide Letter")
            {
                pat ="C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\DEM.docx";
            }
            if (comboBox3.SelectedItem == "Demand Letter")
            {
                pat="C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\DEM1.docx";
            }
            if (comboBox3.SelectedItem == "Letter Head")
            {
                pat="C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\DEM2.docx";
            }*/
            var doc = new Document(pat);
            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = true;
            options.Direction = FindReplaceDirection.Forward;
            options.FindWholeWordsOnly = true;
            doc.Range.Replace("@reg", text, new FindReplaceOptions(FindReplaceDirection.Forward));
            doc.Range.Replace("@name", text1, options);
            doc.Range.Replace("@fname", text2, options);
            doc.Range.Replace("@date", longDateValue, new FindReplaceOptions(FindReplaceDirection.Forward));
            doc.Range.Replace("@1date", longDateValue1, new FindReplaceOptions(FindReplaceDirection.Forward));
            doc.Range.Replace("@year", dateValue1, new FindReplaceOptions(FindReplaceDirection.Forward));
            doc.Range.Replace("@dep", selected, options);
            doc.Range.Replace("@gadf", text3, options);
            doc.Range.Replace("@gecf", text4, options);
            if (selected.Contains("M.Tech") || selected.Contains("MBA") || selected.Contains("MCA") || selected.Contains("MPT"))
            {
                i = "2"; it="2";
            }
            else if (selected.Contains("B.Tech") || selected.Contains("BPT"))
            {
                i = "4"; it="4";
            }
            else if (selected.Contains("BBA") || selected.Contains("BCA"))
            {
                i = "3";it="3";
            }
            //else if(selected.Contains("BPT")) {
            //i = "4.5";it="4.5" ;
            //}
            if (selected.Contains("M.Tech"))
            {
                if (selected.Contains(" CE")) { doc.Range.Replace("@degree", "M.Tech Civil Engineering", options); }
                else if (selected.Contains("CSE (")) { doc.Range.Replace("@degree", "M.Tech CSE (AI+ML)", options); }
                else if (selected.Contains("CSE")) { doc.Range.Replace("@degree", "M.Tech Computer Science & Engineering", options); }
                else if (selected.Contains("ME")) { doc.Range.Replace("@degree", "M.Tech Mechanical Engineering", options); }
                else if (selected.Contains("EE")) { doc.Range.Replace("@degree", "M.Tech Electrical Engineering", options); }
                if (selected.Contains("ECE")) { doc.Range.Replace("@degree", "M.Tech Electronics & Communication Engineering", options); }
            }
            else if (selected.Contains("B.Tech"))
            {
                if (selected.Contains(" CE")) { doc.Range.Replace("@degree", "B.Tech Civil Engineering", options); }
                else if (selected.Contains("CSE (")) { doc.Range.Replace("@degree", "B.Tech CSE (AI+ML)", options); }
                else if (selected.Contains("CSE")) { doc.Range.Replace("@degree", "B.Tech Computer Science & Engineering", options); }
                else if (selected.Contains("ME")) { doc.Range.Replace("@degree", "B.Tech Mechanical Engineering", options); }
                else if (selected.Contains("EE")) { doc.Range.Replace("@degree", "B.Tech Electrical Engineering", options); }
                if (selected.Contains("ECE")) { doc.Range.Replace("@degree", "B.Tech Electronics & Communication Engineering", options); }
            }
            else if (selected.Contains("MCA")) { doc.Range.Replace("@degree", "MCA Master of Computer Application", options); }
            else if (selected.Contains("BCA")){ doc.Range.Replace("@degree", "BCA Bachelors's in Computer Application", options); }
            else if (selected.Contains("MPT")){ doc.Range.Replace("@degree", "MPT Master of Physiotherapy", options); }
            else if (selected.Contains("BPT")){ doc.Range.Replace("@degree", "BPT Bachelor of Physiotherapy", options); }
            else if (selected.Contains("MBA")) { doc.Range.Replace("@degree", "MBA Master of Business Administration", options); }
            else if (selected.Contains("BBA")) { doc.Range.Replace("@degree", "BBA Bachelor of Business Administration", options); }
            string selected1 = this.comboBox2.GetItemText(this.comboBox2.SelectedItem);
                if (selected1.Contains("1st")) { y = 4; }
                else if (selected1.Contains("2nd")) { y = 3; }
                else if (selected1.Contains("3rd")) { y = 2; }
                else if (selected1.Contains("4th")) { y = 1; }
                doc.Range.Replace("@ady", selected1, options);
                if (i == "2")
                {
                    if (y == 3)
                    {
                        it = "1";
                        doc.Range.Replace("@2adf", text3, options);
                        doc.Range.Replace("@2ecf", text4, options);
                        doc.Range.Replace("@1adf", " 0 ", options);
                        doc.Range.Replace("@1ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", "  ", options);
                        doc.Range.Replace("@3ecf", "  ", options);
                        doc.Range.Replace("@4adf", "  ", options);
                        doc.Range.Replace("@4ecf", "  ", options);
                        i1 = dateTime.AddDays(-365).ToString("yyyy");
                        i2 = dateTime.AddDays(365).ToString("yyyy");
                        doc.Range.Replace("@1ef", "  ", options);
                        doc.Range.Replace("@2ef", text5, options);
                        doc.Range.Replace("@3ef", "  ", options);
                        doc.Range.Replace("@4ef", "  ", options);

                        doc.Range.Replace("@1tf", "  ", options);
                        doc.Range.Replace("@2tf", text6, options);
                        doc.Range.Replace("@3tf", "  ", options);
                        doc.Range.Replace("@4tf", "  ", options);

                        doc.Range.Replace("@1hf", "  ", options);
                        doc.Range.Replace("@2hf", text7, options);
                        doc.Range.Replace("@3hf", "  ", options);
                        doc.Range.Replace("@4hf", "  ", options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 2 * int.Parse(text5);
                        gtf = 2 * int.Parse(text6);
                        ghf = 2 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@2st", st1.ToString(), options);
                        doc.Range.Replace("@st11", "  ", options);
                        doc.Range.Replace("@1st", "  ", options);
                        doc.Range.Replace("@st33", "  ", options);
                        doc.Range.Replace("@3st", "  ", options);
                        doc.Range.Replace("@st44", "  ", options);
                        doc.Range.Replace("@4st", "  ", options);
                        doc.Range.Replace("@st22", st.ToString(), options);
                    }
                    else
                    {
                        doc.Range.Replace("@1adf", text3, options);
                        doc.Range.Replace("@1ecf", text4, options);
                        doc.Range.Replace("@2adf", " 0 ", options);
                        doc.Range.Replace("@2ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", "  ", options);
                        doc.Range.Replace("@3ecf", "  ", options);
                        doc.Range.Replace("@4adf", "  ", options);
                        doc.Range.Replace("@4ecf", "  ", options);
                        i1 = dateValue1;
                        i2 = dateTime.AddDays(2 * 365).ToString("yyyy");
                        doc.Range.Replace("@3ef", "  ", options);
                        doc.Range.Replace("@2ef", text5, options);
                        doc.Range.Replace("@1ef", text5, options);
                        doc.Range.Replace("@4ef", "  ", options);

                        doc.Range.Replace("@3tf", "  ", options);
                        doc.Range.Replace("@2tf", text6, options);
                        doc.Range.Replace("@1tf", text6, options);
                        doc.Range.Replace("@4tf", "  ", options);

                        doc.Range.Replace("@3hf", "  ", options);
                        doc.Range.Replace("@2hf", text7, options);
                        doc.Range.Replace("@1hf", text7, options);
                        doc.Range.Replace("@4hf", "  ", options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 4 * int.Parse(text5);
                        gtf = 4 * int.Parse(text6);
                        ghf = 4 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@1st", st1.ToString(), options);
                        doc.Range.Replace("@st11", st.ToString(), options);
                        doc.Range.Replace("@2st", st.ToString(), options);
                        doc.Range.Replace("@st22", st.ToString(), options);
                        doc.Range.Replace("@st33", "  ", options);
                        doc.Range.Replace("@3st", "  ", options);
                        doc.Range.Replace("@st44", "  ", options);
                        doc.Range.Replace("@4st", "  ", options);
                    }
                }
                else if (i == "3")
                {
                    if (y == 3)
                    {
                        it = "2";
                        doc.Range.Replace("@2adf", text3, options);
                        doc.Range.Replace("@2ecf", text4, options);
                        doc.Range.Replace("@1adf", " 0 ", options);
                        doc.Range.Replace("@1ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", " 0 ", options);
                        doc.Range.Replace("@3ecf", " 0 ", options);
                        doc.Range.Replace("@4adf", "  ", options);
                        doc.Range.Replace("@4ecf", "  ", options);
                        i1 = dateTime.AddDays(-365).ToString("yyyy");
                        i2 = dateTime.AddDays(2 * 365).ToString("yyyy");
                        doc.Range.Replace("@1ef", "  ", options);
                        doc.Range.Replace("@2ef", text5, options);
                        doc.Range.Replace("@3ef", text5, options);
                        doc.Range.Replace("@4ef", "  ", options);

                        doc.Range.Replace("@1tf", "  ", options);
                        doc.Range.Replace("@2tf", text6, options);
                        doc.Range.Replace("@3tf", text6, options);
                        doc.Range.Replace("@4tf", "  ", options);

                        doc.Range.Replace("@1hf", "  ", options);
                        doc.Range.Replace("@2hf", text7, options);
                        doc.Range.Replace("@3hf", text7, options);
                        doc.Range.Replace("@4hf", "  ", options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 4 * int.Parse(text5);
                        gtf = 4 * int.Parse(text6);
                        ghf = 4 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@2st", st1.ToString(), options);
                        doc.Range.Replace("@st22", st.ToString(), options);
                        doc.Range.Replace("@3st", st.ToString(), options);
                        doc.Range.Replace("@st33", st.ToString(), options);
                        doc.Range.Replace("@st11", "  ", options);
                        doc.Range.Replace("@1st", "  ", options);
                        doc.Range.Replace("@st44", "  ", options);
                        doc.Range.Replace("@4st", "  ", options);
                    }
                    else if (y == 2)
                    {
                        it = "1";
                        doc.Range.Replace("@3adf", text3, options);
                        doc.Range.Replace("@3ecf", text4, options);
                        doc.Range.Replace("@1adf", " 0 ", options);
                        doc.Range.Replace("@1ecf", " 0 ", options);
                        doc.Range.Replace("@2adf", " 0 ", options);
                        doc.Range.Replace("@2ecf", " 0 ", options);
                        doc.Range.Replace("@4adf", "  ", options);
                        doc.Range.Replace("@4ecf", "  ", options);
                        i1 = dateTime.AddDays(-2 * 365).ToString("yyyy");
                        i2 = dateTime.AddDays(365).ToString("yyyy");
                        doc.Range.Replace("@1ef", "  ", options);
                        doc.Range.Replace("@2ef", "  ", options);
                        doc.Range.Replace("@3ef", text5, options);
                        doc.Range.Replace("@4ef", "  ", options);

                        doc.Range.Replace("@1tf", "  ", options);
                        doc.Range.Replace("@2tf", "  ", options);
                        doc.Range.Replace("@3tf", text6, options);
                        doc.Range.Replace("@4tf", "  ", options);

                        doc.Range.Replace("@1hf", "  ", options);
                        doc.Range.Replace("@2hf", "  ", options);
                        doc.Range.Replace("@3hf", text7, options);
                        doc.Range.Replace("@4hf", "  ", options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 2 * int.Parse(text5);
                        gtf = 2 * int.Parse(text6);
                        ghf = 2 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@3st", st1.ToString(), options);
                        doc.Range.Replace("@st33", st.ToString(), options);
                        doc.Range.Replace("@2st", "  ", options);
                        doc.Range.Replace("@st22", "  ", options);
                        doc.Range.Replace("@st11", "  ", options);
                        doc.Range.Replace("@1st", "  ", options);
                        doc.Range.Replace("@st44", "  ", options);
                        doc.Range.Replace("@4st", "  ", options);
                    }
                    else
                    {
                        doc.Range.Replace("@1adf", text3, options);
                        doc.Range.Replace("@1ecf", text4, options);
                        doc.Range.Replace("@2adf", " 0 ", options);
                        doc.Range.Replace("@2ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", " 0 ", options);
                        doc.Range.Replace("@3ecf", " 0 ", options);
                        doc.Range.Replace("@4adf", "  ", options);
                        doc.Range.Replace("@4ecf", "  ", options);
                        i1 = dateValue1;
                        i2 = dateTime.AddDays(3 * 365).ToString("yyyy");
                        doc.Range.Replace("@1ef", text5, options);
                        doc.Range.Replace("@2ef", text5, options);
                        doc.Range.Replace("@3ef", text5, options);
                        doc.Range.Replace("@4ef", "  ", options);

                        doc.Range.Replace("@1tf", text6, options);
                        doc.Range.Replace("@2tf", text6, options);
                        doc.Range.Replace("@3tf", text6, options);
                        doc.Range.Replace("@4tf", "  ", options);

                        doc.Range.Replace("@1hf", text7, options);
                        doc.Range.Replace("@2hf", text7, options);
                        doc.Range.Replace("@3hf", text7, options);
                        doc.Range.Replace("@4hf", "  ", options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 6 * int.Parse(text5);
                        gtf = 6 * int.Parse(text6);
                        ghf = 6 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@1st", st1.ToString(), options);
                        doc.Range.Replace("@st11", st.ToString(), options);
                        doc.Range.Replace("@2st", st.ToString(), options);
                        doc.Range.Replace("@st22", st.ToString(), options);
                        doc.Range.Replace("@st33", st.ToString(), options);
                        doc.Range.Replace("@3st", st.ToString(), options);
                        doc.Range.Replace("@st44", "  ", options);
                        doc.Range.Replace("@4st", "  ", options);
                    }
                }
                else if (i == "4") // || i == "4.5")
                { if (y == 3)
                    {
                        /*if (i == "4") { */it = "3";
                        //}
                        //else { it = "3.5";
                        //}
                        doc.Range.Replace("@2adf", text3, options);
                        doc.Range.Replace("@2ecf", text4, options);
                        doc.Range.Replace("@1adf", " 0 ", options);
                        doc.Range.Replace("@1ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", " 0 ", options);
                        doc.Range.Replace("@3ecf", " 0 ", options);
                        doc.Range.Replace("@4adf", " 0 ", options);
                        doc.Range.Replace("@4ecf", " 0 ", options);
                        i1 = dateTime.AddDays(-365).ToString("yyyy");
                        i2 = dateTime.AddDays(3 * 365).ToString("yyyy");
                        doc.Range.Replace("@1ef", "  ", options);
                        doc.Range.Replace("@2ef", text5, options);
                        doc.Range.Replace("@3ef", text5, options);
                        doc.Range.Replace("@4ef", text5, options);

                        doc.Range.Replace("@1tf", "  ", options);
                        doc.Range.Replace("@2tf", text6, options);
                        doc.Range.Replace("@3tf", text6, options);
                        doc.Range.Replace("@4tf", text6, options);

                        doc.Range.Replace("@1hf", "  ", options);
                        doc.Range.Replace("@2hf", text7, options);
                        doc.Range.Replace("@3hf", text7, options);
                        doc.Range.Replace("@4hf", text7, options);

                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 6 * int.Parse(text5);
                        gtf = 6 * int.Parse(text6);
                        ghf = 6 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@2st", st1.ToString(), options);
                        doc.Range.Replace("@st33", st.ToString(), options);
                        doc.Range.Replace("@3st", st.ToString(), options);
                        doc.Range.Replace("@st22", st.ToString(), options);
                        doc.Range.Replace("@st11", "  ", options);
                        doc.Range.Replace("@1st", "  ", options);
                        doc.Range.Replace("@st44", st.ToString(), options);
                        doc.Range.Replace("@4st", st.ToString(), options);
                    }
                    else if (y == 2)
                    {
                        /*if (i == "4") {*/ it = "2";
                        //}
                        //else { it = "2.5";
                        //}
                        doc.Range.Replace("@3adf", text3, options);
                        doc.Range.Replace("@3ecf", text4, options);
                        doc.Range.Replace("@1adf", " 0 ", options);
                        doc.Range.Replace("@1ecf", " 0 ", options);
                        doc.Range.Replace("@2adf", " 0 ", options);
                        doc.Range.Replace("@2ecf", " 0 ", options);
                        doc.Range.Replace("@4adf", " 0 ", options);
                        doc.Range.Replace("@4ecf", " 0 ", options);
                        i1 = dateTime.AddDays(-2 * 365).ToString("yyyy");
                        i2 = dateTime.AddDays(2 * 365).ToString("yyyy");
                        doc.Range.Replace("@1ef", "  ", options);
                        doc.Range.Replace("@2ef", "  ", options);
                        doc.Range.Replace("@3ef", text5, options);
                        doc.Range.Replace("@4ef", text5, options);

                        doc.Range.Replace("@1tf", "  ", options);
                        doc.Range.Replace("@2tf", "  ", options);
                        doc.Range.Replace("@3tf", text6, options);
                        doc.Range.Replace("@4tf", text6, options);

                        doc.Range.Replace("@1hf", "  ", options);
                        doc.Range.Replace("@2hf", "  ", options);
                        doc.Range.Replace("@3hf", text7, options);
                        doc.Range.Replace("@4hf", text7, options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 4 * int.Parse(text5);
                        gtf = 4 * int.Parse(text6);
                        ghf = 4 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@3st", st1.ToString(), options);
                        doc.Range.Replace("@st33", st.ToString(), options);
                        doc.Range.Replace("@4st", st.ToString(), options);
                        doc.Range.Replace("@st44", st.ToString(), options);
                        doc.Range.Replace("@st11", "  ", options);
                        doc.Range.Replace("@1st", "  ", options);
                        doc.Range.Replace("@st22", "  ", options);
                        doc.Range.Replace("@2st", "  ", options);
                    }
                    else if (y == 1)
                    {
                        /*if (i == "4") {*/ it = "1";
                        //}
                        //else { it = "1.5";
                        //}
                        doc.Range.Replace("@4adf", text3, options);
                        doc.Range.Replace("@4ecf", text4, options);
                        doc.Range.Replace("@1adf", " 0 ", options);
                        doc.Range.Replace("@1ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", " 0 ", options);
                        doc.Range.Replace("@3ecf", " 0 ", options);
                        doc.Range.Replace("@2adf", " 0 ", options);
                        doc.Range.Replace("@2ecf", " 0 ", options);
                        i1 = dateTime.AddDays(-3 * 365).ToString("yyyy");
                        i2 = dateTime.AddDays(365).ToString("yyyy");
                        doc.Range.Replace("@1ef", "  ", options);
                        doc.Range.Replace("@2ef", "  ", options);
                        doc.Range.Replace("@3ef", "  ", options);
                        doc.Range.Replace("@4ef", text5, options);

                        doc.Range.Replace("@1tf", "  ", options);
                        doc.Range.Replace("@2tf", "  ", options);
                        doc.Range.Replace("@3tf", "  ", options);
                        doc.Range.Replace("@4tf", text6, options);

                        doc.Range.Replace("@1hf", "  ", options);
                        doc.Range.Replace("@2hf", "  ", options);
                        doc.Range.Replace("@3hf", "  ", options);
                        doc.Range.Replace("@4hf", text7, options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 2 * int.Parse(text5);
                        gtf = 2 * int.Parse(text6);
                        ghf = 2 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@4st", st1.ToString(), options);
                        doc.Range.Replace("@st44", st.ToString(), options);
                        doc.Range.Replace("@2st", "  ", options);
                        doc.Range.Replace("@st22", "  ", options);
                        doc.Range.Replace("@st33", "  ", options);
                        doc.Range.Replace("@3st", "  ", options);
                        doc.Range.Replace("@st11", "  ", options);
                        doc.Range.Replace("@1st", "  ", options);
                    }
                    else
                    {
                        doc.Range.Replace("@1adf", text3, options);
                        doc.Range.Replace("@1ecf", text4, options);
                        doc.Range.Replace("@2adf", " 0 ", options);
                        doc.Range.Replace("@2ecf", " 0 ", options);
                        doc.Range.Replace("@3adf", " 0 ", options);
                        doc.Range.Replace("@3ecf", " 0 ", options);
                        doc.Range.Replace("@4adf", " 0 ", options);
                        doc.Range.Replace("@4ecf", " 0 ", options);
                        i1 = dateValue1;
                        i2 = dateTime.AddDays(4 * 365).ToString("yyyy");
                        doc.Range.Replace("@1ef", text5, options);
                        doc.Range.Replace("@2ef", text5, options);
                        doc.Range.Replace("@3ef", text5, options);
                        doc.Range.Replace("@4ef", text5, options);

                        doc.Range.Replace("@1tf", text6, options);
                        doc.Range.Replace("@2tf", text6, options);
                        doc.Range.Replace("@3tf", text6, options);
                        doc.Range.Replace("@4tf", text6, options);

                        doc.Range.Replace("@1hf", text7, options);
                        doc.Range.Replace("@2hf", text7, options);
                        doc.Range.Replace("@3hf", text7, options);
                        doc.Range.Replace("@4hf", text7, options);
                        st1 = int.Parse(text3) + int.Parse(text4) + int.Parse(text5) + int.Parse(text6) + int.Parse(text7) + int.Parse(text8);
                        st = int.Parse(text5) + int.Parse(text6) + int.Parse(text7);
                        gef = 8 * int.Parse(text5);
                        gtf = 8 * int.Parse(text6);
                        ghf = 8 * int.Parse(text7);
                        gt = gef + gtf + ghf + int.Parse(text3) + int.Parse(text4) + int.Parse(text8);
                        doc.Range.Replace("@1st", st1.ToString(), options);
                        doc.Range.Replace("@st11", st.ToString(), options);
                        doc.Range.Replace("@2st", st.ToString(), options);
                        doc.Range.Replace("@st22", st.ToString(), options);
                        doc.Range.Replace("@st33", st.ToString(), options);
                        doc.Range.Replace("@3st", st.ToString(), options);
                        doc.Range.Replace("@st44", st.ToString(), options);
                        doc.Range.Replace("@4st", st.ToString(), options);
                    }
                }
                /*if (int.Parse(text6) < 50000) { */doc.Range.Replace("@type", "provisional", options); //}
                //else { doc.Range.Replace("@type", "final", options); }
                doc.Range.Replace("@uf", text8, options);
                doc.Range.Replace("@tf", text6, options);
                doc.Range.Replace("@hf", text7, options);
                doc.Range.Replace("@gef", gef.ToString(), options);
                doc.Range.Replace("@gtf", gtf.ToString(), options);
                doc.Range.Replace("@ghf", ghf.ToString(), options);
                doc.Range.Replace("@gt", gt.ToString(), options);
                doc.Range.Replace("@syear", i1, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("@lyear", i2, new FindReplaceOptions(FindReplaceDirection.Forward));
                doc.Range.Replace("@nyears", i, options);
                doc.Range.Replace("@tyears", it, options);
                string allText = doc.ToString(Aspose.Words.SaveFormat.Text);
                label2.Text = allText;
                //if (comboBox3.SelectedItem == "Bonafide Letter")
                //{
                if (selected.Contains("M.Tech") || selected.Contains("MBA") || selected.Contains("MCA") || selected.Contains("MPT"))
                    {
                //i = "2"; it = "2";
                    if (lll.Contains("Letter 2") || lll.Contains("Bonafide Letter") || lll.Contains("Letter Head"))
                    {
                    pat1 += "RJ" + dateValue1 + text + " "+ comboBox3.SelectedItem;
                        pat2 += "RJ" + dateValue1 + text + " "+pa1+".pdf";
                        doc.Save(pat1);
                        doc.Save(pat2);
                    }
                    }
                else if (selected.Contains("B.Tech") || selected.Contains("BPT"))
                    {
                        //i = "4"; it = "4";
                        if (lll.Contains("Letter 4") || lll.Contains("Bonafide Letter") || lll.Contains("Letter Head"))
                        {
                        pat1 += "RJ" + dateValue1 + text + " "+ comboBox3.SelectedItem;
                        pat2 += "RJ" + dateValue1 + text + " "+pa1+".pdf";
                        doc.Save(pat1);
                        doc.Save(pat2);
                        }
                    }
                else if (selected.Contains("BBA") || selected.Contains("BCA"))
                    {
                        //i = "3"; it = "3";
                        if (lll.Contains("Letter 3") || lll.Contains("Bonafide Letter") || lll.Contains("Letter Head"))
                        {
                        pat1 += "RJ" + dateValue1 + text + " "+ comboBox3.SelectedItem;
                        pat2 += "RJ" + dateValue1 + text + " "+pa1+".pdf";
                        doc.Save(pat1);
                        doc.Save(pat2);
                        }
                    }
                /*else if(selected.Contains("BPT")) {
                i = "4.5";it="4.5" ;
                }*/
                   /* pat1 += "RJ" + dateValue1 + text + " "+ comboBox3.SelectedItem;
                    pat2 += "RJ" + dateValue1 + text + " "+pa1+".pdf"; */
                /*}
                if (comboBox3.SelectedItem == "Demand Letter")
                {
                    pat1 = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\Word Letters\\RJ" + dateValue1 + text + " Demand Letter.docx";
                    pat2 = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\Printable Letters\\RJ" + dateValue1 + text + " Demand Letter.pdf";
                }
                if (comboBox3.SelectedItem == "Letter Head")
                {
                    pat1 = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\Word Letters\\RJ" + dateValue1 + text + " Letter Head.docx";
                    pat2 = "C:\\Users\\karan\\source\\repos\\DocsRead\\DocsRead\\Printable Letters\\RJ" + dateValue1 + text + " Letter Letter.pdf";
                }*/
                
                

            }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            if (comboBox1.SelectedItem == "BCA" || comboBox1.SelectedItem == "BBA")
            {
                comboBox2.Items.Add("1st");
                comboBox2.Items.Add("2nd");
                comboBox2.Items.Add("3rd");
            }
            else if (comboBox1.SelectedItem == "M.Tech CE" || comboBox1.SelectedItem == "MBA" || comboBox1.SelectedItem == "MCA" || comboBox1.SelectedItem == "MPT" || comboBox1.SelectedItem == "M.Tech CSE" || comboBox1.SelectedItem == "M.Tech EE" || comboBox1.SelectedItem == "M.Tech ECE" || comboBox1.SelectedItem == "M.Tech ME")
            {
                comboBox2.Items.Add("1st");
                comboBox2.Items.Add("2nd");
            }
            else
            {
                comboBox2.Items.Add("1st");
                comboBox2.Items.Add("2nd");
                comboBox2.Items.Add("3rd");
                comboBox2.Items.Add("4th");
                
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox10.Text = openFileDialog.FileName;
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Workbook wb = new Workbook(textBox10.Text);
            WorksheetCollection collection = wb.Worksheets;
            Worksheet worksheet = collection[0];
            int rows = worksheet.Cells.MaxDataRow;
            int cols = comboBox3.Items.Count;
            for (int j=1; j<= rows; j++) {
                for (int i=0; i< cols; i++) {
                    textBox1.Text = worksheet.Cells[j, 0].StringValue;
                    textBox2.Text = worksheet.Cells[j, 1].StringValue;
                    textBox3.Text = worksheet.Cells[j, 2].StringValue;
                    comboBox1.Text = worksheet.Cells[j, 3].StringValue;
                    comboBox2.Text = worksheet.Cells[j, 4].StringValue;
                    textBox4.Text = worksheet.Cells[j, 5].StringValue;
                    textBox5.Text = worksheet.Cells[j, 6].StringValue;
                    textBox6.Text = worksheet.Cells[j, 7].StringValue;
                    textBox7.Text = worksheet.Cells[j, 8].StringValue;
                    textBox8.Text = worksheet.Cells[j, 9].StringValue;
                    textBox9.Text = worksheet.Cells[j, 10].StringValue;
                    comboBox3.SelectedIndex = i;
                    string selected = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
                    string lll = this.comboBox3.GetItemText(this.comboBox3.SelectedItem);
                    if (selected.Contains("M.Tech") || selected.Contains("MBA") || selected.Contains("MCA") || selected.Contains("MPT"))
                    {
                        //i = "2"; it = "2";
                        if (lll.Contains("Letter 2") || lll.Contains("Bonafide Letter") || lll.Contains("Letter Head"))
                        {
                            if (ij == 0)
                            {
                                button1.PerformClick();
                            }
                            else
                            {
                                readword();
                            }
                        }
                    }
                    else if (selected.Contains("B.Tech") || selected.Contains("BPT"))
                    {
                        //i = "4"; it = "4";
                        if (lll.Contains("Letter 4") || lll.Contains("Bonafide Letter") || lll.Contains("Letter Head"))
                        {
                            if (ij == 0)
                            {
                                button1.PerformClick();
                            }
                            else
                            {
                                readword();
                            }
                        }
                    }
                    else if (selected.Contains("BBA") || selected.Contains("BCA"))
                    {
                        //i = "3"; it = "3";
                        if (lll.Contains("Letter 3") || lll.Contains("Bonafide Letter") || lll.Contains("Letter Head"))
                        {
                            if (ij == 0)
                            {
                                button1.PerformClick();
                            }
                            else
                            {
                                readword();
                            }
                        }
                    }
                    
                }
            }
        }
    }
}
