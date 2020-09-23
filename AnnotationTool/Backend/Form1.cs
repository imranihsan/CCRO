using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Data.OleDb;
using SharpEntropy;
using java.util;
using java.io;
using edu.stanford.nlp.pipeline;
using edu.stanford.nlp.parser.lexparser;
using LemmaSharp;
using edu.stanford.nlp.trees;
using edu.stanford.nlp.ling;
using edu.stanford.nlp.process;
using edu.stanford.nlp.tagger.maxent;
using OpenNLP;
using Microsoft.CSharp;
using Microsoft.Office.Interop;
using System.Text.RegularExpressions;
using MoreLinq;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;




namespace GraphRepresentation
{
    public partial class Form1 :MetroFramework.Forms.MetroForm
    {

        public Form1()
        {
            InitializeComponent();
        }
        List<string> list1 = new List<string>();
        List<string> list2 = new List<string>();
        List<string> list3 = new List<string>();
        List<string> list4 = new List<string>();
        List<string> list5 = new List<string>();
        List<string> list6 = new List<string>();
        List<string> list7 = new List<string>();
        List<string> list8 = new List<string>();

        private void Form1_Load(object sender, EventArgs e)
        {
          
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            ExcelPackage ep = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws = ep.Workbook.Worksheets["Sheet1"];

            for (int rw = 1; rw <= ws.Dimension.End.Row; rw++)
            {
                if (ws.Cells[rw, 1].Value != null)
                {
                    list1.Add(ws.Cells[rw, 1].Value.ToString());
                }
            }
            ExcelPackage ep2 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws2 = ep2.Workbook.Worksheets["Sheet2"];

            for (int rw2 = 1; rw2 <= ws2.Dimension.End.Row; rw2++)
            {
                if (ws2.Cells[rw2, 1].Value != null)
                {
                    list2.Add(ws2.Cells[rw2, 1].Value.ToString());
                }
            }
            ExcelPackage ep3 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws3 = ep3.Workbook.Worksheets["Sheet3"];

            for (int rw3 = 1; rw3 <= ws3.Dimension.End.Row; rw3++)
            {
                if (ws3.Cells[rw3, 1].Value != null)
                {
                    list3.Add(ws3.Cells[rw3, 1].Value.ToString());
                }
            }
            ExcelPackage ep4 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws4 = ep4.Workbook.Worksheets["Sheet4"];

            for (int rw4 = 1; rw4 <= ws4.Dimension.End.Row; rw4++)
            {
                if (ws4.Cells[rw4, 1].Value != null)
                {
                    list4.Add(ws4.Cells[rw4, 1].Value.ToString());
                }
            }
            ExcelPackage ep5 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws5 = ep5.Workbook.Worksheets["Sheet5"];

            for (int rw5 = 1; rw5 <= ws5.Dimension.End.Row; rw5++)
            {
                if (ws5.Cells[rw5, 1].Value != null)
                {
                    list5.Add(ws5.Cells[rw5, 1].Value.ToString());
                }
            }
            ExcelPackage ep6 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws6 = ep6.Workbook.Worksheets["Sheet6"];

            for (int rw6 = 1; rw6 <= ws6.Dimension.End.Row; rw6++)
            {
                if (ws6.Cells[rw6, 1].Value != null)
                {
                    list6.Add(ws6.Cells[rw6, 1].Value.ToString());
                }
            }
            ExcelPackage ep7 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws7 = ep7.Workbook.Worksheets["Sheet7"];

            for (int rw7 = 1; rw7 <= ws7.Dimension.End.Row; rw7++)
            {
                if (ws.Cells[rw7, 1].Value != null)
                {
                    list7.Add(ws7.Cells[rw7, 1].Value.ToString());
                }
            }
            ExcelPackage ep8 = new ExcelPackage(new FileInfo(@"E:\WorkCitations.xlsx"));
            ExcelWorksheet ws8 = ep8.Workbook.Worksheets["Sheet8"];

            for (int rw8= 1; rw8<= ws8.Dimension.End.Row; rw8++)
            {
                if (ws8.Cells[rw8, 1].Value != null)
                {
                    list1.Add(ws8.Cells[rw8, 1].Value.ToString());
                }
            }

            

        }
        public static string RemoveSpecialCharacter(string str)
        {
            return Regex.Replace(str, @"[^a-zA-Z0-9 -]", string.Empty);
        }
        string Citing;
        string Cited;
        public static string cite = "";
        public static string citing = "";
        public static string citation = "";
        public static string lematize = "";
        int repitions = 1;
       

       
        List<ContentVerb> showlist = new List<ContentVerb>();
        List<FinalContent> flist = new List<FinalContent>();

      
        
        public string send2(string text)
        {
            string[] exampleWords = text.Split(
                new char[] { ' ', ',', '.', ')', '(' }, StringSplitOptions.RemoveEmptyEntries);

            ILemmatizer lmtz = new LemmatizerPrebuiltCompact(LemmaSharp.LanguagePrebuilt.English);

            StringBuilder sb = new StringBuilder();
            foreach (string word in exampleWords)
            {
                sb.Append(LemmatizeOne(lmtz, word) + " ");
            }
            return sb.ToString();
           
        }
        public string send(string text)
        {
            string[] exampleWords = text.Split(
                new char[] { ' ', ',', '.', ')', '(' }, StringSplitOptions.RemoveEmptyEntries);

            ILemmatizer lmtz = new LemmatizerPrebuiltCompact(LemmaSharp.LanguagePrebuilt.English);

            StringBuilder sb = new StringBuilder();
            foreach (string word in exampleWords)
            {
                sb.Append(LemmatizeOne(lmtz, word) + " ");
            }
         
            string finalstring = sb.ToString();


            var jarRoot = @"E:\stanford-postagger-full-2015-12-09\stanford-postagger-full-2015-12-09";
            var modelsDirectory = jarRoot + @"\models";
            // Loading POS Tagger
            var tagger = new MaxentTagger(modelsDirectory + @"\wsj-0-18-bidirectional-nodistsim.tagger");

            // Text for tagging
            StringBuilder str = new StringBuilder();

            var sentences = MaxentTagger.tokenizeText(new java.io.StringReader(finalstring)).toArray();
            foreach (ArrayList sentence in sentences)
            {
                var taggedSentence = tagger.tagSentence(sentence);
                string sent = SentenceUtils.listToString(taggedSentence, false);

                String[] tokens = sent.Split(' ');
                for (int i = 0; i < tokens.Length; i++)
                {

                    if (tokens[i].Contains("/VB"))
                    {
                        str.Append(tokens[i] + " ");
                    }
                }
                

            }
            return str.ToString();
          
        }

        public string LemmatizeOne(LemmaSharp.ILemmatizer lmtz, string word)
        {
            string wordLower = word.ToLower();
            string lemma = lmtz.Lemmatize(wordLower);
            return lemma;
        }
        

        private void metroTile1_Click(object sender, EventArgs e)
        {
            flist.Clear();
            string senti="";
            string ByAllx = RemoveSpecialCharacter(textBox4.Text);
            Process q = new Process();

            q.StartInfo.FileName = @"C:\\Program Files (x86)\\Invoke Engine\\CSAT\\exe\\test.exe";
            q.StartInfo.Arguments = "\"" + ByAllx + "\"";
            q.StartInfo.RedirectStandardOutput = true;
            q.StartInfo.UseShellExecute = false;
            q.StartInfo.RedirectStandardOutput = true;
            q.StartInfo.CreateNoWindow = true;

            q.Start();

            using (StreamReader reader = q.StandardOutput)
            {


                string sentiment = reader.ReadToEnd();
                string replacement = Regex.Replace(sentiment, @"\t|\n|\r", "");
                double sent = double.Parse(replacement);
                //     MessageBox.Show(sent.ToString());
                if (sent < -0.3)
                {
                    senti= "n";
                }
                if (sent > 0.3)
                {
                    senti = "p";
                }
                if (sent >= -0.3 && sent <= 0.3)
                {
                   senti = "o";
                }


            }

           
                

               
                Citing = textBox1.Text.Replace("-", "");
                Cited = textBox2.Text.Replace("-", "");
         
                string value=send(textBox4.Text);
               
                Char delimiter = ' ';
                String[] substrings = value.Split(delimiter);
                String AllowedChars2 = @"^[a-zA-Z]{3,}";
                
                

                foreach (var substring in substrings)
                {

                    if (Regex.IsMatch(substring, AllowedChars2))
                    {
                        ContentVerb verbdata = new ContentVerb();
                        int indexOfSteam = substring.IndexOf("/");
                        if (indexOfSteam >= 0)
                        {
                            string subs = substring.Remove(indexOfSteam);
                            verbdata.content = subs;
                        }
                        
                        verbdata.repition = repitions;
                        showlist.Add(verbdata);
                    }
                }
                int repeat = showlist.Count;
                for (int pass = 0; pass < showlist.Count; pass++)
                {
                    for (int i = pass + 1; i < showlist.Count; i++)
                    {
                        if (showlist[pass].content == showlist[i].content)
                        {
                            showlist[i].repition = showlist[pass].repition + 1;
                            showlist.RemoveAt(pass);


                        }
                    }
                }
                
                for (int p=0;p<showlist.Count;p++)
                {
                    FinalContent f = new FinalContent();
                    f.content = showlist[p].content;
                    f.repition = showlist[p].repition;
                    flist.Add(f);
                    
                }
                var bindingList = new BindingList<FinalContent>(flist);
                var source = new BindingSource(bindingList, null);
                dataGridView2.DataSource = source;
               
                for (int x = 0; x < list1.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list1[x])
                        {
                            showlist[y].type = "Using_the_Work";
                            showlist[y].typeno2 = "p";
                            showlist[y].piority = 1;

                            

                        }

                    }
                }
                for (int x = 0; x < list2.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list2[x])
                        {
                            showlist[y].type = "Extending_the_work";
                            showlist[y].typeno2 = "p";
                            showlist[y].piority = 2;
                            


                        }

                    }
                }
                for (int x = 0; x < list3.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list3[x])
                        {
                            showlist[y].type = "Based_On";
                            showlist[y].typeno2 = "p";
                            showlist[y].piority=3;


                        }

                    }
                }
                for (int x = 0; x < list4.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list4[x])
                        {
                            showlist[y].type = "Disagree_With_the_Work";
                            showlist[y].typeno2 = "n";
                            showlist[y].piority = 2;


                        }

                    }
                }
                for (int x = 0; x < list5.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list5[x])
                        {
                            showlist[y].type = "Comparison_and_Contrast";
                            showlist[y].typeno2 = "n";
                            showlist[y].piority = 1;



                        }

                    }
                }
                for (int x = 0; x < list6.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list6[x])
                        {
                            showlist[y].type = "Criticizing_the_Work";
                            showlist[y].typeno2 = "n";
                            showlist[y].piority = 3;



                        }

                    }
                }
                for (int x = 0; x < list7.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list7[x])
                        {
                            showlist[y].type = "Related_Work";
                            showlist[y].typeno2 = "o";
                            showlist[y].piority = 1;



                        }

                    }
                }
                for (int x = 0; x < list8.Count; x++)
                {
                    for (int y = 0; y < showlist.Count; y++)
                    {

                        if (showlist[y].content == list8[x])
                        {
                            showlist[y].type = "Background";
                            showlist[y].typeno2 = "o";
                            showlist[y].piority = 2;



                        }

                    }
                }
                
                

                
                
                string path = @"C:\Users\lenovo\Desktop\FYP\FYP\GraphRepresentation\GraphRepresentation\bin\Debug\Example.dot";
                System.IO.File.Create(path).Dispose();
                using (TextWriter tw = new StreamWriter(path))
                {
                    if (textBox3.Text == "n")
                    {
                        tw.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + textBox3.Text + " color=red];}");
                        tw.Close();
                    }
                    if (textBox3.Text == "o")
                    {
                        tw.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + textBox3.Text + " color=blue];}");
                        tw.Close();
                    }
                    if (textBox3.Text == "p")
                    {
                        tw.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + textBox3.Text + " color=green];}");
                        tw.Close();
                    }
                }
            

           

                
                
            ///////////////////////////////////////////////////////////////////

                string path2 = @"C:\Users\lenovo\Desktop\FYP\FYP\GraphRepresentation\GraphRepresentation\bin\Debug\Example2.dot";
                System.IO.File.Create(path2).Dispose();
                using (TextWriter tw2 = new StreamWriter(path2))
                {
                    if (senti == "n")
                    {
                        tw2.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + senti + " color=red];}");
                        tw2.Close();
                    }
                    if (senti == "o")
                    {
                        tw2.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + senti + " color=blue];}");
                        tw2.Close();
                    }
                    if (senti == "p")
                    {
                        tw2.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + senti + " color=green];}");
                        tw2.Close();
                    }
                }


                ////////////////////////////////////////////////////
                string path3 = @"C:\Users\lenovo\Desktop\FYP\FYP\GraphRepresentation\GraphRepresentation\bin\Debug\Example3.dot";
                System.IO.File.Create(path3).Dispose();
                using (TextWriter tw3 = new StreamWriter(path3))
                {
                    if (textBox3.Text == "n")
                    {
                        tw3.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + "Comparison_and_Contrast" + " color=red];}");
                        tw3.Close();
                    }
                    if (textBox3.Text == "o")
                    {
                        tw3.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + "Related_Work" + " color=blue];}");
                        tw3.Close();
                    }
                    if (textBox3.Text == "p")
                    {
                        tw3.WriteLine("digraph G { " + Citing + " -> " + Cited + "[label=_" + "Using_the_Work" + " color=green];}");
                        tw3.Close();
                    }
                }
               
               
               ///////////////////////////////////////////////////////
               
                Process cmd = new Process();
                Process cmd2 = new Process(); 
                Process cmd3 = new Process();
               

                System.Diagnostics.Process.Start("run.exe");
                System.Diagnostics.Process.Start("run2.exe");
                System.Diagnostics.Process.Start("run3.exe");
                System.Threading.Thread.Sleep(2000);
                pictureBox1.ImageLocation = @"C:\Users\lenovo\Desktop\FYP\FYP\GraphRepresentation\GraphRepresentation\bin\Debug\graphname2.png";
                pictureBox2.ImageLocation = @"C:\Users\lenovo\Desktop\FYP\FYP\GraphRepresentation\GraphRepresentation\bin\Debug\graphname3.png";
                pictureBox3.ImageLocation = @"C:\Users\lenovo\Desktop\FYP\FYP\GraphRepresentation\GraphRepresentation\bin\Debug\graphname4.png";
                
                cmd.Close();
                cmd2.Close();
                cmd3.Close();
                showlist.Clear();
            
          
          
        }

        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private void metroTile2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files(*.xls)|*.xlsx";
            openFileDialog1.FileName = "";
            openFileDialog1.Multiselect = false;
            openFileDialog1.ShowDialog();
           
          
           
                   
          

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
        
            string conStr, sheetName;

           

                
             conStr = string.Format(Excel07ConString, filePath, 1);
                  
           

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    System.Data.DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        System.Data.DataTable dt = new System.Data.DataTable();
                        cmd.CommandText = "SELECT * FROM [Sheet1$]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();

                        //Populate DataGridView.
                       dataGridView1.DataSource = dt;
                       this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                       this.dataGridView1.Columns[2].Visible = false;
                    }
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
          
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView1.DataSource != null)
            {
                
                DataGridViewRow dr = dataGridView1.SelectedRows[0];
              
                textBox1.Text = dr.Cells[0].Value.ToString();
                textBox2.Text = dr.Cells[1].Value.ToString();
                textBox3.Text = dr.Cells[2].Value.ToString();
                textBox4.Text = dr.Cells[3].Value.ToString();
                textBox5.Text = send2(textBox4.Text);
            }
            
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
            pictureBox2.Image = null;
            pictureBox3.Image = null;
            dataGridView2.DataSource = null;
            flist.Clear();
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void metroTile4_Click(object sender, EventArgs e)
        {
            about a = new about();
            a.Show();
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {

            cite = textBox1.Text;
            citing = textBox2.Text;
            citation = textBox4.Text;
            lematize = textBox5.Text;
            View a= new View();
           
            a.Show();
            
        }

     

    }

}
class ContentVerb
{

    private string _content;
    private int _repition;
    private string _type;
    private string _typeno2;
    private int _piority;
  
    public string content
    {
        get { return _content; }
        set { _content = value; }
    }
    public int repition
    {
        get { return _repition; }
        set { _repition = value; }
    }
    public string type
    {
        get { return _type; }
        set { _type = value; }
    }
    public string typeno2
    {
        get { return _typeno2; }
        set { _typeno2 = value; }
    }
    public int piority
    {
        get { return _piority; }
        set { _piority = value; }
    }
   
}
class FinalContent
{

    private string _content;
    private int _repition;
   
    public string content
    {
        get { return _content; }
        set { _content = value; }
    }
    public int repition
    {
        get { return _repition; }
        set { _repition = value; }
    }
    
}
