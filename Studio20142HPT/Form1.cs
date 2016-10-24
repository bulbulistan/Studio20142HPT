using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;



namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public static string cesta = "";
        public static XmlDocument X = new XmlDocument();
        public static string vystup = "";
        public static string catLetter = "";
        public static string Wcat = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                listBox1.Items.Clear();
                foreach (String subor in openFileDialog1.FileNames)
                {
                    listBox1.Items.Add(subor);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /* beging check if category is selected */
            if (checkBox1.Checked == true)
            {
                catLetter = "D";
            }
            else if (checkBox2.Checked == true)
            {
                catLetter = "S";
            }
            else if (checkBox3.Checked == true)
            {
                catLetter = "M";
            }
            else if (checkBox1.Checked == false && checkBox2.Checked == false && checkBox3.Checked == false)
            {
                catLetter = "";
            }

            /* if no category is selected, display a warning */
            if (catLetter == "")
            {
                MessageBox.Show("Please select the category");
            }

            else
            {
                Wcat = "W" + catLetter;
                foreach (String subor in listBox1.Items)
                {
                    string ext = Path.GetExtension(subor);

                    if (ext == ".csv")
                    {
                        string[] csvlines = File.ReadAllLines(subor);
                        string[] csvitems = csvlines[csvlines.Length - 1].Split(',');

                        //vystup = subor.Replace("csv", "txt");
                        String allvystup = Path.Combine(Path.GetDirectoryName(subor), "WS-allfiles.csv");

                        if (csvitems.Length != 14)
                        {
                            MessageBox.Show("Check if this is a correct WorldServer analysis CSV file");
                        }
                        else
                        {

                            SortedDictionary<string, int> csvwordcounts = new SortedDictionary<string, int>();
                            csvwordcounts.Add(Wcat + "X", Convert.ToInt32(csvitems[3]));
                            csvwordcounts.Add(Wcat + "P", Convert.ToInt32(csvitems[4]) + Convert.ToInt32(csvitems[10]));
                            csvwordcounts.Add(Wcat + "H", Convert.ToInt32(csvitems[5]));

                            //treat medium fuzzy as low fuzzy lebo Shanshan
                            if (checkBox5.Checked == true)
                            {
                                csvwordcounts.Add(Wcat + "L", Convert.ToInt32(csvitems[6]) + Convert.ToInt32(csvitems[7]));
                            }

                            else
                            {
                                csvwordcounts.Add(Wcat + "M", Convert.ToInt32(csvitems[6]));
                                csvwordcounts.Add(Wcat + "L", Convert.ToInt32(csvitems[7]));
                            }
                            //

                            csvwordcounts.Add(Wcat, Convert.ToInt32(csvitems[8]) + Convert.ToInt32(csvitems[9]));

                            string csvcelkom = Studio20142HPT.PutTogether.PutTogetherDict(csvwordcounts);

                            using (StreamWriter writer = new StreamWriter(allvystup, true))
                            {
                                writer.WriteLine(Path.GetFileName(subor) + "," + csvcelkom);
                                //writer.WriteLine("TOTAL" + "|" + csvcelkom);
                            }                       

                            listBox2.Items.Add("File " + allvystup + " saved successfully");

                        }                        
                    }

                    else if (ext == ".xml")
                    {
                        try
                        {
                            /* load the xml */
                            X.Load(subor);

                            /* prepare the output file, one per each selected file */
                            /* first get the project name from the xml */
                            XmlNode project = X.SelectSingleNode("//project/@name");
                            XmlNode lang = X.SelectSingleNode("//language/@name");
                            String projectname = project.InnerText;
                            /* then put it to the output file name and change the extension to txt */
                            String nazovsuboru = subor.Replace("Analyze Files", projectname);
                            vystup = nazovsuboru.Replace("xml", "csv");

                            String allvystup = Path.Combine(Path.GetDirectoryName(subor), "Studio-allfiles.csv");

                            /* select the appropriate node */
                            XmlNodeList filelist = X.SelectNodes("/*/*[analyse]");

                            /* if it exists */
                            if (!(filelist.Count == 0))
                            {
                                /* process each node */
                                foreach (XmlNode file in filelist)
                                {
                                    //DEBUG MessageBox.Show(file.InnerText);*/
                                    /* first prepare a dictionary to put all the data into */
                                    SortedDictionary<string, int> wordcounts = new SortedDictionary<string, int>();
                                    XmlNodeList decka = file.SelectNodes("analyse/*");
                                    foreach (XmlNode decko in decka)
                                    {
                                        string cat = "";
                                        int val;

                                        /* fuzzy matches are distinguished by min and max attribute */
                                        if (!(decko.Attributes["max"] == null))
                                        {
                                            cat = decko.LocalName + decko.Attributes["max"].InnerText;
                                        }

                                        else
                                        {
                                            cat = decko.LocalName;
                                        }

                                        val = int.Parse(decko.Attributes["words"].InnerText);

                                        wordcounts.Add(cat, val);
                                    }

                                    int xtranslate = wordcounts["inContextExact"] + /*wordcounts.["locked"] +*/ wordcounts["perfect"];
                                    int repetitions = wordcounts["crossFileRepeated"] + wordcounts["repeated"];
                                    int exact = wordcounts["exact"];
                                    int fuzzy84 = wordcounts["fuzzy84"];
                                    int fuzzy94 = wordcounts["fuzzy94"];
                                    int fuzzy99 = wordcounts["fuzzy99"];
                                    int nomatch = wordcounts["new"] + wordcounts["fuzzy74"];

                                    /* let's add this all to a dictionary so that we can work with something usable */
                                    SortedDictionary<string, int> xmlwordcounts = new SortedDictionary<string, int>();

                                    if (checkBox4.Checked == true)
                                    {
                                        xmlwordcounts.Add(Wcat + "X", xtranslate + exact);
                                        xmlwordcounts.Add(Wcat + "P", repetitions);
                                        xmlwordcounts.Add(Wcat + "L", fuzzy84);
                                        xmlwordcounts.Add(Wcat + "M", fuzzy94);
                                        xmlwordcounts.Add(Wcat + "H", fuzzy99);
                                        xmlwordcounts.Add(Wcat, nomatch);
                                    }

                                    else
                                    {
                                        xmlwordcounts.Add(Wcat + "X", xtranslate);
                                        xmlwordcounts.Add(Wcat + "P", repetitions + exact);
                                        xmlwordcounts.Add(Wcat + "L", fuzzy84);
                                        xmlwordcounts.Add(Wcat + "M", fuzzy94);
                                        xmlwordcounts.Add(Wcat + "H", fuzzy99);
                                        xmlwordcounts.Add(Wcat, nomatch);
                                    }

                                    string xmlcelkom = Studio20142HPT.PutTogether.PutTogetherDict(xmlwordcounts);

                                    using (StreamWriter writer = new StreamWriter(vystup, true))
                                    {

                                        if (file.Attributes["name"] != null)
                                        {
                                            writer.WriteLine(file.Attributes["name"].InnerText + "," + xmlcelkom);
                                        }
                                        else
                                        {
                                            writer.WriteLine("TOTAL" + "," + xmlcelkom);
                                        }
                                    }

                                    //write project totals by language
                                    using (StreamWriter writer = new StreamWriter(allvystup, true))
                                    {

                                        if (file.Attributes["name"] != null)
                                        {
                                            //writer.WriteLine(file.Attributes["name"].InnerText + "," + xmlcelkom);
                                        }
                                        else
                                        {
                                            writer.WriteLine(lang.InnerText.Replace(",", "") + " ," + xmlcelkom);
                                        }
                                    }
                                }
                                listBox2.Items.Add("File " + vystup + " saved successfully");
                            }

                            /* if the node /task/file doesn't exist, write an error message to the log */
                            else
                            {
                                listBox2.Items.Add("Are you sure this is a Trados Studio 2014 XML file?");
                            }

                            /* add date and time for logging purposes */
                            Studio20142HPT.PutTogether.PutTogetherAndWriteDateandTime(vystup);

                        }

                        /* if the XmlDocument.Load method cannot parse the file, write an error message to the log */
                        catch (Exception ex)
                        {
                            listBox2.Items.Add("Are you sure this is a Trados Studio 2014 XML file?");
                            listBox2.Items.Add("If yes, please report this error to the author:");
                            listBox2.Items.Add(ex.Message);
                        }
                    }
                }
            }
        }
    }
}