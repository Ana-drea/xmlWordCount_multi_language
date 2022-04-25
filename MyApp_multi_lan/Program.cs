using System;
using System.Collections;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml;

namespace MyApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            System.Console.WriteLine("Please enter the source folder path:");
            String path = System.Console.ReadLine();
            List<string> languagelist = new List<string> { };
            //create an ExcelPackage
            ExcelPackage excelPkg = new ExcelPackage();
            //add a sheet
            ExcelWorksheet ews = excelPkg.Workbook.Worksheets.Add("Sheet1");

            if (Directory.Exists(path))
            { 
                String[] dir = Directory.GetDirectories(path);
                if (dir!= null&&dir.Length!=0){
                    //get a list of the language folder names, it will decide the column names of final excel sheet.
                    foreach (String languagefolderpath in dir){
                        languagelist.Add(languagefolderpath);
                    }
                    //add table head
                    ews.Cells[2, 1].Value = "TOTAL";
                    //iterate each language folder, get the mxliff files and their word counts under this path
                    for(int i=0;i<languagelist.Count;i++)
                    {
                        //for each language, create a filelist to store all the mxliff file names
                        List<string> filelist = new List<string> { };
                        String languagefolderpath = languagelist[i];
                        //get a list of all the mxliff files in this language folder
                        GetMxliff(languagefolderpath, filelist);
                        //add table heads
                        String languagefoldername = Path.GetFileName(languagefolderpath);
                        ews.Cells[2, 2*i+2].Value = languagefoldername;
                        ews.Cells[1, 2*i+3].Value = "word count";
                        //Enumerate the mxliff files, get file name and word count of each file
                        for (int j = 0; j < filelist.Count; j++)
                        {
                            ArrayList result = GetSourceString(filelist[j]);
                            //write file name and word count into excel
                            ews.Cells[j + 3, 2*i+2].Value = result[0];
                            ews.Cells[j + 3, 2*i+3].Value = result[1];

                        }
                        //create the sum formula
                        String column = ""+(char)(2*i+67);
                        String Formula = String.Format("SUM({0}:{1})", column+"3", column + (2 + filelist.Count));
                        //get the total word count in the last cell
                        ews.Cells[2, 2*i+3].Formula = Formula;
                        //set a filter
                        // ews.Cells["A1:" + "B" + (filelist.Count)].AutoFilter = true;


                    }
                }else{
                    System.Console.WriteLine("Sorry, no subfolder under this path");
                }
                //save this excel in the source folder
                excelPkg.SaveAs(new FileInfo(Path.Combine(path, "WordCount.xlsx")));

            }
            else
            {
                System.Console.WriteLine("Invalid file path, please check your input.");
            }

        }


        public static void GetMxliff(String path, List<string> filelist)
        {
            if (Directory.GetDirectories(path) != null)
            {
                foreach (String subdirectory in Directory.GetDirectories(path))
                {
                    GetMxliff(subdirectory, filelist);
                }
            }
            foreach (string file in Directory.EnumerateFiles(path, "*.mxliff"))
            {
                filelist.Add(file);
            }
        }

        public static ArrayList GetSourceString(String path)
        {
            ArrayList nameandcount = new ArrayList();
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(path);
            XmlNodeList list = xmldoc.GetElementsByTagName("trans-unit");
            int wcount = 0;
            foreach (XmlNode node in list)
            {
                if (node.Attributes["m:confirmed"].Value == "1")
                {
                    String source = node.ChildNodes[1].InnerText;
                    wcount += CountWords(source);
                }
            }
            nameandcount.Add(Path.GetFileName(path));
            nameandcount.Add(wcount);
            return nameandcount;
        }

        public static int CountWords(string s)
        {
            MatchCollection collection = Regex.Matches(s, @"[\S]+");
            return collection.Count;
        }

    }
}