using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Xml;

namespace Test_Friedhelm
{
    class Program
    {
        static void Main(string[] args)
        {
            var inset = "\t\t";

            var sb = new StringBuilder();

            var zipFile = @"./print.oxps";
            var zip = ZipFile.OpenRead(zipFile); //unzip oxps file
            foreach (var entry in zip.Entries)
            {
                if (entry.Name.EndsWith(".fpage"))
                {
                    var xmldoc = new XmlDocument();
                    XmlNamespaceManager nsmanager = new XmlNamespaceManager(xmldoc.NameTable);
                    nsmanager.AddNamespace("xsl", "http://schemas.openxps.org/oxps/v1.0");
                    xmldoc.Load(entry.Open());                                                  //Open the file for the current page
                    var someNodes = xmldoc.SelectNodes("//xsl:Glyphs", nsmanager);      //Find all Parts with Text
                    bool flag = false;
                    foreach (XmlNode oneNode in someNodes)
                    {
                        var unicodeString = oneNode.Attributes.GetNamedItem("UnicodeString").Value.Trim();

                        //Filter for specific Texts
                        if (flag)
                        {
                            if (unicodeString.StartsWith("Anmerkung:"))
                            {
                                sb.Append(Environment.NewLine + Environment.NewLine);
                                flag = false;
                            }
                            else
                            {
                                sb.Append(Environment.NewLine + inset + unicodeString);
                            }
                        }
                        else if (unicodeString.EndsWith("Clarholz"))
                        {
                            var i = unicodeString.IndexOf(':', 13);
                            sb.Append(unicodeString[new Range(0, i)].Insert(13, "\t"));
                            sb.Append(Environment.NewLine);
                        }
                        //else if (unicodeString.StartsWith("E-Mail-Adresse"))
                        //{
                        //    sb.Append(inset);
                        //    sb.Append(unicodeString.Remove(0, 33));                   //Email Adress commented out, not needed
                        //    sb.Append(Environment.NewLine);
                        //}
                        else if (unicodeString.Contains("KFZ-Zeichen"))
                        {
                            sb.Append(inset);
                            sb.Append(unicodeString.Remove(0, 37));
                            sb.Append(Environment.NewLine);
                        }
                        else if (unicodeString.Contains("Dank"))
                        {
                            sb.Append(inset);
                            var i = unicodeString.IndexOf(':');
                            sb.Append(unicodeString.Remove(0, i+2));
                            flag = true;
                        }
                    }

                }
            }
            Console.WriteLine(sb.ToString());

            using (System.IO.StreamWriter file = new StreamWriter(@"./Output.txt"))
            {
                file.WriteLine(sb.ToString());
            }

            Process.Start("notepad.exe" , @".\Output.txt");
        }
    }
}
