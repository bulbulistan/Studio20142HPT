using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Studio20142HPT
{
    public static class PutTogether
    {
        public static string PutTogetherDict(SortedDictionary<string, int> dict)
        {
            var filteredDokopy = from dok in dict
                                             where dok.Value != 0
                                             select dok;

                        string celkom = string.Join(";", filteredDokopy.Select(x => x.Value + x.Key));
                        return celkom;
        }

        public static void PutTogetherAndWriteDateandTime (string target)
        {
            try {             
            using (StreamWriter writer = new StreamWriter(target, true))
            {
                writer.WriteLine("Conversion executed at " + DateTime.Now);
                writer.WriteLine("--------------------------------------");
            }
                }
            catch
            {
                MessageBox.Show("Output file cannot be accessed.");
            }
        }
    }
}
