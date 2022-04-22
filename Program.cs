using System;
using System.Collections.Generic;

namespace Export_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> tablename = new List<string>();
            string filename = "z_out.xlsx";
            string databasename = "example.db";

            Console.WriteLine("Starting up");
            int i = 0;
            while (i < args.Length)
            {
                if (match("--file", args[i]))
                {
                    filename = args[i + 1];
                    i = i + 2;
                }
                else if (match("--table", args[i]))
                {
                    tablename.Add(args[i + 1]);
                    i = i + 2;
                }
                else if (match("--database", args[i]))
                {
                    databasename = args[i + 1];
                    i = i + 2;
                }
                else if (match("--verbose", args[i]))
                {
                    int index = 0;
                    foreach (string s in args)
                    {
                        Console.WriteLine(string.Format("argument {0}: {1}", index, s));
                        index = index + 1;
                    }
                    i = i + 1;
                }
                else
                {
                    i = i + 1;

                }
            }

            excel_export e = new excel_export();
            e.initialize(filename, databasename);
            e.Generate(tablename);
            e.MultiSave();
            e.Cleanup();

        }

        static public bool match(string a, string b)
        {
            return a.Equals(b, StringComparison.CurrentCultureIgnoreCase);
        }
    }
}
