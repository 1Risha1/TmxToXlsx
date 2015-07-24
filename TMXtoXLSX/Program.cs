using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace TMXtoXLSX
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.Error.WriteLine("You need to write <file.tmx> <file.xlsx> as comand line arguments");
                return 1;
            }
            string input = args[0];
            string output = args[1];

            if (File.Exists(output))
            {
                Console.Error.WriteLine("Target file: " + output + "is already exist");
                return 1;
            }
            else if (Path.GetDirectoryName(output) != null && !File.Exists(Path.GetDirectoryName(output)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output));
            }
            try
            {
                XlsxWriter writer = new XlsxWriter(output);
                TmxParser parser = new TmxParser(input, writer);
                parser.scanLanguages();
                parser.makeJob();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Can't transform file");
                return 1;
            }
            return 0;
        }
    }
}
