using System;
using System.IO;

namespace CSVConvert
{
    // simple errror checking.
    class Error
    {

        public Error()
        {
        }

        public void helpMenu()
        {
            Console.WriteLine("Program usage (must have -f and -d)");
            Console.WriteLine("-f //path//filename.xlsx");
            Console.WriteLine("-d //path//output//dir");
            Console.WriteLine("-h this menu");
            Environment.Exit(0);
        }

        public string[] checkArgs(string[] args)
        {
            string[] cleanargs = new string[] {"", ""};
            if(args.Length != 4)
                helpMenu();
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].Equals("-d"))
                    cleanargs[0] = args[i +1];
                if (args[i].Equals("-f"))
                    cleanargs[1] = args[i + 1];
            }
            return cleanargs;
        }

        // very simple directory check, exit program if fails
        public bool directoryExists(String dir)
        {
            if (Directory.Exists(dir))
            {
                return true;
            }
            else
            {
                Console.WriteLine("Directory {0} not found", dir);
                Environment.Exit(1);
                return false;
            }
        }

        // combines two methods to check if file exists and it has correct extension
        public bool FileIsCorrect(String file)
        {
            if (FileExist(file))
                if (HasCorrectExtension(file))
                    return true;
            return false;
        }

        // checks if file exists
        public bool FileExist(String file)
        {
            if (File.Exists(file))
            {
                return true;
            } 
            else
            {
                Console.WriteLine("file {0} not found", file);
                Environment.Exit(1);
                return false;
            }

        }

        // check to make sure we have the correct extension
        public bool HasCorrectExtension(String file)
        {
            string extension = Path.GetExtension(file);
            if (extension.Equals(".xlsx"))
                return true;
            else
            {
                Console.WriteLine("This is not an XLSX file");
                Environment.Exit(1);
                return false;
            }

        }
    }
}
