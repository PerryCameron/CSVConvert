using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSVConvert
{
    class Printer
    {
        List<Product> productList;

        public string OutputDir { get; set; }

        public Printer(List<Product> productList)
        {
            this.productList = productList;
        }

        public void PrintCSVFiles()
        {
            // rows per file
            int Rows = 9000;
            // size of list
            int ListSize = productList.Count;
            // number of files needed not counting remainder of rows
            int NumberOfFiles = ListSize / Rows;
            // find remainder of rows
            int Remainder = ListSize % Rows;
            // true if there is a remainder of rows 
            bool excess = false;
            // set the value of excess 
            if (Remainder > 0) excess = true;
            // have to add one for begin point
            int iterate = 1;
            // if there is excess we need to iterate one more time
            if (excess) iterate++;
            // variable for stop points
            int[] stop = new int[NumberOfFiles + iterate];
            // initial value is goint to be 0
            stop[0] = 0;
            // initialize i before loop so we can use after loop
            int i = 0;
            // lets set the other values with loop
            for (i = 1; i < NumberOfFiles + 1; i++)
                stop[i] = stop[i - 1] + 9000;
            // if there were excess rows we need one more stop point
            if (excess)
                stop[i] = stop[i - 1] + Remainder;
            // initialize string to use for fullpath
            String fullpath;
            for (i = 1; i < NumberOfFiles + iterate; i++)
            {
                fullpath = OutputDir + "output" + i + ".txt";
                // path plus begin and end of List
                PrintOutPutFile(fullpath, stop[i - 1], stop[i]);
            }


        }

        private bool PrintOutPutFile(String fullPath, int begin, int end)
        {
            if (fullPath is null)
            {
                Console.WriteLine("You did not supply a file path.");
                return false;
            }

            try
            {
                StreamWriter writer = new StreamWriter(fullPath);
                writer.WriteLine("PID^Product Id^Mfr Name^Mfr P/N^Price^COO^Short Description^UPC^UOM");

                for (int i = begin; i < end; i++)
                {
                    writer.WriteLine(productList[i].pid + "^"
                         + productList[i].productId + "^"
                         + productList[i].manufacturerName + "^"
                         + productList[i].manufacturerPN + "^"
                         + productList[i].cost + "^"
                     + productList[i].coo + "^"
                     + productList[i].description + "^"
                     + productList[i].upc + "^"
                     + productList[i].uom);
                }
                writer.Close();
                return true;
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("The file or directory cannot be found.");
            }
            catch (DirectoryNotFoundException)
            {
                Console.WriteLine("The file or directory cannot be found.");
            }
            catch (DriveNotFoundException)
            {
                Console.WriteLine("The drive specified in 'path' is invalid.");
            }
            catch (PathTooLongException)
            {
                Console.WriteLine("'path' exceeds the maxium supported path length.");
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("You do not have permission to create this file.");
            }
            catch (IOException e) when ((e.HResult & 0x0000FFFF) == 32)
            {
                Console.WriteLine("There is a sharing violation.");
            }
            catch (IOException e) when ((e.HResult & 0x0000FFFF) == 80)
            {
                Console.WriteLine("The file already exists.");
            }
            catch (IOException e)
            {
                Console.WriteLine($"An exception occurred:\nError code: " +
                                  $"{e.HResult & 0x0000FFFF}\nMessage: {e.Message}");
            }
            return false;
        }

        public static void writeErrorFile(String path)
        {

        }

    }

}

