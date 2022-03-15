using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace robot
{
    class Program
    {
        private static string path;
        static void Main(string[] args)
        {
            //await GetPath();

            cl_Parser Parser = new cl_Parser();
            Parser.OpenFile();
        }

        async Task GetPath()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.ShowDialog();
            path = openDialog.FileName;
        }
    }
}
