using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks; 

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            runCmd(@"""C:\测试中文文件夹\excel2pdf.exe"" -i ""C:\Users\alex\Desktop\Cliente_11_A11UR8903.xlsx"" -o ""C:\Users\alex\Desktop\Cliente_11_AUR8903.pdf""");
            Console.ReadLine();
        }

        public static  void runCmd(string strCMD)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();
            cmd.StandardInput.WriteLine(strCMD);
            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();
            Console.WriteLine(cmd.StandardOutput.ReadToEnd());
        }
    }
}
