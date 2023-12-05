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
            runCmd(@"excel2pdf.exe ""C:\测试中文文件夹\Cliente_11_AUR8903.xlsx""");
            Console.ReadLine();
        }

        public static  void runCmd(string strCMD)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = strCMD;
            //cmd.StartInfo.Arguments = "/c " + strCMD;
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();
            //cmd.StandardInput.WriteLine(strCMD);
            //cmd.StandardInput.Flush();
            //cmd.StandardInput.Close();
            cmd.WaitForExit();
            var res = cmd.StandardOutput.ReadToEnd();
            Console.WriteLine(res);
        }
    }
}
