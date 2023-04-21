using System;
using System.IO;
using System.Text;
using Jint;
using Jint.Native;

namespace SheetJSJint
{
    class Program
    {

        private static void SafeRunFile(ref Jint.Engine eng, string fileToLoad)
        {
            try
            {
                var data = File.ReadAllText(fileToLoad);
                eng.Execute(data);
     
            } catch (Exception e)
            {
                Console.WriteLine("Exception in run file. {0}", e.Message);
            }
        }

        private static String ReadFile(string file)
        {
            try
            {
                
                byte[] b = File.ReadAllBytes(file);

                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < b.Length; i++) sb.Append(Char.ToString((char)(b[i] < 0 ? b[i] + 256 : b[i])));

                return sb.ToString();
            } catch (Exception e)
            {
                Console.WriteLine("Exception in read file. {0}", e.Message);
                return "";
            }
        }

        private static JsValue EvalString(ref Jint.Engine eng, string cmd)
        {
            try
            {
                var v = eng.Evaluate(cmd);
                return v;

            } catch (Exception e)
            {
                Console.WriteLine("Exception in eval string. {0}", e.Message);
                return JsValue.Null;
            }
        }

        private static void WriteType(ref Jint.Engine eng, string type)
        {
            try
            {
                var b64str = EvalString(ref eng, "XLSX.write(wb, {type:'base64', bookType: '" + type + "'})");
                byte[] buf = Convert.FromBase64String(b64str.ToString());
                File.WriteAllBytes("sheetjsw." + type, buf);
            } catch (Exception e)
            {
                Console.WriteLine("Exception in write type: {0}", e.Message);
            }
        }

        static void Main(string[] args)
        {
            /* initialize */
            var eng = new Engine();
            EvalString(ref eng, "var global = (function(){ return this; }).call(null);"); //why?


            /* load library */
            SafeRunFile(ref eng, @"shim.min.js");
            SafeRunFile(ref eng, @"xlsx.full.min.js");


            /* get version string */
            var v = EvalString(ref eng, "XLSX.version");
            Console.WriteLine("SheetJS library version " + v);


            /* read file */
            var data = ReadFile("pres.numbers");
            eng.SetValue("buf", data);


            /* parse workbook */
            EvalString(ref eng, "wb = XLSX.read(buf, {type:'binary'});");
            EvalString(ref eng, "ws = wb.Sheets[wb.SheetNames[0]]");


            /* print CSV */
            var output = EvalString(ref eng, "XLSX.utils.sheet_to_csv(ws)");
            Console.WriteLine(output);


            var statement = Console.ReadLine();


            /* write file */
            WriteType(ref eng, "xlsx");
            var statement2 = Console.ReadLine();

        }
    }
}
