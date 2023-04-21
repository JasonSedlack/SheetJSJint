# SheetJSJint
A demo of using SheetJS Scripts with Jint


# C# + Jint

Jint is a c# implementation of ECMAScript5.1. 
SheetJS is a JavaScript solution for extracting and editing data from complex spreadsheets.

# Integration Details

Initialize JInt

Jint does not provide a global variable. It can be created in one line:

```
/* initialize */
var eng = new Engine();

/* jint does not expose a standard "global" by default */
eng.Evaluate("var global = (function(){ return this; }).call(null);");
```


Load SheetJS Scripts

The shim and main libraries can be loaded by reading the scripts from the file system and evaluating in the Jint context:

```
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

// ...
    SafeRunFile(ref eng, @"shim.min.js");
    SafeRunFile(ref eng, @"xlsx.full.min.js");
```

To confirm the library is loaded, XLSX.version can be inspected:

```
    var v = EvalString(ref eng, "XLSX.version");
    Console.WriteLine("SheetJS library version " + v);
```

# Reading Files


Files can be read into []byte:

```
/* read file */
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
```

This string can be loaded into the JS engine and processed:

```
/* Load into engine */
eng.SetValue("buf", data);

/* parse */
eng.Evaluate("wb = XLSX.read(buf, {type:'binary'});");
```

# Writing Files

"base64" strings can be passed from the JS context to C# code:

```
/* write to Base64 string */
var b64str = eng.Evaluate("XLSX.write(wb, {type:'base64', bookType: 'xlsx'})");

/* pull data back into C# and write to file */
byte[] buf = Convert.FromBase64String(b64str.ToString());
File.WriteAllBytes("sheetjsw.xlsx", buf);
```

# Complete Example

This demo was tested on 2023 March 20.


0) Create a project and install dependencies:
	1) Create project SheetJSJint
	2) References -> Manage Nuget Packages
		1) Add https://www.myget.org/F/jint/api/v3/index.json as package source
		2) Check include prerelease
		3) Install Jint (tested on preview 438)


1) Download sheetjs.jint.cs
	1) Copy file onto Program.cs file


2) Build the application


3) Download the standalone script, shim and test file into debug folder:

-   [xlsx.full.min.js](https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js)
-   [shim.min.js](https://cdn.sheetjs.com/xlsx-latest/package/dist/shim.min.js)
-   [pres.numbers](https://sheetjs.com/pres.numbers)

4) Run the file, when asked for file name, enter pres.numbers. CSV contents will be printed to console. Press enter to exit and the file sheetjsw.xlsx will be created. That file can be opened with Excel.