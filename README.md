# PowerShell-Data-Use
A little text explaining the usage of diferent tipes of data in powershell, made for UES SIGLO XXI


PowerShell Data Basics: File-Based Data
One of the frustrations of anyone beginning with PowerShell is the simple task of getting data in and out. To help out with this, Michael Sorens begins a series that shows you how to import data from most of the common file formats and sources. He also shows how to export data in a range of formats.
Introduction
As I have continued working with PowerShell I realized that at the core is the data (as with any software really). In the case of PowerShell, this boils down to converting external data  into PowerShell objects and vice versa. This is the first of a series of articles that shows you how to importalmost all of the common data formats you are likely to encounter, and how to export to some of them as well.
Part 1: General Data
This article explains how to import data from files of all kinds except for XML, which is covered separately in the next article in this series. The sections below describe a variety of text formats, from fixed-width, variable-width, and ragged-right files to CSV, property lists, INI files, and JSON data, and concludes with a treatment of importing and exporting to Excel.
Part 2 of this series will  illustrate the two principal  technologies available within PowerShell to access XML data, XPath and objects. You will learn how to move XML data into and out of PowerShell along with how to reference, manipulate, and modify it.
Data by Line
Read a Text File in its Entirety
Regular readers will know of my fondness for Lewis Carroll; the current topic again draws me unerringly to this inadvertent comment on software methodology:
‘Begin at the beginning,’ the King said gravely, ‘and go on till you come to the end: then stop.’
–Lewis Carroll, Alice’s Adventures in Wonderland
In the context of PowerShell, the beginning is where you learn to read a file in PowerShell. My sample file consists of fifteen lines containing the text “one”, “two”, etc., on up to “fifteen”. The Get-Contentcmdlet is the primary workhorse for bringing data from a file into PowerShell:
Get-Content -Path .\numbers.txt
one
two
three
four
five
six
seven
eight
nine
ten
eleven
twelve
thirteen
fourteen
fifteen
In the absence of any other direction, Get-Content displays its result on the console. You could redirect this output to a file using exactly the same mechanism as in a DOS or Unix or Linux shell …
	Get-Content -Path .\numbers.txt> newfile.txt 
… or you could send this output through the pipeline to another process, likewise:
	Get-Content -Path .\numbers.txt| ... 
Yet another common outlet is collecting the output into a variable for later processing:
	$numbers = Get-Content-Path .\numbers.txt 
Read a Text File Selectively by Index
PowerShell uses the $ prefix character to indicate a variable. Get-Content conveniently stores its output into $numbers as an array. Here are a couple of expressions to illustrate this:
$numbers.Count # report the cardinality
$numbers[0]    # display the first item
$numbers[14]   # display the last item
15
one
fifteen
PowerShell also offers a very handy range operator allowing you to select a portion of the array as a subset:
$numbers[0..3]
one
two
three
four
Applying that concept back to the original file, you can directly index the content of the file by line number-the parentheses are needed to first materialize the array that you then index with the bracket notation:
(Get-Content -Path .\numbers.txt)[0..3]
one
two
three
four
The array indexing notation gives you the most succinct way to select specific lines.  It is almost as though you are doing a pipeline, i.e. read-a-file | select-lines-from-file. You can, in fact, do the same operation with a full-fledged pipeline:
	Get-Content-Path .\numbers.txt |Select-Object -First 4  
Here you are passing the file contents to the Select-Object cmdlet and instructing it to take the first four elements, just as done in the previous example. That works for lines at the beginning, and you also have a -Last parameter to grab lines from the end of the file, but what about something in the middle? With array notation you can specify an arbitrary range [m..n]. With Select-Object this equates to two pipelined calls to Select-Object. Here is code to demonstrate the equivalence of the two methods:
	$m=5; $n=7
Get-Content .\numbers.txt | select -First ($n+1) | select -Last ($n-$m+1)
(Get-Content .\numbers.txt)[$m..$n] 
Obviously in this case the array indexing notation is much cleaner and more concise, but this example is valuable to introduce you to ways to manipulate data in PowerShell. You might use this technique to skip the header row when reading in a file, for example.
Read a Text File Selectively By Search
Now consider the scenario where rather than wanting to get specific lines by number you want to get all lines that contain a particular string. Use the Select-String cmdlet (similar to the Linux grepcommand or the DOS findstr command).
Get-Content -Path .\numbers.txt | Select-String teen
thirteen
fourteen
fifteen
You are not limited to string constants. In fact, the default search term is a regular expression unless you specify that you want to match just  a simple string with the -SimpleMatch parameter . But if your parameter does not contain any regex metacharacters you can omit the “-SimpleMatch” label as I have done in the above example. So if you want to select the lines ending in teen that also contain an “i” earlier in the string, use this:
Get-Content -Path .\numbers.txt | Select-String -pattern "i.*teen$"
thirteen
fifteen
Data by Fixed Width Fields
The PowerShell Approach
The next logical step is to not just read a file by lines, but to split up those lines into fields. Say you have a text file where each line is a record and each record consists of, for example, 7 characters for the given name, 10 characters for the surname, and 3 characters for an ID of some kind.
A Fixed Width File
12345671234567890123
george jetson    5 
warren buffett   123
horatioalger     -99
	#process the file
Get-Content .\fixedwidth.log | #read each line as a separate object
   select -Property @{name='ID';expression={$_.Substring(17).Trim()}}, #first column
                    @{name='FirstName';expression={$_.Substring(0,7).Trim()}}, #2nd 
                    @{name='LastName';expression={$_.Substring(7,10).Trim()}} #3rd
A Fixed Width File Converted to Objects
Id  FirstName LastName 
--  --------- -------- 
123 1234567   1234567890
5   george    jetson   
123 warren    buffett  
-99 horatio   alger
The Regex Approach
 This approach does use regular expressions and the mere mention of them causes otherwise hearty, courageous, and fearless developers to quake in their boots. Or so I have heard. But for the task at hand what you need to know about regular expressions is simple. Really. For each field in the file you need to give it a name and you need to know its length. Take those two values and plug them into this template:
	(?<field name goes here>.{field length goes here})
Repeat that for each field then lay each one down adjacent to the previous one. Here I am representing names with “n” and lengths with “l”:
	(?<n1>.{l1})(?<n2>.{l2})(?<n3>.{l3})
Finally, add a caret (^) at the front end and a dollar sign ($) at the rear:
	^(?<n1>.{l1})(?<n2>.{l2})(?<n3>.{l3})$
Here is the realization of this for the example at hand:
	$regex= "^(?<FirstName>.{7})(?<LastName>.{10})(?<Id>.{3})$"
After you define the regex, the code to use it to parse the same text file is almost trivial, but this reads the file and extracts all the fields from each line (record) into an array of PowerShell objects-all thanks to the magic of the built-in -match operator.
	function ImportWith-Regex([string]$FilePath, [string]$regex)
{
    Get-Content $FilePath | ForEach-Object {
        if ($PSItem -match $regex)
        {
            New-Object PSObject -Property (Get-RegexNamedGroups $matches)
        }    
    }
}  
Note that this code needs a supplemental function to do its work, shown next, that collects just the parts of the regex match results that are needed. (Don’t worry about exactly how this works).
	function Get-RegexNamedGroups($hash)
{
    $newHash = @{};
    $hash.keys | ? { $_ -notmatch '^\d+$' } | % { $newHash[$_] = $hash[$_] }
    $newHash 
} 
The major advantage of the above approach is that you have complete separation between your data specification (the regex definition) and the code to process it. You can use the same function on a different file by just passing in a different regex and a different file name.
	$regex = "^(?<FirstName>.{7})(?<LastName>.{10})(?<Id>.{3})$"
ImportWith-Regex .\fixedwidth.log $regex
However, it is worth revisiting the earlier code-that used substrings-and seeing how you could do it with regular expressions.  Just make sure that your regex and your explicit field names in the selectstatement remain in sync.
	$regex = "^(?<FirstName>.{7})(?<LastName>.{10})(?<Id>.{3})$"
Get-Content .\fixedwidth.log |
   ForEach-Object {
      if ($_ -match $regex) {$matches} else {Throw 'line $_ was corrupt'}
   } |
   select -Property @{name='FirstName';expression={$_.FirstName}}, 
                    @{name='LastName'; expression={$_.LastName} }, 
                    @{name='Id';       expression={$_.Id}       }
The Ragged Right Variation
One important variation of fixed width files is the ragged right format, which defines all columns by width except for the last column, which simply runs to the end of the line. Accommodating this format requires just a trivial change to the last code sample. Modify the last capture group in the regular expression to use .* instead of .{n} as shown below. I have also renamed it from Id to Description since that is now a more likely field name.
	$regex = "^(?<FirstName>.{7})(?<LastName>.{10})(?<Description>.*)$"
ImportWith-Regex .\fixedwidth.log $regex
Here’s the second form:
	$regex = "^(?<FirstName>.{7})(?<LastName>.{10})(?<Description>.*)$"
Get-Content .\fixedwidth.log |
   ForEach-Object {
      if ($_ -match $regex) {$matches} else {Throw 'line $_ was corrupt'}
   } |
   select -Property @{name='FirstName';   expression={$_.FirstName}  }, 
                    @{name='LastName';    expression={$_.LastName}   }, 
                    @{name='Description'; expression={$_.Description}}
 
Data by Variable Width Fields
Importing from CSV
Pipe your data or your file into ConvertFrom-Csv (if immediate data) or Import-Csv (if file data), to yield an array of PowerShell objects-no muss, no fuss. This first example uses multi-record input from a string constant for clarity; loading data from a file is just as straightforward.
@’
Shape,Color,Count
Square,Green,4
Rectangle,,12
Parallelogram,,0
Trapezoid,Black,100
‘@ | ConvertFrom-Csv
Shape              Color     Count
-----              -----     -----
Square             Green     4   
Rectangle                    12  
Parallelogram                0   
Trapezoid          Black     100  
Things to observe from this example:
•	The first row of data contains column headers that become properties in PowerShell.
•	Each input record generates a PowerShell object with those properties.
•	The actual output of ConvertFrom-Csv (or Import-Csv) is an array of those PowerShell objects.
•	All data are strings even if they look like other types (e.g. integers). That is why the values of the Count column/property are left-justified. See the next section for a workaround.
ConvertFrom-Csv (and Import-Csv) gives you the option to include the header row in your data or not, at your choice. This next example shows how to separate the header row from the data by including it as part of the command invocation using the -Header parameter.
@’
Square,Green,4
Rectangle,,12
Parallelogram,,0
Trapezoid,Black,100
‘@ | ConvertFrom-Csv -Header Shape, Color, Count
Shape              Color     Count
-----              -----     -----
Square             Green     4   
Rectangle                    12  
Parallelogram                0   
Trapezoid          Black     100  
One special case of interest is worth considering here: directly populating a hash table from a CSV file. In this case, your data records should consist of two fields each. (Any additional fields for a given record will be ignored; any record with just a single field will assign a null value to that hash entry.)
 If your CSV file has a header record (i.e. column names in the first record), use this (note that this also gives you the flexibility to select any two fields by name in the record if more than two are present):
$hash = @{};
Import-Csv data.csv |
% { $hash[$_.first] = $_.second } # Assumes header row “first,second”
$hash[“Square”] # output one of the stored values
Green
Exporting to CSV
There are two cmdlets for exporting data to CSV: ConvertTo-Csv sends its output to stdout while Export-Csv sends its output specifically to a file. Otherwise, they operate the same. In the example below you start with the output of the Get-Process cmdlet, filter by row with Where-Object to just those processes beginning with “W”, filter by column with Select-Object , and finally pipe to the ConvertTo-Csv cmdlet to generate the output shown.
Get-Process |
Where-Object { $_.name -like “W*” } |
Select-Object name, path, vm, fileversion, id, handles |
ConvertTo-Csv
#TYPE Selected.System.Diagnostics.Process
"Name","Path","VM","FileVersion","Id","Handles"
"wininit",,"34832384",,"488","77"
"winlogon",,"44924928",,"576","112"
"WLIDSVC",,"64946176",,"1844","337"
"WLIDSVCM",,"27185152",,"2088","53"
"WLTRAY","C:\Windows\System32\WLTRAY.EXE","168869888","4.170.25.12","3400","305"
"WLTRYSVC",,"17956864",,"1928","49"
"WmiPrvSE",,"67387392",,"3980","257"
"wmpnetwk",,"130781184",,"3616","496"
Notice that:
•	Outputs the object type as a comment in the first row. This may be suppressed with the -NoTypeInformation parameter.
•	Outputs the property names as the column headers.
•	Quotes every property as it emits it in the output. CSV format does not require quoting everyvalue, only those where ambiguity would arise-see The Comma Separated Value (CSV) File Format for the list of cases where quotes are required. However, quoting every value does no harm.
Generating CSV output, then, is simply a matter of getting the data into the form you want then piping it to  ConvertTo-Csv or Export-Csv.
Importing from a Log File
There are, of course, countless variations of log files, but one class of log file that is very common is that generated by a web server. The Apache/NCSA  common log format, a standardized format used by Apache web servers, contains fields separated by white space but it also allows whitespace within a field. CSV files handle this case by letting you optionally enclose a field in quotation marks; commas inside such a quoted region are considered normal text characters, not field separators. The Apache log allows this as well; it is most commonly used on the access request field. Other fields however, use different bracketing. The timestamp field, for instance must use required brackets ([ and ]) and treats white space within as plain text. Here are just a few lines from a log using this common log format.
Get-Content .\webserver.log
127.0.0.1 - frank [10/Oct/2012:13:55:36 -0700] "GET /apache_pb.gif HTTP/1.0" 200 2326
111.111.111.111 - martha [18/Oct/2012:01:17:44 -0700] "GET / HTTP/1.0" 200 101
111.111.111.111 - - [18/Oct/2007:11:17:55 -0700] "GET /style.css HTTP/1.1" 200 4525
Each row contains seven fields-here is the first record split apart with each field identified.
Host or IP address	 127.0.0.1
Remote log name	 –
Authenticated
user name	 frank
Timestamp	[10/Oct/2000:13:55:36 -0700]
Access request	GET /apache_pb.gif HTTP/1.0
Result status code	 200
Bytes transferred	 2326
Earlier you saw how to build a complicated-looking regular expression with a simple template to recognize fixed-width data and then pass that regex to ImportWith-Regex. Here’s the regex to recognize the Apache common log format followed by a call to ImportWith-Regex. I have wrapped them together into a function merely for convenience in the subsequent examples.
	function Import-ApacheLog($FileName)
{
    $apacheExtractor = "(?<Host>\S*)",
       "(?<LogName>.*?)",
       "(?<UserId>\S*)",
       "\[(?<TimeStamp>.*?)\]",
      "`"(?<Request>[^`"]*)`"",
       "(?<Status>\d{3})",
       "(?<BytesSent>\S*)" -join "\s+"
    ImportWith-Regex $FileName $apacheExtractor
}
If you just run the above function you get output in PowerShell’s canonical list format (each field is on a separate line and records are separated by an extra blank line). This occurs typically when a record has four or more fields. However, PowerShell’s table format is often more useful-and certainly more concise. To convert the output from the former to the latter, just pipe it to the Format-Table cmdlet. When you do that, however, the width of your screen may cause truncation of the data on screen. The last snippet, then, shows how to select fewer columns with the Select-Object cmdlet to avoid that issue.
Import-ApacheLog .\webserver.log
TimeStamp : 10/Oct/2012:13:55:36 -0700
LogName   : -
Host      : 127.0.0.1
UserId    : frank
Status    : 200
Request   : GET /apache_pb.gif HTTP/1.0
BytesSent : 2326
TimeStamp : 18/Oct/2012:01:17:44 -0700
LogName   : -
Host      : 111.111.111.111
UserId    : martha
Status    : 200
Request   : GET / HTTP/1.0
BytesSent : 101
TimeStamp : 18/Oct/2007:11:17:55 -0600
LogName   : -
Host      : 111.111.111.111
UserId    : -
Status    : 200
Request   : GET /style.css HTTP/1.1
BytesSent : 4525
Import-ApacheLog .\webserver.log | Format-Table -AutoSize
TimeStamp                  LogName Host            UserId Status Request           
---------                  ------- ----            ------ ------ -------           
10/Oct/2012:13:55:36 -0700 -       127.0.0.1       frank  200    GET /apache_pb.g...
18/Oct/2012:01:17:44 -0700 -       111.111.111.111 martha 200    GET / HTTP/1.0    
18/Oct/2007:11:17:55 -0600 -       111.111.111.111 -      200    GET /style.css H...
Import-ApacheLog .\webserver.log | Select Host,UserId,TimeStamp,Status,Request | Format-Table -AutoSize
Host            UserId TimeStamp                  Status Request                   
----            ------ ---------                  ------ -------                   
127.0.0.1       frank  10/Oct/2012:13:55:36 -0700 200    GET /apache_pb.gif HTTP/1.0
111.111.111.111 martha 18/Oct/2012:01:17:44 -0700 200    GET / HTTP/1.0            
111.111.111.111 -      18/Oct/2007:11:17:55 -0600 200    GET /style.css HTTP/1.1
Please note that the TimeStamp column is strictly a text value at this point. A more correct approach would require converting that to an actual DateTime object.  That could be done either by making a custom version of ImportWith-Regex or by going back to the other familiar method of importing you have seen in this article:
	$apacheExtractor = "(?<Host>\S*)",
       "(?<LogName>.*?)",
       "(?<UserId>\S*)",
       "\[(?<TimeStamp>.*?)\]",
      "`"(?<Request>[^`"]*)`"",
       "(?<Status>\d{3})",
       "(?<BytesSent>\S*)" -join "\s+"
 
Get-Content .\webserver.log |   #read each line as a separate object
   Foreach-Object{if ($_ -match $apacheExtractor) {$matches} else {throw 'Bad record $_ '}} |
     select -Property @{n='UserID';e={$_.UserId}}, #first column
           @{n='LogName';  e={$_.LogName}},
           @{n='Time';     e={[DateTime]::ParseExact($_.TimeStamp,
                              "dd/MMM/yyyy:HH:mm:ss zzz",
                              [System.Globalization.CultureInfo]::InvariantCulture)}},
           @{n='TimeStamp';e={$_.TimeStamp}},
           @{n='Host';     e={$_.Host}},
           @{n='Status';   e={$_.Status}},
           @{n='Request';  e={$_.Request}},
           @{n='BytesSent';e={$_.BytesSent}} |
Format-Table -AutoSize 
String Data Formats
Hash Table or Property List
A hash table or dictionary is often a very handy data structure to use. Say, for example, you want to maintain a list of configuration settings within your script. The next example shows three equivalent ways to do this. The last approach-with the ConvertFrom-StringData cmdlet-minimizes the use of punctuation requiring neither brackets, quotes, nor semicolons within the data.
$Options = @{}
$Options["height"]=1200
$Options["width"]=1600
$Options["aspect"]="4:3"
$Options["depth"]="24-bit"
# dump the contents
$Options	$Options = @{
  "height" = 1200;
  "width" = 1600;
  "aspect" = "4:3";
  "depth "="24-bit"
}
# dump the contents
$Options	$Options = @"
  height = 1200
  width = 1600
  aspect = 4:3
  depth = 24-bit
"@ | ConvertFrom-StringData
# dump the contents
$Options
Name       Value
----       -----
width      1600 
depth      24-bit
height     1200 
aspect     4:3	Name       Value
----       -----
width      1600 
depth      24-bit
height     1200 
aspect     4:3	Name     Value
----     -----
aspect   4:3
depth    24-bit
height   1200
width    1600
Such a list of configuration properties could be even more useful if you put them in a separate configuration file so you can edit the configuration file independently of the program. Here, for example, are four properties given some initial value:
Get-Content .\properties.txt
color=green
food=biscuit
flavor=bittersweet
voice=mellifluous
The ConvertFrom-StringData cmdlet operates on a single string containing multiple lines of text strings rather than a file, so your sequence begins by importing the file (Get-Content or gc) and converting it to a single string (Out-String). Store the result into a variable and you have a ready-made dictionary of configuration values.
$myConfig = gc.\properties.txt | Out-String | ConvertFrom-StringData
Name                           Value                   
----                           -----
color                          green 
food                           biscuit
flavor                         bittersweet
voice                          mellifluous
$myConfig["color"]
green                                                                                                 
INI Files
The INI file format is an old though still popular standard for configuration files used by Windows applications. INI files are simple text files composed of properties grouped into sections. In the example, there are two section, Install (with four properties) and Extras (with two properties). If you import the file with Get-Content, as shown, you just get lines of text.
Get-Content .\sample.ini
[Install]
Options=22744
Ignore=65534
Hardware=640|480|4|0
Software=0|0|0|0|0|0|0
[Extras]
Options=10
DllPath=0
However, if you instead import it with the Get-IniFile cmdlet that I’ll describe in a moment, you get a hash table indexed by section names, whose entries are themselves hash tables indexed by property names. Let’s see that in slow motion. The first command reads the INI file and displays it to the console. The second sequence stores it to a variable for convenience and then displays a value from the hash table. The final sequence shows a reference to one of the nested hash tables.
Get-IniFile .\sample.ini
Name                           Value
----                           -----                                                                 
Install                        {Ignore, Software, Options, Hardware}
Extras                         {Options, DllPath}
$inifile = Get-IniFile .\sample.ini
$inifile["Install"]
Name                           Value                                                                                                 
----                           -----      
Ignore                         65534
Software                       0|0|0|0|0|0|0
Options                        22744
Hardware                       640|480|4|0
$inifile["Install"]["Hardware"]
640|480|4|0
Here is the code for Get-IniFile (adapted from this StackOverflow post). Note that if your file has properties occurring before any section is defined, those properties are put in a section labeled “-unknown-“.
	Function Get-IniFile ([string]$fileName) {
       $ini = @{}
       switch -regex -file $fileName {
              "^\[(.+)\]$" {                # recognize a section
                     $section = $matches[1]
                     $ini[$section] = @{}
              }
              "^\s*([^#]+?)\s*=\s*(.*)" {   # recognize a property
                     $name,$value = $matches[1..2]
                     if (!(Test-path variable:\section)) {
                           $section = "-unknown-"
                           $ini[$section] = @{}
                     }
                     $ini[$section][$name] = $value.trim()
              }
       }
       $ini
} 
JSON Data
The JSON standard for data interchange derives from JavaScript notation (hence the nameJavaScript Object Notation) but it is language-independent. It serves much the same purpose as XML and has a similar expressive power as XML. Depending on the how you represent data in the two formats, a JSON representation may be shorter than one in XML (primarily due to closing tags on XML elements). JSON is in some ways less burdensome than XML, though, as aptly described in JSON: The Fat-Free Alternative to XML.
PowerShell (with the advent of V3) provides direct support for JSON with the ConvertFrom-Jsonand ConvertTo-Json cmdlets. So let’s convert some simple JSON to a PowerShell object. The fields are an excerpt from a .NET DateTime object, showing both simple properties and nested properties.
$dateObject = @"
{
    Day:  16,
    DayOfWeek:  3,
    DayOfYear:  16,
    Hour:  15,
    Minute:  56,
    Month:  1,
    Second:  58,
    Ticks:  634939486185604791,
    TimeOfDay:  {
        TotalDays:  0.66456667221180554,
        TotalHours:  15.949600133083333,
        TotalMilliseconds:  57418560.479100004,
        TotalMinutes:  956.976007985,
        TotalSeconds:  57418.5604791
        },
    Year:  2013
}
"@ | ConvertFrom-Json
$dateObject
Day Month Year
--- ----- ----
 16     1 2013
$dateObject.Year
2013
$dateObject.TimeOfDay.TotalSeconds
57418.5604791
If you want to try the same thing with live data instead of a string constant, you can take the output of Get-Date as a list of properties and convert it to JSON :
	Get-Date| Select-Object -Property* | ConvertTo-Json
Tack on the same ConvertFrom-Json as a final command in the sequence to mimic the earlier results (though in this case it is not terribly productive!):
	Get-Date| Select-Object -Property* | ConvertTo-Json| ConvertFrom-Json 
To demonstrate something a bit more useful, the Invoke-WebRequest cmdlet fetches content of a web page or web service. Here you see successive steps to fetch a web response, unwrap its JSON content, and convert that to PowerShell objects so that you can directly address its elements. (Note that the actual JSON data-the output of the second command-was manually run through the JSON Formatter and Validator to pretty-print it for this article; otherwise you would see everything on one line, making it very difficult to see what is there.)
# Web response as a PowerShell object
$url = "http://search.twitter.com/search.json?q=PowerShell"
Invoke-WebRequest $url
StatusCode        : 200
StatusDescription : OK
Content           : {"completed_in":0.037,"max_id":291699484801527808,
Forms             : {}
Headers           : {[X-Transaction, 89f7072a7bd683a0], [X-Frame-Options, SAMEORIGIN]...}
Images            : {}
InputFields       : {}
Links             : {}
ParsedHtml        : mshtml.HTMLDocumentClass
RawContentLength  : 11985
. . .
# JSON response
(Invoke-WebRequest $url).content
{
   "completed_in":0.015,
   "max_id":291699484801527808,
   "max_id_str":"291699484801527808",
   "page":1,
   "query":"PowerShell",
   "results":[
      {
         "created_at":"Thu, 17 Jan 2013 00:12:31 +0000",
         "from_user":"mjolinor",
         "from_user_id":226782418,
         "from_user_id_str":"226782418",
         "from_user_name":"Rob Campbell",
         "geo":null,
         . . .
# JSON response converted to a PowerShell object
$jsonContent = (Invoke-WebRequest $url).content |  ConvertFrom-Json
$ jsonContent
completed_in     : 0.023
max_id           : 291699484801527808
max_id_str       : 291699484801527808
page             : 1
query            : PowerShell
results          : {@{created_at=Thu, 17 Jan 2013 00:12:31 +0000; from_user=mjolinor;
                   from_user_id=226782418; from_user_id_str=226782418;
                   from_user_name=Rob Campbell; geo=; ...
. . .
$ jsonContent.completed_in
0.013
$ jsonContent.results[0].created_at
Thu, 17 Jan 2013 00:12:31 +0000
See JSON.org for more on JSON.
Excel
Reading and writing Excel from PowerShell is done fairly easily as well, though it is much more involved than everything else you have read thusfar. Chances are your needs fall into one of two camps: reading Excel on a system that has Excel installed, and reading Excel on a system that does not (perhaps because your application needs to be used by all your customer service reps by their machines are not set up with Excel).
Excel with Excel
To read Excel, Robert M. Toups, Jr. in his blog entry Speed Up Reading Excel Files in PowerShellexplains that while loading a spreadsheet in PowerShell is fast, actually reading its cells is very slow. On the other hand, PowerShell can read a text file very quickly, so his solution is to load the spreadsheet in PowerShell, use Excel’s native CSV export process to save it as a CSV file, then use PowerShell’s standard Import-Csv cmdlet to process the data blazingly fast. He reports that this has given him up to a 20 times faster import process! Leveraging Toups’ code, I created an Import-Excel function that lets you import spreadsheet data very easily:
$spreadsheetData = Import-Excel "datadir\sample.xlsx"
$spreadsheetData
name   id
----   --
foo    3 
bar    25
alpha  -99
# Display name in first row (0-based index)
$spreadsheetData[0].name
foo
My code adds the capability to select a specific worksheet within an Excel workbook, rather than just using the default worksheet (i.e. the active sheet at the time you saved the file). If you omit the -SheetName parameter, it uses the default worksheet.
	function Import-Excel([string]$FilePath, [string]$SheetName = "")
{
    $csvFile = Join-Path $env:temp ("{0}.csv" -f (Get-Item -path $FilePath).BaseName)
    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }
 
    # convert Excel file to CSV file
    $xlCSVType = 6 # SEE: http://msdn.microsoft.com/en-us/library/bb241279.aspx
    $excelObject = New-Object -ComObject Excel.Application  
    $excelObject.Visible = $false 
    $workbookObject = $excelObject.Workbooks.Open($FilePath)
    SetActiveSheet $workbookObject $SheetName | Out-Null
    $workbookObject.SaveAs($csvFile,$xlCSVType) 
    $workbookObject.Saved = $true
    $workbookObject.Close()
 
     # cleanup 
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) |
        Out-Null
    $excelObject.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) |
        Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
 
    # now import and return the data 
    Import-Csv -path $csvFile
}
These supplemental functions are used by Import-Excel:
	function FindSheet([Object]$workbook, [string]$name)
{
    $sheetNumber = 0
    for ($i=1; $i -le $workbook.Sheets.Count; $i++) {
        if ($name -eq $workbook.Sheets.Item($i).Name) { $sheetNumber = $i; break }
    }
    return $sheetNumber
}
 
function SetActiveSheet([Object]$workbook, [string]$name)
{
    if (!$name) { return }
    $sheetNumber = FindSheet $workbook $name
    if ($sheetNumber -gt 0) { $workbook.Worksheets.Item($sheetNumber).Activate() }
    return ($sheetNumber -gt 0)
}
To write to Excel, Robert M. Toups, Jr. in his blog entry Write Excel Spreadsheets Fast in PowerShellsuggests that if you have a lot of data to load into Excel, doing this directly in PowerShell is much more time-consuming than converting the data to a CSV file than letting Excel’s native CSV import process-controlled through PowerShell-do the data loading. I adapted his code to be suitable for a generic Excel exporter, allowing you to specify the title and author of the Excel workbook, and the name of the single worksheet it creates in the file:
	function Export-Excel([string]$FilePath, [string]$Title, [string]$Author, [string]$SheetName, [Object[]]$Data)
{
    # Specify to save in a standard .XSLX format.
    $xlOpenXMLType = 51 # SEE: http://msdn.microsoft.com/en-us/library/bb241279.aspx
 
    $csvFile = Join-Path $env:temp `
        ("{0}.csv" -f ([System.IO.FileInfo]$FilePath).BaseName)
    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }
    if (Test-Path -path $FilePath) { Remove-Item -path $FilePath }
 
    $Data | Export-Csv -path $csvFile -noTypeInformation
 
    $excelObject = New-Object -comObject Excel.Application
    $excelObject.Visible = $false 
    $workbookObject = $excelObject.Workbooks.Open($csvFile)
    $workbookObject.Title = $Title
    $workbookObject.Author = $Author
    $worksheetObject = $workbookObject.Worksheets.Item(1)
    $worksheetObject.UsedRange.Columns.Autofit() | Out-Null
    $worksheetObject.Name = $SheetName
    $workbookObject.SaveAs($FilePath, $xlOpenXMLType)
    $workbookObject.Saved = $true
    $workbookObject.Close()
 
    # cleanup
    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) |
        Out-Null
    $excelObject.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) |
        Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
 
Of course there is so much more customization you could do when creating an Excel file. The MSDN documentation for the Microsoft.Office.Interop.Excel namespace is your starting point for digging into the wealth of methods and properties available on an Excel.Application object. For completeness, I will also mention Jeff Hicks’ Integrating Microsoft Excel with PowerShell article that shows how to do direct access to worksheet cells, rather than going through the intermediate CSV steps discussed above. It is quite suitable when you do not have a lot of data to process.
Excel without Excel
There are certainly situations where you might want to read an Excel spreadsheet on a machine that does not have Excel installed. In this case you must make do without the automation capabilities of Office, instead opting for a data access approach using OLEDB or ODBC. A very clean approach to this scenario using OLEDB is the publically available Get-OLEDBData created by Chad Miller, avid PowerShell aficionado. Here is the entire function:
	function Get-OLEDBData ($connectstring, $sql) {
   $OLEDBConn = New-Object System.Data.OleDb.OleDbConnection($connectstring)
   $OLEDBConn.open()
   $readcmd = New-Object system.Data.OleDb.OleDbCommand($sql,$OLEDBConn)
   $readcmd.CommandTimeout = '300'
   $da = New-Object system.Data.OleDb.OleDbDataAdapter($readcmd)
   $dt = New-Object system.Data.datatable
   [void]$da.fill($dt)
   $OLEDBConn.close()
   return $dt
} 
This function returns a DataTable, one of the object types readily handled by PowerShell. The catch, of course, is that you need to know what to provide for the $connectstring and $sql parameters. The commentary of Miller’s code in the PowerShell Code Repository details connection strings for commonly used data sources: Excel 2007 (or higher), Excel 2003, Informix, Oracle, and SQL Server. From there, you can see that Excel should use this:
	Provider=Microsoft.ACE.OLEDB.12.0;Data Source="C:\path\to\your\file.xlsx";Extended Properties="Excel 12.0 Xml;HDR=YES"
But to understand what really goes into an Excel connection string, I borrowed this format from the PowerShell Scripting Guy in his post How Can I Read from Excel Without Using Excel? Breaking it out this way makes things immediately obvious:
	$strFileName      = "C:\usr\tmp\sample.xlsx"
$strProvider      = "Provider=Microsoft.ACE.OLEDB.12.0"
$strDataSource    = "Data Source=`"$strFileName`""
$strExtend        = "Extended Properties=`"Excel 12.0 Xml;HDR=YES`""
$connectionString = "$strProvider;$strDataSource;$strExtend"
 
$strSheetName     = 'Sheet1'
$strQuery         = "Select * from [$strSheetName$]"
 
Get-OLEDBData -connectstring $connectionString -sql $strQuery 
One caveat: on a 64-bit machine you can install either 32-bit Office or 64-bit Office. (The latter provides additional capacity-for example, handling spreadsheets larger than 2GB-but at a potential cost of compatibility with add-ins not yet having 64-bit versions.) If you have installed 32-bit Office you will then need to run 32-bit PowerShell to use the Microsoft.ACE.OLEDB.12.0 or Microsoft.Jet.OLEDB.4.0 provider. If you attempt to use 64-bit PowerShell you will get an error stating: The ‘Microsoft.ACE.OLEDB.12.0’ provider is not registered on the local machine. While not technically a bug, this is a known issue and this post on the MSDN forums shows one developer’s journey to find a workaroun
