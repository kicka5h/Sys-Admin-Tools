#Get the required params
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [String]$Username,
  
    [Parameter(Mandatory=$true)]
    [String]$Password,

    [Parameter(Mandatory=$true)]
	[String]$To
)

# Set the style for the email
$CSS = @"
<style>
    table { margin: auto; font-family: Brandon Grotesque; border-collapse: collapse; }
    table, th, td { border: 1px solid #E3E3E3; }
    th { background: #AFC2C4; color: #fff; max-width: 400px; padding: 5px 10px; }
    td { font-size: 11px; padding: 5px 20px; color: #000; }
    tr { background: #fff; }
</style>
"@

# Formatting the first table to use in the email
[array]$Object1 = $null
[array]$Object1 += [PSCustomObject]@{
    "Column Name 1" = "Value 1" 
    "Column Name 2" = "Value 2"
	"Column Name 3" = "Value 3"
}

# Formatting the second table to use in the email
[array]$ValueList = $null
[array]$ValueList += "Value 1"
[array]$ValueList += "Value 2"
[array]$ValueList += "Value 3"

$OneValue = "Value 1"

[array]$Object2 = $null
[array]$Object2 += [PSCustomObject]@{         
     "Column Name 1" = (@($ValueList) -join ', ')
     "Column Name 2" = $OneValue
}

# Prepare the tables in the email. Use Fragment to join the two tables.
$Table1 = ($Object1 | Sort-Object "Column Name 1", "Column Name 2", "Column Name 3" | Select-Object "Column Name 1", "Column Name 2", "Column Name 3" | ConvertTo-Html -Fragment)
$Table2 = ($Object2 | Sort-Object "Column Name 1", "Column Name 2" | Select-Object "Column Name 1", "Column Name 2" | ConvertTo-Html -Fragment)

# Create the body of the email
$Body = @"
$CSS

<h1> Heading 1 </h1>
<h2> Heading 2 </h2>
<h3> Heading 3 </h3>
<h4> Heading 4 </h4>
<h5> Heading 5 </h5>
<h6> Heading 6 </h6>

<br>

Regular text <br>
<B> Bold text </B> <br>
<i> Itallic text </i> <br>

<br>

<ol type="1">
    <li>A</li>
    <li>Numbered</li>
    <li>List</li>
</ol>

<br>

<ol type="A">
    <li>A</li>
    <li>Lettered</li>
    <li>List</li>
</ol>

<br>

<ul style="list-style-type:disc;">
    <li>A</li>
    <li>Bulleted</li>
    <li>List</li>
</ul>

<br>

<script src="https://cdn.jsdelivr.net/gh/google/code-prettify@master/loader/run_prettify.js"></script>

<pre class="prettyprint">
    <code>
        This is a code block. This code block will be styled and displayed as plain text. 
    </code>
</pre>

$Table1

<br>
<br>

$Table2

<br>
<br>
"@

# Get credentials to send the email
[string][ValidateNotNullOrEmpty()]$passwd = $Password
$secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
$cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)

# Set up the email params
$Date = Get-Date -Format 'MMMM dd yyyy'
$From = $Username
$Subject = "Email Formatting Test " + $Date
$SMTPPort = "587"
$SMTPServer = "smtp.office365.com"

# Send the email
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHTML -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred