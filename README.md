# CSV-to-HTML-Table
a Windows PowerShell script that converts a CSV file to a basic HTML table, with a simple GUI window

The script has been inspired by my friend Lorenzo. In his quest to learn Polish, he has decided to run a blog and have its readers submit corrections.

It may be useful, if you need a quick HTML table for a blog post and you are a Windows user.
The script shows a simple window that lets you:
<ul>
  <li>open a file,</li>
  <li>choose delimiter and character encoding,</li>
  <li>preview the content (to see the options are correct for the input file),</li>
  <li>save the file in the preferred location,</li>
  <li>set up default delimiter and encoding.</li>
</ul>

The functionalities also include error message popups.

<h2>Requirements</h2>
The script requires PowerShell interpreter and .NET platform.
It should run on a reasonably new Windows system without any additional installation.
It has been tested on Windows 10 with PowerShell version 5.1 and .NET Framework 4.8.1.
It is not known if it will run on any non-Windows operating system and what packages need to be installed on such systems.

<h2>How to run it?</h2>
Download the ps1 file.

If your execution policy allows running scripts, you can right click it and click "Run with PowerShell".
You can also:
<ol>
  <li>Open the file in PowerShell ISE (also right click -> "Edit").</li>
  <li>Select the entire code (Ctrl + A).</li>
  <li>Click Run Selection.</li>
</ol>

If run with PowerShell ISE, the script will try reading its configuration file from the current folder and write it there.
