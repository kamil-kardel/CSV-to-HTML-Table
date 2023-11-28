<# 
This script displays a GUI window that allows the user to import a .csv file and export an HTML table.


Author: kamil.kardel@gmail.com
#>


# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework



#region GUI

$WindowForm = New-Object System.Windows.Forms.Form
$WindowForm.ClientSize = '534,320'
$WindowForm.Text = '.CSV File to HTML table converter'
$WindowForm.FormBorderStyle = 1

$GUIXOffset = 16
$GUIYOffset = 16
$GUIRowHeight = 20
$GUIRowGap = 8
$GUIColumnGap = 12

#source file label

$SourceFileLabel = New-Object System.Windows.Forms.Label
$SourceFileLabel.Location = New-Object System.Drawing.Point($GUIXOffset,$GUIYOffset)
$SourceFileLabel.Text = 'Source file (.csv) :'
$SourceFileLabel.AutoSize = $true

$WindowForm.Controls.Add($SourceFileLabel)

#source file box
$SourceFileBox = New-Object System.Windows.Forms.TextBox
$SourceFileBox.Location = New-Object System.Drawing.Point($($GUIXOffset), $($GUIYOffset + $GUIRowHeight + $GUIRowGap))
$SourceFileBox.Size = New-Object System.Drawing.Size(400,$GUIRowHeight)
$SourceFileBox.ReadOnly = $true # the user is supposed to use the open file dialog and using it will fill in this box

$WindowForm.Controls.Add($SourceFileBox)

#Open file button
$OpenFileButton = New-Object System.Windows.Forms.Button
$OpenFileButton.Location = New-Object System.Drawing.Point($($SourceFileBox.Right + $GUIColumnGap), $($SourceFileBox.Top - 2))
$OpenFileButton.Text = 'Open file'
$OpenFileButton.AutoSize = $true

$WindowForm.Controls.Add($OpenFileButton)

#Preview label
$PreviewLabel = New-Object System.Windows.Forms.Label
$PreviewLabel.Location = New-Object System.Drawing.Point($GUIXOffset,$($OpenFileButton.Bottom + $GUIRowGap))
$PreviewLabel.Text = 'Input file preview :'
$PreviewLabel.AutoSize = $true

$WindowForm.Controls.Add($PreviewLabel)

#Preview box
$PreviewBox = New-Object System.Windows.Forms.RichTextBox
$PreviewBox.Location = New-Object System.Drawing.Point($GUIXOffset, $($PreviewLabel.Bottom + $GUIRowGap))
$PreviewBox.Size = New-Object System.Drawing.Size($($OpenFileButton.Right), 160)
$PreviewBox.ReadOnly = $true #not to be edited
$PreviewBox.Font = New-Object System.Drawing.Font('consolas',10) #monospace font, so columns are visible
$PreviewBox.WordWrap = $false

$WindowForm.Controls.Add($PreviewBox)

#encoding selection label
$EncodingLabel = New-Object System.Windows.Forms.Label
$EncodingLabel.Location = New-Object System.Drawing.Point($GUIXOffset, $($PreviewBox.Bottom + $GUIRowGap))
$EncodingLabel.Text = 'Source file character encoding :'
$EncodingLabel.AutoSize = $true

$WindowForm.Controls.Add($EncodingLabel)

#encoding selection box
$EncodingBox = New-Object System.Windows.Forms.ComboBox
$EncodingBox.Location = New-Object System.Drawing.Point($($EncodingLabel.Right + $GUIColumnGap), $($EncodingLabel.Top - 2))
$EncodingBox.Size = New-Object System.Drawing.Size(96, $GUIRowHeight)
$EncodingBox.DropDownStyle = 2 #only fixed set of possible encoding systems is available
$EncodingBox.Items.AddRange(@('Unicode';'UTF7';'UTF8';'ASCII';'UTF32';'BigEndianUnicode';'Default';'OEM'))

$WindowForm.Controls.Add($EncodingBox)

#delimiter selection label

$DelimiterLabel = New-Object System.Windows.Forms.Label
$DelimiterLabel.Location = New-Object System.Drawing.Point($($EncodingBox.Right + $GUIColumnGap),$($EncodingLabel.Top))
$DelimiterLabel.Text = 'Delimiter :'
$DelimiterLabel.AutoSize = $true

$WindowForm.Controls.Add($DelimiterLabel)

#delimiter selection box

$DelimiterBox = New-Object System.Windows.Forms.ComboBox
$DelimiterBox.Location = New-Object System.Drawing.Point($($DelimiterLabel.Right + $GUIColumnGap), $($EncodingBox.Top))
$DelimiterBox.Size = New-Object System.Drawing.Size(32, $GUIRowHeight)
$DelimiterBox.MaxLength = 1
$DelimiterBox.DropDownStyle = 1
$DelimiterBox.Items.AddRange(@(',',';',':'))

$WindowForm.Controls.Add($DelimiterBox)

#save file button

$SaveFileButton = New-Object System.Windows.Forms.Button
$SaveFileButton.Location = New-Object System.Drawing.Point($GUIXOffset, $($EncodingBox.Bottom + $GUIRowGap))
$SaveFileButton.Text = 'Save HTML table'
$SaveFileButton.AutoSize = $true

$WindowForm.Controls.Add($SaveFileButton)

#save default settings

$SaveDefaultsButton = New-Object System.Windows.Forms.Button
$SaveDefaultsButton.Location = New-Object System.Drawing.Point($($DelimiterBox.Right + $GUIColumnGap), $($EncodingBox.Top - 2))
$SaveDefaultsButton.Text = 'Save default settings'
$SaveDefaultsButton.Autosize = $true

$WindowForm.Controls.Add($SaveDefaultsButton)

#endregion

#region loadDefaultSettings
$iniFile = "`.`\CSV to HTML Table.ini"


function Get-SavedDefaults {
    try {
        
        $validationRegex = [regex]'^[\s]*((?<keyword>Encoding)[\s]*=[\s]*(?<value>(Unicode|UTF7|UTF8|ASCII|UTF32|BigEndianUnicode|Default|OEM))|(?<keyword>Delimiter)[\s]*=[\s]*(?<value>.))[\s]*$'
        $loadedSettings = Get-Content -Path $iniFile | Where-Object {$_ -match $validationRegex}
    }
    catch {
        [System.Windows.MessageBox]::Show($Error[0], 'Error', 'OK', 'Error')
        break
    }

    try {
        $loadedSettings | ForEach-Object {
            $_ -match $validationRegex
            switch($Matches['keyword']) {
                'Encoding' { $EncodingBox.SelectedIndex = $EncodingBox.Items.IndexOf($Matches['value'])
                            break }
                'Delimiter' { $DelimiterBox.Text = $Matches['value']
                            break}
            }

        }
    }
    catch {
        [System.Windows.MessageBox]::Show($Error[0], 'Error', 'OK', 'Error')
        $DelimiterBox.Text = ','
    }
    finally {
        if($EncodingBox.SelectedIndex -lt 0 -or $EncodingBox.SelectedIndex -ge $EncodingBox.Items.Count) {
            $EncodingBox.SelectedIndex = $EncodingBox.Items.IndexOf('Default')
        }
        if($DelimiterBox.Text -eq '' -or $DelimiterBox.Text.Length -gt 1) {
            $DelimiterBox.Text = ','
        }
    }

}

#endregion

#region functionalities

#this function will reload preview if a different encoding or delimiter is selected

function Get-UpdatedPreview {
    if($SourceFileBox.Text -eq '') {
        return #to prevent error popups if delimiter and character encoding boxes are changed w/ no file loaded
    }
    try {
        $script:ImportedData = Import-Csv -Path $SourceFileBox.Text -Encoding $EncodingBox.Text -Delimiter $DelimiterBox.Text #variable scope is required to make the variable visible outside the function
        $PreviewBox.Text = $ImportedData | 
        Format-Table |  #to avoid formatting as a list if the number of columns in the file would make PS format data as a list
        Out-String
    }
    catch {
        [System.Windows.MessageBox]::Show($Error[0], 'Error', 'OK', 'Error')
    }
}

#open file

$OpenFileButton.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $Result = $OpenFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $SourceFileBox.Text = $OpenFileDialog.FileName #populate the box with input file path
        Get-UpdatedPreview #update the preview
    }
})

#update previews
$EncodingBox.Add_SelectedIndexChanged({Get-UpdatedPreview})
$DelimiterBox.Add_SelectedIndexChanged({Get-UpdatedPreview})
$DelimiterBox.Add_TextChanged({Get-UpdatedPreview})

#save file
function ConvertTo-HTML {
    #
    $outputString += "<table>`n<tr>"
    forEach($property in $ImportedData[0].PSObject.Properties) {
        Write-Host $property.Name
        $outputString += "<th>$($property.Name)</th>"
    }

    $outputString += "</tr>`n"

    forEach($item in $ImportedData) {
        $outputString += '<tr>'
        forEach($property in $item.PSObject.Properties) {
            $outputString += "<td>$($property.Value)</td>"
        }
        $outputString += "</tr>`n"
    }
    $outputString += "</table>"

    return $outputString
}

$SaveFileButton.Add_Click({
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "HTML Files (*.htm. *.html)|*.htm; *.html"
    $result = $saveFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        #$PreviewBox.SaveFile($saveFileDialog.FileName, [System.Windows.Forms.RichTextBoxStreamType]::PlainText)
        ConvertTo-HTML | Out-File -Encoding UTF8 -FilePath $SaveFileDialog.FileName
    }
})


#save default settings

$SaveDefaultsButton.Add_Click({
    $Configuration = "Encoding = $($EncodingBox.Text)`nDelimiter = $($DelimiterBox.Text)"
    try {
        $Configuration | Out-File -FilePath $iniFile -Encoding UTF8
    }
    catch {
        [System.Windows.MessageBox]::Show($Error[0], 'Error', 'OK', 'Error')
    }
})


#endregion

#read default settings

Get-SavedDefaults


#show the form
$WindowForm.ShowDialog()
