#Created by: Maxwell Kobe
#Date created: 6/3/2015
#Read classroom use xlsx file and generate email lists for each room
# Adding PS Snapin

Function Generate-Window()
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Email List Generator"
    $form.Size = New-Object System.Drawing.Size(400,300)
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(125,200)
    #$OKButton.Anchor = New-Object System.Windows.Forms.AnchorStyles.Bottom
    $OKButton.Size = New-Object System.Drawing.Size(75,25)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKbutton)


    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(200,200)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(390,30) 
    $label.Text = "Please select the file you wish to generate the email lists from."
    $form.Controls.Add($label)

    $SelectionLabel = New-Object System.Windows.Forms.Label
    $SelectionLabel.Location = New-Object System.Drawing.Point(10, 150)
    $SelectionLabel.Size = New-Object System.Drawing.Size(390, 25)
    $form.Controls.Add($SelectionLabel)

    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Location = New-Object System.Drawing.Point(155,80)
    $BrowseButton.Size = New-Object System.Drawing.Size(90,25)
    $BrowseButton.Text = "Select File"
    $Browser = New-Object System.Windows.Forms.OpenFileDialog
    $BrowseButton.Add_Click(
        {
            $Browser.ShowDialog()
            $SelectionLabel.Text = $Browser.Filename

        })
    $form.Controls.Add($BrowseButton)


    $form.Topmost = $True #show form as topmost window
    #$form.Add_Shown({$input.Select()})

    $result = $form.ShowDialog()

    If($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        #Parse-XLSX -filename $Browser.FileName
        Parse-CSV -FileName $Browser.FileName
    }

}

Function Parse-CSV($FileName)
{
    $dir = ($FileName).Replace('.csv','')
    New-Item -ItemType directory -path $dir -Force

    $doc = Import-Csv $FileName
    #$doc | Where-Object {$_.BLDG -eq "AGS-E" -and $_.ROOM -eq "1001"} | Select Email -Unique
    $bldgs = $doc | select BLDG -Unique
    Foreach($b in $bldgs)
    {
        $rms = $doc | Where-Object{$_.BLDG -eq $b.BLDG} | Select ROOM -Unique

        Foreach($rm in $rms)
        {
            
            $tofile = ($doc | Where-Object{$_.BLDG -eq $b.BLDG -and $_.ROOM -eq $rm.ROOM -and $_.EMAIL -ne "#N/A"} | Select EMAIL -Unique)
            [string]$toFileName = $b.BLDG.toString()+"_"+$rm.ROOM.toString() 
            Out-File -FilePath $toFileName -Append $tofile.toString() -Force
        }
    }
}

Function Create-Outlook-Group($grouplist)
{
    
}


Function Parse-XLSX($filename)
{

    Test-Path $filename
    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.Visible = $True
    $wb = $objExcel.Workbooks.Open($filename)
    $worksheet = $wb.Sheets.item(($wb.Sheets | select Name -First 1).Name)
    
    
}


Generate-Window