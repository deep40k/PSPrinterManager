Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$CurrentComputer


#Initialize the Form
$Form = New-Object System.Windows.Forms.Form
$Form.Width = 800
$Form.Height = 450
$Form.MinimumSize = New-Object System.Drawing.Size(800,450)
$Form.Text = "Printer Tool"

#Misc Labels
$EnterIPLabel = New-Object System.Windows.Forms.Label
$EnterIPLabel.Text = "Enter the Printer's IP Address"
$EnterIPLabel.Location = New-Object System.Drawing.Size(30,250)
$EnterIPLabel.Size = New-Object System.Drawing.Size(155,25)
$EnterPortLabel = New-Object System.Windows.Forms.Label
$EnterPortLabel.Text = "Enter the Name of the Printer"
$EnterPortLabel.Location = New-Object System.Drawing.Size(220,250)
$EnterPortLabel.Size = New-Object System.Drawing.Size(160,25)
$SelectDriverLabel = New-Object System.Windows.Forms.Label
$SelectDriverLabel.Text = "Select the printer driver"
$SelectDriverLabel.Location = New-Object System.Drawing.Size(230,319)
$SelectDriverLabel.Size = New-Object System.Drawing.Size(160,30)

#IP Address Text Box
$EnterIPBox = New-Object System.Windows.Forms.TextBox
$EnterIPBox.Location = New-Object System.Drawing.Size(30,275)
$EnterIPBox.Size = New-Object System.Drawing.Size(150,25)

#PC Name TextBox
$EnterNameBox = New-Object System.Windows.Forms.TextBox
$EnterNameBox.Location = New-Object System.Drawing.Size(220,275)
$EnterNameBox.Size = New-Object System.Drawing.Size(150,25)

#Enter PC Name
$PrinterLabel = New-Object System.Windows.Forms.TextBox
$PrinterLabel.Text = 'Enter PC Name'
$PrinterLabel.Location = New-Object System.Drawing.Size(75,10)
$PrinterLabel.Size = New-Object System.Drawing.Size(250,30)

#Retrieve Ports Button
$RetrievePortsButton = New-Object System.Windows.Forms.Button
$RetrievePortsButton.Text = "Retrieve Ports"
$RetrievePortsButton.Location = New-Object System.Drawing.Size(30,50)
$RetrievePortsButton.Size = New-Object System.Drawing.Size(150,50)

#Add Printer Button
$AddPrinter = New-Object System.Windows.Forms.Button
$AddPrinter.Text = "Add Printer"
$AddPrinter.Location = New-Object System.Drawing.Size(30,320)
$AddPrinter.Size = New-Object System.Drawing.Size(150,50)

#Retrieve Printers Button
$RetrievePrintersButton = New-Object System.Windows.Forms.Button
$RetrievePrintersButton.Text = "Retrieve Printers"
$RetrievePrintersButton.Location = New-Object System.Drawing.Size(220,50)
$RetrievePrintersButton.Size = New-Object System.Drawing.Size(150,50)

#Uninstall Button
$UninstallButton = New-Object System.Windows.Forms.Button
$UninstallButton.Text = "Uninstall Selected Printer/Delete Selected Port"
$UninstallButton.Location = New-Object System.Drawing.Size(220,130)
$UninstallButton.Size = New-Object System.Drawing.Size(150,50)

#Printer Grid List
$PrinterList = New-Object System.Windows.Forms.DataGridView
$PrinterList.ReadOnly = $true
$PrinterList.ColumnCount = 2
$PrinterList.Columns[0].Name = "Printers"
$PrinterList.Columns[0].Width = 177
$PrinterList.Columns[1].Name = "Ports"
$PrinterList.Columns[1].Width = 180
$PrinterList.Size = New-Object System.Drawing.Size(400,400)
$PrinterList.Location = New-Object System.Drawing.Size(380,5)
$PrinterList.Anchor = 'Top, Bottom, Left, Right'

#Driver List Dropdown
$PrinterDrivers = New-Object System.Windows.Forms.ComboBox
$PrinterDrivers.DropDownStyle = "DropDownList"
$PrinterDrivers.Location = New-Object System.Drawing.Size(220,335)
$PrinterDrivers.Size = New-Object System.Drawing.Size(150,20)

#List of Drivers Installed


#Button Events
$RetrievePrintersButton.Add_Click(
    {
        $PrinterThing = Get-Printer -ComputerName $PrinterLabel.Text
        $PrintDrivers = Get-PrinterDriver -ComputerName $PrinterLabel.Text
        $Global:CurrentComputer = $PrinterLabel.Text
        $PrinterDrivers.Items.Clear()
        foreach($Item in $PrintDrivers.Name){
        $PrinterDrivers.Items.Add($Item)
        }
        $PrinterList.RowCount = $PrinterThing.Count
        $PrinterList.Rows.Clear()
        $PrinterList.Columns[0].Name = "Printers on $Global:CurrentComputer"
        $PrinterList.Columns[1].Name = "Ports on $Global:CurrentComputer"
        foreach($Thing in $PrinterThing){
        $PrinterList.Rows.Add($Thing.Name, $Thing.PortName)
        }
    }
)
$RetrievePortsButton.Add_Click(
    {
        $PrinterPorts = Get-PrinterPort -ComputerName $PrinterLabel.Text
        $Global:CurrentComputer = $PrinterLabel.Text
        $PrintDrivers = Get-PrinterDriver -ComputerName $PrinterLabel.Text
        $PrinterDrivers.Items.Clear()
        foreach($Item in $PrintDrivers.Name){
        $PrinterDrivers.Items.Add($Item)
        }
        $PrinterList.RowCount = $PrinterPorts.Count
        $PrinterList.Rows.Clear()
        $PrinterList.Columns[1].Name = "Ports on $Global:CurrentComputer"
        foreach($Port in $PrinterPorts){
        $PrinterList.Rows.Add("",$Port.Name)
        }
    }
)

$AddPrinter.Add_Click(
    {
        $Global:CurrentComputer = $PrinterLabel.Text
        Add-PrinterPort -Name $EnterIPBox.Text -PrinterHostAddress $EnterIPBox.Text
        Add-Printer -ComputerName $Global:CurrentComputer -Name $EnterNameBox.Text -DriverName $PrinterDrivers.SelectedItem -PortName $EnterIPBox.Text
    }
)

$UninstallButton.Add_Click(
    {
        $SelectedCell = $PrinterList.SelectedCells
        switch($SelectedCell.ColumnIndex){
        0{
        Remove-Printer -ComputerName $Global:CurrentComputer -Name $SelectedCell.Value
        }
        1{
        Remove-PrinterPort -ComputerName $Global:CurrentComputer -Name $SelectedCell.Value
        }
        }
    }
)

#Add the Controls to the Form
$Form.Controls.Add($PrinterDrivers)
$Form.Controls.Add($PrinterLabel)
$Form.Controls.Add($RetrievePrintersButton)
$Form.Controls.Add($UninstallButton)
$Form.Controls.Add($PrinterList)
$Form.Controls.Add($RetrievePortsButton)
$Form.Controls.Add($EnterIPLabel)
$Form.Controls.Add($EnterPortLabel)
$Form.Controls.Add($EnterIPBox)
$Form.Controls.Add($EnterNameBox)
$Form.Controls.Add($AddPrinter)
$Form.Controls.Add($SelectDriverLabel)
$Form.ShowDialog()
