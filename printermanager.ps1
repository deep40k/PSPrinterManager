[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="PSPrinterManager" Height="450" Width="800" MinHeight="450" MinWidth="800">
    <Grid>
        <GroupBox Header="Printer Operations" HorizontalAlignment="Left" Height="158" Margin="10,10,0,0" VerticalAlignment="Top" Width="215">
            <Grid HorizontalAlignment="Left" Height="148" Margin="0,0,-2,-12" VerticalAlignment="Top" Width="205">
                <Label Content="PC Name:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
                <TextBox Name="PrinterLabel" HorizontalAlignment="Left" Height="23" Margin="77,13,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                <Button Name="RetrievePortsButton" Content="Get Ports" HorizontalAlignment="Left" Margin="10,46,0,0" VerticalAlignment="Top" Width="90" Height="30"/>
                <Button Name="RetrievePrintersButton" Content="Get Printers" HorizontalAlignment="Left" Margin="107,46,0,0" VerticalAlignment="Top" Width="90" Height="30"/>
                <Button Name="UninstallButton" Content="Uninstall Printer / Delete Port" HorizontalAlignment="Left" Margin="10,99,0,0" VerticalAlignment="Top" Width="187" Height="30"/>
            </Grid>
        </GroupBox>
        <GroupBox Header="New Printer Operations" HorizontalAlignment="Left" Height="185" Margin="10,183,0,0" VerticalAlignment="Top" Width="215">
            <Grid HorizontalAlignment="Left" Height="175" VerticalAlignment="Top" Width="205" Margin="0,0,-2,-12">
                <Label Content="IP Address:" HorizontalAlignment="Left" Margin="0,10,0,0" VerticalAlignment="Top"/>
                <Label Content="Name:" HorizontalAlignment="Left" Margin="0,41,0,0" VerticalAlignment="Top"/>
                <Label Content="Driver:" HorizontalAlignment="Left" Margin="0,72,0,0" VerticalAlignment="Top"/>
                <TextBox Name="EnterIPBox" HorizontalAlignment="Left" Height="23" Margin="73,13,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="129"/>
                <TextBox Name="EnterNameBox" HorizontalAlignment="Left" Height="23" Margin="49,44,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="153"/>
                <ComboBox Name="PrinterDrivers" HorizontalAlignment="Left" Margin="49,76,0,0" VerticalAlignment="Top" Width="153"/>
                <Button Name="AddPrinter" Content="Add Printer" HorizontalAlignment="Left" Margin="49,115,0,0" VerticalAlignment="Top" Width="109" Height="29"/>
            </Grid>
        </GroupBox>
        <DataGrid Name="PrinterList" Margin="252,10,10,10" IsReadOnly="True">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Printer" Binding="{Binding Printer}" Width="177"/>
                <DataGridTextColumn Header="Port" Binding="{Binding Port}" Width="180"/>
            </DataGrid.Columns>
        </DataGrid>
    </Grid>
</Window>
'@

$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

$RetrievePrintersButton.Add_Click(
    {
        $PrinterThing = Get-Printer -ComputerName $PrinterLabel.Text
        $PrintDrivers = Get-PrinterDriver -ComputerName $PrinterLabel.Text
        $Global:CurrentComputer = $PrinterLabel.Text
        $PrinterDrivers.Items.Clear()
        foreach($Item in $PrintDrivers.Name){
        $PrinterDrivers.Items.Add($Item)
        }
        $PrinterList.Items.Clear()
        $PrinterList.Columns[0].Header = "Printers on $Global:CurrentComputer"
        $PrinterList.Columns[1].Header = "Ports on $Global:CurrentComputer"
        foreach($Thing in $PrinterThing){
        $PrinterList.Items.Add([pscustomobject]@{"Printer" = $Thing.Name; "Port" = $Thing.PortName})
        }
        $PrinterList.Items.Refresh()
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
        $PrinterList.Items.Clear()
        $PrinterList.Columns[1].Header = "Ports on $Global:CurrentComputer"
        foreach($Port in $PrinterPorts){
        $PrinterList.Items.Add([pscustomobject]@{"Printer" = ""; "Port" = $Port.Name})
        }
        $PrinterList.Items.Refresh()
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

$Form.ShowDialog() | Out-Null
