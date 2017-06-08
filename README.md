# PSPrinterManager
Manages printers with Powershell, allowing one to remotely delete and add printers and ports to a domain machine.

# Screenshots
![Alt Text](http://i.imgur.com/5yQIU2j.png)

# How to use
Clone Repo or Download Zip and extract to a folder, then run the script with Powershell.

You can click on Retrieve Ports to create a list of Ports on the computer specified in the Enter PC Name text box.
Retrieve Printers retrieves a list of printers on the machine you enter as well as their associated port.
The Uninstall button uninstalls the selected printer or port that you have selected in the list on the right, note that you cannot delete a port that is in use by a printer.
You can add a printer to the specified machine by using the bottom area, currently you can only use drivers that are installed on the machine.

# Features to Add
Remote installation of printer drivers as well.
