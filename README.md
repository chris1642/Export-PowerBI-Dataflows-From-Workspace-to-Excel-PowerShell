# The PowerShell script does the following:

1. Downloads every Dataflow within a workspace and saves them as a txt (Text) file in a dated-backup folder
2. Parses out the file and extracts the individual query information
3. Combines the query information to an Excel table and saves the Table in the dated backup folder
4. Appends the data to a second master backup excel file that allows you to have a history of your Dataflows for each time you run the script.

The PowerShell script is set to auto-install every required module, set the correct permissions, and create the folders. The only input required is the workspace ID. There is no API/security key required, as it just uses a pop-up MS login.

The only 'requirement' is replacing the XXXXXXXX-XXXX-XXXX-XXXXXXXXXXXX with the Workspace ID you are trying to export the dataflows from. 

You must have read access to the dataflow (if you can refresh a model with a dataflow, you have read access). Depending on if you can run scripts directly, I've included a text file and PS1 copy.

Example excel output

![image](https://github.com/user-attachments/assets/30ec67fb-f059-427e-a878-110bcf1ab066)

