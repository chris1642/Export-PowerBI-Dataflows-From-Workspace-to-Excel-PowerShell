

1. Every Dataflow is downloaded and saved as a text format. 
2. The informarion is parsed and organized within an Excel file, detailing the dataflows, individual queries, and the query steps. 
3. This also appends to a master Excel file, allowing a dated history of the changes between each run. 
 
This allows you to not only easily understand what each dataflow is doing, but also helps have an easy backup method and historical view of dataflows if you don’t use github/devops (we do, but still, this is easier to digest compared to the json file).
 
The PowerShell script is set to auto-install every required module, set the correct permissions, and create the folders. The only input genuinely needed is the workspace ID. There is no API/security key required, as it just uses a pop-up MS login.
