# mailkit
storing code from https://adamtheautomator.com/powershell-email/ to help send secure e-mail from PowerShell
Since the powershell function send-mailmessage is no longer recommended. There are some alternatives and
this repository is to help me with my projects that may need a secure way to send email.
This is still a work in progress, but just as the author from above website anyone can use this code and or add to this. 
Thank you
.
.
.
Although it was thought this should work in almost any version or powershell it only works in powershell version 7 and higher from what I can tell. 

My version 1
* **Install_mailkit.ps1** - This will confirm if mailkit is installed, add the nuget repository if needed and install mailkit and mimekit (dependency) 

* **mailkit.ps1** - will load the function into the shell. Function called Send-Mailkitmessage.

* **Configparser.ps1** - is another function that can read a file (typically ini) with a name=value format. This isn't needed, but is recommended so no passwords or other settings are in a script. 
