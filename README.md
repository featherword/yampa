# Yampa
Simple Powershell Scripts


Here lies the more motivated side of my brain.

Random --

  Update-Putty: Updates putty, since it doesnt auto update. It is also able to unisntall older versions before 0.68 when they switched to having .msi installers available.
  
  Move-Old-PCs: Moves pcs that havent signed into the domain for a specified number of months (12 by default) to a new ou named Disabled Computers. This is created if it doesnt exist.





WDS & MDT -- Used during deployment through Microsoft Deployment Toolkit

  Activate-PC: Attempts to active windows with the bios key, if not able to or key is not for that version of windows it looks for a .csv files with MAK/KMS keys.
  
  Install-Automate: Installs Connectwise Automate pulling from a list of locations store in a .csv file. Query the Automate db to keep the .csv locations file updated.
