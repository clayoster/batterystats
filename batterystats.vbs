'This script determines the percentage of battery wear and battery model
' of one or two installed batteries and writes the output to a popup box,
'csv file on a file server, or registry values under a specified key.

ON ERROR RESUME NEXT
Set objWMIService = GetObject("winmgmts:\\.\root\WMI")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set shell = WScript.CreateObject( "WScript.Shell" )
Set objNetwork = CreateObject("Wscript.Network")
Set oShellEnv = Shell.Environment("Process")
computerName  = oShellEnv("ComputerName")

'Define variables
Dim devicename, fullchargedcapacity, designedcapacity, batterywear
Dim devicename1, fullchargedcapacity1, designedcapacity1, batterywear1
Dim devicename2, fullchargedcapacity2, designedcapacity2, batterywear2
Dim devicenameArray, designedcapacityArray, fullchargedcapacityArray
Dim driveLetter, fileSharePath, outputFilePath, fsUserName, fsPassword
Dim battery1ModelPath, battery2ModelPath, battery1WearPath, battery2WearPath

'****Set variables specific to your needs here!****

'Set variables for the fnWriteFileshare function here:
'Since this info is set static in the script, use a very limited user account.
'Uncomment the variables and fill out with appropriate information

'driveLetter = "I:"
'fileSharePath = "\\example.domain.com\example"
'outputFilePath = "\batterystats\example.txt"
'fsUserName = ""
'fsPassword = ""

'Set variables for the fnRegWrite function here:
battery1ModelPath = "HKLM\Software\BatteryStats\Battery1Model"
battery2ModelPath = "HKLM\Software\BatteryStats\Battery2Model"
battery1WearPath = "HKLM\Software\BatteryStats\Battery1Wear"
battery2WearPath = "HKLM\Software\BatteryStats\Battery2Wear"

'****On to data gathering and functions!****

'Query WMI for BatteryStaticData to output items DeviceName, SerialNumber,
    'ManufactureName, UniqueID, and Designed Capacity
Set batterydata = objWMIService.ExecQuery("Select * From BatteryStaticData")
For Each objItem in batterydata
    devicename = devicename & objItem.devicename & ","
    designedcapacity = designedcapacity & objItem.DesignedCapacity & ","
Next

'Query WMI for BatteryFullChargedCapacity to output items FullChargedCapacity
Set FullCharge = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")
For Each objItem in FullCharge
    fullchargedcapacity = fullchargedcapacity & objItem.fullchargedcapacity & ","
Next

'Split DeviceName array to get values for both batteries (if applicable)
DeviceNameArray = Split(DeviceName, ",", -1, 1)
DeviceName1 = DeviceNameArray(0)
DeviceName2 = DeviceNameArray(1)

'Split designedcapacity array to get values for both batteries (if applicable)
designedcapacityArray = Split(designedcapacity, ",", -1, 1)
designedcapacity1 = designedcapacityArray(0)
designedcapacity2 = designedcapacityArray(1)

'Split fullchargedcapacity array to get values for both batteries (if applicable)
fullchargedcapacityArray = Split(fullchargedcapacity, ",", -1, 1)
fullchargedcapacity1 = fullchargedcapacityArray(0)
fullchargedcapacity2 = fullchargedcapacityArray(1)

If NOT designedcapacity1 = "" Then
    batterywear1 = round(((fullchargedcapacity1 / designedcapacity1) * 100),2)
End If
If NOT designedcapacity2 = "" Then
    batterywear2 = round(((fullchargedcapacity2 / designedcapacity2) * 100),2)
End If

'Determine if there were arguments passed to the script and execute functions accordingly
If Wscript.Arguments.Count = 0 Then
    fnDisplayPopup
Else
    For i = 0 to Wscript.Arguments.Count - 1
        If Wscript.Arguments(i) = "/f" Then
            fnWriteFileshare
        ElseIf Wscript.Arguments(i) = "/r" Then
            fnWriteReg
        ElseIf Wscript.Arguments(i) = "/?" Then
            fnHelp
            ElseIf Wscript.Arguments(i) = "-?" Then
            fnHelp
    End If
    Next
End If

'****Data output functions!****

'Function to display output in a Popup Window
Function fnDisplayPopup ()
    If NOT designedcapacity1 = "" Then
        popupoutput1 = "Battery1 Model: " & devicename1 & " | Remaining Capacity: " _
                    & batterywear1 & "%"
    End If
    If NOT designedcapacity2 = "" Then
        popupoutput2 = "Battery2 Model: " & devicename2 & " | Remaining Capacity: " _
                    & batterywear2 & "%"
    End If

    wscript.echo popupoutput1 & vbCrLf & popupoutput2 & vbCrLf & vbCrLf _
        & "Use the /? switch to see more options"
end function

'Write values to the registry to be gathered by LANDesk Inventory Scan
Function fnWriteReg ()
    If NOT designedcapacity1 = "" Then
        shell.RegWrite battery1ModelPath, devicename1, "REG_SZ"
        shell.RegWrite battery1WearPath, batterywear1, "REG_SZ"
    End If
    If NOT designedcapacity2 = "" Then
        shell.RegWrite battery2ModelPath, devicename2, "REG_SZ"
        shell.RegWrite battery2WearPath, batterywear2, "REG_SZ"
    End If
end function

'Write values to a file on a network file share in CSV format
Function fnWriteFileshare ()
    objNetwork.MapNetworkDrive driveLetter, fileSharePath, "False", fsUserName, fsPassword
    set LOGFILE = objFSO.opentextfile (driveLetter & outputFilePath, 8, true)
    Logfile.writeline(computerName & "," & DeviceName1 & "," & batterywear1 & "," _
            & DeviceName2 & "," & batterywear2 & ",")
    Logfile.Close
    objNetwork.RemoveNetworkDrive driveLetter
end function

'Output description of available command arguments
Function fnHelp ()
wscript.echo "Available command arguments: " & vbCrLf _
    & "/f - Write Battery Stats to a file share. (Set related variables inside script)" & vbCrLf _
    & "/r - Write Battery Stats to the registry. (Set related variables inside script)" & vbCrLf _
    & "/? - Displays this help output"
end function

wscript.quit