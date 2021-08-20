'# 8/20/21
'# Script created by Richard Martinez
'# This script will rename computer via cscript.exe during imaging process
'#
'#
'############################################################################################
'Assign value to variables outside of functions
strChassisTypeName = GetChassisTypeName()
strComputerSerial = GetBIOSSerial()
'String functions together to get final computer name
strComputerName = "COB-"+strChassisTypeName+"-"+strComputerSerial


'############################################################################################
'This function takes care of getting Chassis Type from computer.
Function GetChassisTypeName()
		Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set objResults = objWMI.InstancesOf("Win32_SystemEnclosure")
		For Each objInstance In objResults
		Select Case objInstance.ChassisTypes(0)
			Case "8", "9", "10", "11", "12", "14", "18", "21", "30", "31", "32"
				strChassisTypeName = "L"
			Case "3", "4", "5", "6", "7", "15", "16", "35", "36"
				strChassisTypeName = "D"
		End Select
		Exit For
	Next
	'Return chassis type to function and store
	GetChassisTypeName = strChassisTypeName
	Exit Function
End Function

'This function takes care of harvesting computer serial (Dell service tag in this case).
Function GetBIOSSerial()
		Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set objBSS = objWMI.ExecQuery("Select * from Win32_BIOS",,48)
		For Each ObjSerial In objBSS
		strComputerSerial = ObjSerial.SerialNumber
		Exit For
	Next
	'Return computer serial to function and store
	GetBIOSSerial = strComputerSerial
	Exit Function
End Function

'# Deprecated code, here for the reminder of failures.
'#
'#	Function Win32Rename()
'#		Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
'#		Set colCName = objWMI.ExecQuery("Select * from Win32_ComputerSystem")
'#			For Each objComputerR In colCName
'#			strComputerName = objComputerR.Rename("COB-"+strChassisTypeName+"-"+strComputerSerial)
'#			Exit For
'#		Next
'#	Win32Rename = strComputerName
'#		Exit Function
'# 	End Function

'Store strComputerName in a much simpler named variable to pass on to operation
Name = strComputerName
'I think these fields can be deprecated
Password = ""
Username = ""

Set objWMIService = GetObject("Winmgmts:root\cimv2")

' Call always gets only one Win32_ComputerSystem object.
For Each objComputer in _
    objWMIService.InstancesOf("Win32_ComputerSystem")

        Return = objComputer.rename(Name,Password,Username)
        'If Return <> 0 Then 'Deprecating IF because it is not needed for the purposes that this scrip will be used for.
        '   WScript.Echo "Rename failed. Error = " & Err.Number
        'Else
        '   WScript.Echo "Rename succeeded." & _
        '       " Reboot for new name to go into effect"
        'End If

Next
'This is here to test and make sure the value of strComputerName is the sum of strChassisTypeName and strComputerSerial.
'Wscript.echo(strComputerName)
