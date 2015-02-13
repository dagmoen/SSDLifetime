strComputer = "."
SSDInstanceName = "SCSI\Disk&Ven_&Prod_OCZ-AGILITY4\4&15828421&0&010000_0"
set wbemServices = GetObject("winmgmts:\\" & strComputer & "\root\wmi")
set wbemObjectSetVendor = wbemServices.InstancesOf("MSStorageDriver_ATAPISmartData")

strReturn = getLifeTimeRemain()

On Error Resume Next
wscript.Echo strReturn	'for cmd line
Echo strReturn		'for BGInfo
on error goto 0

FUNCTION getLifeTimeRemain()
For Each wbemObject In wbemObjectSetVendor

  lifetimeremain = 0
 
  arrVendorSpecific = wbemObject.VendorSpecific

  IF wbemObject.InstanceName = SSDInstanceName  then
    for i=0 to 359
	    if ((arrVendorSpecific(i) = 0) OR (arrVendorSpecific(i) = 16)) then 
		    i2 = i+1
		    if arrVendorSpecific(i2) = 0 then 
			    i3 = i2+1 
				  i6 = i2+4 
				  if arrVendorSpecific(i3) = 233 then
					  getLifeTimeRemain = arrVendorSpecific(i6)
				  End If	
			  End If
		  End if
    Next
  END IF
Next
END FUNCTION
