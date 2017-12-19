# $language = "VBScript"
# $interface = "1.0"

  crt.Screen.Synchronous = True

sub Main

  ' If folder doesn't exist create it
  '
  Dim filesys, newfolder
  set filesys=CreateObject("Scripting.FileSystemObject")
  If  Not filesys.FolderExists("C:\00 - SecureCRT Logfiles") Then
     newfolder = filesys.CreateFolder ("C:\00 - SecureCRT Logfiles")
  End If

  ' Get hostname to be used in the logfile filename
  '
   crt.Screen.Send "show system" & vbCr : crt.Screen.WaitForString vbCr
   strHNPre = crt.Screen.ReadString ("#")
   strHN = Split(Mid(strHNPre, 32), vbcrlf)(0)
   
  ' Set logfile name & turn on logging
  '
  logfile = "C:\00 - SecureCRT Logfiles\%Y.%M.%D-%h.%m.%s-" & strHN & ".log"
  crt.Session.LogFileName = logfile
  crt.Session.Log True


  '
  '
  ' do stuff here
  '
  '

  ' Stop logging
  '
  crt.Session.Log False

End sub