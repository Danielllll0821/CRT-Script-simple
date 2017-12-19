#$language = "VBScript"
#$interface = "1.0"
'本脚本示范：从一个文件里面自动读取设备IP地址，密码等，自动将设备配置备份

Sub Main
    '打开保存设备管理地址以及密码的文件
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso,file1,line,logfile,params,hn
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file1 = fso.OpenTextFile("d:\device.txt",Forreading, False)    
    crt.Screen.Synchronous = True
    DO While file1.AtEndOfStream <> True
       '读出每行
       line = file1.ReadLine
       '分离每行的参数 IP地址 密码 En密码
       params = Split (line)
       'Telnet到这个设备上
       crt.Session.Connect "/TELNET " & params(0)
	   '创建目录存放日志文件
	   'logfile = "d:\logfile\" & params(0) & "-%Y%M%D%h%m%s.txt"
	   	   
	   
 
       '输入telnet密码
	   crt.Screen.WaitForString "Username:"
       crt.Screen.Send params(1) & vbcr
       crt.Screen.WaitForString "Password:"
       crt.Screen.Send params(2) & vbcr
       '进特权模式
       crt.Screen.Send "enable" & vbcr
       crt.Screen.WaitForString "Password:"
       crt.Screen.Send params(3) & vbcr
       crt.Screen.waitForString "#"

	   crt.Screen.Send "show run | in host" & vbCr 
	   content = crt.Screen.ReadString("#")	   
       'hn = split(content,vbCr)(2)	

	   strHN = Split(Mid(content, 19),vbCr)(1)
	   msgbox  strHN
	   HN = Mid(StrHN,10)
	   msgbox  HN

       ' Set logfile name & turn on logging
       logfile = "d:\logfile\" & HN &"-%Y.%M.%D-%h.%m.%s.log"
	   crt.Session.LogFileName = logfile
	   crt.Session.Log(true)
	   
	   loop
  End Sub