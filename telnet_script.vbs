#$language = "VBScript"
#$interface = "1.0"
'本脚本示范：从一个文件里面自动读取设备IP地址，密码等，自动输入巡检命令，并记录日志到文件。
'  此版本使用的文件名是通过调用参数，读取device.txt文件终端ip地址作为文件名。

Sub Main
    '打开保存设备管理地址以及密码的文件
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso,file1,line,logfile,params
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

	   '创建目录存放日志文件,根据ip地址命名文件	 
	   logfile = "d:\logfile\" & params(0) & "_%Y%M%D%_h%m%s.txt"

	   
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
	  
	   '开启记录日志
	   crt.Session.LogFileName = logfile
	   crt.Session.Log(true) 	   
	   
	   crt.Screen.Send "terminal length 0" & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show ver" & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show env ala" & vbcr	   
       crt.Screen.waitForString "#"
       crt.Screen.Send "show env stat" & vbcr 
       crt.Screen.waitForString "#"
       crt.Screen.Send "show process cpu " & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show process memory" & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show module" & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show logging" & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show clock" & vbcr
       crt.Screen.waitForString "#"
       crt.Screen.Send "show ntp status" & vbcr

	   
       '备份完成后退出
       crt.Screen.waitForString "#",3
       crt.Session.Disconnect
	   'crt.Quit
	 
       loop
    crt.Screen.Synchronous = False
	'crt.Quit 	
End Sub