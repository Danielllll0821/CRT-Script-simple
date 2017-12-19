#$language = "VBScript"
#$interface = "1.0"
'本脚本示范：从一个文件里面自动读取设备IP地址，密码等，自动输入巡检命令，并记录日志到文件。
'  此版本使用的文件名是提取hostname信息，并作为文件名。并且针对因不能正确提取用户名，取消使用Mid()函数分割结果。
'  1、在用户模式下，通过show ver | in uptime来提取hostname信息；此处因cisco系统版本不同，
'     会有部分结果的第一个字符为空，导致获取不到正确的用户名。
'  2、使用Split(str,vbCr)获得hostname ;取消使用Mid()函数分割结果。
'  3、使用Mid()函数获得hostname。注意Split(str,vbCr)函数获得的结果,第一个元素为空""，因此不能之间应用到文件名中。

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
	   '----------------------------------------------------------
	   '创建目录存放日志文件,根据ip地址命名文件	 
	   'logfile = "d:\logfile\" & params(0) & "-%Y%M%D%h%m%s.txt"
	   '----------------------------------------------------------
	   

	   
       '输入telnet密码
	   crt.Screen.WaitForString "Username:"
       crt.Screen.Send params(1) & vbcr
       crt.Screen.WaitForString "Password:"
       crt.Screen.Send params(2) & vbcr
       '进特权模式
       'crt.Screen.Send "enable" & vbcr
       'crt.Screen.WaitForString "Password:"
       'crt.Screen.Send params(3) & vbcr
       crt.Screen.waitForString ">"
	   
	   '从"show run | in host" 的结果中提取用户名，并代入文件中
	   crt.Screen.Send "show ver | in uptime" & vbCr 
	   Result = crt.Screen.ReadString(">")	   
       
	   '第一种用换行来分割结果，获取hostname R1,但是返回的是数组的元素，无法加入到文件名中，需要Mid函数再次提取
	   strHN = Split(Result,vbCr)(2) 
       'msgbox strHN
	   HN = Mid(strHN,2)  '''hn中包含两行，第一行为空，但是第二行为HOSTNAME，但是第一行要占用一个字符，第二行从2开始。
  	   'msgbox HN
	   
       '设置文件存放目录及文件名，文件名中包含日期
       logfile = "d:\logfile\" & HN & "_%Y.%M.%D_%h.%m.%s_.log"
	  
	   '开启记录日志
	   crt.Session.LogFileName = logfile
	   crt.Session.Log(true) 	   
	   
	   crt.Screen.Send "terminal length 0" & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show ver" & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show env ala" & vbcr	   
       crt.Screen.waitForString ">"
       crt.Screen.Send "show env stat" & vbcr 
       crt.Screen.waitForString ">"
       crt.Screen.Send "show process cpu " & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show process memory" & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show module" & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show logging" & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show clock" & vbcr
       crt.Screen.waitForString ">"
       crt.Screen.Send "show ntp status" & vbcr

	   
       '备份完成后退出
       crt.Screen.waitForString ">",3
       crt.Session.Disconnect
	   'crt.Quit
	 
       loop
    crt.Screen.Synchronous = False
	'crt.Quit 	
End Sub