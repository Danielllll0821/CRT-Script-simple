#$language = "VBScript"
#$interface = "1.0"


	crt.Screen.Send "show memo statis" & vbcr
	   Result = crt.Screen.ReadString ("#")	   
       result = Trim(Result)
	   msgbox result
	   line = Split(result, vbcr)(2)
	   msgbox line
	   s1 = Replace(line, " ", "#")
	   msgbox s1
	   msgbox Len(s1)
	   a = left(Trim(s1),10)
	   b = right(s1,(Len(s1)-10))
	   s2 = a & "#" & b
	   msgbox s2
	   
	  
	   REM line11 = Trim(line1)
	   REM line111 = LTrim(line1)
	   REM line2 = Split(Result, vbcr)(3)
	   REM msgbox line11
	    REM msgbox line111
	   processor_total = Split(s2,"####")(2)
	   processor_used = Split(s2,"####")(3)
	   processor_free = Split(s2,"####")(4)
	   msgbox processor_total
	   msgbox processor_used
	   msgbox processor_free
	   used_per = processor_used / processor_total
	   u_p = round(used_per,2) 
	    msgbox u_p
	   REM msgbox line2
	   
	   
	   REM strHN = Split(Result,":")(1) 
       REM msgbox strHN
	   REM '''msgbox Mid(Result, 21)
	   REM '''strHN = Split(Mid(Result, 21),vbCr)(1)	'第2种方法，用Mid函数提取，在用Split提取R1   
	   REM '''msgbox strHN
	   REM hn = Split(strHN)(0)
	   REM '''msgbox hn	   
	   REM HN = Mid(hn,2)  