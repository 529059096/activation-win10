Dim WshShell,fso,file
Set WshShell = Wscript.CreateObject("WScript.Shell")
set fso = createobject("scripting.filesystemobject")
Rem 管理员权限
If WScript.Arguments.Length = 0 Then 
  Set ObjShell = CreateObject("Shell.Application") 
  ObjShell.ShellExecute "wscript.exe" _ 
  , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1 
  WScript.Quit 
End If 
REM 服务器赋值
Dim KMSServer(20)
 KMSServer(0)="web.529059096.xyz:16888"
 KMSServer(1)="kms.cary.tech"
 KMSServer(2)="kms.jm33.me"
 KMSServer(3)="kms.sdit163.com"
 KMSServer(4)="kms.sixyin.com"
 KMSServer(5)="kms.zhuxiaole.org"
 KMSServer(6)="kms.cgtsoft.com"
 KMSServer(7)="kms.zhi.fun"
 KMSServer(8)="windows.kms.app"
 KMSServer(9)="kms.000606.xyz"
 KMSServer(10)="win.kms.pub"
 KMSServer(11)="xincheng213618.cn"
 KMSServer(12)="kms.kmzs123.cn"
 KMSServer(13)="kms.sevenxu.cc"
 KMSServer(14)="kms.catqu.com"
 KMSServer(15)="kms.moeyuuko.top"
 KMSServer(16)="kms.bearlele.top"
 KMSServer(17)="kms.moeyuuko.com"
 KMSServer(18)="kms.yunyize.com"
 KMSServer(19)="kms.yangyingyun.cn"
 KMSServer(20)="kms.hmg.pw"
REM 密钥赋值
Dim Core,CoreCountrySpecific,CoreN,CoreSingleLanguage,ProfessionalStudent,ProfessionalStudentN,Professional,ProfessionalN,ProfessionalSN,ProfessionalWMC,Enterprise,EnterpriseN,Education,EducationN,EnterpriseS,EnterpriseSN
 Core="TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
 CoreCountrySpecific="PVMJN-6DFY6-9CCP6-7BKTT-D3WVR"
 CoreN="3KHY7-WNT83-DGQKR-F7HPR-844BM"
 CoreSingleLanguage="7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"
 ProfessionalStudent="YNXW3-HV3VB-Y83VG-KPBXM-6VH3Q"
 ProfessionalStudentN="8G9XJ-GN6PJ-GW787-MVV7G-GMR99"
 Professional="W269N-WFGWX-YVC9B-4J6C9-T83GX"
 ProfessionalN="MH37W-N47XK-V7XM9-C7227-GCQG9"
 ProfessionalSN="8Q36Y-N2F39-HRMHT-4XW33-TCQR4"
 ProfessionalWMC="NKPM6-TCVPT-3HRFX-Q4H9B-QJ34Y"
 Enterprise="NPPR9-FWDCX-D2C8J-H872K-2YT43"
 EnterpriseN="DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
 Education="NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
 EducationN="2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
 EnterpriseS="WNMTR-4C88C-JK8YV-HQ7T2-76DF9"
 EnterpriseSN="2F77B-TNFGY-69QQF-B8YKP-D69TJ"
REM 系统版本获取
OSVersion = CreateObject("Wscript.Shell").RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\EditionID")
Rem 安装密钥
If OSVersion = "Core" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& Core
ElseIf OSVersion = "CoreCountrySpecific" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& CoreCountrySpecific
ElseIf OSVersion = "CoreN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& CoreN
ElseIf OSVersion = "CoreSingleLanguage" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& CoreSingleLanguage
ElseIf OSVersion = "ProfessionalStudent" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& ProfessionalStudent
ElseIf OSVersion = "ProfessionalStudentN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& ProfessionalStudentN
ElseIf OSVersion = "Professional" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& Professional
ElseIf OSVersion = "ProfessionalN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& ProfessionalN
ElseIf OSVersion = "ProfessionalSN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& ProfessionalSN
ElseIf OSVersion = "ProfessionalWMC" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& ProfessionalWMC
ElseIf OSVersion = "Enterprise" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& Enterprise
ElseIf OSVersion = "EnterpriseN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& EnterpriseN
ElseIf OSVersion = "Education" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& Education
ElseIf OSVersion = "EducationN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& EducationN
ElseIf OSVersion = "EnterpriseS" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& EnterpriseS
ElseIf OSVersion = "EnterpriseSN" Then
	WshShell.Run "slmgr.vbs -ipk" &" "& EnterpriseSN
End If
Rem 链接激活服务器
on error resume next
For i = 0 TO 20
	WshShell.Run "cmd /k cscript //Nologo %windir%\system32\slmgr.vbs -skms"&" "&KMSServer(i)&" "&">c:\KMS.ini &&exit",0,TRUE
	WshShell.Run "cmd /c cscript //Nologo %windir%\system32\slmgr.vbs /ato >c:\KMS.ini",0,TRUE
	'MsgBox "循环：" & i
	set file = fso.opentextfile("c:/KMS.ini",1,true)
	Do Until file.AtEndOfStream 
		content2 = file.ReadLine
		content = file.readall
	Loop
	'MsgBox content
	code = InStrRev(content,"成功")
	'MsgBox code
	If code = 1 Then
		MsgBox "成功地激活了产品。"
		WScript.Quit
	End If
Next
MsgBox "激活未成功，请检查网络链接或稍后再试"