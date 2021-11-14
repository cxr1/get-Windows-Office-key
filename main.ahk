RunWith("admin")
;当前AHK版本 := (!A_IsUnicode) ? "ANSI" : (A_PtrSize=4) ? "Unicode 32" : "Unicode 64"
;当前权限    := (A_IsAdmin=1)  ? "管理员权限" : "普通权限"
;MsgBox, % "当前AHK版本: " 当前AHK版本 "`n`n当前权限: " 当前权限
office_licence = 
(%
 cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /dstatus
)
recmd := RunCmd(office_licence)
;MsgBox, % recmd
FileAppend,
(

#################################################################################

-----------------------------------------------OFFICE licence------------------------------------------------------

#################################################################################

), OFFICE&WINDOWS密钥信息.txt
FileAppend, % recmd, OFFICE&WINDOWS密钥信息.txt
windows_licence =
(%
 wmic path softwarelicensingservice get OA3xOriginalProductKey
)
recmd := RunCmd(windows_licence)
FileAppend,
(

#################################################################################

-----------------------------------------------WINDOWS licence--------------------------------------------------
                                 
#################################################################################

), OFFICE&WINDOWS密钥信息.txt
FileAppend, % recmd, OFFICE&WINDOWS密钥信息.txt

RunCmd(CmdLine, WorkingDir:="", Cp:="CP0") { ; Thanks Sean!  SKAN on D34E @ tiny.cc/runcmd 
  Local P8 := (A_PtrSize=8),  pWorkingDir := (WorkingDir ? &WorkingDir : 0)                                                
  Local SI, PI,  hPipeR:=0, hPipeW:=0, Buff, sOutput:="",  ExitCode:=0,  hProcess, hThread
                   
  DllCall("CreatePipe", "PtrP",hPipeR, "PtrP",hPipeW, "Ptr",0, "UInt",0)
, DllCall("SetHandleInformation", "Ptr",hPipeW, "UInt",1, "UInt",1)
    
  VarSetCapacity(SI, P8? 104:68,0),      NumPut(P8? 104:68, SI)
, NumPut(0x100, SI,  P8? 60:44,"UInt"),  NumPut(hPipeW, SI, P8? 88:60)
, NumPut(hPipeW, SI, P8? 96:64)   
, VarSetCapacity(PI, P8? 24:16)               
  If not DllCall("CreateProcess", "Ptr",0, "Str",CmdLine, "Ptr",0, "UInt",0, "UInt",True
              , "UInt",0x08000000 | DllCall("GetPriorityClass", "Ptr",-1,"UInt"), "UInt",0
              , "Ptr",pWorkingDir, "Ptr",&SI, "Ptr",&PI )  
     Return Format( "{1:}", "" 
          , DllCall("CloseHandle", "Ptr",hPipeW)
          , DllCall("CloseHandle", "Ptr",hPipeR)
          , ErrorLevel := -1 )
  DllCall( "CloseHandle", "Ptr",hPipeW)
, VarSetCapacity(Buff, 4096, 0), nSz:=0   
  While DllCall("ReadFile",  "Ptr",hPipeR, "Ptr",&Buff, "UInt",4094, "PtrP",nSz, "UInt",0)
    sOutput .= StrGet(&Buff, nSz, Cp)
  hProcess := NumGet(PI, 0),  hThread := NumGet(PI,4)
, DllCall("GetExitCodeProcess", "Ptr",hProcess, "PtrP",ExitCode)
, DllCall("CloseHandle", "Ptr",hProcess),    DllCall("CloseHandle", "Ptr",hThread)
, DllCall("CloseHandle", "Ptr",hPipeR),      ErrorLevel := ExitCode  
Return sOutput  
}


RunWith(RunAsAdmin:="Default", ANSI_U32_U64:="Default")
{
	; 格式化预期的模式
	switch, RunAsAdmin
	{
		case "Normal","Standard","No","0":		RunAsAdmin:=0
		case "Admin","Yes","1":								RunAsAdmin:=1
		case "Default":												RunAsAdmin:=A_IsAdmin
		default:															RunAsAdmin:=A_IsAdmin
	}
	switch, ANSI_U32_U64
	{
		case "A32","ANSI","A":								ANSI_U32_U64:="AutoHotkeyA32.exe"
		case "U32","X32","32":								ANSI_U32_U64:="AutoHotkeyU32.exe"
		case "U64","X64","64":								ANSI_U32_U64:="AutoHotkeyU64.exe"
		case "Default":												ANSI_U32_U64:="AutoHotkey.exe"
		default:															ANSI_U32_U64:="AutoHotkey.exe"
	}
	; 获取传递给 “.ahk” 的用户参数（不是 /restart 之类传递给 “.exe” 的开关参数）
	for k, v in A_Args
	{
		if (RunAsAdmin=1)
		{
			; 转义所有的引号与转义符号
			v:=StrReplace(v, "\", "\\")
			v:=StrReplace(v, """", "\""")
			; 无论参数中是否有空格，都给参数两边加上引号
			; Run       的内引号是 "
			ScriptParameters  .= (ScriptParameters="") ? """" v """" : A_Space """" v """"
		}
		else
		{
			; 转义所有的引号与转义符号
			; 注意要转义两次 Run 和 RunAs.exe
			v:=StrReplace(v, "\", "\\")
			v:=StrReplace(v, """", "\""")
			v:=StrReplace(v, "\", "\\")
			v:=StrReplace(v, """", "\""")
			; 无论参数中是否有空格，都给参数两边加上引号
			; RunAs.exe 的内引号是 \"
			ScriptParameters .= (ScriptParameters="") ? "\""" v "\""" : A_Space "\""" v "\"""
		}
	}
	; 判断当前 exe 是什么版本
	if (!A_IsUnicode)
		RunningEXE:="AutoHotkeyA32.exe"
	else if (A_PtrSize=4)
		RunningEXE:="AutoHotkeyU32.exe"
	else if (A_PtrSize=8)
		RunningEXE:="AutoHotkeyU64.exe"
	; 运行模式与预期相同，则直接返回。 ANSI_U32_U64="AutoHotkey.exe" 代表不对 ahk 版本做要求。
	if (A_IsAdmin=RunAsAdmin and (ANSI_U32_U64="AutoHotkey.exe" or ANSI_U32_U64=RunningEXE))
		return
	; 如果当前已经是使用 /restart 参数重启的进程，则报错避免反复重启导致死循环。
	else if (RegExMatch(DllCall("GetCommandLine", "str"), " /restart(?!\S)"))
	{
		预期权限:=(RunAsAdmin=1) ? "管理员权限" : "普通权限"
		当前权限:=(A_IsAdmin=1)  ? "管理员权限" : "普通权限"
		ErrorMessage=
		(LTrim
		预期使用: %ANSI_U32_U64%
		当前使用: %RunningEXE%
		预期权限: %预期权限%
		当前权限: %当前权限%
		程序即将退出。
		)
		MsgBox 0x40030, 运行状态与预期不一致, %ErrorMessage%
		ExitApp
	}
	else
	{
		; 获取 AutoHotkey.exe 的路径
		SplitPath, A_AhkPath, , Dir
		if (RunAsAdmin=0)
		{
			; 强制普通权限运行
			switch, A_IsCompiled
			{
				; %A_ScriptFullPath% 必须加引号，否则含空格的路径会被截断。%ScriptParameters% 必须不加引号，因为构造时已经加了。
				; 工作目录不用单独指定，默认使用 A_WorkingDir 。
				case, "1": Run, RunAs.exe /trustlevel:0x20000 "\"%A_ScriptFullPath%\" /restart %ScriptParameters%",, Hide
				default: Run, RunAs.exe /trustlevel:0x20000 "\"%Dir%\%ANSI_U32_U64%\" /restart \"%A_ScriptFullPath%\" %ScriptParameters%",, Hide
			}
		}
		else
		{
			; 强制管理员权限运行
			switch, A_IsCompiled
			{
				; %A_ScriptFullPath% 必须加引号，否则含空格的路径会被截断。%ScriptParameters% 必须不加引号，因为构造时已经加了。
				; 工作目录不用单独指定，默认使用 A_WorkingDir 。
				case, "1": Run, *RunAs "%A_ScriptFullPath%" /restart %ScriptParameters%
				default: Run, *RunAs "%Dir%\%ANSI_U32_U64%" /restart "%A_ScriptFullPath%" %ScriptParameters%
			}
		}
		ExitApp
	}
}
