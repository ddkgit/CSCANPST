;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         josh levine (http://josh.com/cmdscanpst)
;                 minor tweak by GBM to fix some carriage returns, add date/time to log, and add 'end' call for log
;
; Script Function:
;	Allows you to run the Outlook Repair Tool (SCANPST) from a commmand line batch file
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Must use vars as can't use %1% etc in most places
prm1 = %1%	;scanpst path\file
prm2 = %2%	;pst file
prm3 = %3%	;N (no backup) if reqd

;GBM: call this at end of all calls to this script to log end of all repairs. As must pass 2 prms, call using "cscanpst end ."
if prm1=end
{
	FileAppend, %A_DD%/%A_MM% %A_Hour%:%A_Min% end of run`n , cscanpst.log
	ExitApp 0
}

StringTrimRight, strBakFile, prm2, 3
strBakFile = %strBakFile%bak
FileDelete, %strBakFile%
FileAppend, %A_DD%/%A_MM% %A_Hour%:%A_Min% Launched on %prm2%... , cscanpst.log

if 0 < 2
{
	FileAppend, ERROR: Not enough prmeters specified`n , cscanpst.log
	MsgBox CSCANPST requires at least 2 incoming prmeters but it only received %0%.
	ExitApp 11
}

if !FileExist(prm1)
{
	FileAppend, ERROR: SCANPST.EXE not found`n , cscanpst.log
	MsgBox SCANPST.EXE not found at [%prm1%].
	ExitApp 12
}

if !FileExist(prm2)
{
	FileAppend, ERROR: PST file [%prm2%] not found`n , cscanpst.log
	MsgBox No PST file found at [%prm2%].
	ExitApp 13
}
	
if prm3=N
{
	FileAppend, (INFO: No backup PST file will be created)...  , cscanpst.log
}

SetTitleMatchMode 2

IfWinExist Microsoft Inbox Repair Tool
{
	FileAppend, ERROR: Repair Tool is already running`n , cscanpst.log
	ExitApp 3
}

;Run the SCANPST exe file...
Run %prm1%

WinWaitActive Inbox Repair Tool

;Enter the filename to be scanned
Send %prm2%
Send !S

Loop 
{
	ifWinExist, Inbox Repair Tool, been canceled
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
		FileAppend, ERROR: User cancelled`n , cscanpst.log
		exitapp 4
	}
	
	ifWinExist, Inbox Repair Tool, error prevented access
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
		FileAppend, ERROR: Could not open file`n , cscanpst.log
		exitapp 5
	}

	ifWinExist, Inbox Repair Tool, in use by another
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
		FileAppend, ERROR: File already in use`n , cscanpst.log
		exitapp 6
	}

	ifWinExist, Inbox Repair Tool, does not exist
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
		FileAppend, ERROR: File not found`n , cscanpst.log
		exitapp 7
	}

	ifWinExist, Inbox Repair Tool, does not recognize the file
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
		FileAppend, ERROR: File type not recognized`n , cscanpst.log
		exitapp 8
	}

	ifWinExist, Inbox Repair Tool, error has occurred
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
		FileAppend, ERROR: An error has occurred`n , cscanpst.log
		exitapp 9
	}

	ifWinExist, Inbox Repair Tool, is read-only
	{
		WinActivate
		WinWaitActive
		Send {ESC}
		Send !C
		FileAppend, ERROR: PST file is read only`n , cscanpst.log
		exitapp 10
	}

	ifWinExist, Inbox Repair Tool, No errors were found
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose
		FileAppend, No errors found`n , cscanpst.log
		exitapp 0
	}

	ifWinExist, Inbox Repair Tool, To repair these errors
	{
		WinActivate
		WinWaitActive
		if prm3=N
		   Send !M
		Send !R

		Loop 
		{
			ifWinExist, Inbox Repair Tool, The backup file 
			{
				WinActivate
				WinWaitActive
				Send !Y
				WinWaitClose
			}
		
			ifWinExist, Inbox Repair Tool, Repair complete 
			{
				WinActivate
				WinWaitActive
				Send {ENTER}
				WinWaitClose
				FileAppend, File repaired`n , cscanpst.log
				exitapp 2
			}
		}			
	}	

	ifWinExist, Inbox Repair Tool, Only minor inconsistencies were found
	{
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose
		FileAppend, Minor inconsistencies not repaired`n , cscanpst.log
	 	exitapp 1
	}	
}


