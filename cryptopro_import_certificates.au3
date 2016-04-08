#include <GuiListView.au3>
#include <Array.au3>

$kriptoprocsp = 'C:\Windows\System32\rundll32.exe shell32.dll,Control_RunDLL "C:\Program Files\Crypto Pro\CSP\cpconfig.cpl"'
Run($kriptoprocsp)
WinWait("КриптоПро CSP","")
BlockInput($BI_DISABLE)
If Not WinActive("КриптоПро CSP","") Then WinActivate("КриптоПро CSP","")
WinWaitActive("КриптоПро CSP","")
Send("{CTRLDOWN}{TAB}{CTRLUP}")
Send("{CTRLDOWN}{TAB}{CTRLUP}")
Send("{ALTDOWN}{к}{ALTUP}")
BlockInput($BI_ENABLE)

WinWait("Сертификаты в контейнере закрытого ключа","")
BlockInput($BI_DISABLE)
If Not WinActive("КриптоПро CSP","") Then WinActivate("Сертификаты в контейнере закрытого ключа","")
WinWaitActive("Сертификаты в контейнере закрытого ключа","")
Send("{ALTDOWN}{б}{ALTUP}")
$sTitle = 'КриптоПро CSP'
$hWin = WinWait($sTitle)
Sleep(2000)
$hListView = ControlGetHandle($hWin, '', '')
$iCountRow = _GUICtrlListView_GetItemCount($hListView) ;кол-во строк

if $iCountRow > 0 Then
    While Not ControlCommand($sTitle, '', 'OK', 'IsEnabled', '')
        Sleep(500)
    WEnd
    $hListView = ControlGetHandle($hWin, '', '')
    $iCountRow = _GUICtrlListView_GetItemCount($hListView) ;кол-во строк
EndIf

Sleep(500)
Send("{ESC}")
WinWaitActive("Сертификаты в контейнере закрытого ключа","")

;MsgBox(64, "Information", "Item Count: " & $iCountRow)

For $i = 0 to $iCountRow-1
    Send("{ALTDOWN}{б}{ALTUP}")
    $hWin = WinWait($sTitle)
    While Not ControlCommand($sTitle, '', 'OK', 'IsEnabled', '')
        Sleep(500)
	 WEnd
    $hListView = ControlGetHandle($hWin, '', '[CLASS:SysListView32; INSTANCE:1]')
    _GUICtrlListView_SetItemFocused($hListView, $i)

    Send("{ENTER}")
    If Not WinActive("КриптоПро CSP","") Then WinActivate("Сертификаты в контейнере закрытого ключа","")
    WinWaitActive("Сертификаты в контейнере закрытого ключа","")
    Send("{ALTDOWN}{д}{ALTUP}")
	$hWnd = WinWait("Сертификаты в контейнере закрытого ключа", "Установить", 5)
	If Not $hWnd Then
	    Send("{ENTER}") ; error
        If Not WinActive("Сертификаты в контейнере закрытого ключа","") Then Send("{ENTER}") ; other error?
    Else
        Send("{SHIFTDOWN}{TAB}{TAB}{SHIFTUP}{ENTER}")

        BlockInput($BI_ENABLE)

        WinWaitActive("КриптоПро CSP","")
        Send("{ENTER}")
        WinWaitActive("Сертификаты в контейнере закрытого ключа","",1)
        If Not WinActive("Сертификаты в контейнере закрытого ключа","") Then Send("{ENTER}") ; if already exist
    EndIf
    WinWaitActive("Сертификаты в контейнере закрытого ключа","")
    BlockInput($BI_DISABLE)
	Send("{ALTDOWN}{н}{ALTUP}")
 Next
Send("{ESC}")
If Not WinActive("КриптоПро CSP","") Then WinActivate("КриптоПро CSP","")
WinWaitActive("КриптоПро CSP","")
Send("{ESC}")
BlockInput($BI_ENABLE)