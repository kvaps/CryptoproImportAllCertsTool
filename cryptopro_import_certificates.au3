#include <GuiListView.au3>
#include <Array.au3>

$kriptoprocsp = 'C:\Windows\System32\rundll32.exe shell32.dll,Control_RunDLL "C:\Program Files\Crypto Pro\CSP\cpconfig.cpl"'
Run($kriptoprocsp)
WinWait("КриптоПро CSP","Версия ядра СКЗИ")
BlockInput($BI_DISABLE)
If Not WinActive("КриптоПро CSP","Версия ядра СКЗИ") Then WinActivate("КриптоПро CSP","Версия ядра СКЗИ")
WinWaitActive("КриптоПро CSP","Версия ядра СКЗИ")

Send("^{TAB}")

WinWaitActive("КриптоПро CSP","Считыватели закрытых ключей")

Send("^{TAB}")

WinWaitActive("КриптоПро CSP","Сертификаты в контейнере закрытого ключа")

Send("{ALT}{!к}")
BlockInput($BI_ENABLE)

WinWait("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера")
If Not WinActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера") Then WinActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера")
BlockInput($BI_DISABLE)
WinWaitActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера")

Send("{ALT}{!б}")
$sTitle = 'КриптоПро CSP'
$sText = 'Выбор ключевого контейнера'
$hWin = WinWait($sTitle, $sText)
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

Send("{ESC}")
WinWaitActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера", 1)
If Not WinActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера") Then WinActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера")

;MsgBox(64, "Information", "Item Count: " & $iCountRow)

For $i = 0 to $iCountRow-1
    Send("{ALT}{!б}")
    $hWin = WinWait($sTitle)
    While Not ControlCommand($sTitle, '', 'OK', 'IsEnabled', '')
        Sleep(500)
	 WEnd
    $hListView = ControlGetHandle($hWin, '', '[CLASS:SysListView32; INSTANCE:1]')
    _GUICtrlListView_SetItemFocused($hListView, $i)

    Send("{ENTER}")
    If Not WinActive("КриптоПро CSP","Выбор ключевого контейнера") Then WinActivate("Сертификаты в контейнере закрытого ключа","")
    WinWaitActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера")
    Send("{ALT}{!д}")
	$hWnd = WinWait("Сертификаты в контейнере закрытого ключа", "Установить", 5)
	If Not $hWnd Then
	    Send("{ENTER}") ; error
        If Not WinActive("Сертификаты в контейнере закрытого ключа","Имя ключевого контейнера") Then Send("{ENTER}") ; other error?
    Else
        Send("{ALT}{!н}")

        BlockInput($BI_ENABLE)

        WinWaitActive("КриптоПро CSP","Выбор ключевого контейнера")
        Send("{ENTER}")
        WinWaitActive("Сертификаты в контейнере закрытого ключа","Просмотрите и выберите сертификат",1)
        If Not WinActive("Сертификаты в контейнере закрытого ключа","Просмотрите и выберите сертификат") Then Send("{ENTER}") ; if already exist
    EndIf
    WinWaitActive("Сертификаты в контейнере закрытого ключа","")
    BlockInput($BI_DISABLE)
	Send("{ALT}{!н}")
 Next
Send("{ESC}")
If Not WinActive("КриптоПро CSP","") Then WinActivate("КриптоПро CSP","")
WinWaitActive("КриптоПро CSP","")
Send("{ESC}")
BlockInput($BI_ENABLE)