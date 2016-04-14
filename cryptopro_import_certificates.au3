#include <GuiListView.au3>
#include <Array.au3>

$CriptoPro = 'rundll32.exe shell32.dll,Control_RunDLL "C:\Program Files\Crypto Pro\CSP\cpconfig.cpl"'
$hCriptoPro = "КриптоПро CSP"
$tCriptoProTab1 = "Версия ядра СКЗИ"
$tCriptoProTab2 = "Считыватели закрытых ключей"
$tCriptoProTab3 = "Сертификаты в контейнере закрытого ключа"
$hCertsInPrivContainer = "Сертификаты в контейнере закрытого ключа"
$tCertsInPrivContainerStep1 = "Имя ключевого контейнера"
$tCertsInPrivContainerStep2 = "Просмотрите и выберите сертификат"
$tCriptoProCSPSelectContainer = "Выбор ключевого контейнера"
$tInstallButton = "Установить"

$kViewCertsInContainerButton = "к"
$kBrowseButton = "б"
$kNextButton = "д"
$kBackButton = "н"




Run($CriptoPro)
WinWait($hCriptoPro,$tCriptoProTab1)
BlockInput($BI_DISABLE)
If Not WinActive($hCriptoPro,$tCriptoProTab1) Then WinActivate($hCriptoPro,$tCriptoProTab1)
WinWaitActive($hCriptoPro,$tCriptoProTab1)

Send("^{TAB}")

WinWaitActive($hCriptoPro,$tCriptoProTab2)

Send("^{TAB}")

WinWaitActive($hCriptoPro,$tCriptoProTab3)

Send("{ALTDOWN}{" & $kViewCertsInContainerButton & "}{ALTUP}")
BlockInput($BI_ENABLE)

WinWait($hCertsInPrivContainer,$tCertsInPrivContainerStep1)
If Not WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1) Then WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)
BlockInput($BI_DISABLE)
WinWaitActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)

Send("{ALTDOWN}{" & $kBrowseButton & "}{ALTUP}")
$sTitle = $hCriptoPro
$sText = $tCriptoProCSPSelectContainer
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
WinWaitActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1, 1)
If Not WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1) Then WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)

;MsgBox(64, "Information", "Item Count: " & $iCountRow)

For $i = 0 to $iCountRow-1
    Send("{ALTDOWN}{" & $kBrowseButton & "}{ALTUP}")
    $hWin = WinWait($sTitle)
    While Not ControlCommand($sTitle, '', 'OK', 'IsEnabled', '')
        Sleep(500)
	 WEnd
    $hListView = ControlGetHandle($hWin, '', '[CLASS:SysListView32; INSTANCE:1]')
    _GUICtrlListView_SetItemFocused($hListView, $i)

    Send("{ENTER}")
    If Not WinActive($hCriptoPro,$tCriptoProCSPSelectContainer) Then WinActivate($hCertsInPrivContainer,"")
    WinWaitActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)
    Send("{ALTDOWN}{" & $kNextButton & "}{ALTUP}")
	$hWnd = WinWait($hCertsInPrivContainer, $tInstallButton, 5)
	If Not $hWnd Then
	    Send("{ENTER}") ; error
        If Not WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1) Then Send("{ENTER}") ; other error?
    Else
        Send("{ALTDOWN}{" & $kBackButton & "}{ALTUP}")

        BlockInput($BI_ENABLE)

        WinWaitActive($hCriptoPro,$tCriptoProCSPSelectContainer)
        Send("{ENTER}")
        WinWaitActive($hCertsInPrivContainer,$tCertsInPrivContainerStep2,1)
        If Not WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep2) Then Send("{ENTER}") ; if already exist
    EndIf
    WinWaitActive($hCertsInPrivContainer,"")
    BlockInput($BI_DISABLE)
	Send("{ALTDOWN}{" & $kBackButton & "}{ALTUP}")
 Next
Send("{ESC}")
If Not WinActive($hCriptoPro,"") Then WinActivate($hCriptoPro,"")
WinWaitActive($hCriptoPro,"")
Send("{ESC}")
BlockInput($BI_ENABLE)