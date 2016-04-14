#include <GuiListView.au3>
#include <Array.au3>

$CryptoPro = 'rundll32.exe shell32.dll,Control_RunDLL "C:\Program Files\Crypto Pro\CSP\cpconfig.cpl"'
$hCryptoPro = ''
$tCryptoProTab1 = ''
$tCryptoProTab2 = ''
$tCryptoProTab3 = ''
$hCertsInPrivContainer = ''
$tCertsInPrivContainerStep1 = ''
$tCertsInPrivContainerStep2 = ''
$tCryptoProCSPSelectContainer = ''
$tInstallButton = ''

$kViewCertsInContainerButton = ''
$kBrowseButton = ''
$kNextButton = ''
$kBackButton = ''
$kUniqueNamesRadio = ''
$tOKButton = ''
$kKeyContainerNameField = ''

Func English()
    $hCryptoPro = "CryptoPro CSP"
    $tCryptoProTab1 = "CSP core version"
    $tCryptoProTab2 = "Private key readers"
    $tCryptoProTab3 = "Certificates in private key container"
    $hCertsInPrivContainer = "Certificates in private key container"
    $tCertsInPrivContainerStep1 = "Key container name"
    $tCertsInPrivContainerStep2 = "View and choose certificate"
    $tCryptoProCSPSelectContainer = "Select key container"
	$tOKButton = "OK"
    $tInstallButton = "Install"

    $kViewCertsInContainerButton = "v"
    $kBrowseButton = "r"
    $kNextButton = "n"
    $kBackButton = "b"
    $kUniqueNamesRadio = "u"
	$kKeyContainerNameField = 'k'
EndFunc

Func Russian()
    $hCryptoPro = "КриптоПро CSP"
    $tCryptoProTab1 = "Версия ядра СКЗИ"
    $tCryptoProTab2 = "Считыватели закрытых ключей"
    $tCryptoProTab3 = "Сертификаты в контейнере закрытого ключа"
    $hCertsInPrivContainer = "Сертификаты в контейнере закрытого ключа"
    $tCertsInPrivContainerStep1 = "Имя ключевого контейнера"
    $tCertsInPrivContainerStep2 = "Просмотрите и выберите сертификат"
    $tCryptoProCSPSelectContainer = "Выбор ключевого контейнера"
	$tOKButton = "ОК"
    $tInstallButton = "Установить"

    $kViewCertsInContainerButton = "к"
    $kBrowseButton = "б"
    $kNextButton = "д"
    $kBackButton = "н"
    $kUniqueNamesRadio = "у"
	$kKeyContainerNameField = 'и'
EndFunc


Run($CryptoPro)

$hWnd = WinWaitActive("[REGEXPTITLE:(CryptoPro CSP|КриптоПро CSP)]")

If WinGetTitle($hWnd) = "CryptoPro CSP" then
    English()
Else
    Russian()
EndIf

If Not WinActive($hCryptoPro,$tCryptoProTab1) Then WinActivate($hCryptoPro,$tCryptoProTab1)
WinWaitActive($hCryptoPro,$tCryptoProTab1)

Send("^{TAB}")

WinWaitActive($hCryptoPro,$tCryptoProTab2)

Send("^{TAB}")

WinWaitActive($hCryptoPro,$tCryptoProTab3)

Send("{ALTDOWN}{" & $kViewCertsInContainerButton & "}{ALTUP}")

WinWait($hCertsInPrivContainer,$tCertsInPrivContainerStep1)
If Not WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1) Then WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)
WinWaitActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)

Send("{ALTDOWN}{" & $kBrowseButton & "}{ALTUP}")
$hWin = WinWait($hCryptoPro, $tCryptoProCSPSelectContainer)

$hListView = ControlGetHandle($hWin, '', '')
Send("{ALTDOWN}{" & $kUniqueNamesRadio & "}{ALTUP}")
Sleep(2000)

$iCountRow = _GUICtrlListView_GetItemCount($hListView) ;кол-во строк

if $iCountRow > 0 Then
    While Not ControlCommand($hCryptoPro, $tCryptoProCSPSelectContainer, $tOKButton, 'IsEnabled', '')
        Sleep(500)
    WEnd
    $iCountRow = _GUICtrlListView_GetItemCount($hListView) ;кол-во строк
	Global $containers[$iCountRow]

    For $i = 0 to $iCountRow-1
        $aItem = _GUICtrlListView_GetItemTextArray($hListView, $i)
	    $containers[$i] = $aItem[2]
    Next
EndIf

Send("{ESC}")
WinWaitActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1, 1)
If Not WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1) Then WinActive($hCertsInPrivContainer,$tCertsInPrivContainerStep1)

For $i = 0 to $iCountRow-1
    BlockInput($BI_DISABLE)
    Send("{ALTDOWN}{" & $kKeyContainerNameField & "}{ALTUP}")
	Send($containers[$i])
    Send("{ALTDOWN}{" & $kNextButton & "}{ALTUP}")
	BlockInput($BI_ENABLE)

    $stillLooking = True
    While $stillLooking
        $activeWindowTitle = WinGetTitle(WinActive(""))
		$activeWindowText = WinGetText(WinActive(""))
        If $activeWindowTitle == $hCryptoPro Then
			Send("{ENTER}") ; close error
			sleep(500)
            If WinActive($hCryptoPro) Then Send("{ENTER}") ; other error?
            $stillLooking = False
	    ElseIf $activeWindowTitle == $hCertsInPrivContainer Then
            If StringInStr($activeWindowText, $tInstallButton) Then
                BlockInput($BI_DISABLE)
                Send("{SHIFTDOWN}{TAB}{TAB}{SHIFTUP}{ENTER}")
		        BlockInput($BI_ENABLE)

                WinWaitActive($hCryptoPro)
	            Send("{ENTER}") ; close info
		        sleep(500)
                If WinActive($hCryptoPro) Then Send("{ENTER}") ; other info?
                $stillLooking = False
		    EndIf
        EndIf
        sleep(5)
    WEnd

    WinWaitActive($hCertsInPrivContainer,"")
	Send("{ALTDOWN}{" & $kBackButton & "}{ALTUP}")
 Next
Send("{ESC}")
If Not WinActive($hCryptoPro,"") Then WinActivate($hCryptoPro,"")
WinWaitActive($hCryptoPro,"")
Send("{ESC}")