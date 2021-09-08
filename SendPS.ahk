; AHKプログラムをインストールしてからお使いください
; このコードを　SJIS　で 【SendPS.ahk】として保存してから使用すること！
;
;//////////////////////////////////////////////
;【署名】【URL】【ログインID PW】 管理プログラム
;  By skojima@kuhen.org

Vertext=Ver.2021-09-08_2107

;//////////////////////////////////////////////


;///////////////
; 初期設定：登録可能数を設定
;///////////////

Gvarkosu=25

;///////////////

MsgBox ,,　【署名】【URL】【ログインID PW】 管理プログラム    %Vertext%,【F12】　１回：メニュー　２回：設定ファイル直接編集　長押し：設定再読込`n`nこのプログラムは無保証です自己責任でお使いください   By skojima@kuhen.org , 3


;/// マイドキュメント内に　\SendPS.ini　が存在しないなら＞初期設定
IfNotExist, %A_MyDocuments%\SendPS.ini
{
FileAppend,,%A_MyDocuments%\SendPS.ini
msgbox,【重要】今から署名、パスワード保存用として`n`n【%A_MyDocuments%\SendPS.ini】`n`nを作成＆使用します！`n`n保存されるデータは暗号化されませんのでご注意ください！
}

;---------------------------------------------------------------
;/// 署名の最大個数を読込
IniRead, sig_kosu,%A_MyDocuments%\SendPS.ini, sig, kosu

;/// エラーなら＞初期設定
IfEqual, sig_kosu,ERROR
{
IniWrite, %Gvarkosu% , %A_MyDocuments%\SendPS.ini, sig, kosu
IniRead, sig_kosu,%A_MyDocuments%\SendPS.ini, sig, kosu

  MsgBox, 登録可能「署名」数 %sig_kosu%

Loop , %Gvarkosu% {
IniWrite, "" , %A_MyDocuments%\SendPS.ini, sig, sigid%A_index%
IniWrite,"", %A_MyDocuments%\SendPS.ini, sig, sigdata%A_index%

}

}

;/// 署名出力用メニュー作成
Loop , %sig_kosu% {
IniRead, sig_id,%A_MyDocuments%\SendPS.ini, sig, sigid%A_index%
IniRead, sig_data,%A_MyDocuments%\SendPS.ini, sig, sigdata%A_index%
Menu , kuhensig , Add , 【sigid%A_index%】 %sig_id% (出力: %sig_data% ),SENDSIG
}

;/// 署名設定用メニュー作成
Loop , %sig_kosu% {
IniRead, sig_id,%A_MyDocuments%\SendPS.ini, sig, sigid%A_index%
Menu , kuhensigset , Add ,【sigid%A_index%】%sig_id%,SAVESIG
}


;---------------------------------------------------------------
;/// パスワードの最大個数を読込
IniRead, pw_kosu,%A_MyDocuments%\SendPS.ini, pw, kosu

;/// エラーなら＞初期設定
IfEqual, pw_kosu,ERROR
{
IniWrite, %Gvarkosu% , %A_MyDocuments%\SendPS.ini, pw, kosu
IniRead, pw_kosu,%A_MyDocuments%\SendPS.ini, pw, kosu

  MsgBox, 登録可能「パスワード」数 %pw_kosu%

Loop , %Gvarkosu% {
IniWrite, "" , %A_MyDocuments%\SendPS.ini, pw, pwid%A_index%
IniWrite,"", %A_MyDocuments%\SendPS.ini, pw, pwdata%A_index%
}

}

;/// パスワード出力用メニュー作成
Loop , %pw_kosu% {
IniRead, pw_id,%A_MyDocuments%\SendPS.ini, pw, pwid%A_index%
Menu , kuhenpw , Add ,【pwid%A_index%】 %pw_id%,SENDPW
}

;/// パスワード設定用メニュー作成
Loop , %pw_kosu% {
IniRead, pw_id,%A_MyDocuments%\SendPS.ini, pw, pwid%A_index%
Menu , kuhenpwset , Add ,【pwid%A_index%】%pw_id%,SAVEPW

}


;===============================================================================================
;Menu , kuhen , Add ,　【重要】保存パスワードは暗号化されてません！　 %Vertext%,nothing
;Menu , kuhen , Add ,　　　　無保証です自己責任でお使いください　By skojima@kuhen.org,nothing
;Menu , kuhen , Add ,　【F12】　１回：メニュー表示　２回：設定ファイル直接編集　長押し：設定再読込,nothing
;Menu , kuhen , Add ;

Menu , kuhen , Add ,■【署名】【URL】等を出力, :kuhensig
Menu , kuhen , Add ;
Menu , kuhen , Add ,■【パスワードのID、PWを出力】, :kuhenpw
Menu , kuhen , Add ;
Menu , kuhen , Add ,　　　◆【署名等設定】　　「登録名」と「署名本文」、パスワードで使う「URL」や「メモ」等を設定, :kuhensigset
Menu , kuhen , Add ;
Menu , kuhen , Add ,　　　◆【パスワード設定】　　「ID」 と 「パスワード」 を設定, :kuhenpwset
Menu , kuhen , Add ;
Menu , kuhen , Add ,(&!v)　[ Alt+ V ]　【ペイント起動】＋【貼り付け】, ALT_V

;===============================================================================================

return


;---------------------------------------------------------------
;/// Menuのコメント用（何も処理しない）
;------------
nothing:
Return


;------------
!v:: ; Alt+v
;------------

ALT_V:

IMEGetstateOFF()
Sleep,300

run, C:\Windows\System32\mspaint.exe
Sleep,1000

Send ^v ; Ctrl+v

Return


;---------------------------------------------------------------
;登録している「署名」「URL」等を出力
;
SENDSIG(Item){

  ichi := InStr(Item,"】")-7
  StringMid, id , Item , 7 , %ichi%

Sleep,300

IMEGetstateOFF()

  IniRead, sig_id,%A_MyDocuments%\SendPS.ini, sig, sigid%id%
Sleep,200
  IniRead, sig_data,%A_MyDocuments%\SendPS.ini, sig, sigdata%id%
Sleep,200

SIGVar=%sig_data%
SIGVar:=RegExReplace(SIGVar,"\\n","`n")
Clipboard =%SIGVar%

If StrLen(sig_data) = 0{
MsgBox,,,エラー：署名が「空白」です！`n`n◆【署名設定】で登録を！
}else{

MsgBox,,,【 %sig_id% 】`n`n%SIGVar%`n`n「貼り付け」で使えます,3

;URLなら開く
 If SubStr( sig_data,1,4)="http" {
  MsgBox,4,,【 %sig_id% 】`n`n%SIGVar%`n`nを開きますか？
   IfMsgBox, Yes
    Run ,  %sig_data%
 }

}

}


;---------------------------------------------------------------
;登録している「ID」「PW」を出力
;
SENDPW(Item){

  ichi := InStr(Item,"】")-6
  StringMid, id , Item , 6 , %ichi%

Sleep,200

IMEGetstateOFF()

  IniRead, pw_id,%A_MyDocuments%\SendPS.ini, pw, pwid%id%

Sleep,200
Clipboard =%pw_id%

  IniRead, pw_data,%A_MyDocuments%\SendPS.ini, pw, pwdata%id%

If StrLen(pw_data) = 0{
MsgBox,,,エラー：パスワードが「空白」です！`n`n◆【パスワード設定】で登録を！
}else{
MsgBox,,,現在【 %pw_id% 】が「貼り付け」で使えます`n`nこの後は「パスワード」が「貼り付け」で使えます
}


Sleep,300
Clipboard =%pw_data%


}


;---------------------------------------------------------------
;「署名」「URL」を保存
;
SAVESIG(Item){

  ichi := InStr(Item,"】")-7
  StringMid, id , Item , 7 , %ichi%

SIGSET(id)

}


;---------------------------------------------------------------
;「ID」と「PW」を保存
;
SAVEPW(Item){
  ichi := InStr(Item,"】")-6
  StringMid, id , Item , 6 , %ichi%

PWSET(id)

}


;---------------------------------------------------------------
; F12 を押す回数で処理を選択
;
;https://pouhon.net/ahk-keywait/2848/
#UseHook                 ;UseHookをOnにした状態で使用
F12::
key := "F12"
KeyWait, %key%, T0.3
If(ErrorLevel){          ;長押しした場合
msgbox, 今から「設定ファイル」を再読み込みします！
Reload
    KeyWait, %key%
    return
}
KeyWait, %key%, D, T0.2
If(!ErrorLevel){         ;2度押しした場合
Run , %A_MyDocuments%\SendPS.ini
    KeyWait, %key%
    return
}else{                   ;短押しした場合

Menu , kuhen , show

    KeyWait, %key%
    return
}
return


;---------------------------------------------------------------
;「署名」「URL」等を登録
;
SIGSET(No){

IMEGetstateON()

IniRead, sig_sigid,%A_MyDocuments%\SendPS.ini, sig, sigid%No%

InputBox, new_sig_sigid,【署名】「表示名」を登録,【2-1】現在の【表示名】→【　%sig_sigid%　】`n`n登録したい　「表示名」を入力`n`n変更 しない→【Cancel】,,600,200

If ErrorLevel = 0
{
IniWrite, %new_sig_sigid% , %A_MyDocuments%\SendPS.ini, sig, sigid%No%
sleep 300
IniRead, sig_sigid2,%A_MyDocuments%\SendPS.ini, sig, sigid%No%
sleep 300
  msgbox,現在の【表示名】は`n`n【　%sig_sigid2%　】`n`nです。
}
Else
{
  MsgBox,キャンセル：「表示名」は更新しませんでした！
}


IniRead, sig_sigdata,%A_MyDocuments%\SendPS.ini, sig, sigdata%No%

InputBox, new_sig_sigdata,【署名】「本文」を登録,【2-2】現在の【署名「本文」】→【　%sig_sigdata%　】`n`n登録したい　【署名】「本文」、「URL」等　を入力`n`n【改行】は → \n を入力してください`n`n変更 しない→【Cancel】,,800,250

If ErrorLevel = 0
{
IniWrite, %new_sig_sigdata% , %A_MyDocuments%\SendPS.ini, sig, sigdata%No%
sleep 300
IniRead, sig_sigdata2,%A_MyDocuments%\SendPS.ini, sig, sigdata%No%
sleep 300
  msgbox,現在の【「署名」本文】は`n`n【　%sig_sigdata2%　】`n`nです。
  
}
Else
{
  MsgBox,キャンセル：「署名」本文は更新しませんでした！
}

Reload

return
}


;---------------------------------------------------------------
;「ID」「PW」を登録
;
PWSET(No){

IMEGetstateOFF()

IniRead, pw_pwid,%A_MyDocuments%\SendPS.ini, pw, pwid%No%

InputBox, new_pw_pwid,【パスワード】のIDを登録,【2-1】現在の【ID】→【　%pw_pwid%　】`n`n登録ログイン用　「ID」　を入力`n`n変更 しない→【Cancel】,,400,250

If ErrorLevel = 0
{
IniWrite, %new_pw_pwid% , %A_MyDocuments%\SendPS.ini, pw, pwid%No%
sleep 300
IniRead, pw_pwid2,%A_MyDocuments%\SendPS.ini, pw, pwid%No%
sleep 300
  msgbox,現在の【ID】は`n`n【　%pw_pwid2%　】`n`nです。
}
Else
{
  MsgBox,キャンセル：【ID】は更新しませんでした！
}


IniRead, pw_pwdata,%A_MyDocuments%\SendPS.ini, pw, pwdata%No%

InputBox, new_pw_pwdata,【パスワード】の登録,【2-2】現在の【パスワード】→【　%pw_pwdata%　】`n`n登録ログイン用　「パスワード」　を入力`n`n変更 しない→【Cancel】,,400,300

If ErrorLevel = 0
{
IniWrite, %new_pw_pwdata% , %A_MyDocuments%\SendPS.ini, pw, pwdata%No%
sleep 300
IniRead, pw_pwdata2,%A_MyDocuments%\SendPS.ini, pw, pwdata%No%
sleep 300
  msgbox,現在の【パスワード】は`n`n【　%pw_pwdata2%　】`n`nです。
}
Else
{
  MsgBox,キャンセル：【パスワード】は更新しませんでした！
}

Reload

return

}


;===============================================================================================
;Windows IME On Off 制御
;https://qiita.com/neokix/items/83ac2ac8e1dd405528b0
;

IMEGetstateOFF(){
    DetectHiddenWindows, ON
    WinGet, vcurrentwindow, ID, A
    vgetdefault := DllCall("imm32.dll\ImmGetDefaultIMEWnd", "Uint", vcurrentwindow)
    vimestate := DllCall("user32.dll\SendMessageA", "UInt", vgetdefault, "UInt", 0x0283, "Int", 0x0005, "Int", 0)
    DetectHiddenWindows, Off

    If (vimestate=0) ;imeがoffなら
    {
    ;Offだから何もしない
    }
    else
    {
    Send, {vkf3}
    }

    return

}


IMEGetstateON(){
    DetectHiddenWindows, ON
    WinGet, vcurrentwindow, ID, A
    vgetdefault := DllCall("imm32.dll\ImmGetDefaultIMEWnd", "Uint", vcurrentwindow)
    vimestate := DllCall("user32.dll\SendMessageA", "UInt", vgetdefault, "UInt", 0x0283, "Int", 0x0005, "Int", 0)
    DetectHiddenWindows, Off

    If (vimestate=0) ;imeがoffなら
    {
    Send, {vkf3}
    }
    else
    {
    ;Onだから何もしない
    }

    return

}
;===============================================================================================
