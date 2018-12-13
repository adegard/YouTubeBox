; https://support.google.com/youtube/answer/7631406?hl=en

#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
FileEncoding, UTF-8


MYAPP_PROTOCOL:="myapp"
Menu Tray, Icon, shell32.dll, 169

Gui 1:+Resize 
Gui 1:Add, Edit, vsearch x370 y5 w178 h30  -WantReturn -VScroll, megadeth ;x450 y3 w261 h30
Gui 1:Add, Button, x560 y6 w35 h30 gsearch hwndsearch ;, &SEARCH x716 y3 w35 h30
GuiButtonIcon(search, "shell32.dll", 56, "s32 a4")

Gui 1:Add, Edit, vvlc x21 y380 w275 h21
Gui 1:Add, Button, x303 y379 w80 h23 gvlc , Send to VLC

Gui Add, Link, x15 y10 w187 h15, <a href="https://autohotkey.com/boards/viewtopic.php?f=6&t=59438">Forum Thread ahk</a>

;VIDEO PLAYER:
Gui, 1:Add, ActiveX, x15 y37 w350 h364 vpwb, Shell.Explorer
pwb.Navigate("https://www.youtube.com/")
pwb.statusbar := false
WinGetTitle, WinTitle, A

GuiControl,, vlc, https://www.youtube.com/watch?v=UD_mry_DN-s

;VIDEO LIST :
Gui 1:Add, ActiveX, x370 y37 w270 h364 vWB, Shell.Explorer  ; The final parameter is the name of the 
WinGetTitle, WinTitleWB, A

;**********************************
Gui, 1:Show, w658 h412, YoutubeBox ;w811 h412
WinGetPos,,, Width, Height, YoutubeBox
WinMove, YoutubeBox,, (A_ScreenWidth)-(Width), (A_ScreenHeight)-(Height*1.2)
while pwb.busy
	sleep 10
;
window := pwb.document.parentWindow

WinGet, active_id, ID, A
  
return


64Bit() {
    Return (FileExist("C:\Program Files (x86)")) ? 1 : 0
}

search:

;Waiting...
Gui, 2: +Toolwindow +Disabled +AlwaysOnTop
Gui, 2: Add, Text,, The script is working. Please be patient.
Gui, 2: Show,, Please Wait
WinGetPos,,, Width, Height, Please Wait
WinMove, Please Wait,, (A_ScreenWidth)-(Width) ;, (A_ScreenHeight)-(Height)


GuiControlGet, keyword,, search
;MsgBox, %keyword%

; JSON PARSING METHOD from API YOUTUBE *****************
; example here, by SirRFI (thanks!): https://autohotkey.com/boards/viewtopic.php?f=76&t=38059&p=175017&hilit=parsing+html#p175017
;and JSON.ack by cocobelgica (thanks!) https://github.com/cocobelgica/AutoHotkey-JSON/blob/master/JSON.ahk

;ADD YOUR API KEY HERE:
YTurl :="https://www.googleapis.com/youtube/v3/search?part=snippet&q=" keyword "&type=video%20&videoCaption=closedCaption&key=***APIKEY***"


;MsgBox, %YTurl%
;extract urls

FileDelete, YourJsonFile.json
UrlDownloadToFile, %YTurl%, YourJsonFile.json  ; 01
FileRead, JsonContent, YourJsonFile.json	; ensure the file is saved in UTF-8 with BOM, if there are UNICODE symbols You want to display properly
MyJsonInstance := new JSON()
JsonObject := MyJsonInstance.Load(JsonContent)

;reset lista
Lista=
     
;MsgBox, %JsonObject%
; or simply JsonObject := JSON.Load(JsonContent), since You want to use the class' method only
for id,items in JsonObject["items"]{	; go through "JsonObject" array (series[1], series[2], and so on)
    for key1,val1 in items.snippet{		; go through "JsonObject" array (series[1], series[2], and so on)
        if (key1="Title")
            Title := val1
     }
    for key3,val3 in items.snippet.thumbnails.default{	
	;MsgBox, %key3%
            if (key3="url")
                thumb := val3
     }
     
     
    for key2,val2 in items.id{		; go through "JsonObject" array (series[1], series[2], and so on)
        if (key2="videoId")
            Linka := val2
     }

    Lista  .= "<tr><td style='text-align: left; width: 87px;'><a href=myapp://https://www.youtube.com/watch?v="Linka "><img border='0' src='"thumb "' width='87' ></a></td><td style='text-align: left; width: 134px;'><a href=myapp://https://www.youtube.com/watch?v="Linka ">" Title "</a></td></tr>"     
    ;MsgBox, %Lista%
    
 }        

/* OTHER WAY TO SEARCH: PLAYLISTS:
search?part=snippet
                     &q=soccer
                     &type=playlist
                     &key={YOUR_API_KEY}
             */
;ADD YOU API KEY HERE:	     
YTurl2 :="https://www.googleapis.com/youtube/v3/search?part=snippet&q=" keyword "&type=playlist&key=***APIKEY***"
                    

FileDelete, playlistJsonFile.json
UrlDownloadToFile, %YTurl2%, playlistJsonFile.json  ; 01
FileRead, JsonContent2, playlistJsonFile.json	; ensure the file is saved in UTF-8 with BOM, if there are UNICODE symbols You want to display properly
MyJsonInstance2 := new JSON()
JsonObject2 := MyJsonInstance2.Load(JsonContent2)

Lista  .=   "<tr><td style='text-align: left; width: 97px;'>PlayLists:</td><td style='text-align: left; width: 134px;'></td></tr>"

for id,items in JsonObject2["items"]{	; go through "JsonObject" array (series[1], series[2], and so on)
    for key4,val4 in items.snippet{		; go through "JsonObject" array (series[1], series[2], and so on)
        if (key4="Title")
            Title := val4
     }
    for key5,val5 in items.snippet.thumbnails.default{	
	;MsgBox, %key3%
            if (key5="url")
                thumb := val5
     }
     
    for key6,val6 in items.id{		; go through "JsonObject" array (series[1], series[2], and so on)
        if (key6="playlistId")
            Linka := val6
     }

    Lista  .= "<tr><td style='text-align: left; width: 87px;'><a href=myapp://https://www.youtube.com/playlist?list="Linka "><img border='0' src='"thumb "' width='87' ></a></td><td style='text-align: left; width: 134px;'><a href=myapp://https://www.youtube.com/playlist?list="Linka ">" Title "</a></td></tr>"     
    
 }     

HTML_page_head =
( Ltrim Join
  <html>
   <body>
<table style="height: 106px; width: 243px;">
<tbody>

 )

 HTML_page_foot =
( Ltrim Join            
</tbody>
</table>
  
  </body>
  </html>
)
HTML_page = %HTML_page_head%%Lista%%HTML_page_foot%


WB.silent := true ;Surpress JS Error boxes
Display(WB,HTML_page)
Gui, 2: destroy
ComObjConnect(WB, WB_events)  ; Connect WB's events to the WB_events class object.
firsttime=1
return

;VLC OPEN
vlc:

    GuiControlGet, vlcurl,, vlc

    Filename1=VLCPlugin & ActiveX Test
    Gui,3:default
    GUI,3:Font,s14 cGray,Lucida Console
    Gui,3: -DPIScale
    Gui,3: Color, Black,Black 


    wa:=A_screenwidth
    ha:=A_screenHeight
    xx:=105
    LW  :=(wa*88) /xx 
    LH  :=(ha*88) /xx  
    GW  :=(wa*90) /xx 
    GH  :=(ha*92) /xx  
    ;vlc1        =%A_programfiles%\VideoLAN\VLC\vlc.exe
    
        if (64Bit()=1)
    {
        vlc1        =C:\Program Files (x86)\VideoLAN\VLC\vlc.exe
    }
    else
    {
        vlc1        =%A_programfiles%\VideoLAN\VLC\vlc.exe
    } 
    
    
    ifnotexist,%vlc1% 
      {
      msgbox,needs=`n%vlc1%
      exitapp
      }
    xxa=VideoLAN.VLCPlugin.2
    Gui,3:Add,ActiveX, x20    y0     w%lw%  h%lh%  vVlcx,%xxa%
    Gui,3:Show,x0 y0 w%gw% h%gh%,%filename1%
    gosub,aa1
    return
    3Guiclose:
    Gui,3:Destroy
    ;------------
return


aa1:
;MsgBox, %vlcurl%
    ;id1=_b9R_x_imBM
    ;id1=CK-pDtdW4Ug
    ;id1=0LTZRYJzqB4    ;- not works
    ;F1=https://www.youtube.com/watch?v=%id1%
    F1= %vlcurl%
    ;MsgBox, %F1%
         vlcx.playlist.add(F1,"","""""")
         vlcx.playlist.next()
return



/*
GuiSize:
    if ErrorLevel = 1  ; The window has been minimized.  No action needed.
        return
    ; Otherwise, the window has been resized or maximized. Resize the Edit control to match.
    NewWidth := A_GuiWidth - 390
    NewHeight := A_GuiHeight - 80
    GuiControl, 1: Move, pwb, W%NewWidth% H%NewHeight%

return
*/
FileExit:     ; User chose "Exit" from the File menu.
GuiClose:  ; User closed the window.
ExitApp

;Esc::ExitApp


    class WB_events
{
	;for more events and other, see http://msdn.microsoft.com/en-us/library/aa752085
	
	NavigateComplete2(WB) {
		WB.Stop() ;blocked all navigation, we want our own stuff happening
	}
	DownloadComplete(WB, NewURL) {
		WB.Stop() ;blocked all navigation, we want our own stuff happening
	}
	DocumentComplete(WB, NewURL) {
		WB.Stop() ;blocked all navigation, we want our own stuff happening
	}
	
	BeforeNavigate2(WB, NewURL)
	{
		WB.Stop() ;blocked all navigation, we want our own stuff happening
		;parse the url
		global MYAPP_PROTOCOL
        
        ;%MYAPP_PROTOCOL%://https://www.youtube.com/watch?v=SyRHeyFdl0I 
        ;myapp://https://www.youtube.com/watch?v=SyRHeyFdl0I 
		;if (InStr(NewURL,MYAPP_PROTOCOL "://")==1) { ;if url starts with "myapp://"
			what := SubStr(NewURL,Strlen(MYAPP_PROTOCOL)+4) ;get stuff after "myapp://"
            ;MsgBox %what%
            
            pwb := PWB_Init(WinTitle) 
			pwb.Navigate(what)
            ComObjConnect(pwb)
            
            GuiControl,, vlc, %what% 
            
        ;}
		;else do nothing
	}
}


Display(WB,html_str) {
	Count:=0
	while % FileExist(f:=A_Temp "\" A_TickCount A_NowUTC "-tmp" Count ".DELETEME.html")
		Count+=1
	FileAppend,%html_str%,%f%
	WB.Navigate("file://" . f)
}

;*****************************IE POINTER LIB*******************************
;thanks to @berban lib: https://autohotkey.com/boards/viewtopic.php?t=19889
PWB_Get(WinTitle="A", Svr#=1) ; Jethrow - http://www.autohotkey.com/board/topic/47052-basic-webpage-controls-with-javascript-com-tutorial/
{
	Static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
	, IID := "{0002DF05-0000-0000-C000-000000000046}" ; IID_IWebBrowserApp
	;,IID := "{332C4427-26CB-11D0-B483-00C04FD90119}" ; IID_IHTMLWindow2
	SendMessage, msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
	If (ErrorLevel != "FAIL") {
		lResult := ErrorLevel, VarSetCapacity(GUID, 16, 0)
		If (DllCall("ole32\CLSIDFromString", "wstr", "{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr", &GUID) >= 0) {
			DllCall("oleacc\ObjectFromLresult", "ptr", lResult, "ptr", &GUID, "ptr", 0, "ptr*", pdoc)
			Return ComObj(9, ComObjQuery(pdoc, IID, IID), 1), ObjRelease(pdoc)
		}
	}
	MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error,  Unable to obtain browser object (PWB) from window:`n`n%WinTitle%
}
PWB_Init(WinTitle="")
{
	Global PWB, Element
	PWB_Clear(False), ComObjError(False)
	If !PWB or (WinTitle <> "") {
		TitleMatchMode := A_TitleMatchMode, Element := ""
		SetTitleMatchMode, RegEx
		HWND := WinExist("ahk_class IEFrame")
		SetTitleMatchMode, %TitleMatchMode%
		If !HWND {
			MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, No internet explorer windows exist.
			Return
		}
		If (WinTitle <> "") and WinExist(WinTitle " ahk_class IEFrame")
			PWB := PWB_Get(WinTitle)
		Else IfWinActive, ahk_class IEFrame
			PWB := PWB_Get("A")
		Else
			PWB := PWB_Get("ahk_id " HWND)
	}
	Return PWB
}


PWB_Clear(Set=True)
{
	PWB_DefaultTimeout = 1500
	;------------------------------------------------------------------------------------------------------------------------
	Global Element, PWB, PWB_Timeout
	Static Enabled, Initialized
	If !Initialized {
		If (PWB_Timeout = "")
			PWB_Timeout := PWB_DefaultTimeout
		Initialized := True
	}
	If Set in On,Off
		Enabled := (Set = "On")
	Else If (Enabled = "")
		Enabled := True
	If PWB_Timeout and Enabled
		If Set
			SetTimer, %A_ThisFunc%, %PWB_Timeout%
	Else
		SetTimer, %A_ThisFunc%, Off
	Return
	PWB_Clear:
	Element := PWB := ""
	Return
}




;StrX() :: Auto-Parser for XML / HTML***********************
; https://autohotkey.com/board/topic/47368-strx-auto-parser-for-xml-html/
; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
;StrX( H, BS,BO,BT, ES,EO,ET, NextOffset )


StrX(H,BS="",BO=0,BT=1,ES="",EO=0,ET=1,ByRef N="") { ; By Skan | 19-Nov-2009
Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
 <0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
 +(0-ET)):(X+P)):(X)))-P) 
}


;UnHTM() :: Remove HTML formatting from a String [Updated]
;https://autohotkey.com/board/topic/47356-unhtm-remove-html-formatting-from-a-string-updated/

UnHTM( HTM ) {   ; Remove HTML formatting / Convert to ordinary text   by SKAN 19-Nov-2009
 Static HT,C=";" ; Forum Topic: www.autohotkey.com/forum/topic51342.html  Mod: 16-Sep-2010
 IfEqual,HT,,   SetEnv,HT, % "ááââ´´ææàà&ååãã&au"
 . "mlä&bdquo„¦¦&bull•çç¸¸¢¢&circˆ©©¤¤&dagger†&dagger‡°"
 . "°÷÷ééêêèèððëë&euro€&fnofƒ½½¼¼¾¾>>&h"
 . "ellip…ííîî¡¡ìì¿¿ïï««&ldquo“&lsaquo‹&lsquo‘<<&m"
 . "acr¯&mdash—µµ··  &ndash–¬¬ññóóôô&oeligœòò&or"
 . "dfªººøøõõöö¶¶&permil‰±±££"""»»&rdquo”
 . "®&rsaquo›&rsquo’&sbquo‚&scaronš§§­ ¹¹²²³³ßßþþ&tilde˜&tim"
 . "es×&trade™úúûûùù¨¨üüýý¥¥ÿÿ"
 $ := RegExReplace( HTM,"<[^>]+>" )               ; Remove all tags between  "<" and ">"
 Loop, Parse, $, &`;                              ; Create a list of special characters
   L := "&" A_LoopField C, R .= (!(A_Index&1)) ? ( (!InStr(R,L,1)) ? L:"" ) : ""
 StringTrimRight, R, R, 1
 Loop, Parse, R , %C%                               ; Parse Special Characters
  If F := InStr( HT, L := A_LoopField )             ; Lookup HT Data
    StringReplace, $,$, %L%%C%, % SubStr( HT,F+StrLen(L), 1 ), All
  Else If ( SubStr( L,2,1)="#" )
    StringReplace, $, $, %L%%C%, % Chr(((SubStr(L,3,1)="x") ? "0" : "" ) SubStr(L,3)), All
Return RegExReplace( $, "(^\s*|\s*$)")            ; Remove leading/trailing white spaces
}


;https://autohotkey.com/boards/viewtopic.php?t=23987
GetFocusedControlClassNN( )
{
GuiWindowHwnd := WinExist("A")		;stores the current Active Window Hwnd id number in "GuiWindowHwnd" variable
				;"A" for Active Window

ControlGetFocus, FocusedControl, ahk_id %GuiWindowHwnd%	;stores the  classname "ClassNN" of the current focused control from the window above in "FocusedControl" variable
						;"ahk_id" searches windows by Hwnd Id number

return, FocusedControl
}







/**;https://autohotkey.com/boards/viewtopic.php?f=6&t=627&hilit=parsing+json

 * Lib: JSON.ahk
 *     JSON lib for AutoHotkey.
 * Version:
 *     v2.1.3 [updated 04/18/2016 (MM/DD/YYYY)]
 * License:
 *     WTFPL [http://wtfpl.net/]
 * Requirements:
 *     Latest version of AutoHotkey (v1.1+ or v2.0-a+)
 * Installation:
 *     Use #Include JSON.ahk or copy into a function library folder and then
 *     use #Include <JSON>
 * Links:
 *     GitHub:     - https://github.com/cocobelgica/AutoHotkey-JSON
 *     Forum Topic - http://goo.gl/r0zI8t
 *     Email:      - cocobelgica <at> gmail <dot> com
 */


/**
 * Class: JSON
 *     The JSON object contains methods for parsing JSON and converting values
 *     to JSON. Callable - NO; Instantiable - YES; Subclassable - YES;
 *     Nestable(via #Include) - NO.
 * Methods:
 *     Load() - see relevant documentation before method definition header
 *     Dump() - see relevant documentation before method definition header
 */
 
 

 
class JSON
{
	/**
	 * Method: Load
	 *     Parses a JSON string into an AHK value
	 * Syntax:
	 *     value := JSON.Load( text [, reviver ] )
	 * Parameter(s):
	 *     value      [retval] - parsed value
	 *     text    [in, ByRef] - JSON formatted string
	 *     reviver   [in, opt] - function object, similar to JavaScript's
	 *                           JSON.parse() 'reviver' parameter
	 */
     
     
	class Load extends JSON.Functor
	{
		Call(self, ByRef text, reviver:="")
		{
			this.rev := IsObject(reviver) ? reviver : false
		; Object keys(and array indices) are temporarily stored in arrays so that
		; we can enumerate them in the order they appear in the document/text instead
		; of alphabetically. Skip if no reviver function is specified.
			this.keys := this.rev ? {} : false

			static quot := Chr(34), bashq := "\" . quot
			     , json_value := quot . "{[01234567890-tfn"
			     , json_value_or_array_closing := quot . "{[]01234567890-tfn"
			     , object_key_or_object_closing := quot . "}"

			key := ""
			is_key := false
			root := {}
			stack := [root]
			next := json_value
			pos := 0

			while ((ch := SubStr(text, ++pos, 1)) != "") {
				if InStr(" `t`r`n", ch)
					continue
				if !InStr(next, ch, 1)
					this.ParseError(next, text, pos)

				holder := stack[1]
				is_array := holder.IsArray

				if InStr(",:", ch) {
					next := (is_key := !is_array && ch == ",") ? quot : json_value

				} else if InStr("}]", ch) {
					ObjRemoveAt(stack, 1)
					next := stack[1]==root ? "" : stack[1].IsArray ? ",]" : ",}"

				} else {
					if InStr("{[", ch) {
					; Check if Array() is overridden and if its return value has
					; the 'IsArray' property. If so, Array() will be called normally,
					; otherwise, use a custom base object for arrays
						static json_array := Func("Array").IsBuiltIn || ![].IsArray ? {IsArray: true} : 0
					
					; sacrifice readability for minor(actually negligible) performance gain
						(ch == "{")
							? ( is_key := true
							  , value := {}
							  , next := object_key_or_object_closing )
						; ch == "["
							: ( value := json_array ? new json_array : []
							  , next := json_value_or_array_closing )
						
						ObjInsertAt(stack, 1, value)

						if (this.keys)
							this.keys[value] := []
					
					} else {
						if (ch == quot) {
							i := pos
							while (i := InStr(text, quot,, i+1)) {
								value := StrReplace(SubStr(text, pos+1, i-pos-1), "\\", "\u005c")

								static tail := A_AhkVersion<"2" ? 0 : -1
								if (SubStr(value, tail) != "\")
									break
							}

							if (!i)
								this.ParseError("'", text, pos)

							  value := StrReplace(value,  "\/",  "/")
							, value := StrReplace(value, bashq, quot)
							, value := StrReplace(value,  "\b", "`b")
							, value := StrReplace(value,  "\f", "`f")
							, value := StrReplace(value,  "\n", "`n")
							, value := StrReplace(value,  "\r", "`r")
							, value := StrReplace(value,  "\t", "`t")

							pos := i ; update pos
							
							i := 0
							while (i := InStr(value, "\",, i+1)) {
								if !(SubStr(value, i+1, 1) == "u")
									this.ParseError("\", text, pos - StrLen(SubStr(value, i+1)))

								uffff := Abs("0x" . SubStr(value, i+2, 4))
								if (A_IsUnicode || uffff < 0x100)
									value := SubStr(value, 1, i-1) . Chr(uffff) . SubStr(value, i+6)
							}

							if (is_key) {
								key := value, next := ":"
								continue
							}
						
						} else {
							value := SubStr(text, pos, i := RegExMatch(text, "[\]\},\s]|$",, pos)-pos)

							static number := "number", integer :="integer"
							if value is %number%
							{
								if value is %integer%
									value += 0
							}
							else if (value == "true" || value == "false")
								value := %value% + 0
							else if (value == "null")
								value := ""
							else
							; we can do more here to pinpoint the actual culprit
							; but that's just too much extra work.
								this.ParseError(next, text, pos, i)

							pos += i-1
						}

						next := holder==root ? "" : is_array ? ",]" : ",}"
					} ; If InStr("{[", ch) { ... } else

					is_array? key := ObjPush(holder, value) : holder[key] := value

					if (this.keys && this.keys.HasKey(holder))
						this.keys[holder].Push(key)
				}
			
			} ; while ( ... )

			return this.rev ? this.Walk(root, "") : root[""]
		}

		ParseError(expect, ByRef text, pos, len:=1)
		{
			static quot := Chr(34), qurly := quot . "}"
			
			line := StrSplit(SubStr(text, 1, pos), "`n", "`r").Length()
			col := pos - InStr(text, "`n",, -(StrLen(text)-pos+1))
			msg := Format("{1}`n`nLine:`t{2}`nCol:`t{3}`nChar:`t{4}"
			,     (expect == "")     ? "Extra data"
			    : (expect == "'")    ? "Unterminated string starting at"
			    : (expect == "\")    ? "Invalid \escape"
			    : (expect == ":")    ? "Expecting ':' delimiter"
			    : (expect == quot)   ? "Expecting object key enclosed in double quotes"
			    : (expect == qurly)  ? "Expecting object key enclosed in double quotes or object closing '}'"
			    : (expect == ",}")   ? "Expecting ',' delimiter or object closing '}'"
			    : (expect == ",]")   ? "Expecting ',' delimiter or array closing ']'"
			    : InStr(expect, "]") ? "Expecting JSON value or array closing ']'"
			    :                      "Expecting JSON value(string, number, true, false, null, object or array)"
			, line, col, pos)

			static offset := A_AhkVersion<"2" ? -3 : -4
			throw Exception(msg, offset, SubStr(text, pos, len))
		}

		Walk(holder, key)
		{
			value := holder[key]
			if IsObject(value) {
				for i, k in this.keys[value] {
					; check if ObjHasKey(value, k) ??
					v := this.Walk(value, k)
					if (v != JSON.Undefined)
						value[k] := v
					else
						ObjDelete(value, k)
				}
			}
			
			return this.rev.Call(holder, key, value)
		}
	}

	/**
	 * Method: Dump
	 *     Converts an AHK value into a JSON string
	 * Syntax:
	 *     str := JSON.Dump( value [, replacer, space ] )
	 * Parameter(s):
	 *     str        [retval] - JSON representation of an AHK value
	 *     value          [in] - any value(object, string, number)
	 *     replacer  [in, opt] - function object, similar to JavaScript's
	 *                           JSON.stringify() 'replacer' parameter
	 *     space     [in, opt] - similar to JavaScript's JSON.stringify()
	 *                           'space' parameter
	 */
    
	class Dump extends JSON.Functor
	{
		Call(self, value, replacer:="", space:="")
		{
			this.rep := IsObject(replacer) ? replacer : ""

			this.gap := ""
			if (space) {
				static integer := "integer"
				if space is %integer%
					Loop, % ((n := Abs(space))>10 ? 10 : n)
						this.gap .= " "
				else
					this.gap := SubStr(space, 1, 10)

				this.indent := "`n"
			}

			return this.Str({"": value}, "")
		}

		Str(holder, key)
		{
			value := holder[key]

			if (this.rep)
				value := this.rep.Call(holder, key, ObjHasKey(holder, key) ? value : JSON.Undefined)

			if IsObject(value) {
			; Check object type, skip serialization for other object types such as
			; ComObject, Func, BoundFunc, FileObject, RegExMatchObject, Property, etc.
				static type := A_AhkVersion<"2" ? "" : Func("Type")
				if (type ? type.Call(value) == "Object" : ObjGetCapacity(value) != "") {
					if (this.gap) {
						stepback := this.indent
						this.indent .= this.gap
					}

					is_array := value.IsArray
				; Array() is not overridden, rollback to old method of
				; identifying array-like objects. Due to the use of a for-loop
				; sparse arrays such as '[1,,3]' are detected as objects({}). 
					if (!is_array) {
						for i in value
							is_array := i == A_Index
						until !is_array
					}

					str := ""
					if (is_array) {
						Loop, % value.Length() {
							if (this.gap)
								str .= this.indent
							
							v := this.Str(value, A_Index)
							str .= (v != "") ? v . "," : "null,"
						}
					} else {
						colon := this.gap ? ": " : ":"
						for k in value {
							v := this.Str(value, k)
							if (v != "") {
								if (this.gap)
									str .= this.indent

								str .= this.Quote(k) . colon . v . ","
							}
						}
					}

					if (str != "") {
						str := RTrim(str, ",")
						if (this.gap)
							str .= stepback
					}

					if (this.gap)
						this.indent := stepback

					return is_array ? "[" . str . "]" : "{" . str . "}"
				}
			
			} else ; is_number ? value : "value"
				return ObjGetCapacity([value], 1)=="" ? value : this.Quote(value)
		}

		Quote(string)
		{
			static quot := Chr(34), bashq := "\" . quot

			if (string != "") {
				  string := StrReplace(string,  "\",  "\\")
				; , string := StrReplace(string,  "/",  "\/") ; optional in ECMAScript
				, string := StrReplace(string, quot, bashq)
				, string := StrReplace(string, "`b",  "\b")
				, string := StrReplace(string, "`f",  "\f")
				, string := StrReplace(string, "`n",  "\n")
				, string := StrReplace(string, "`r",  "\r")
				, string := StrReplace(string, "`t",  "\t")

				static rx_escapable := A_AhkVersion<"2" ? "O)[^\x20-\x7e]" : "[^\x20-\x7e]"
				while RegExMatch(string, rx_escapable, m)
					string := StrReplace(string, m.Value, Format("\u{1:04x}", Ord(m.Value)))
			}

			return quot . string . quot
		}
	}

	/**
	 * Property: Undefined
	 *     Proxy for 'undefined' type
	 * Syntax:
	 *     undefined := JSON.Undefined
	 * Remarks:
	 *     For use with reviver and replacer functions since AutoHotkey does not
	 *     have an 'undefined' type. Returning blank("") or 0 won't work since these
	 *     can't be distnguished from actual JSON values. This leaves us with objects.
	 *     Replacer() - the caller may return a non-serializable AHK objects such as
	 *     ComObject, Func, BoundFunc, FileObject, RegExMatchObject, and Property to
	 *     mimic the behavior of returning 'undefined' in JavaScript but for the sake
	 *     of code readability and convenience, it's better to do 'return JSON.Undefined'.
	 *     Internally, the property returns a ComObject with the variant type of VT_EMPTY.
	 */
     
     
	Undefined[]
	{
		get {
			static empty := {}, vt_empty := ComObject(0, &empty, 1)
			return vt_empty
		}
	}

	class Functor
	{
		__Call(method, ByRef arg, args*)
		{
		; When casting to Call(), use a new instance of the "function object"
		; so as to avoid directly storing the properties(used across sub-methods)
		; into the "function object" itself.
			if IsObject(method)
				return (new this).Call(method, arg, args*)
			else if (method == "")
				return (new this).Call(arg, args*)
		}
	}
}




GuiButtonIcon(Handle, File, Index := 1, Options := "")
{
	RegExMatch(Options, "i)w\K\d+", W), (W="") ? W := 16 :
	RegExMatch(Options, "i)h\K\d+", H), (H="") ? H := 16 :
	RegExMatch(Options, "i)s\K\d+", S), S ? W := H := S :
	RegExMatch(Options, "i)l\K\d+", L), (L="") ? L := 0 :
	RegExMatch(Options, "i)t\K\d+", T), (T="") ? T := 0 :
	RegExMatch(Options, "i)r\K\d+", R), (R="") ? R := 0 :
	RegExMatch(Options, "i)b\K\d+", B), (B="") ? B := 0 :
	RegExMatch(Options, "i)a\K\d+", A), (A="") ? A := 4 :
	Psz := A_PtrSize = "" ? 4 : A_PtrSize, DW := "UInt", Ptr := A_PtrSize = "" ? DW : "Ptr"
	VarSetCapacity( button_il, 20 + Psz, 0 )
	NumPut( normal_il := DllCall( "ImageList_Create", DW, W, DW, H, DW, 0x21, DW, 1, DW, 1 ), button_il, 0, Ptr )	; Width & Height
	NumPut( L, button_il, 0 + Psz, DW )		; Left Margin
	NumPut( T, button_il, 4 + Psz, DW )		; Top Margin
	NumPut( R, button_il, 8 + Psz, DW )		; Right Margin
	NumPut( B, button_il, 12 + Psz, DW )	; Bottom Margin	
	NumPut( A, button_il, 16 + Psz, DW )	; Alignment
	SendMessage, BCM_SETIMAGELIST := 5634, 0, &button_il,, AHK_ID %Handle%
	return IL_Add( normal_il, File, Index )
}



