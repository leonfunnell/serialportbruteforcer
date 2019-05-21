
Option Explicit

Dim objComport, str, response, chargen, notExit, passfile, passlength, fso

passlength=8
'chargen="1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ;'#~@:[]{}!$£%^&*()_+\|/"
chargen="1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
passfile="C:\Program Files\ActiveXperts\Serial Port Component\Samples\VBScript\QueryDevice\passfile.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
Set objComport       = CreateObject( "AxSerial.ComPort" )
' Clear (good practise)
objComport.Clear()
objComport.Device = "COM3"

' Optionally override defaults for direct COM ports
If( Left( objComport.Device, 3 ) = "COM" ) Then
  objComport.BaudRate  = "115200"
' objComport.HardwareFlowControl  = True
' objComport.SoftwareFlowControl  = False
End If

' Open the port
objComport.Open
wscript.echo "Open, result:" & objComport.LastError & " (" & objComport.GetErrorDescription( objComport.LastError ) & ")"

If( objComport.LastError <> 0 ) Then
  WScript.Quit
End If

' Set all Read functions (e.g. ReadString) to timeout after a specified number of millisconds
objComport.ComTimeout = 50  ' Timeout after 500msecs 

notExit=TRUE
WriteCommand objComport,"" 
while notExit
response=ReadResponse(objComport)
'response="Password:"
WScript.StdOut.Write response

select case response
	Case "VMG1312-B10D login:"
		WriteCommand objComport,"admin" 

	Case "Password:"
		WriteCommand objComport,Passgen(passfile,passlength,chargen) 

	Case ""
		' do nothing
		
	Case "Please press Enter to activate this console."
		WriteCommand objComport,"" 
	Case ">"
		debug "did we find it???"
		str = WScript.StdIn.ReadLine
	
	Case "#"
		debug "did we find it???"
		str = WScript.StdIn.ReadLine
	
	Case else
		WriteCommand objComport,""
end select
	
WEnd

objComport.Close()
debug "Close, result: " & objComport.LastError & " (" & objComport.GetErrorDescription(objComport.LastError) & ")"

debug "Ready."



' ********************************************************************
' Sub Routines
' ********************************************************************

function ReadResponse(ByVal objComport)
  Dim str

  str = "notempty"
  objComport.Sleep(50)
  While (str <> "")
    str = objComport.ReadString()
    If (str <> "") Then
      ReadResponse = str
    End If

  WEnd
End function


' ********************************************************************

sub WriteCommand(ByVal objComport,ByVal str)
  wscript.echo str
  objComport.WriteString(str)
  If( objComport.LastError = 0 ) Then
  Else
    wscript.echo "Write failed, result: " & objComport.LastError & " (" & objComport.GetErrorDescription(objComport.LastError) & ")"
  End If

End sub

Function Passgen(passfile,passlength,chargen)
	dim objFile, strFileText, LenFileText, wkg, i, incflag, oldchar, newchar
	strFileText=" " ' initialise file
	if fso.FileExists(passfile) then
		Set objFile = fso.OpenTextFile(passfile,1)
		strFileText=objFile.ReadAll()
		objFile.close
		set objFile = nothing
		fso.DeleteFile passfile, True
	end if
	LenFileText=len(strFileText)
	incflag=true
		For i= LenFileText to 1 step -1
		if incflag=true then
			wkg=InStr(1,chargen,mid(strFileText,i,1),0)
			if wkg>=Len(chargen) then 
				wkg=1
				if i>=LenFileText-1 then 
					'nothing
				else 
					incflag=false
				end if
			else 
				wkg=wkg+1
				incflag=false
			end if
			oldchar=mid(strFileText,i,1)
			newchar=mid(chargen,wkg,1)
			if LenFileText>1 then strFileText=left(strFileText,i-1) & replace(strFileText,oldchar,newchar,i,1) else strFileText = newchar
		end if
	next
	if incflag then 
		strFileText = mid(chargen,1,1) & strFileText
	end if
	
	Set objFile = fso.CreateTextFile(passfile,true)
	objFile.Write strFileText
	objFile.close
	set objFile = nothing
	Passgen=strFileText
end function
	
