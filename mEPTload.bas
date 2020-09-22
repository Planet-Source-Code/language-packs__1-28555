Attribute VB_Name = "ModuleEPTLoad"
' ---------------------------------------------------
' External Properties Tables(EPT) file format loader
' Both this code and the file format are copyrighted
' by Pedro Fayolle (a.k.a. piLaF-kun)
'
' E-mail:     · pedrofayolle@fibertel.com.ar
'
' Home pages: · http://pcsoft.cjb.net
'             · http://pilafkun.cjb.net
'
' This example code was made especially for Planet
' Source Code · http://www.planetsourcecode.com
'
' ---------------------------------------------------
' The main objective of this file format and code
' is to provide VB developers with a method to enable
' the use of Language Packs for their apps that does
' not suck.
'
' Feel free to use this module with any of your
' applications. Of course I would be very thankful if
' you could mention me somewhere in the credits of
' your app if you do. Also an e-mail would be nice :)
'
' If you feel there's something that has to be added/
' removed/fixed, you can contact me about it or do it
' yourself(in which case it would be very cool if you
' could send me the modification).
'
' Please enjoy!
' ---------------------------------------------------

' Debug options
Option Compare Text
Option Explicit

' Constants used for verifications and parsing
Const EPT_ID As String = "EPT"
Const EPT_COMMENT_MARK As String = ";"
Const EPT_HEADER_MARK As String = ">"
Const EPT_PROPERTY_MARK As String = "."
Const EPT_ASSIGN_MARK As String = "="
Const EPT_GLOBAL_MODEL As String = "*" & EPT_ASSIGN_MARK & "*"
Const EPT_FORM_MODEL As String = "*" & EPT_PROPERTY_MARK & "*" & EPT_PROPERTY_MARK & "???"
Const EPT_COMMENT As String = EPT_COMMENT_MARK & "*"
Const EPT_HEADER As String = EPT_HEADER_MARK & "*"

' Constants used for parser specifications
Const EPT_ARRAY_RANGE As String = "[0-9]-[0-9]"
Const EPT_ARRAY_ALL As String = "!"

' This constant is replaced by a linebreak
' during the parsing process.
Const EPT_REPLACE_LINEBREAK As String = "//BR"

' This constants are used to target control
' properties.
Const EPT_PROP_CAPTION As String = "CAP"
Const EPT_PROP_TOOLTIP As String = "TIP"
Const EPT_PROP_TAG As String = "TAG"
Const EPT_PROP_TEXT As String = "TXT"

' This function returns True if the file pointed
' is a valid EPT file. There isn't much to explain
' here.
Function IsEPT(FileName As String) As Boolean
    If Dir(FileName) = "" Then Exit Function
    
    Dim FileID As Integer
    Dim TempLine As String
    
    FileID = FreeFile
    
    Open FileName For Input As #FileID
    
    Line Input #FileID, TempLine
    
    IsEPT = (TempLine = EPT_ID)
    
    Close #FileID
End Function

' This function enumerates all headers in an EPT
' file and returns all their names in an array.
Function EPTEnumHeaders(FileName As String) As Variant

    ' Check if the file is a valid EPT file.
    If Not IsEPT(FileName) Then Exit Function
    
    Dim FileID As Integer
    Dim TempLine As String
    Dim Headers() As String
    Dim NumHeaders As Integer
    
    NumHeaders = -1
    
    FileID = FreeFile
    
    Open FileName For Input As #FileID
    
    ' Loop through each line of the file searching
    ' for headers.
    Do Until EOF(FileID)
        
        Line Input #FileID, TempLine
        
        TempLine = Trim(TempLine)
        
        If TempLine Like EPT_HEADER And Not Len(TempLine) = 1 Then
        
            NumHeaders = NumHeaders + 1
            
            ReDim Preserve Headers(NumHeaders)
            
            Headers(NumHeaders) = Right(TempLine, Len(TempLine) - 1)
            
        End If
        
    Loop
    
    If Not NumHeaders = -1 Then EPTEnumHeaders = Headers
    
    Close #FileID
End Function

' This function enumerates all properties found
' under a specified header in the pointed file.
Function EPTGetProperties(FileName As String, Header As String) As Variant

    Header = Trim(Header)
    
    If Not IsEPT(FileName) Or Header = "" Then Exit Function
    
    Dim Lines() As String
    Dim TempLine As String
    Dim FoundHeader As Boolean
    Dim FileID As Integer
    Dim NumLines As Long
    
    NumLines = -1
    FileID = FreeFile
    
    Open FileName For Input As #FileID
    
    ' Loop through each line under the header
    ' checking if it's a valid properties line.
    Do Until EOF(FileID)
        
        Line Input #FileID, TempLine
        
        TempLine = Trim(TempLine)
        
        If TempLine Like EPT_COMMENT Or TempLine = "" Then GoTo NextLine
        
        If FoundHeader Then
            
            If TempLine Like EPT_HEADER Then Exit Do
            
            If TempLine Like EPT_GLOBAL_MODEL Then
                NumLines = NumLines + 1
                
                ReDim Preserve Lines(NumLines) As String
                
                Lines(NumLines) = TempLine
            End If
            
        End If
        
        If TempLine = EPT_HEADER_MARK & Header Then FoundHeader = True
        
NextLine:

    Loop
    
    Close #FileID
    
    EPTGetProperties = Lines
    
    Erase Lines
End Function

' The major function, this is what loads the
' values from an EPT file into each object.
' There's two kind of objects that can be passed
' to the Object parameter of this function:
' Forms and Collections.
'
' When a Form is passed, this function will load
' all properties found under the specified Header
' into the controls of the form.
'
' Collections can be used for translation of
' message boxes(MsgBox(...)) and other things that
' aren't permanently shown.
'
' If you think comments are making the code hard
' to read just erase them :P
Sub LoadEPT(FileName As String, Header As String, Object As Object)
    'On Error Resume Next

    ' This string array will store all lines under
    ' the header, except by comments.
    Dim Lines() As String
    
    ' This string array will store each part of
    ' a property line.
    Dim Parse() As String
    
    ' This variables are used for the For-Next loops.
    Dim A As Integer
    Dim B As Integer
    
    ' Enumerate properties. This also verifies
    ' the file is valid.
    Lines = EPTGetProperties(FileName, Header)
    
    ' If the object passed is a collection, set it up.
    If TypeOf Object Is Collection And Not IsEmpty(Object) Then
        Set Object = Nothing
        Set Object = New Collection
    End If
    
    ' Start the mother loop. This loops through all
    ' the property lines.
    For A = 0 To UBound(Lines)
        
        ' This splits each part of a property line
        ' into many variables.
        '
        ' All but the last variables in the array
        ' will be the properties which to assign the
        ' value. The value to assign is stored in the
        ' last variable of the array.
        '
        ' Normally there's just a property and a value,
        ' but by using this more than one property can
        ' receive the same value.
        '
        ' A normal property line of these looks like:
        '
        '   mycontrol..CAP = [value]
        '
        ' or with multiple recipients...:
        '
        '   mycontrol1..CAP = mycontrol2..CAP = [value]
        '
        ' There can be as many recipients as you want,
        ' of course.
        Parse = Split(Lines(A), EPT_ASSIGN_MARK)
        
        ' Clean all spaces before and after each part.
        For B = 0 To UBound(Parse)
            Parse(B) = Trim(Parse(B))
        Next B
        
        ' Place linebreaks if needed.
        Parse(UBound(Parse)) = Replace(Parse(UBound(Parse)), EPT_REPLACE_LINEBREAK, vbCrLf)
        
        ' First child loop. Loops through each recipient
        ' of the property line and applies the value.
        For B = 0 To UBound(Parse) - 1

            If TypeOf Object Is Form Then
                
                ' Control propertiy lines look slightly different
                ' from common property lines. This is why we need
                ' a new variables array where to store the new
                ' parsing.
                Dim SubParse() As String
                
                ' This is a pointer to the controls whose properties
                ' will be changed. There might be more than one
                ' control, that's why it's an array.
                Dim MyControl() As Object
                
                ' This is used for the third-class loops.
                Dim C As Integer
                
                With Object
                    If Parse(B) Like EPT_FORM_MODEL Then
                        ' Parse once again.
                        SubParse = Split(Parse(B), EPT_PROPERTY_MARK)

                        ReDim MyControl(0)
                        
                        ' If the owner of the property is the form
                        ' itself, put it to the array and go on... :)
                        If SubParse(0) = .Name Then
                            Set MyControl(0) = Object
                            
                        ' ...if not and...
                        Else
                            ' ...if the control is an array then,
                            ' get all the controls into the array :P
                            '
                            ' I don't really feel like explaining this part,
                            ' sorry :/
                            If IsControlArray(.Controls(SubParse(0))) Then
                                If SubParse(1) Like EPT_ARRAY_RANGE Then
                                    Dim RangeParse() As String
                                    RangeParse = Split(SubParse(1), "-")
                                    ReDim Preserve MyControl(CInt(RangeParse(1)) - CInt(RangeParse(0)))
                                    For C = 0 To CInt(RangeParse(1)) - CInt(RangeParse(0))
                                        Set MyControl(C) = .Controls(SubParse(0)).Item(C + CInt(RangeParse(0)))
                                    Next C
                                    Erase RangeParse
                                ElseIf SubParse(1) Like EPT_ARRAY_ALL Then
                                    ReDim Preserve MyControl(.Controls(SubParse(0)).UBound - .Controls(SubParse(0)).LBound)
                                    For C = 0 To .Controls(SubParse(0)).UBound - .Controls(SubParse(0)).LBound
                                        Set MyControl(C) = .Controls(SubParse(0)).Item(C + .Controls(SubParse(0)).LBound)
                                    Next C
                                Else
                                    Set MyControl(0) = .Controls(SubParse(0)).Item(SubParse(1))
                                End If
                            Else
                                Set MyControl(0) = .Controls(SubParse(0))
                            End If
                        End If
                        
                        ' Loop through every collected control
                        ' and apply a value to the corresponding
                        ' property. And.. voilá! :D
                        For C = 0 To UBound(MyControl)
                            Select Case Trim(SubParse(2))
                                Case EPT_PROP_CAPTION
                                    MyControl(C).Caption = Parse(UBound(Parse))
                                Case EPT_PROP_TOOLTIP
                                    MyControl(C).ToolTipText = Parse(UBound(Parse))
                                Case EPT_PROP_TAG
                                    MyControl(C).Tag = Parse(UBound(Parse))
                                Case EPT_PROP_TEXT
                                    MyControl(C).Text = Parse(UBound(Parse))
                            End Select
                        Next C
                        
                        ' Close...
                        Erase MyControl
                        Erase SubParse
                    ' Close...
                    End If
                ' Close...
                End With
            
            ' Before we forget... If instead of a form,
            ' the object is a collection, create all the
            ' items and assign all the values.
            ElseIf TypeOf Object Is Collection Then
                Object.Add Parse(UBound(Parse)), Parse(B)
            ' Close...
            End If
        ' Close...
        Next B
    ' Close...
    Next A
    
    ' Close...
    Erase Lines
    Erase Parse
' Close...
End Sub

' A little function used by LoadEPT(...). It verifies
' wether or not a control belongs to a controls array.
Function IsControlArray(Control As Object) As Boolean
    On Error GoTo NotArray
    
    Dim TMP As Integer
    
    TMP = Control.UBound
    IsControlArray = True
    
    Exit Function
    
NotArray:
    IsControlArray = False
End Function
