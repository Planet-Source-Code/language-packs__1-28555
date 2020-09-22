Attribute VB_Name = "mEPTLoad"
Option Compare Text
Option Explicit

Const EPT_ID As String = "EPT"

Const EPT_COMMENT_MARK As String = ";"
Const EPT_HEADER_MARK As String = ">"
Const EPT_PROPERTY_MARK As String = "."
Const EPT_ASSIGN_MARK As String = "="

Const EPT_GLOBAL_MODEL As String = "*" & EPT_ASSIGN_MARK & "*"
Const EPT_FORM_MODEL As String = "*" & EPT_PROPERTY_MARK & "*" & EPT_PROPERTY_MARK & "???" ' & EPT_ASSIGN_MARK & "*"
'Const EPT_FORM_MODEL As String = "*" & EPT_PROPERTY_MARK & "*" & EPT_PROPERTY_MARK & "???"
Const EPT_COMMENT As String = EPT_COMMENT_MARK & "*"
Const EPT_HEADER As String = EPT_HEADER_MARK & "*"

'Const EPT_CONSTANT_MARK As String = "%%"
'Const EPT_CONSTANT As String = EPT_CONSTANT_MARK & "*"

Const EPT_ARRAY_RANGE As String = "[0-9]-[0-9]"
Const EPT_ARRAY_ALL As String = "!"

Const EPT_REPLACE_LINEBREAK As String = "//BR"
'Const EPT_REPLACE_INDEX As String = "//IND"

Const EPT_PROP_CAPTION As String = "CAP"
Const EPT_PROP_TOOLTIP As String = "TIP"
Const EPT_PROP_TAG As String = "TAG"
Const EPT_PROP_TEXT As String = "TXT"

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

Function EPTEnumHeaders(FileName As String) As Variant
    If Not IsEPT(FileName) Then Exit Function
    
    Dim FileID As Integer
    Dim TempLine As String
    Dim Headers() As String
    Dim NumHeaders As Integer
    
    NumHeaders = -1
    
    FileID = FreeFile
    
    Open FileName For Input As #FileID
    
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

Sub LoadEPT(FileName As String, Header As String, Object As Object)
    On Error Resume Next
   
    Dim Lines() As String
    Dim Parse() As String
    Dim A As Integer
    Dim B As Integer
    
    Lines = EPTGetProperties(FileName, Header)
    
    If TypeOf Object Is Collection And Not IsEmpty(Object) Then
        Set Object = Nothing
        Set Object = New Collection
    End If
    
    For A = 0 To UBound(Lines)
        
        Parse = Split(Lines(A), EPT_ASSIGN_MARK)
        
        For B = 0 To UBound(Parse)
            Parse(B) = Trim(Parse(B))
        Next B
        
        Parse(UBound(Parse)) = Replace(Parse(UBound(Parse)), EPT_REPLACE_LINEBREAK, vbCrLf)
        
        For B = 0 To UBound(Parse) - 1
            If TypeOf Object Is Form Then
                
                Dim SubParse() As String
                Dim MyControl() As Object
                Dim C As Integer
                
                With Object
                    If Parse(B) Like EPT_FORM_MODEL Then
                        SubParse = Split(Parse(B), EPT_PROPERTY_MARK)
        
                        ReDim MyControl(0)
                        
                        If SubParse(0) = .Name Then
                            Set MyControl(0) = Object
                        Else
                            If IsControlArray(.Controls(SubParse(0))) Then
                                If SubParse(1) Like EPT_ARRAY_RANGE Then
                                    Dim RangeParse() As String
                                    RangeParse = Split(SubParse(1), "-")
                                    ReDim Preserve MyControl(CInt(RangeParse(1)) - CInt(RangeParse(0)))
                                    For C = 0 To CInt(RangeParse(1)) - CInt(RangeParse(0))
                                        Set MyControl(B) = .Controls(SubParse(0)).Item(B + CInt(RangeParse(0)))
                                    Next C
                                    Erase RangeParse
                                ElseIf SubParse(1) Like EPT_ARRAY_ALL Then
                                    ReDim Preserve MyControl(.Controls(SubParse(0)).UBound - .Controls(SubParse(0)).LBound)
                                    For C = 0 To .Controls(SubParse(0)).UBound - .Controls(SubParse(0)).LBound
                                        Set MyControl(B) = .Controls(SubParse(0)).Item(B + .Controls(SubParse(0)).LBound)
                                    Next C
                                Else
                                    Set MyControl(0) = .Controls(SubParse(0)).Item(SubParse(1))
                                End If
                            Else
                                Set MyControl(0) = .Controls(SubParse(0))
                            End If
                        End If
                        
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
                        
                        Erase MyControl
                        Erase SubParse
                    End If
                End With
            ElseIf TypeOf Object Is Collection Then
                Object.Add Parse(UBound(Parse)), Parse(B)
            End If
        Next B
    Next A
    
    Erase Lines
    Erase Parse
End Sub

Function IsControlArray(Control As Object) As Boolean
    On Error GoTo NotArray
    
    Dim TMP As Integer
    
    TMP = Control.UBound
    IsControlArray = True
    
    Exit Function
    
NotArray:
    IsControlArray = False
End Function
