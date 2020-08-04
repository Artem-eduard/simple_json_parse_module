VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strFileName As String
strFileName = App.Path & "\sample.json"


  Set p = JSON.parse(ReadTextFile(strFileName))
      If Not (p Is Nothing) Then
         If JSON.GetParserErrors <> "" Then
            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
         Else
         
          Set p2 = p.Item("result")
          Dim count As Integer
          count = p.Item("result").count
          For i = 1 To count
                   
            Debug.Print "hotel_id = " & JSON.toString(p.Item("result").Item(i).Item("hotel_id"))
            
            Dim count_of_roomdata As Integer
            count_of_roomdata = p.Item("result").Item(i).Item("room_data").count
            Dim max_price As Double
            max_price = 0
            Dim min_price As Double
            min_price = 9999999
            Dim str_max_price As String
            Dim str_min_price As String
            
            For j = 1 To count_of_roomdata
                         
                str_max_price = JSON.toString(p.Item("result").Item(i).Item("room_data").Item(j).Item("room_info").Item("max_price"))
                str_min_price = JSON.toString(p.Item("result").Item(i).Item("room_data").Item(j).Item("room_info").Item("min_price"))
                If max_price < CDbl(Val(str_max_price)) Then
                 max_price = CDbl(Val(str_max_price))
                 End If
                If min_price > CDbl(Val(str_min_price)) Then
                 min_price = CDbl(Val(str_min_price))
                End If
                
                
            Next j
            
            Debug.Print "max_price = " & max_price
            Debug.Print "min_price = " & min_price
           ' Debug.Print "main_photo = " & JSON.toString(p.Item("result").Item(i).Item("hotel_data").Item(j).Item("main_photo"))
          Next i
           

         End If
      Else
         MsgBox "An error occurred parsing " & cd.FileName
      End If
End Sub


Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
   
      handle = FreeFile
      Open sFilePath For Binary As #handle
      ReadTextFile = Space$(LOF(handle))
      Get #handle, , ReadTextFile
      Close #handle
      
   End If
   
End Function

Private Sub Command2_Click()
Dim strFileName As String
strFileName = App.Path & "\sample2.json"


  Set p = JSON.parse(ReadTextFile(strFileName))
      If Not (p Is Nothing) Then
         If JSON.GetParserErrors <> "" Then
            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
         Else
         
          Set p2 = p.Item("result")
          Dim count As Integer
          count = p.Item("result").count
          For i = 1 To count
                   
            Debug.Print "hotel_id = " & JSON.toString(p.Item("result").Item(i).Item("hotel_id"))
            
            Dim count_of_hotel_photos As Integer
            count_of_hotel_photos = p.Item("result").Item(i).Item("hotel_data").Item("hotel_photos").count
      
            Dim strphotoflag As Boolean
            For j = 1 To count_of_hotel_photos
                         
              ' Debug.Print "main_photo = " & JSON.toString(p.Item("result").Item(i).Item("hotel_data").Item("hotel_photos").Item(j).Item("main_photo"))
              strphotoflag = p.Item("result").Item(i).Item("hotel_data").Item("hotel_photos").Item(j).Item("main_photo")
              If strphotoflag = True Then
              
              Debug.Print JSON.toString(p.Item("result").Item(i).Item("hotel_data").Item("hotel_photos").Item(j).Item("url_original"))
              
              End If
              
                
                
            Next j
            
           
          
          Next i
           

         End If
      Else
         MsgBox "An error occurred parsing " & cd.FileName
      End If
End Sub
