VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin DDSharpGridOLEDB2.SGGrid SGGrid1 
      Height          =   8535
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   7335
      _cx             =   12938
      _cy             =   15055
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   1
      HeadingColCount =   1
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   -2147483633
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   0
      ColorEven       =   -2147483628
      ColorOdd        =   -2147483624
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   2
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   1
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   0
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1200
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   0
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   -1  'True
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   225
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   1
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   -1  'True
      GroupByBoxVisible=   -1  'True
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"Form2.frx":0000
      ColumnsCollection=   $"Form2.frx":1DDF
      ValueItems      =   $"Form2.frx":2C38
   End
   Begin VB.CommandButton Command 
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SGGrid1_ColHeadClick(ByVal ColIndex As Long)
   Dim sColKey As String, lRowKey As Long
   Dim vValue As Variant, row As SGRow
   Dim lStartRowPos As Long
   'get column and row keys
   sColKey = SGGrid1.Columns(ColIndex).Key
   lRowKey = SGGrid1.Rows.At(1).Key
   'get the first row value
   vValue = SGGrid1.Cell(lRowKey, sColKey).value
   'check whether the top row is the last one
   If SGGrid1.TopRow + 1 = SGGrid1.RowCount Then
      lStartRowPos = 1
   Else
      lStartRowPos = SGGrid1.TopRow + 1
   End If
   'try to find a row using the value from the first row
   If ColIndex = 1 Then
      Set row = SGGrid1.Rows.Find(sColKey, sgOpContains, vValue, , lStartRowPos)
   Else
      Set row = SGGrid1.Rows.Find(sColKey, sgOpEqual, vValue, , lStartRowPos)
   End If
   
   If row Is Nothing Then
      'there is not a row that meets the criteria
      SGGrid1.TopRow = 1
   Else
      'there is a row that meets the criteria so move it to the top
      SGGrid1.TopRow = row.Position
      Set row = Nothing
   End If
End Sub

Private Sub SGGrid1_KeyPressEdit _
   (ByVal RowKey As Long, ByVal ColIndex As Long, KeyASCII As Integer)
   'exit cell-editing mode on the return key
   If KeyASCII = vbKeyReturn And SGGrid1.row = 1 Then _
      SGGrid1.Editing = False
End Sub

Private Sub SGGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   'use the return key to find a row
   If KeyCode = vbKeyReturn And SGGrid1.row = 1 Then _
      SGGrid1_ColHeadClick SGGrid1.Columns.Current.ColIndex
End Sub

Private Sub SGGrid1_OnInit()
   Dim sFile As String
   Dim sgCol As SGColumn
   Dim sWinDir As String
   Dim i As Integer
   'disable redraw to speed up process
   SGGrid1.RedrawEnabled = False
   'remove all columns
   SGGrid1.Columns.RemoveAll False
   SGGrid1.RowNumbering = True
   'use tab key to navigate columns
   SGGrid1.TabKeyBehavior = sgTabColumns
   SGGrid1.NavigationWrapMode = sgNavigationWrapSame
   'Create columns
   SGGrid1.Columns.Add "File"
   Set sgCol = SGGrid1.Columns.Add("FileDate")
   sgCol.DataType = sgtDateTime
   sgCol.SortType = sgSortTypeDateTime
   Set sgCol = SGGrid1.Columns.Add("Size")
   sgCol.DataType = sgtLong
   
   sWinDir = Environ("WINDIR") & "\"
   i = 1
   sFile = Dir(sWinDir & "*.*", vbNormal)
   'add files to the grid
   Do Until Len(sFile) = 0
      i = i + 1
      SGGrid1.DataRowCount = i
      SGGrid1.Rows.At(i).Cells(1).value = sFile
      SGGrid1.Rows.At(i).Cells(2).value = _
         FormatDateTime(FileDateTime(sWinDir & sFile), vbShortDate)
      SGGrid1.Rows.At(i).Cells(3).value = FileLen(sWinDir & sFile)
      sFile = Dir
   Loop
   
   SGGrid1.Rows.At(1).Frozen = True
   SGGrid1.RedrawEnabled = True
End Sub


