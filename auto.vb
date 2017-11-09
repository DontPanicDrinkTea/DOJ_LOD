Sub AllTheThings()
	Application.ScreenUpdating = False
	Call changecolour
	Call addborder
	Call fixDate
	Call fixAddressblock1
	Call fixAddressblock2
	Call fixAddressblock3
	Call unhideNumbers
	Call alignLastColText
	Call alignTableCols
	Call JustifyTextAlignment
	Application.ScreenUpdating = True
End Sub

Sub changecolour()''#1' change colour of all text to black'
	ActiveDocument.Range.Font.Color = -587137025
End Sub

Sub addborder()
'#2' add border to two lines on first page'
	With ActiveDocument.Content.Find
		.Text = "HER MAJESTY THE QUEEN^p^pRespondent^p^p"
		.Forward = True
		.Execute
		If .Found = True Then .Parent.Select
	End With

	With Selection.Borders(wdBorderBottom)
		.LineStyle = Options.DefaultBorderLineStyle
		.LineWidth = Options.DefaultBorderLineWidth
		.Color = Options.DefaultBorderColor
	End With

	With ActiveDocument.Content.Find
		.Text = "RESPONDENT'S LIST OF DOCUMENTS^p(Partial Disclosure)^p^p"
		.Forward = True
		.Execute
		If .Found = True Then .Parent.Select
	End With

	With Selection.Borders(wdBorderBottom)
		.LineStyle = Options.DefaultBorderLineStyle
		.LineWidth = Options.DefaultBorderLineWidth
		.Color = Options.DefaultBorderColor
	End With
End Sub

Sub fixDate()
'#3'find and replace date section with a slightly different one
    With ActiveDocument.Content.Find
		.MatchWildcards = True
		.Text = "DATED at Halifax, Nova Scotia*" & Year(Date)
		.Forward = True
		.Execute
		If .Found = True Then .Parent.Text = "DATED at Halifax, Nova Scotia on " & MonthName(Month(Date)) & " ___ " & Year(Date)
	End With
End Sub

Sub fixAddressblock1()
'#4'find and replace AGC address section with a slightly different one
    With ActiveDocument.Content.Find
		.Text = "Department of Justice Canada^pAtlantic Regional Office^pTax Law Services Section^pSuite 1400, Duke Tower^p5251 Duke Street^pHalifax, Nova Scotia^pB3J 1P3^pFax:  (902) 426-8802^p^p"
		.Forward = True 
		.Execute
		If .Found = True Then .Parent.Text = "Department of Justice Canada" & Chr(13) & "Atlantic Regional Office" & Chr(13) & "Tax Law Services Section" & Chr(13) & "Suite 1400, Duke Tower" & Chr(13) & "5251 Duke Street" & Chr(13) & "Halifax, NS   B3J 1P3" & Chr(13) & "Fax:  (902) 426-8802" & Chr(13) & Chr(13)
	End With
End Sub

Sub fixAddressblock2()
'#4-2'find and replace Registrar address section with a slightly different one
    With ActiveDocument.Content.Find
		.Text = "Tax Court of Canada^p200 Kent Street, 3rd Floor^pOttawa, Ontario^pK1A 0M1"
		.Forward = True 
		.Execute
		If .Found = True Then .Parent.Text = "Tax Court of Canada" & Chr(13) & "200 Kent Street, 3rd Floor" & Chr(13) & "Ottawa, ON   K1A 0M1"
	End With
End Sub

Sub fixAddressblock3()
'#4-3'look to see if there is an additional recipient (there should be at least one) and format the address block
Dim addrTable As Table
Dim addrTableCur As Table
Dim rng1 As Word.Range
Dim rng2 As Word.Range
Dim prov As String
Dim provAbrv As String
Dim provArray(12) As String
Dim provAbrvArray(12) As String
'find instance of recipient "AND TO:" and work within that table
For i = 1 To ActiveDocument.Tables.Count
Set addrTable = ActiveDocument.Tables(i)
With addrTable.Range.Cells(1).Range.Find
.Text = "AND TO"
.Forward = True
.Execute
If .Found = True Then Set addrTableCur = addrTable
If (Not (.Found = True)) Then GoTo Next i
End With

Set rng1 = addrTableCur.Cell(1, 2).Range.Characters(1)rng1.MoveEnd wdCharacter, addrTableCur.Cell(1, 2).Range.Characters.Count - 2
Set rng2 = addrTableCur.Cell(2, 2).Range.Characters(1)rng2.MoveEnd wdCharacter, addrTableCur.Cell(1, 2).Range.Characters.Count - 2

'If the name of the recipient appears twice, highlight the second occurrence for review
If rng1.Text = rng2.Text Then
rng2.HighlightColorIndex = wdYellow
End If

'Search for Province and replace with abbreviation
provArray(0) = "Alberta" & Chr(13)
provArray(1) = "British Columbia" & Chr(13)
provArray(2) = "Manitoba" & Chr(13)
provArray(3) = "New Brunswick" & Chr(13)
provArray(4) = "Newfoundland and Labrador" & Chr(13)
provArray(5) = "Northwest Territories" & Chr(13)
provArray(6) = "Nova Scotia" & Chr(13)
provArray(7) = "Nunavut" & Chr(13)
provArray(8) = "Ontario" & Chr(13)
provArray(9) = "Prince Edward Island" & Chr(13)
provArray(10) = "Quebec" & Chr(13)
provArray(11) = "Saskatchewan" & Chr(13)
provArray(12) = "Yukon" & Chr(13)
provAbrvArray(0) = "AB   "
provAbrvArray(1) = "BC   "
provAbrvArray(2) = "MB   "
provAbrvArray(3) = "NB   "
provAbrvArray(4) = "NL   "
provAbrvArray(5) = "NT   "
provAbrvArray(6) = "NS   "
provAbrvArray(7) = "NU   "
provAbrvArray(8) = "ON   "
provAbrvArray(9) = "PE   "
provAbrvArray(10) = "QC   "
provAbrvArray(11) = "SK   "
provAbrvArray(12) = "YT   "

'nts for later: find out if French or any other names should be added!

For j = 0 To 12
With addrTable.Range.Cells(4).Range.Find
.Text = provArray(j)
.Forward = True
.Execute
If .Found = True Then .Parent.Select
If (Not (.Found = True)) Then GoTo next j
End With
Selection.Text = provAbrvArray(j)

nextj:
Next j

Nexti:
Next i
End Sub

Sub unhideNumbers()
'#5'unhide numbers in column 3 of LOD
ActiveDocument.Tables(ActiveDocument.Tables.Count - 1).Columns(3).Select    With Selection.Font        If .Hidden = True Then .Hidden = False        If .Hidden = False Then .Bold = True    End With
End Sub

Sub alignLastColText()
'#6'change alignment of last column in LoD
ActiveDocument.Tables(ActiveDocument.Tables.Count - 1).Columns(4).Select    With Selection        .ParagraphFormat.Alignment = wdAlignParagraphLeft        .Cells.VerticalAlignment = wdCellAlignVerticalTop    End With
End Sub

Sub alignTableCols()''#7' align table column width''Dim lodTable As RangeDim schedBTable As Range
Set lodTable = ActiveDocument.Tables(ActiveDocument.Tables.Count - 1).RangeSet schedBTable = ActiveDocument.Tables(ActiveDocument.Tables.Count).Range
    With lodTable.ParagraphFormat        .RightIndent = InchesToPoints(0)        .SpaceBefore = 6        .SpaceBeforeAuto = False        .SpaceAfter = 6        .SpaceAfterAuto = False        .LineSpacingRule = wdLineSpaceSingle        .Alignment = wdAlignParagraphLeft        .WidowControl = True        .KeepWithNext = False        .KeepTogether = False        .PageBreakBefore = False        .NoLineNumber = False        .Hyphenation = True        .OutlineLevel = wdOutlineLevelBodyText        .CharacterUnitRightIndent = 0        .LineUnitBefore = 0        .LineUnitAfter = 0        .MirrorIndents = False        .TextboxTightWrap = wdTightNone        .CollapsedByDefault = False    End With
With lodTable
.Columns(2).SetWidth _ColumnWidth:=InchesToPoints(4.06), _RulerStyle:=wdAdjustNone
.Columns(3).SetWidth _ColumnWidth:=InchesToPoints(0.5), _RulerStyle:=wdAdjustNone
.Columns(4).SetWidth _ColumnWidth:=InchesToPoints(1.75), _RulerStyle:=wdAdjustNone
End With
With schedBTable
.Columns(2).SetWidth _ColumnWidth:=InchesToPoints(4.06), _RulerStyle:=wdAdjustNone
.Columns(3).SetWidth _ColumnWidth:=InchesToPoints(0.5), _RulerStyle:=wdAdjustNone
.Columns(4).SetWidth _ColumnWidth:=InchesToPoints(1.75), _RulerStyle:=wdAdjustNone
End With
End Sub

Sub JustifyTextAlignment()
'#8'Justify those few sections where it's not justified but maybe should be?
    With ActiveDocument.Content.Find    .Text = "TAKE NOTICE that the documents referred to in Schedule A below may be inspected and copies taken at 5251 Duke Street, Suite 1400, Halifax, Nova Scotia, on any weekday, by appointment, between the hours of 8:30 a.m. and 4:30 p.m."    .Forward = True    .Execute    If .Found = True Then .Parent.Select    End With        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    With ActiveDocument.Content.Find    .Text = "Documents of which the Respondent has knowledge but which are not in the control or power of the Respondent:"    .Forward = True    .Execute    If .Found = True Then .Parent.Select    End With
End Sub
