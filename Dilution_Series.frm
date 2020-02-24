VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Dilution Series"
   ClientHeight    =   5495
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4165
   OleObjectBlob   =   "Dilution_Series.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdAccept_Click()
        
        'defining variables and setting equal to typed data
        Dim vol1 As Integer
        Dim vol2 As Integer
        Dim vol3 As Integer
        Dim vol4 As Integer
        vol1 = UserForm1.volume1.Value
        vol2 = UserForm1.volume2.Value
        vol3 = UserForm1.volume3.Value
        vol4 = UserForm1.volume4.Value
        
        'volume1: aspiration commands for media in first column
        
        'wells filled per aspiration (a1)
        Dim a1 As Double
        a1 = 1000 / vol1
        a1 = Application.WorksheetFunction.RoundDown(a1, 0)
        
        'number of aspiration commands for column1 (a2)
        Dim a2 As Double
        a2 = 6 / a1
        a2 = Application.WorksheetFunction.RoundUp(a2, 0)
        
        'current row of aspiration command (b0)
        Dim b0 As Integer
        b0 = 1
        
        'filling in aspiration commands to script
        With Sheets("Script")
        For x = 1 To a2
            .Cells(b0, "A") = "A"
            .Cells(b0, "B") = "Media"
            .Cells(b0, "D") = "Falcon 50ml"
            .Cells(b0, "E") = 1
            'required aspiration volume
            If a2 - x = 0 Then
            .Cells(b0, "G") = vol1 * (6 - a1 * (x - 1))
            Else
            .Cells(b0, "G") = vol1 * a1
            End If
            b0 = b0 + a1 + 1
        Next x
        
        'dispense commands for media in first column
        
        'current row of dispense command (b1)
        Dim b1 As Integer
        b1 = 2
        
        'current well name (y1)
        Dim y1 As Integer
        y1 = 2
        
        'current well out of the 6 in the first column (y2)
        Dim y2 As Integer
        y2 = 1
        
        'filling in media dispense commands to script
        For x2 = 1 To a2
            For x1 = 1 To a1
            If y2 <= 6 Then
                .Cells(b1, "A") = "D"
                .Cells(b1, "B") = "96WP"
                .Cells(b1, "D") = "96 Well Microplate"
                .Cells(b1, "E") = y1
                .Cells(b1, "G") = vol1
                'well out of 6
                y2 = y2 + 1
                'well name
                y1 = y1 + 1
                'next row
                b1 = b1 + 1
            End If
            Next x1
        'add 1 to b1 to skip the row with aspiration command
        If x2 <> a2 Then
        b1 = b1 + 1
        End If
        Next x2
        
        'volume3: filling well B2 to G12 with media
        
        'wells filled per aspiration command
        Dim c1 As Double
        c1 = 1000 / vol3
        c1 = Application.WorksheetFunction.RoundDown(c1, 0)
        
        'aspiration commands required to fill wells B2 to G12
        Dim c2 As Double
        c2 = 66 / c1
        c2 = Application.WorksheetFunction.RoundUp(c2, 0)
        
        'next free row in sheet 2
        uRow = .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
        
        'current aspiration command
        Dim d0 As Integer
        d0 = uRow
        
        'filling in aspiration commands for well B2 to G12 into script
        For x3 = 1 To c2
            .Cells(d0, "A") = "A"
            .Cells(d0, "B") = "Media"
            .Cells(d0, "D") = "Falcon 50ml"
            .Cells(d0, "E") = 1
            'required aspiration volume
            If c2 - x3 = 0 Then
            .Cells(d0, "G") = vol3 * (66 - c1 * (x3 - 1))
            Else
            .Cells(d0, "G") = vol3 * c1
            End If
            d0 = d0 + c1 + 1
        Next x3
        
        'dispense commands for filling in media to wells B2 to G12
        
        'current excel sheet row
        Dim e0 As Integer
        e0 = uRow + 1
        
        'current dispense command
        Dim nw As Double
        nw = 0
        
        'defining pi
        Dim pi As Double
        pi = WorksheetFunction.pi()
        
        'dispense commands for wells B2 to G12
        For x4 = 10 To 90 Step 8
            For x5 = x4 To x4 + 5
                .Cells(e0, "A") = "D"
                .Cells(e0, "B") = "96WP"
                .Cells(e0, "D") = "96 Well Microplate"
                .Cells(e0, "E") = x5
                .Cells(e0, "G") = vol3
                e0 = e0 + 1
                nw = nw + 1
                'skip row whenever nw/a1 equals integer
                If Round(Sin((nw * pi) / c1), 2) = 0 Then
                e0 = e0 + 1
                End If
            Next x5
        Next x4
        
        'inserting wash step after every third aspiration in the media dispensing section
        Dim lCount As Long
        
        Set rFound = .Cells(1, "A")
        
        For q1 = 1 To (c2 + a2) / 3
        For lCount = 1 To 3
            Set rFound = .Range("a1:a200").Find("A", after:=rFound, LookIn:=xlValues)
        Next lCount
        rFound.EntireRow.Insert
        rFound.Offset(-1, 0).Value = "W"
        Next q1
        
        'volume2: Culture into column 1 wells
        
        'next free row in sheet 2
        nRow = .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
        Dim b2 As Double
        b2 = nRow
        
        'wash-step to change pipette tip before using culture
        .Cells(b2, "A") = "W"
        
        'excel sheet row for first culture aspiration command
        b2 = b2 + 1
        
        'aspirate from well q2 of 6 Well Plate
        Dim q2 As Integer
        Dim q3 As Integer
        q2 = 1
        q3 = 2
        
        For x6 = 1 To 6
            'aspirate from 6 Well Plate
            .Cells(b2, "A") = "A"
            .Cells(b2, "B") = "6WP"
            .Cells(b2, "D") = "6 Well Microplate"
            .Cells(b2, "E") = q2
            .Cells(b2, "G") = vol4
            
            q2 = q2 + 1
            b2 = b2 + 1
        
            'dispense into first well of corresponding row on 96 Well Plate
            .Cells(b2, "A") = "D"
            .Cells(b2, "B") = "96WP"
            .Cells(b2, "D") = "96 Well Microplate"
            .Cells(b2, "E") = q3
            .Cells(b2, "G") = vol4
            
            q3 = q3 + 1
            b2 = b2 + 1
            
        Next x6
        
        .Cells(b2, "A") = "W"
        
        'aspiration commands for wells B2 to G12
        
        'can we run half of the script in one csv file using 1000ul pipette
        'and the other half using 350ul pipettes?
        'aspirate vol4 from each well along the horizontal axis, with a well inbetween
        
        'next free row in excel sheet
        uRow = .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
        Dim b4 As Double
        b4 = uRow
        
        'filling in aspiration commands for culture transfer B2 to G12
        For y5 = 2 To 7
            For x9 = y5 To y5 + 80 Step 8
                .Cells(b4, "A") = "A"
                .Cells(b4, "B") = "96WP"
                .Cells(b4, "D") = "96 Well Microplate"
                .Cells(b4, "E") = x9
                .Cells(b4, "G") = vol4
                b4 = b4 + 2
                If x9 = y5 + 80 Then
                .Cells(b4, "A") = "W"
                b4 = b4 + 1
                End If
            Next x9
        Next y5
        
        'dispense commands for culture transfer wells B2 to G12
        b5 = uRow + 1
        For y6 = 10 To 15
            For x9 = y6 To y6 + 80 Step 8
                .Cells(b5, "A") = "D"
                .Cells(b5, "B") = "96WP"
                .Cells(b5, "D") = "96 Well Microplate"
                .Cells(b5, "E") = x9
                .Cells(b5, "G") = vol4
                b5 = b5 + 2
                If x9 = y6 + 80 Then
                b5 = b5 + 1
                End If
            Next x9
        Next y6
        
        
        UserForm1.Hide
        Worksheets.Item(2).Activate
        End With
        
        
        
End Sub

Private Sub cmdAdd_Click()

    'deleting content in cells A2 to D10
    Range("A2:L10").Delete
    
    'changing name of sheet one to Dilution Factor Overview
    Sheets.Item(1).Name = "Dilution Factor Overview"
    
    'deleting all sheets except first
    Dim sht As Worksheet
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name <> "Dilution Factor Overview" Then
            sht.Delete
        End If
    Next sht
    
    'define variables
    Dim iRow As Long
    Dim ws As Worksheet
    Set ws = Sheets.Item(1)
    Dim vol1 As Integer
    Dim vol2 As Integer
    Dim vol3 As Integer
    Dim vol4 As Integer
    
    'find first empty row in database
    iRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'set variables equal to typed data
    vol1 = UserForm1.volume1.Value
    vol2 = UserForm1.volume2.Value
    vol3 = UserForm1.volume3.Value
    vol4 = UserForm1.volume4.Value
    
    'copy the data to the sheet
    ws.Cells(iRow, 1).Value = vol1
    ws.Cells(iRow, 2).Value = vol2
    ws.Cells(iRow, 3).Value = vol3
    ws.Cells(iRow, 4).Value = vol4
    
    'calculate dilution factors
    col1 = vol2 / (vol1 + vol2)
    col2 = (vol4 / (vol3 + vol4)) * col1
    col3 = (vol4 / (vol3 + vol4)) ^ 2 * col1
    col4 = (vol4 / (vol3 + vol4)) ^ 3 * col1
    col5 = (vol4 / (vol3 + vol4)) ^ 4 * col1
    col6 = (vol4 / (vol3 + vol4)) ^ 5 * col1
    col7 = (vol4 / (vol3 + vol4)) ^ 6 * col1
    col8 = (vol4 / (vol3 + vol4)) ^ 7 * col1
    col9 = (vol4 / (vol3 + vol4)) ^ 8 * col1
    col10 = (vol4 / (vol3 + vol4)) ^ 9 * col1
    col11 = (vol4 / (vol3 + vol4)) ^ 10 * col1
    col12 = (vol4 / (vol3 + vol4)) ^ 11 * col1
    
    'insert row names
    For colx = 1 To 12
        ws.Cells(iRow + 1, colx).Value = "column" & colx
    Next colx
    Range("A3:L3").Font.Bold = True
    
    'Insert dilution factor next to row name
    ws.Cells(iRow + 2, 1).Value = col1
    ws.Cells(iRow + 2, 2).Value = col2
    ws.Cells(iRow + 2, 3).Value = col3
    ws.Cells(iRow + 2, 4).Value = col4
    ws.Cells(iRow + 2, 5).Value = col5
    ws.Cells(iRow + 2, 6).Value = col6
    ws.Cells(iRow + 2, 7).Value = col7
    ws.Cells(iRow + 2, 8).Value = col8
    ws.Cells(iRow + 2, 9).Value = col9
    ws.Cells(iRow + 2, 10).Value = col10
    ws.Cells(iRow + 2, 11).Value = col11
    ws.Cells(iRow + 2, 12).Value = col12
    
    'create new sheet
    Application.ScreenUpdating = False
    Dim script As Worksheet
    Set script = ActiveWorkbook.Sheets.Add(after:=Sheets.Item("Dilution Factor Overview"))
    script.Name = "Script"
    Worksheets.Item(1).Activate
    Application.ScreenUpdating = True
    
End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub volume1_Change()

End Sub
