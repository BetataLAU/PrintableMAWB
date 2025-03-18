Option Explicit

Public wbOrigin As Workbook
Public wbMAWB As Workbook
Public Const MYC As Double = 2.4  'Apply for 2024/NOV



Private Sub CainiaoGeneratePrintableMAWB()


    'Set opened booking xls into obj var.
    Dim wbName As String
    wbName = ActiveWorkbook.Name
    
    'Public wbOrigin As Workbook
    Set wbOrigin = ActiveWorkbook
    
    Dim wsOrigin As Worksheet
    Set wsOrigin = Worksheets(1)
        
    
    
    'OpenExistingWorkbook
    ' Replace the path below with the actual path to your .xlsx file.
    Dim desktopPath As String
    Dim folderPath As String
    Dim fileNames As String
    
    Dim filePath As String
    
    Dim currentPath As String
    currentPath = wbOrigin.Path
    
    desktopPath = Environ("USERPROFILE") & "\Desktop\"
    folderPath = "CAINIAO - HC 223\"
    fileNames = "223 AGL FORMAT.xlsx"
        
    'filePath = desktopPath & folderPath & fileNames <<testing done, all locations setting.
    filePath = currentPath & "\223 AGL FORMAT.xlsx"
    
    'Open the workbook and set it to the wb object.
    'Public wbMAWB As Workbook
    Set wbMAWB = Workbooks.Open(filePath)
    
    ' Set the author name of the workbook
    wbMAWB.BuiltinDocumentProperties("Author") = "Beta LAU"
    
    Dim wsMAWB As Worksheet
    Set wsMAWB = Worksheets(1)
    'MsgBox wsMAWB.Name
    
    
    'MsgBox "open success"
    
    
    'Getting current cursor's Co-ordinate.
    Dim cellRow As Long
    Dim cellColumn As Long
    
    
    Workbooks(wbName).Activate
    
    cellRow = ActiveCell.row
    cellColumn = ActiveCell.Column

    'Testing
    'MsgBox cellRow & "," & cellColumn
    
    
    
    'Main
    Call ReadMAWB(cellRow, cellColumn)
    Call ReadParty(cellRow, cellColumn)
    Call AssignCarrierNameAddr
    Call IssuingCarrierInfo
    Call FlightInfo(cellRow, cellColumn)
    Call CurrencySection
    Call HandlingInfo(cellRow, cellColumn)
    Call PCSandRate(cellRow, cellColumn)
    Call LoadingInfoBlock(cellRow, cellColumn)
    Call Commodity(cellRow, cellColumn)
    Call OtherCharges(cellRow, cellColumn)
    Call TotalFreightAmount
    Call Signature
    Call GenLoading(cellRow, cellColumn)
    Call SaveFile(cellRow, cellColumn)
    Call SaveAsPDFMinimized
    Call MoveFile  'Move file "MAWBBASE_compressed.pdf"
    Call AddChopToCompressedPDF
    
    'MsgBox "Finished."
    'wbMAWB.Close SaveChanges:=True
    wbMAWB.Close
    
    Set wbOrigin = Nothing
    Set wbMAWB = Nothing
    Set wsMAWB = Nothing
    
    
End Sub
    
    
Sub ReadMAWB(cellRow, cellColumn)

    'Read MAWB#
    Dim MAWBnum As String
    
    MAWBnum = Cells(cellRow, 3)
    
    'Remove "-"
    MAWBnum = Replace(MAWBnum, "-", "")
    
    'Remove "Space"
    MAWBnum = Replace(MAWBnum, " ", "")
        
    'Checking MAWB#.
        'Checking the whole MAWB length
        If Len(MAWBnum) <> 11 Then
            MsgBox "invalid MAWB#, " & MAWBnum & "pls check again."
            Exit Sub
        End If
        
        'Testing.
        'MsgBox MAWBnum
                
        'Get MAWB prefix
        Dim MAWBPrefix As String
        
        MAWBPrefix = Mid(MAWBnum, 1, 3)
           
        'Testing.
        'MsgBox MAWBPrefix
    
    
        'Get MAWB suffix
        Dim MAWBSuffix As String
            
        MAWBSuffix = Mid(MAWBnum, 4, 8)
        
        'Testing.
        'MsgBox "MAWB Suffix is " & MAWBSuffix
        'MsgBox "Suffix 1st 7 number is " & CLng(Mid(MAWBSuffix, 1, 7))
        'MsgBox "Suffix last number is " & Right(CLng(MAWBSuffix), 1)
        'MsgBox "Suffix MOD 7 is " & (CLng(Mid(MAWBSuffix, 1, 7)) Mod 7)
                        
        'Verifty the suffix
        If (CLng(Mid(MAWBSuffix, 1, 7)) Mod 7) <> Right(CLng(MAWBSuffix), 1) Then
            MsgBox "Wrong MAWB suffix, " & MAWBSuffix & " Exit SUB"
            Exit Sub
        End If
    
    'Assign MAWB# to wsMAWB.
        
    
    With wbMAWB.Worksheets("MAWB")
    'Assign Prefix & Suffix.
    .Range("A1") = MAWBPrefix
    .Range("C1") = "HKG"
    .Range("E1") = MAWBSuffix
    
    .Range("AH1") = MAWBPrefix
    .Range("AJ1") = MAWBSuffix
    
    .Range("AH62") = MAWBPrefix
    .Range("AH62") = MAWBPrefix
    End With
    
End Sub


Sub ReadParty(cellRow, cellColumn)

           
    'Get MAWB Shipper & Consignee & Notify.
    Dim Shipper As String
    Dim Consignee As String
    Dim Notify As String
            
    With wbOrigin.Worksheets(1)
        Shipper = .Range("N" & cellRow)
        Consignee = .Range("O" & cellRow)
        Notify = .Range("Q" & cellRow)
    End With
    
    'Testing
    'MsgBox Shipper
    'MsgBox Consignee
    'MsgBox Notify
    
    'Manipulate MAWB Shipper & Consignee & Notify.
    'Skipping 1st two vbCrLF, then the rest of them will be deleted, in order to lessen the # of row.
    'count set to 2 consideration: 1st line CN com name, 2nd line C/O HKG com name.
    'If count set to 1: many case show that everything are too long.
    Dim count As Long
    Dim i As Long
    
    'For Shipper
    count = 0
    
    For i = 1 To Len(Shipper)
        If Mid(Shipper, i, 1) = vbLf Then
            count = count + 1
            If count > 2 Then
                Mid(Shipper, i, 1) = " "
            End If
        End If
    Next i
    
    'For Consignee
    count = 0
    
    For i = 1 To Len(Consignee)
        If Mid(Consignee, i, 1) = vbLf Then
            count = count + 1
            If count > 2 Then
                Mid(Consignee, i, 1) = " "
            End If
        End If
    Next i
    
    'For Notify
    count = 0
    
    For i = 1 To Len(Notify)
        If Mid(Notify, i, 1) = vbLf Then
            count = count + 1
            If count > 2 Then
                Mid(Notify, i, 1) = " "
            End If
        End If
    Next i
    
    
    'Testing, # of lines should be less than previous one.
    
    
'        MsgBox Shipper
'        MsgBox Consignee
'        MsgBox Notify
    
    'Assign Shipper / Consignee / Notify to wbMAWB
    wbMAWB.Worksheets("MAWB").Range("A40") = ""
    
    With wbMAWB.Worksheets("MAWB")
        .Range("A3") = Shipper
        .Range("A9") = Consignee
        
        If Notify = "/" Then
            Notify = ""
        Else
            .Range("A40") = "CNEE Notify Party:" & vbLf & Notify
        End If
        
        'MsgBox "New Notify: " & Notify
        
        .Range("U14") = vbNewLine & "*FREIGHT PREPAID*" & vbNewLine & "**C/O AIR GLOBAL LIMITED." & vbNewLine & "ROOM 503, 5/F, HARBOUR CENTRE, TOWER 2, 8 HOK CHEUNG STREET, HUNGHOM, KOWLOON, HONG KONG."
        
    End With
        
End Sub

    
    
Sub AssignCarrierNameAddr()

    'Determine Airline from Prefix.
    Dim prefix As String
    prefix = wbMAWB.Worksheets("MAWB").Range("A1")
    
    'Carrier selection.
    Select Case prefix
        
        Case "65"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "SAUDI ARABIAN AIRLINES"
        Case "71"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "ETHIOPIAN AIRLINES GROUP" & vbNewLine & "ADDIS ABABA, ETHIOPIA"
        Case "157"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "QATAR AIRWAYS"
        Case "172"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "CARGOLUX AIRLINES INTERNATIONAL" & vbNewLine & "LUXEMBOURG L-2990 LUXEMBOURG"
        Case "180"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "KOREAN AIRLINES CO LTD"
        Case "223"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "ONE AIR LTD" & vbNewLine & "BECKETTS PLACE, 1 HAMPTON WICK KINGSTON-UPON-THAMES, SURREY KT1 4EQ, UNITED KINGDOM"
        Case "317"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "AIR ATLANTA ICELANDIC"
        Case "406"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "UPS AIR CARGO" & vbNewLine & "1400 NORTH HURSTBOURNE PARKWAY LOUISVILLE KY 40223 US"
        Case "485"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "ASTRAL AVIATION"
        Case "501"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "SILK WAY WEST AIRLINES LLC" & vbNewLine & "HEYDAR ALIYEV INTERNATIONAL AIRPORT" & vbNewLine & "BAKU, AZ1044, AZERBAIJAN"
        Case "574"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "AIR ATLANTA ICELANDIC"
        Case "586"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "EUROAVIA AIRLINES"
'        Case "574"
'            wbMAWB.Worksheets("MAWB").Range("Z3") = "ALLIED AIR LIMITED" & vbNewLine & "NAHCO CARGO OFFICE COMPLEX (2ND FLOOR) MURTALA MUHAMMED INT. AIRPORT, LAGOS, NIGERIA"
        Case "756"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "ASL AIRLINES BELGIUM"
        Case "763"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "MAERSK AIR CARGO A/S" & vbNewLine & "LYNGBY HOVEDGADE 85 2800, KOGENS LYNGBY, DENMARK"
        Case "828"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "HONG KONG AIR CARGO CARRIER LTD" & vbNewLine & "UNIT 1210-1218, 12/F, TOWER 2, KOWLOON COMMERCE CENTRE, 51 KWAI CHEONG ROAD, KWAI CHUNG, NEW TERRORIES, HONG KONG"
        Case "933"
            wbMAWB.Worksheets("MAWB").Range("Z3") = "NIPPON CARGO AIRLINES CO LTD"
        
    End Select

End Sub


Sub IssuingCarrierInfo()

    'By default, 2 selections only, AGL / DHL WH.
    
    With wbMAWB.Worksheets("MAWB")
        Select Case .Range("A1")
        
            Case "65"
                'Issuing Carrier.
                .Range("A15") = "WORLDWIDE PARTNER LOGISTICS CO LTD / HKG"
                'IATA code.
                .Range("A19") = "1337344-0006"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "71"
                'Issuing Carrier.
                .Range("A15") = "DSV AIR & SEA LTD / HKG"
                'IATA code.
                .Range("A19") = "13-3 0621/0002"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "157"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "172"
                'Issuing Carrier.
                .Range("A15") = "CARGO-PARTNER LOGISTICS LTD / HKG"
                'IATA code.
                .Range("A19") = "1330815"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "180"
                'Issuing Carrier.
                .Range("A15") = "UPS SCS (ASIA) LTD / HKG"
                'IATA code.
                .Range("A19") = "13-3-0333"
                'IATA A/C code.
                .Range("K19") = ""
            
            Case "223"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "317"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                                
            Case "406"
                'Issuing Carrier.
                .Range("A15") = "UPS SCS (ASIA) LTD / HKG"
                'IATA code.
                .Range("A19") = "13-3-0333/0003"
                'IATA A/C code.
                .Range("K19") = ""
            
            Case "485"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "501"
                'Issuing Carrier.
                .Range("A15") = "DSV AIR & SEA LTD / HKG"
                'IATA code.
                .Range("A19") = "13-3 0621/0002"
                'IATA A/C code.
                .Range("K19") = ""
            
            Case "574"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "586"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
'            Case "574"
'                'Issuing Carrier.
'                .Range("A15") = "CARGO-PARTNER LOGISTICS LTD / HKG"
'                'IATA code.
'                .Range("A19") = "13-3 0815/0004"
'                'IATA A/C code.
'                .Range("K19") = "0800368HKG"
                
            Case "756"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "763"
                'Issuing Carrier.
                .Range("A15") = "AIR GLOBAL LIMITED / HKG"
                'IATA code.
                .Range("A19") = "13-3-0182 0000"
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "828"
                'Issuing Carrier.
                .Range("A15") = "DONGNAM WAREHOUSE LIMITED / HKG"
                'IATA code.
                .Range("A19") = ""
                'IATA A/C code.
                .Range("K19") = ""
                
            Case "933"
                'Issuing Carrier.
                .Range("A15") = "UPS SCS (ASIA) LTD / HKG"
                'IATA code.
                .Range("A19") = "13-3-0333"
                'IATA A/C code.
                .Range("K19") = ""
                
        End Select
    End With
    
End Sub


Sub FlightInfo(cellRow, cellColumn)
    
    With wbMAWB.Worksheets("MAWB")
    
        'Aiport of Origin.
        .Range("A21") = "HONG KONG"
        
        'Dest code, 3 letters.
        .Range("A23") = wbOrigin.Worksheets(1).Range("F" & cellRow)
        
        'Carrier code, 2 letters.
        .Range("D23") = Mid(wbOrigin.Worksheets(1).Range("G" & cellRow), 1, 2)
        
        'Dest in full name.
        Dim destFullName As String
        
        Select Case .Range("A23")
            Case "BHX"
                destFullName = "BIRMINGHAM"
            Case "BUD"
                destFullName = "BUDAPEST"
            Case "EMA"
                destFullName = "NOTTINGHAM"
            Case "JFK"
                destFullName = "NEW YORK"
            Case "LAX"
                destFullName = "LOS ANGELES"
            Case "LGG"
                destFullName = "LIEGE"
            Case "LHR"
                destFullName = "LONDON"
            Case "MAD"
                destFullName = "MADRID"
            Case "ORD"
                destFullName = "CHICAGO"
            Case "STN"
                destFullName = "STANSTED"
            Case "TLV"
                destFullName = "TEL AVIV YAFO"
        End Select
        
        .Range("A25") = destFullName
        
        'Flight Number.
        .Range("J25") = wbOrigin.Worksheets(1).Range("G" & cellRow)
        
        'Flight Date (Converting format to dd/Mmm/yyyy).
        Dim flightDate As Date
        
        flightDate = wbOrigin.Worksheets(1).Range("H" & cellRow) 'String to date convertion.
        
        .Range("P25") = Format(flightDate, "dd/Mmm/yyyy")
        
    End With

End Sub


Sub CurrencySection()
    
    With wbMAWB.Worksheets("MAWB")
    
        'Currency
        .Range("T23") = "HKD"
        
        'Wt PPD
        .Range("X23") = "PP"
        
        'Other PPD
        .Range("Z23") = "PP"
        
        'N.V.D.
        .Range("AC23") = "N.V.D."
    
        'AS PER INVOICE.
        .Range("AI23") = "AS PER INV."
        
        'Insurance Amount
        .Range("U25") = "NIL"
    
    End With
    
End Sub


Sub HandlingInfo(cellRow, cellColumn)

    'Get # of PKGS.
    Dim pkgs As String
    pkgs = CStr(wbOrigin.Worksheets(1).Range("I" & cellRow))
    
    'Gen PKGS statement.
    Dim pkgsSteatment As String
    pkgsSteatment = "TOTAL: (" & pkgs & ") PACKAGES ONLY."
    
    'Get special instructions.
    Dim specialInstructions As String
    
    specialInstructions = wbOrigin.Worksheets(1).Range("P" & cellRow)

    If InStr(specialInstructions, "BUP") > 0 Then  '0 means not found, >1 means is found.
        specialInstructions = "BUP DO NOT BREAKDOWN."
    Else
        specialInstructions = ""
    End If
        
    'Others info.
    Dim othInfo As String
    othInfo = "NO S.W.P.M. "
    
    Dim othInfo2 As String
    othInfo2 = ""
    If InStr(wbOrigin.Worksheets(1).Range("O" & cellRow), "FTL-SERVICE") > 0 Then
        othInfo2 = " -CONSIGNEE WILL PICK UP CARGO AT DEST."
    End If
    
    Dim prefix As String
    prefix = wbMAWB.Worksheets("MAWB").Range("A1")

    If prefix = "71" Then
        othInfo = "NO S.W.P.M. "
        othInfo2 = ""
    End If
    
    If prefix = "501" Then
        othInfo = "NO S.W.P.M. "
        othInfo2 = ""
    End If
    
    'Gen s whole statement.
    wbMAWB.Worksheets("MAWB").Range("A27") = pkgsSteatment & vbNewLine & specialInstructions & vbNewLine & othInfo & othInfo2
    
    
    'Some special case.
'    If (wbOrigin.Worksheets(1).Range("F" & cellRow) = "LGG") And InStr(wbOrigin.Worksheets(1).Range("O" & cellRow), "A PLUS") Then
'        wbMAWB.Worksheets("MAWB").Range("A27") = pkgsSteatment & " " & specialInstructions & vbNewLine & othInfo & _
'            "-CNEE Custom office code: GB000081 (UK North Auth Consignor/nees OR Leeds North Auth Consignor/nees)"
'    End If
    
    'Some special case.
    If (wbOrigin.Worksheets(1).Range("F" & cellRow) = "LGG") And InStr(wbOrigin.Worksheets(1).Range("O" & cellRow), "SOLUTION UNLIMITED SRL") Then
        wbMAWB.Worksheets("MAWB").Range("A27") = pkgsSteatment & " " & specialInstructions & vbNewLine & othInfo & _
            "-CNEE Email: sirui.wang@wahoo-freight.com; Info@wallitrans.com; sirui.wang@solution-unlimited.eu; z276762641@163.com"
    End If
    
    'Some special case.
    If (wbOrigin.Worksheets(1).Range("F" & cellRow) = "LGG") And InStr(wbOrigin.Worksheets(1).Range("O" & cellRow), "BREAK&BUILD SRL") Then
        wbMAWB.Worksheets("MAWB").Range("A27") = pkgsSteatment & " " & specialInstructions & vbNewLine & othInfo & _
            "-CNEE TEL:+32 494348802"
    End If
    
        
End Sub


Sub PCSandRate(cellRow, cellColumn)

    
    With wbMAWB.Worksheets("MAWB")
        
        Dim specialInstructions As String
        specialInstructions = wbOrigin.Worksheets(1).Range("P" & cellRow)
    
        If InStr(specialInstructions, "BUP") > 0 Then  '0 means not found, >1 means is found.
            .Range("A32") = 1
        Else
            .Range("A32") = wbOrigin.Worksheets(1).Range("I" & cellRow)
        End If
        
        'Assign est G.W. onto MAWB.
        .Range("C32") = "e" & wbOrigin.Worksheets(1).Range("K" & cellRow)
        
        
        'Assign est C.W. onto MAWB.
        'Since the stiuation should not very complex, it will base onto G.W.
        .Range("M32") = "=C32"
        
        '****Test area.
        
        'Below 1 line is original statement.
        '.Range("M32") = "=C32"
        

        
        '****End of test area
        
        .Range("H32") = "Q"
    
        'Determine the TACT rate, no consideration abt the wt break since it is always in ULD.
        Dim Rate As Double  'The Double data type occupies 8 bytes (64 bits) of memory, which is a reasonable amount of memory usage
        Dim dest As String
        
        Select Case .Range("A23")
            Case "BHX"
                Rate = 67.03
            Case "BUD"
                Rate = 36#
            Case "EMA"
                Rate = 67.03
            Case "JFK"
                Rate = 48.12
            Case "LAX"
                Rate = 41.9
            Case "LGG"
                Rate = 36#
            Case "LHR"
                Rate = 66.74
            Case "MAD"
                Rate = 36#
            Case "ORD"
                Rate = 47.06
            Case "STN"
                Rate = 66.74
            Case "TLV"
                Rate = 49.98
            Case Else
                MsgBox "No record for port: " & .Range("A23").value & ", pls input manually."
        End Select
        
        .Range("Q32") = Rate
        
        If .Range("A15") = "CARGO-PARTNER LOGISTICS LTD / HKG" And .Range("A1") = "574" Then
            .Range("Q32") = ""
        End If
        
        'Assign formula for freightage.
        .Range("V32") = "=M32*Q32"
    
        
        '***Special Cases.***
        
        '*******
        '* 485 *
        '*******
        If .Range("A1") = "485" Then
            .Range("Q32") = ""
            .Range("V32") = "AS ARRANGED"
        End If
    
    
    End With
    
        
End Sub


Sub LoadingInfoBlock(cellRow, cellColumn)

    With wbMAWB.Worksheets("MAWB")
        
        'Add the T/S statement based on Dest.
        .Range("A34") = "T/S CARGO FM CHINA TO " & .Range("A23") & " VIA HKG BY TRUCK."
    
        'Add SLAC pieces in the next line.
        Dim pkgs As String
        pkgs = CStr(wbOrigin.Worksheets(1).Range("I" & cellRow))
        
        Dim otherWording As String
        If InStr(.Range("A27"), "BUP") > 0 Then  '0 means not found, >1 means is found.
            otherWording = "(BUP)"
        End If

        .Range("A35") = "(" & pkgs & ")" & otherWording
        
        'Add the loading heading sentence.
        .Range("A36") = "LOADING INFO:"
        
    End With

End Sub


Sub Commodity(cellRow, cellColumn)

    Dim srcCommodity As String
    Dim actualCommodity As String
    Dim commodities() As String
    Dim i As Long

    ' Sample input string
    srcCommodity = wbOrigin.Worksheets(1).Range("S" & cellRow)

    ' Split the input string into individual commodities
    commodities = Split(srcCommodity, vbLf)

    ' Initialize variables to store the extracted commodities
    Dim commodityStrings() As String
    ReDim commodityStrings(0 To UBound(commodities))

    ' Loop through the commodities and extract the commodity names using regular expressions
    For i = 0 To UBound(commodities)
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "^\s*(\w+(?:\s\w+)*)"
        Dim match As Object
        Set match = regex.Execute(commodities(i))
        If match.count > 0 Then
            commodityStrings(i) = UCase(match.item(0).value)
        End If
    Next i

    ' Display the extracted commodity strings
    For i = 0 To UBound(commodities)
        If Len(commodityStrings(i)) > 0 Then
            'Debug.Print "Commodity: " & commodityStrings(i)
            
            Select Case commodityStrings(i)

                Case "CONSOLIDATION SHIPMENT AS PER ATTACHED CARGO MANIFEST"
                    actualCommodity = actualCommodity & "CONSOL SHPT" & vbNewLine
                Case "CLOTHES"
                    actualCommodity = actualCommodity & "CLOTHES(T-SHIRT/DRESS)-HS:610332" & vbNewLine
                Case "CROCHET"
                    actualCommodity = actualCommodity & "CROCHET -HS:392690" & vbNewLine
                Case "SHOE"
                    actualCommodity = actualCommodity & "SHOE -HS:640291" & vbNewLine
                Case "HAIRPIN"
                    actualCommodity = actualCommodity & "HAIRPIN -HS:961519" & vbNewLine
                Case "KEY CHAIN"
                    actualCommodity = actualCommodity & "KEY CHAIN -HS:830810" & vbNewLine
                Case "LAUNDRY DETERGENT"
                    actualCommodity = actualCommodity & "LAUNDRY DETERGENT (SUPER OIL DETERGENT) -HS:340250" & vbNewLine & "-NOT RESTRICTED" & vbNewLine
                Case "CAMERA"
                    actualCommodity = actualCommodity & "CAMERA -HS:852589" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "HEADLAMP"
                    actualCommodity = actualCommodity & "HEADLAMP -HS:851220" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "LAPTOP"
                    actualCommodity = actualCommodity & "LAPTOP -HS:847130" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "LED LIGHT"
                    actualCommodity = actualCommodity & "LED LIGHT -HS:940542" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "MINI PC"
                    actualCommodity = actualCommodity & "MINI PC -HS:847130" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "MASSAGER"
                    actualCommodity = actualCommodity & "MASSAGER -HS:901910" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "NETWORK ATTACHED STORAGE"
                    actualCommodity = actualCommodity & "NETWORK ATTACHED STORAGE -HS:847330" & vbNewLine & _
                    "LITHIUM METAL BATTERIES IN COMPLIANCE WITH SECTION II OF PI970 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "SMARTWATCH"
                    actualCommodity = actualCommodity & "SMARTWATCH -HS:851762" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "TABLE LAMP"
                    actualCommodity = actualCommodity & "TABLE LAMP -HS:940511" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
                Case "TABLET PC"
                    actualCommodity = actualCommodity & "TABLET PC -HS:847130" & vbNewLine & _
                    "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967 " & _
                    "-" & CStr(wbOrigin.Worksheets(1).Range("I" & cellRow)) & " PKGS"
            End Select
        End If
    Next i
    
    'Final output with cbm value.
    wbMAWB.Worksheets("MAWB").Range("AD32") = actualCommodity & vbNewLine & vbNewLine & _
        wbOrigin.Worksheets(1).Range("L" & cellRow) & " CBM"
            
    Dim prefix As String
    prefix = wbMAWB.Worksheets("MAWB").Range("A1")
    
    If prefix = "71" Then
        wbMAWB.Worksheets("MAWB").Range("AD32") = _
            "CLOTHES HS620331" & vbNewLine & _
            "SHOES HS640359" & vbNewLine & _
            "-NO BATTERY" & vbNewLine & _
            "" & vbNewLine & _
            "CCTV CAMERA HS852589" & vbNewLine & _
            "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI966"
    End If
        
    If prefix = "485" Then
        wbMAWB.Worksheets("MAWB").Range("AD32") = srcCommodity & vbNewLine & vbNewLine & wbOrigin.Worksheets(1).Range("L" & cellRow) & " cbm"
    End If
        
    If prefix = "501" Then
        wbMAWB.Worksheets("MAWB").Range("AD32") = _
            "DRESS HS392620" & vbNewLine & _
            "COAT HS420310" & vbNewLine & _
            "-NO BATTERY" & vbNewLine & _
            "" & vbNewLine & _
            "LED LIGHT HS940521" & vbNewLine & _
            "LITHIUM ION BATTERIES IN COMPLIANCE WITH SECTION II OF PI967"
    End If
        
    Set regex = Nothing
    Set match = Nothing

End Sub


Sub OtherCharges(cellRow, cellColumn)
    
    'Set TC value bases on B, X, P.
    ' Open the workbook as read-only
    Dim TC As Double
    
    Dim currentPath As String
    currentPath = wbOrigin.Path
    
    ' Open the workbook as read-only
    Dim wbLoadingData As Workbook
    Set wbLoadingData = Workbooks.Open(currentPath & "\HC HIN LISTING.xlsx", ReadOnly:=True)
    
    ' Set the value to search for
    Dim searchValue As String
    searchValue = wbOrigin.Worksheets(1).Range("C" & cellRow)
    searchValue = Mid(searchValue, 1, 8) & " " & Mid(searchValue, 9, 4) 'Since dest target ve number formatting like this.

    ' Set the range (column) to search
    Dim rng As Range
    Dim foundCell As Range
    Set rng = wbLoadingData.Worksheets(1).Range("A:A") 'Change "A:A" to your column
    
    'testing
'    MsgBox TypeName(wbLoadingData.Worksheets(1).Range("A:A").value)

    ' Use the Find method to search for the value
    Set foundCell = rng.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlPart)

    ' Check if the value was found
'    If Not foundCell Is Nothing Then
'        MsgBox "Value found at " & foundCell.row
'    Else
'        MsgBox "Value not found"
'    End If
    
    'Save loading / contour / pkgs / wt value.
    Dim preDeclType As String
    preDeclType = wbLoadingData.Worksheets(1).Range("F" & foundCell.row)
    
    
    Select Case preDeclType
        Case Is = "P"
            TC = 1.48
        Case Is = "X"
            TC = 1.48
        Case Is = "B"
            TC = 1.68
    End Select
    
    Dim prefix As String
    prefix = wbMAWB.Worksheets("MAWB").Range("A1")
    
    If prefix = "71" Then
        Select Case preDeclType
        Case Is = "P"
            TC = 1.5
        Case Is = "X"
            TC = 1.5
        Case Is = "B"
            TC = 1.7
        End Select
    End If
    
    With wbMAWB.Worksheets("MAWB")
    
        Select Case .Range("A1")
        
        
            Case "65"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'TC
                .Range("AB47") = "TCC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                            
                'MYC
                .Range("Q48") = "MYC"
                .Range("U48") = "=ROUNDUP(M32*" & MYC & ",1)"
                
                'MSC
                .Range("AA48") = "MSC"
                .Range("AF48") = "=ROUNDUP(M32*" & 0.13 & ",1)"
                
                'SSC statement.
                .Range("P51") = "WEIGHT CHARGE INCLUDES A SECURITY CHARGES OF HKD 1.80 PER KG"
                
            Case "71"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'TC
                .Range("Q47") = "TCC:"
                .Range("U47") = "=ROUNDUP(C32*" & TC & ",1)"
                            
                'AW
                .Range("Q48") = "AWC"
                .Range("U48") = 78#
                
                'CG
                .Range("Q49") = "CGC"
                .Range("U49") = 78#
                
            Case "157"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "AWC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "FEC:"
                .Range("AF47") = "=ROUNDUP(C32*" & TC & ",1)"
                
                'CG
                .Range("Q48") = "CGC:"
                .Range("U48") = 18
                
                'MYC
                .Range("AA48") = "MYC:"
                .Range("AF48") = "=ROUNDUP(M32*2.8,1)"
                
                'RIC
                .Range("Q49") = "RAC:"
                .Range("U49") = "500"
                
                'SSC
                .Range("AB49") = "XBC:"
                .Range("AF49") = "=ROUNDUP(M32*2.5,1)"
                
            Case "172"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'TC
                .Range("Q47") = "TC:"
                .Range("U47") = "=ROUNDUP(C32*1.48,1)"
                
                'ADC
                .Range("AB47") = "ADC:"
                .Range("AF47") = 13#
                
                'MYC
                .Range("Q48") = "MYC:"
                .Range("U48") = "=ROUNDUP(M32*" & MYC & ",1)"
                
                'CG
                .Range("AA48") = "CG:"
                .Range("AF48") = 15.6
                
                'MW
                .Range("Q49") = "MW:"
                .Range("U49") = "=ROUNDUP(M32*2.25,1)"
                
                'XD
                .Range("AB49") = "XD:"
                .Range("AF49") = "=ROUNDUP(M32*.8,1)"
                
            Case "180"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
                'CG
                .Range("Q48") = "CG:"
                .Range("U48") = 5
                
                'MYC
                .Range("AA48") = "MYC:"
                .Range("AF48") = "=ROUNDUP(M32*" & MYC & ",1)"
                
                'SSC statement
                .Range("AA48") = "WEIGHT CHARGE INCLUDES A SECURITY CHARGE OF HKD 2.5/K"
                
            Case "223"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
            Case "317"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""

                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#

                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
            Case "406"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
                'CG
                .Range("Q48") = "CG:"
                .Range("U48") = 5#
                
            
            Case "485"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'TC
                .Range("Q47") = ""
                .Range("U47") = "AS ARRANGED"
                

            Case "501"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'TC
                .Range("Q47") = "TC:"
                .Range("U47") = "=ROUNDUP(C32*" & TC & ",1)"
                
                'ADC
                .Range("AB47") = "AWC:"
                .Range("AF47") = 13#
                
                'MYC
                .Range("Q48") = "MYC:"
                .Range("U48") = "=ROUNDUP(M32*4.2,1)"
                
                'CG
                .Range("AA48") = "SC:"
                .Range("AF48") = "=ROUNDUP(M32*2.5,1)"
                
            Case "574"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""

                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#

                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*" & TC & ",1)"
                
            Case "586"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""

                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#

                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*" & TC & ",1)"
'            Case "574"
'                'Clear contents.
'                .Range("Q47:AF50").value = ""
'                .Range("P51").value = ""
'
'                'ADC
'                .Range("Q47") = ""
'                .Range("U47") = "AS AGREED"
                                
            Case "756"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
                'MYC
                .Range("Q48") = "MYC:"
                .Range("U48") = "=ROUNDUP(M32*" & MYC & ",1)"
            
            Case "763"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
            Case "828"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "AWC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TCC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                            
                'MYC
                .Range("Q48") = "MYC"
                .Range("U48") = "=ROUNDUP(M32*" & MYC & ",1)"
                
                'SSC statement.
                .Range("P51") = "WEIGHT CHARGE INCLUDES A SECURITY CHARGES OF HKD 3.0 PER KG."
                
            Case "933"
                'Clear contents.
                .Range("Q47:AF50").value = ""
                .Range("P51").value = ""
                
                'ADC
                .Range("Q47") = "ADC:"
                .Range("U47") = 13#
                
                'TC
                .Range("AB47") = "TC:"
                .Range("AF47") = "=ROUNDUP(C32*1.48,1)"
                
                'CG
                .Range("Q48") = "CG:"
                .Range("U48") = 5
                
                'MYC
                .Range("AA48") = "MYC:"
                .Range("AF48") = "=ROUNDUP(M32*" & MYC & ",1)"
                
                'SSC statement
                .Range("P51") = "WEIGHT CHARGE INCLUDES A SECURITY CHARGE OF HKD 2.0/K"
                
                'KZ MAWB statement
                .Range("A40") = "THIS IS A SHIPPER'S/AGENTS LOADED&COUNTED CONSIGNMENT THE SHPR/AGENT RELEASE THE CARRIERS RESPONSIBILITY & LIABILITY OF ANY DAMAGE & CORRECTNESS  OF TTL NO OF PIECE CONTAINED EXCEPT RESULTING FM THE  WILLFUL MISCONDUCT/GROSS NEGIGENCE OF NCA."
                
        End Select
    
    End With
    
    Set wbLoadingData = Nothing
    
    
End Sub

Sub TotalFreightAmount()

    With wbMAWB.Worksheets("MAWB")

        'Freightage.
        .Range("A47") = "=V32"
        
        'Total Other Charges.
        .Range("A55") = "=SUM(U47:U50) + SUM(AF47:AF50)"
        
        'Total Prepaid.
        .Range("A59") = "=A47+A55"
        
        
        '**SPECIAL CASE. "AS ARRANGED"
        If CStr(.Range("A47")) = "AS ARRANGED" Then
            
            .Range("A59") = "AS ARRANGED"
        
        End If
        
        'Special case for CPL.
        If .Range("A15") = "CARGO-PARTNER LOGISTICS LTD / HKG" And .Range("A1") = "574" Then
        
            .Range("V32") = "AS AGREED"
            .Range("A47") = "AS AGREED"
            .Range("A55") = "AS AGREED"
            .Range("A59") = "AS AGREED"
        
        End If
        
    End With

End Sub


Sub Signature()
    
    With wbMAWB.Worksheets("MAWB")
        
        Dim issuingCarrier As String
        Dim RAcode As String
        
        issuingCarrier = .Range("A15")
        
        Select Case issuingCarrier
            Case "AIR GLOBAL LIMITED / HKG"
                RAcode = "RA20371"
                'Signature.
                .Range("P55") = "SPX"
                .Range("P56") = "RICH-SALE INTERNATIONAL TRANSPORTATION CO LTD."
            Case "CARGO-PARTNER LOGISTICS LTD / HKG"
                RAcode = "RA11370"
                'Signature.
                .Range("P55") = "SPX"
                .Range("P56") = "RICH-SALE INTERNATIONAL TRANSPORTATION CO LTD."
            Case "DONGNAM WAREHOUSE LIMITED / HKG"
                RAcode = "RA32805"
                'Signature.
                .Range("P55") = "SPX"
                .Range("P56") = "RICH-SALE INTERNATIONAL TRANSPORTATION CO LTD."
            Case "DSV AIR & SEA LTD / HKG"
                RAcode = "RA03854"
                'Signature.
                .Range("P55") = "SPX"
                .Range("P56") = "AIR GLOBAL LIMITED."
            Case "WORLDWIDE PARTNER LOGISTICS CO LTD / HKG"
                RAcode = "RA28490"
                'Signature.
                .Range("P55") = "SPX"
                .Range("P56") = "RICH-SALE INTERNATIONAL TRANSPORTATION CO LTD."
            Case "UPS SCS (ASIA) LTD / HKG"
                RAcode = "RA01325"
                'Signature.
                .Range("P55") = "SPX"
                .Range("P56") = "AIR GLOBAL LIMITED"
            Case Else
                MsgBox "No RA code found, pls input manually."
        End Select
        
       
        'MAWB issue date.
        .Range("Q60") = "=today()"
        
        'Place.
        .Range("X60") = "HONG KONG"
        
        'RA code.
        .Range("AI60") = RAcode
                
    End With

End Sub


Sub GenLoading(cellRow, cellColumn)
    
    Dim currentPath As String
    currentPath = wbOrigin.Path
    
    ' Open the workbook as read-only
    Dim wbLoadingData As Workbook
    Set wbLoadingData = Workbooks.Open(currentPath & "\HC HIN LISTING.xlsx", ReadOnly:=True)
    
    ' Set the value to search for
    Dim searchValue As String
    searchValue = wbOrigin.Worksheets(1).Range("C" & cellRow)
    searchValue = Mid(searchValue, 1, 8) & " " & Mid(searchValue, 9, 4) 'Since dest target ve number formatting like this.

    ' Set the range (column) to search
    Dim rng As Range
    Dim foundCell As Range
    Set rng = wbLoadingData.Worksheets(1).Range("A:A") 'Change "A:A" to your column
    
    'testing
'    MsgBox TypeName(wbLoadingData.Worksheets(1).Range("A:A").value)

    ' Use the Find method to search for the value
    Set foundCell = rng.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlPart)

    ' Check if the value was found
'    If Not foundCell Is Nothing Then
'        MsgBox "Value found at " & foundCell.row
'    Else
'        MsgBox "Value not found"
'    End If
    
    'Save loading / contour / pkgs / wt value.
    Dim ULD As String
    Dim contour As String
    Dim pkgs As String
    Dim wt As String
    Dim preDeclType As String
    
    If wbLoadingData.Worksheets(1).Range("C" & foundCell.row) = wbOrigin.Worksheets(1).Range("I" & cellRow) Then
        wt = wbLoadingData.Worksheets(1).Range("D" & foundCell.row)
    Else
        MsgBox "Diff pcs between 2 tabels, pls close & check."
        Exit Sub
    End If
    
    ULD = wbLoadingData.Worksheets(1).Range("G" & foundCell.row)
    contour = wbLoadingData.Worksheets(1).Range("H" & foundCell.row)
    pkgs = wbLoadingData.Worksheets(1).Range("C" & foundCell.row)
    wt = wbLoadingData.Worksheets(1).Range("D" & foundCell.row)
    preDeclType = wbLoadingData.Worksheets(1).Range("F" & foundCell.row)
    
    'Check ULD exist?
    If ULD = "" Then
        Exit Sub
    End If
    
    'Gen completed loading sentence.
    Dim completeLoading As String
    
    Select Case preDeclType
        Case Is = "P"
            completeLoading = pkgs & " PKGS LDD ON " & ULD & " / " & contour & " (PRE-PACK)"
        Case Is = "X"
            completeLoading = pkgs & " PKGS LDD ON " & ULD & " / " & contour & " (MIX-LOAD)"
        Case Is = "B"
            wbMAWB.Worksheets(1).Range("A36").value = ""
            completeLoading = ""
    End Select
    
    
    
    
'    If InStr(wbMAWB.Worksheets(1).Range("A27"), "BUP") = 0 Then  '0 means not found, >1 means is found.
'        completeLoading = pkgs & " PKGS LDD ON " & ULD & " / " & contour & " (MIX-LOAD)"
'    Else
'        completeLoading = pkgs & " PKGS LDD ON " & ULD & " / " & contour & " (PRE-PACK)"
'        'MsgBox completeLoading
'    End If
    

    
    'Assign WT & LOADING to wbMAWB.
    wbMAWB.Worksheets(1).Range("C32").value = wt
    wbMAWB.Worksheets(1).Range("A37").value = completeLoading
    
    wbLoadingData.Close
    
    Set rng = Nothing
    Set foundCell = Nothing
    Set wbLoadingData = Nothing

End Sub



Sub SaveFile(cellRow, cellColumn)

    Dim currentPath As String
    Dim newFolderPath As String
    

    ' Get the current path of the workbook
    currentPath = wbOrigin.Path

    ' Specify the name of the new folder
    newFolderPath = currentPath & "\MAWB xls files"

    ' Create the new folder
    On Error Resume Next ' Ignore the error if the folder already exists
    MkDir newFolderPath
    On Error GoTo 0 ' Turn off error handling

    ' Specify the name of the new workbook
    Dim newWorkbookName As String
    Dim newFileName As String
    
    'Stroe a MAWB# xxx-xxxxxxxx as new file name.
    newFileName = wbOrigin.Worksheets(1).Range("C" & cellRow) & " MAWB"
    
    newWorkbookName = newFolderPath & "\" & newFileName & ".xlsx"

    ' Save the active workbook to the new folder with the .xlsx format
    wbMAWB.SaveAs fileName:=newWorkbookName, FileFormat:=xlOpenXMLWorkbook

End Sub


Sub SaveAsPDFMinimized()
   
   Dim pdfPath As String
   
    ' Define the path to save the PDF
    pdfPath = wbMAWB.Path & "\TEMP" & Left(wbMAWB.Name, 12) & ".pdf"
    
    ' Save the worksheet as PDF with minimized size
    wbMAWB.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, _
                                             fileName:=pdfPath, _
                                             Quality:=xlQualityMinimum, _
                                             IncludeDocProperties:=True, _
                                             IgnorePrintAreas:=False, _
                                             OpenAfterPublish:=False
End Sub

Sub MoveFile()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define the source file and destination path
    Dim sourceFile As String
    Dim destinationFile As String
    
    sourceFile = "C:\Users\user\Desktop\CAINIAO - HC 223\MSDS\TESTING\MAWBBASE_compressed.pdf"
    destinationFile = wbMAWB.Path & "\MAWBBASE_compressed.pdf"
    
    If Not fso.fileExists(destinationFile) Then
        ' Move the file if it doesn't exist in the destination
        fso.CopyFile sourceFile, destinationFile
        'MsgBox "File moved successfully!"
    Else
        'MsgBox "File already exists in the destination folder. Skipping move."
    End If
    
    Set fso = Nothing
    
End Sub

'Sub AddChopToCompressedPDF()
'
'    Dim command As String
'
''    'Construct the command.
''    command = "pdftk.exe " & _
''              """\MAWB xls files\MAWBBASE_compressed.pdf"" stamp " & _
''              """\MAWB xls files\TEMP" & Left(wbMAWB.Name, 12) & ".pdf"" output " & _
''              """" & wbMAWB.Path & "\" & Left(wbMAWB.Name, 12) & " MAWB.pdf"""
'
'    'Construct the command.
'    command = "pdftk.exe " & _
'              "MAWBBASE_compressed.pdf stamp " & _
'              "TEMP" & Left(wbMAWB.Name, 12) & ".pdf output " & _
'              """" & Left(wbMAWB.Name, 12) & " MAWB.pdf"""
'
'
'    'Generate a batch file.
'    Dim filePath As String
'    Dim fileNumber As Integer
'    Dim batchCommands As String
'
'    ' Define the path where the batch file will be saved
'    filePath = wbMAWB.Path & "\run_pdftk.bat"
'
'    ' Define the commands to be written to the batch file    "@echo off" & "cd ""MAWB xls files""" & _  vbCrLf & _
'
'    batchCommands = command & _
'                    vbCrLf & _
'                    "del """ & wbMAWB.Path & "\TEMP" & Left(wbMAWB.Name, 12) & ".pdf""" & _
'                    vbCrLf & _
'                    "pause"
'
'
'    ' Get a free file number
'    fileNumber = FreeFile
'
'    ' Open the file for output
'    Open filePath For Output As #fileNumber
'
'    ' Write the commands to the file
'    Print #fileNumber, batchCommands
'
'    ' Close the file
'    Close #fileNumber
'
'    'MsgBox "Batch file created successfully at " & filePath
'
'
''    Dim wsh As Object
''    Set wsh = CreateObject("WScript.Shell")
''    Dim CMDcommand As String
''    command = "cmd.exe /c " "C:\Path\To\Your\BatchFile\example.bat"""
''    wsh.Run CMDcommand, 1, True
'
'
'    'command = "cmd.exe /c """ & wbMAWB.Path & "\run_pdftk.bat"""
'    command = wbMAWB.Path & "\run_pdftk.bat"
''    shell command, vbHide
'    shell command
'
'End Sub


Sub AddChopToCompressedPDF()

    Dim command As String
    Dim pdftkPath As String
    
    ' Set the path to PDFtk executable
    pdftkPath = "pdftk.exe"
    
    'Construct the command.
'    command = "pdftk.exe " & _
'              """" & wbMAWB.Path & "\MAWBBASE_compressed.pdf"" stamp " & _
'              """" & wbMAWB.Path & "\TEMP" & Left(wbMAWB.Name, 12) & ".pdf"" output " & _
'              """" & wbMAWB.Path & "\" & Left(wbMAWB.Name, 12) & "+001.pdf"""
              
    command = Chr(34) & pdftkPath & Chr(34) & " " & _
              Chr(34) & wbMAWB.Path & "\MAWBBASE_compressed.pdf" & Chr(34) & " stamp " & _
              Chr(34) & wbMAWB.Path & "\TEMP" & Left(wbMAWB.Name, 12) & ".pdf" & Chr(34) & " output " & _
              Chr(34) & wbMAWB.Path & "\" & Left(wbMAWB.Name, 12) & "+001.pdf" & Chr(34)
              
    'Generate a batch file.
    Dim filePath As String
    Dim fileNumber As Integer
    Dim batchCommands As String
    
    ' Define the path where the batch file will be saved
    filePath = wbMAWB.Path & "\run_pdftk.bat"
              
                    
    ' Define the commands to be written to the batch file
    batchCommands = "@echo off" & vbCrLf & _
                    "cd " & Chr(34) & wbMAWB.Path & Chr(34) & vbCrLf & _
                    command & vbCrLf & _
                    "del " & Chr(34) & wbMAWB.Path & "\TEMP" & Left(wbMAWB.Name, 12) & ".pdf" & Chr(34) & vbCrLf & _
                    "exit"
                    
    ' Get a free file number
    fileNumber = FreeFile

    ' Open the file for output
    Open filePath For Output As #fileNumber

    ' Write the commands to the file
    Print #fileNumber, batchCommands

    ' Close the file
    Close #fileNumber

    'MsgBox "Batch file created successfully at " & filePath
    
    'Run the batch file
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & filePath & """", 1, True
    
End Sub

