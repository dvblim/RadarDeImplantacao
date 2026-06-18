Attribute VB_Name = "Módulo4"
Sub GerarDashboardProjetos_Paginado()
    Dim pptApp As Object, pptPres As Object, pptSlide As Object, pptTable As Object
    Dim ws As Worksheet: Dim lo As ListObject
    Dim i As Long, j As Integer, rowIdx As Long
    Dim totalEquipe As Double, totalPendencias As Double, globalPercent As Double
    Dim currentEquipe As Double, currentPend As Double, vDataLanc As Variant
    
    ' --- CONFIGURAÇÃO DE PAGINAÇÃO ---
    Dim rowsPerSlide As Integer: rowsPerSlide = 10
    Dim totalProjetos As Long, numSlides As Integer, s As Integer
    Dim startRow As Long, endRow As Long, rowsOnThisSlide As Integer

    ' Constantes PowerPoint
    Const ppLayoutBlank As Integer = 12: Const msoTrue As Integer = -1: Const ppAlignCenter As Integer = 2

    ' --- CONFIGURAÇÃO DE DESIGN ---
    Dim fName As String: fName = "Nunito"
    Dim cGrafite As Long: cGrafite = RGB(44, 48, 62)
    Dim cCoral As Long: cCoral = RGB(255, 94, 87)
    Dim cAmarelo As Long: cAmarelo = RGB(250, 224, 114)
    Dim cVerde As Long: cVerde = RGB(0, 176, 80)
    
    ' --- REFERÊNCIA DA PLANILHA E TABELA ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Project Deployment Tracker")
    Set lo = ws.ListObjects("tblProjects")
    On Error GoTo 0
    
    If lo Is Nothing Then
        MsgBox "Erro: Planilha 'Project Deployment Tracker' ou Tabela 'tblProjects' não encontrada!", vbCritical
        Exit Sub
    End If

    ' --- CÁLCULOS TOTAIS PARA KPI GLOBAL ---
    totalProjetos = lo.ListRows.Count
    If totalProjetos > 0 Then
        totalEquipe = Application.Sum(lo.ListColumns(7).DataBodyRange)
        totalPendencias = Application.Sum(lo.ListColumns(10).DataBodyRange)
        If totalEquipe > 0 Then globalPercent = totalPendencias / totalEquipe Else globalPercent = 0
    Else
        MsgBox "A tabela está vazia.": Exit Sub
    End If
    
    ' Iniciar PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(class:="PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    
    Set pptPres = pptApp.Presentations.Add
    pptApp.Visible = msoTrue

    ' --- LOOP DE CRIAÇÃO DE SLIDES ---
    numSlides = Application.WorksheetFunction.RoundUp(totalProjetos / rowsPerSlide, 0)
    
    For s = 1 To numSlides
        startRow = ((s - 1) * rowsPerSlide) + 1
        endRow = s * rowsPerSlide
        If endRow > totalProjetos Then endRow = totalProjetos
        rowsOnThisSlide = (endRow - startRow) + 1
        
        Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)

        ' 1. BARRA DE KPI NO TOPO
        With pptSlide.Shapes.AddShape(1, 0, 0, 960, 60)
            .Fill.ForeColor.RGB = cGrafite: .Line.Visible = 0
        End With
        
        ' Título (Tamanho 16 conforme solicitado)
        With pptSlide.Shapes.AddTextbox(1, 30, 15, 600, 40)
            With .TextFrame.TextRange
                .Text = "ACOMPANHAMENTO DE PROJETOS E IMPLANTAÇÕES • JUNHO 2026"
                .Font.Name = fName: .Font.Size = 16: .Font.Bold = msoTrue: .Font.Color.RGB = vbWhite
            End With
        End With

        ' Métricas Consolidadas
        With pptSlide.Shapes.AddTextbox(1, 580, 18, 350, 30)
            With .TextFrame.TextRange
                .Text = totalProjetos & " PROJETOS  |  " & totalPendencias & " PENDÊNCIAS  |  " & Format(globalPercent, "0%") & " CRÍTICO"
                .Font.Name = fName: .Font.Size = 11: .Font.Bold = msoTrue: .Font.Color.RGB = vbWhite
                .ParagraphFormat.Alignment = 3
            End With
        End With

        ' 2. CRIAR TABELA
        Set pptTable = pptSlide.Shapes.AddTable(rowsOnThisSlide + 1, 13, 10, 80, 940, 350)
        
        Dim h As Variant
        h = Array("UNIT", "ID", "NOME DO PROJETO", "CIDADE", "UF", "CRIAÇÃO", "TIME", "ST. RH", "ST. OPS", "PEND.", "% TAXA", "DIAS P/", "LANÇAM.")
        
        For j = 0 To 12
            With pptTable.Table.Cell(1, j + 1).Shape
                .Fill.ForeColor.RGB = RGB(60, 65, 80)
                With .TextFrame.TextRange
                    .Text = h(j): .Font.Name = fName: .Font.Size = 8: .Font.Bold = msoTrue: .Font.Color.RGB = vbWhite
                    .ParagraphFormat.Alignment = ppAlignCenter
                End With
            End With
        Next j

        ' 3. PREENCHER DADOS
        Dim currentPPTLine As Integer: currentPPTLine = 2
        
        For i = startRow To endRow
            currentPend = Val(lo.DataBodyRange(i, 10).Value)
            vDataLanc = lo.DataBodyRange(i, 13).Value
            
            For j = 1 To 13
                Dim celula As Object: Set celula = pptTable.Table.Cell(currentPPTLine, j).Shape
                celula.TextFrame.TextRange.Font.Name = fName: celula.TextFrame.TextRange.Font.Size = 8
                celula.TextFrame.TextRange.Font.Color.RGB = cGrafite
                celula.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                
                Select Case j
                    Case 1 To 9
                        celula.TextFrame.TextRange.Text = lo.DataBodyRange(i, j).Text
                    Case 10 ' PEND.
                        celula.TextFrame.TextRange.Text = currentPend
                        If currentPend > 0 Then celula.TextFrame.TextRange.Font.Color.RGB = vbRed: celula.TextFrame.TextRange.Font.Bold = msoTrue
                    Case 11 ' % TAXA
                        celula.Fill.ForeColor.RGB = cCoral
                        celula.TextFrame.TextRange.Text = lo.DataBodyRange(i, j).Text
                        celula.TextFrame.TextRange.Font.Color.RGB = vbWhite: celula.TextFrame.TextRange.Font.Bold = msoTrue
                    Case 12 ' DIAS P/
                        celula.Fill.ForeColor.RGB = cAmarelo
                        celula.TextFrame.TextRange.Text = lo.DataBodyRange(i, j).Text
                        celula.TextFrame.TextRange.Font.Color.RGB = RGB(100, 70, 0): celula.TextFrame.TextRange.Font.Bold = msoTrue
                    Case 13 ' LANÇAM.
                        celula.Fill.ForeColor.RGB = cVerde
                        celula.TextFrame.TextRange.Text = Format(vDataLanc, "dd/mm/yy")
                        celula.TextFrame.TextRange.Font.Color.RGB = vbWhite: celula.TextFrame.TextRange.Font.Bold = msoTrue
                End Select
            Next j
            currentPPTLine = currentPPTLine + 1
        Next i

        ' Rodapé
        With pptSlide.Shapes.AddTextbox(1, 30, 510, 500, 20)
            .TextFrame.TextRange.Text = "Página " & s & " de " & numSlides & " | Project Deployment Tracker"
            .TextFrame.TextRange.Font.Name = fName: .TextFrame.TextRange.Font.Size = 7: .TextFrame.TextRange.Font.Color.RGB = RGB(150, 150, 150)
        End With
    Next s

    MsgBox "Dashboard 'tblProjects' gerado com sucesso!", vbInformation
End Sub
