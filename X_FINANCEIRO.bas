Attribute VB_Name = "X_FINANCEIRO"
Option Explicit
Option Base 1
      
      Global Conecta_Financeiro As New ADODB.Connection
      Function Banco_Financeiro(Conecta_Financeiro As ADODB.Connection)
      Conecta_Financeiro.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
      "Data Source=" & ThisWorkbook.Path & "\DATADUH_2024_FINANCEIRO.accdb;" '& _
      ";jet oledb:database password=123456PT"
      End Function
      Function dataduh_financeiro_botaoum()
      
            dataduh_financeiro_pagina
            dataduh_financeiro_filtros
            dataduh_financeiro_novoregistro1
            dataduh_financeiro_salarios
            dataduh_financeiro_excluir
            dataduh_financeiro_clonar
            dataduh_financeiro_alterardados
            dataduh_financeiro_resumo
            dataduh_financeiro_resumosemana
            dataduh_financeiro_resumomes
            dataduh_financeiro_detalhes
            dataduh_financeiro_documentos
            dataduh_financeiro_autorizacoes
            dataduh_financeiro_fixarpagamento
            dataduh_financeiro_novoregistrofavorito
            dataduh_ferramenta_exportarexcel
            dataduh_financeiro_calculos
            dataduh_financeiro_diagnostico
            dataduh_dashboard_layout
'            dataduh_financeiro_resumo
      
      End Function
      Function dataduh_financeiro_botaotres()
      
            dataduh_financeiro_filtrar1
            dataduh_financeiro_salariosexe
            dataduh_financeiro_clonarexe
            dataduh_financeiro_editaralterar
            dataduh_financeiro_novoregistroavancar
            dataduh_financeiro_novoregistrovoltar
            dataduh_financeiro_novoregistrogravar
            dataduh_financeiro_alterardadosexe
            dataduh_financeiro_recibo
            dataduh_financeiro_cheque
            dataduh_financeiro_canhoto
            dataduh_financeiro_textos
            dataduh_financeiro_fixarpagamentoexe
            dataduh_financeiro_favoritarpagamentoexe
            dataduh_financeiro_novoregistrofavoritoexe
            dataduh_financeiro_novoregistrofavoritover
            dataduh_financeiro_calculofolha
            dataduh_dashboard_filtros
      
      End Function
      Function dataduh_financeiro_mascaras()
      
            dataduh_financeiro_filtrosmasc
            dataduh_financeiro_editarmasc
            dataduh_financeiro_novoregistro1masc
            dataduh_financeiro_novoregistro2masc
            dataduh_financeiro_novoregistro3masc
            dataduh_financeiro_novoregistro4masc
            dataduh_financeiro_novoregistro5masc
            dataduh_financeiro_novoregistro6masc
            dataduh_financeiro_novoregistro7masc
            dataduh_financeiro_novoregistro8masc
            dataduh_financeiro_novoregistro9masc
            dataduh_financeiro_alterardadosmasc
            dataduh_financeiro_documentosmasc
            dataduh_financeiro_calculosmasc
            dataduh_dashboard_masc
      
      End Function
      Function dataduh_financeiro_double1()
      
             dataduh_financeiro_editar
      
      End Function
      Function dataduh_financeiro_double2()
'
            dataduh_financeiro_editarcriterio
      
      End Function
      Function dataduh_financeiro_pagina()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "" Then Exit Function
            dataduh_menu_permissao
            If CAIXAS(1) = "" Then Exit Function
            
            For i = 2 To 5
                  CAMADAS(i).Visible = False
            Next i
            
            With CAIXAS(2)
                  .Clear
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Novo Registro"
                  .AddItem "Novo Registro Favorito"
                  .AddItem "Pagar Colaboradores"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Filtro"
                  .AddItem "Exportar"
                  .AddItem "Exportar Excel"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Resumir"
                  .AddItem "Resumir Semana"
                  .AddItem "Resumir Mês"
                  .AddItem "Configurar Página"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Detalhes"
                  .AddItem "Documentos"
                  .AddItem "Autorizações"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Alterar Dados"
                  .AddItem "Clonar"
                  .AddItem "Excluir"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Dashboard"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Classificar Pagamento"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Calculos"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Calculadora Windows"
                  .AddItem "----------------------------------------------------------"
                  .AddItem "Relatório de Erros"
                  .AddItem "----------------------------------------------------------"
            End With
            
            dataduh_financeiro_gerarcontafixa
            
            CAIXAS(2) = "Relatório de Erros"
            dataduh_financeiro_diagnostico
            CAIXAS(2) = ""
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            
            MEMFILTRO(1, 1) = "Vencimento"
            MEMFILTRO(1, 2) = CDate(Date)
            MEMFILTRO(1, 3) = CDate(Date)
            MEMFILTRO(1, 4) = ""
            MEMFILTRO(1, 5) = ""
 
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            CAIXAS(1) = Nome
            X = 1: dataduh_ferramenta_operacao
      End Function
      Function dataduh_financeiro_filtros()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Filtro" Then Exit Function
            CAMADAS(5).Visible = False
            dataduh_formulario_filtros
             
            With CONTROLCX(2)
                  .Clear
                  .AddItem "Data"
                  .AddItem "Vencimento"
                  .Value = "Vencimento"
            End With
            
            CONTROLCX(5).Clear: dataduh_financeiro_assistente
            
            For i = 1 To UBound(RECORDVBA)
            CONTROLCX(5).AddItem RECORDVBA(i, 1)
            Next i
            
            ReDim RECORDVBA(20, 10)
            For i = 1 To 20
                  RECORDVBA(i, 2) = ROTULOS(3)
                  RECORDVBA(i, 3) = CDate(Date)
                  RECORDVBA(i, 4) = Format(Now, "hh:mm")
                  RECORDVBA(i, 5) = CAIXAS(1)
            Next
                                    
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            dataduh_financeiro_filtrosmasc
      End Function
     
      Function dataduh_financeiro_filtrosmasc()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Filtro" Then Exit Function
            If CAIXAS(3) = "Sair" Then Exit Function
            
            CAIXAS(2) = "Filtro"
        
            If COMPARADOR(1) <> CONTROLCX(1) Then
                  For i = 1 To 20
                        RECORDVBA(i, 6) = ""
                        RECORDVBA(i, 7) = ""
                        RECORDVBA(i, 8) = ""
                  Next
            End If
      
            If CONTROLCX(2) <> "Vencimento" And CONTROLCX(2) <> "Data" Then CONTROLCX(2) = "Vencimento"
            
            dataduh_buscador_relatorio
            
            If CONTROLCX(1) <> "" Then
                  If MATRIZ1(1, 1) <> "" Then
                        For i = 1 To 20
                              RECORDVBA(i, 6) = ""
                              RECORDVBA(i, 7) = ""
                              RECORDVBA(i, 8) = ""
                        Next
                        For i = 1 To UBound(MATRIZ1)
                              RECORDVBA(i, 6) = MATRIZ1(i, 6)
                              RECORDVBA(i, 7) = MATRIZ1(i, 7)
                              RECORDVBA(i, 8) = MATRIZ1(i, 8)
                        Next i
                  End If
            End If

            For i = 1 To 20
                  If CONTROLCX(1) <> "" Then RECORDVBA(i, 6) = CONTROLCX(1)
                  If CONTROLCX(3) <> "" Then RECORDVBA(i, 9) = CONTROLCX(3)
                  If CONTROLCX(4) <> "" Then RECORDVBA(i, 10) = CONTROLCX(4)
            Next
            
            If COMPARADOR(5) <> CONTROLCX(5) Then
                  If CONTROLCX(5) <> "" Then
                        CODIGO = CONTROLCX(5): CONTROLCX(6).Clear
                        dataduh_financeiro_abastececombo
                        For i = 1 To UBound(MATRIZ1)
                              On Error Resume Next
                              If MATRIZ1(i) <> 0 Then CONTROLCX(6).AddItem MATRIZ1(i)
                        Next i
                  End If
            End If
            
            
            If CONTROLCX(5) <> "" And CONTROLCX(6) <> "" Then
                  CAIXAS(3) = "Adicionar"
            Else
                  CAIXAS(3) = ""
            End If
            
            If COMPARADOR(1) <> CONTROLCX(1) Then
                  If CONTROLCX(1) <> "" Then
                        X = 3: dataduh_ferramenta_semanadata
                        CONTROLCX(3) = DATA1: CONTROLCX(4) = DATA2
                        CAIXAS(3) = "Filtrar"
                  End If
            End If
             
            
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            dataduh_formulario_filtroassistente
      End Function
      Function dataduh_financeiro_filtrarexe()
     
            DATA1 = MEMFILTRO(1, 2): DATA2 = MEMFILTRO(1, 3)
            dataduh_ferramenta_invertedata
            MEMFILTRO(1, 2) = DATA1: MEMFILTRO(1, 3) = DATA2

            For k = 1 To 3
            
                  SQL = "SELECT * FROM CAIXA WHERE " & MEMFILTRO(1, 1) & " BETWEEN #" & MEMFILTRO(1, 2) & "# AND #" & MEMFILTRO(1, 3) & "#"
                  For i = 1 To 20
                        On Error Resume Next
                        If MEMFILTRO(i, 4) <> "" Then
                              SQL = SQL & " AND " & MEMFILTRO(i, 4) & " LIKE '" & MEMFILTRO(i, 5) & "%'"
                        End If
                        If MEMFILTRO(i, 4) = "" Then Exit For
                        If MEMFILTRO(i, 4) = 0 Then Exit For
                  Next i
                  SQL = SQL & " ORDER BY Nome"
                                    
                  Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
                  
                  a = 0: a = CDbl(RS.RecordCount)
                  ReDim DATADUHLIST1(1, 1): DATADUHLIST1(1, 1) = "VAZIO"
                  ReDim VIRTUALTAB(1, 1): VIRTUALTAB(1, 1) = "VAZIO"
                  ReDim BDTITULOS(1): BDTITULOS(1) = "VAZIO"
                  TOTAL = 0: TITULO = 0

                  SEPARADOR = "|"
                  If a <> 0 Then
                        ReDim VIRTUALTAB(a, 40)
                        ReDim BDTITULOS(40)
                        ReDim DATADUHLIST1(a, 6)
                        
                        For i = 1 To 40
                              BDTITULOS(i) = RS.Fields(i - 1).Name
                        Next i
                                          
                        For j = 1 To a
                              'If j = 1 Then
                                    For i = 1 To 40
                                          BDTITULOS(i) = RS.Fields(i - 1).Name
                                          VIRTUALTAB(j, i) = RS(i - 1)
                                    Next i
                              'End If
  
                              DATADUHLIST1(j, 1) = RS(0)
                              DATADUHLIST1(j, 2) = CDbl(RS(24))
                              DATADUHLIST1(j, 3) = RS(23)
                                              
                              'codigos identificadores pagadores-----------------
                              If RS(8) = 4 Then CODIGO = "AC"
                              If RS(8) = 6 Then CODIGO = "AV"
                              If RS(8) = 125 Then CODIGO = "JV"
                              If RS(8) = 138 Then CODIGO = "JN"
                              If RS(8) = 165 Then CODIGO = "NV"
                              If RS(8) = 624 Then CODIGO = "JT"
                              If RS(8) = 635 Then CODIGO = "TV"
                              
                              'codigos identificadores fazendas-----------------
                              If RS(6) = "FAZENDA APARECIDA" Then CODIGO = "AP" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA CRILULI" Then CODIGO = "CR" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA DAS POSSES" Then CODIGO = "PO" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA RETORNO" Then CODIGO = "RE" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA SANTA RITA" Then CODIGO = "SR" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA SÃO JOSÉ" Then CODIGO = "SJ" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA SÃO JOSE" Then CODIGO = "SJ" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA SÃO LÁZARO" Then CODIGO = "SL" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA SÃO LAZARO" Then CODIGO = "SL" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA UITGEESTERMEER" Then CODIGO = "UI" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA VALE VERDE" Then CODIGO = "VV" & SEPARADOR & CODIGO
                              If RS(6) = "FAZENDA ZAMARIOLI" Then CODIGO = "ZA" & SEPARADOR & CODIGO
                              If RS(6) = "AGRÍCOLA VELDT" Then CODIGO = "AV" & SEPARADOR & CODIGO
                              If RS(6) = "AGRÍCOLA CASTRICUM" Then CODIGO = "AC" & SEPARADOR & CODIGO
                              
                              'codigos identificadores classe-----------------
                              If RS(5) = "AGENDA" Then CODIGO = "AG" & SEPARADOR & CODIGO
                              If RS(5) = "ANOTAÇÃO" Then CODIGO = "AN" & SEPARADOR & CODIGO
                              If RS(5) = "CONTÁBIL" Then CODIGO = "CO" & SEPARADOR & CODIGO
                              If RS(5) = "FLUXO" Then CODIGO = "FL" & SEPARADOR & CODIGO

                              If RS(28) = "PARC. 1/1" Then
                              DATADUHLIST1(j, 4) = "[" & CODIGO & "] " & RS(15) & " | " & RS(30)
                              Else
                                    If RS(28) = 0 Then
                                         DATADUHLIST1(j, 4) = "[" & CODIGO & "] " & RS(15) & " | " & RS(30)
                                    Else
                                         DATADUHLIST1(j, 4) = "[" & CODIGO & "] " & RS(15) & " | " & RS(30) & " | " & RS(28)
                                    End If
                              End If
                              
                              DATADUHLIST1(j, 5) = RS(17)
                              DATADUHLIST1(j, 6) = RS(25) 'Format(MATRIZ1(i, 26), "#,##0.00")

                              TOTAL = TOTAL + CDbl(RS(25))
                              If RS(25) = 0 Then TITULO = TITULO + 1
                              
                              RS.MoveNext
                        Next j
                  End If
                  RS.Close: Conecta_Financeiro.Close

                  If a = 0 Then
                        MEMFILTRO(1, 2) = CDate(MEMFILTRO(1, 2)) - 1
                        MEMFILTRO(1, 3) = CDate(MEMFILTRO(1, 3)) + 1
                  Else
                        Exit For
                  End If
            Next k

      If UBound(DATADUHLIST1) > 1 Then
            'Classificando Tipo
            ReDim MATRIZ1(UBound(DATADUHLIST1), 1)
            For i = 1 To UBound(DATADUHLIST1)
                  On Error Resume Next
                  MATRIZ1(i, 1) = DATADUHLIST1(i, 5)
            Next i
            DATADUHLIST1 = WorksheetFunction.SortBy(DATADUHLIST1, MATRIZ1, -1)
      
            'Classificando Datas
            ReDim MATRIZ1(UBound(DATADUHLIST1), 1)
            For i = 1 To UBound(DATADUHLIST1)
                  On Error Resume Next
                  MATRIZ1(i, 1) = DATADUHLIST1(i, 2)
            Next i
            DATADUHLIST1 = WorksheetFunction.SortBy(DATADUHLIST1, MATRIZ1, 1)
      End If

      End Function
      Function dataduh_financeiro_filtrar1()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Filtro" Then Exit Function
            If CAIXAS(3) <> "Filtrar" Then Exit Function
            
            ReDim MEMFILTRO(UBound(RECORDVBA), 5)
            For i = 1 To UBound(RECORDVBA)
            MEMFILTRO(i, 1) = CONTROLCX(2)
            MEMFILTRO(i, 2) = CONTROLCX(3)
            MEMFILTRO(i, 3) = CONTROLCX(4)
            MEMFILTRO(i, 4) = RECORDVBA(i, 7)
            MEMFILTRO(i, 5) = RECORDVBA(i, 8)
            Next i
            
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
      End Function
      Function dataduh_financeiro_filtrar2()
            dataduh_formulario_listbox
            
            If DATADUHLIST1(1, 1) = "VAZIO" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Nada a filtrar !"
                  Exit Function
            End If
    
            ReDim DIMENSOES(7)
                  DIMENSOES(1) = LISTAS(1).Width * 4.3 / 100
                  DIMENSOES(2) = LISTAS(1).Width * 7 / 100
                  DIMENSOES(3) = LISTAS(1).Width * 8 / 100
                  DIMENSOES(4) = LISTAS(1).Width * 63 / 100
                  DIMENSOES(5) = LISTAS(1).Width * 6 / 100
                  DIMENSOES(6) = LISTAS(1).Width * 7 / 100

            l = LISTAS(1).Left
            For i = 5 To 10
                  With ROTULOS(i)
                  .Left = l
                  .Width = DIMENSOES(i - 4)
                  End With
            l = l + ROTULOS(i).Width
            Next i
                
            If BDTITULOS(1) <> "VAZIO" Then
                  ROTULOS(5) = BDTITULOS(1)
                  ROTULOS(6) = BDTITULOS(25)
                  ROTULOS(7) = BDTITULOS(24)
                  ROTULOS(8) = BDTITULOS(16)
                  ROTULOS(9) = BDTITULOS(18)
                  ROTULOS(10) = BDTITULOS(26)

                  For i = 1 To 6
                        If i < 6 Then
                              DIMENSOES(7) = DIMENSOES(7) & DIMENSOES(i) & ";"
                        Else
                              DIMENSOES(7) = DIMENSOES(7) & DIMENSOES(i)
                        End If
                  Next i
                  
                  With LISTAS(1)
                        .Clear
                        .ColumnCount = 6
                        .ColumnWidths = DIMENSOES(7)
                        .List = DATADUHLIST1
                  End With
                                    
                  For i = 0 To DATADUH.ListBox1.ListCount - 1
                  DATADUH.ListBox1.List(i, 1) = FormatDateTime(DATADUH.ListBox1.List(i, 1), vbGeneralDate)
                  Next i
            Else
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Nada encontrado !"
            End If
            
            ROTULOS(2) = FormatCurrency(TOTAL, 2)
            If TITULO <> 0 Then
            
                  If MEMFILTRO(1, 5) <> "FAVORITA" Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        If TITULO = 1 Then
                              ROTULOS(24) = "Aviso !": ROTULOS(25) = TITULO & " Título sem valor !"
                        Else
                              ROTULOS(24) = "Aviso !": ROTULOS(25) = TITULO & " Títulos sem valores !"
                        End If
                  End If
                  
            End If

            If CAIXAS(2) = "Filtro" Then
                  If UBound(DATADUHLIST1) < 501 Then dataduh_financeiro_relatorio
            Else
                  Cells.Clear
            End If
            
      End Function
      Function dataduh_financeiro_abastececombo()
            SQL = "SELECT DISTINCT " & CODIGO & " FROM CAIXA "
            SQL = SQL & " ORDER BY " & CODIGO
            
            On Error Resume Next
            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
            
            a = 0: a = CDbl(RS.RecordCount)
            If a <> 0 Then
                  ReDim MATRIZ1(a): a = 1
                  
                  Do While Not RS.EOF
                        MATRIZ1(a) = RS(0)
                        RS.MoveNext: a = a + 1
                  Loop
            End If
            
            RS.Close: Conecta_Financeiro.Close
      End Function
      Function dataduh_financeiro_abastececombomemoria()
            SQL = "SELECT DISTINCT " & CODIGO & " FROM MEMORIA "
            SQL = SQL & " ORDER BY " & CODIGO
            
            On Error Resume Next
            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
            
            X = 0: a = 0: a = CDbl(RS.RecordCount)
            If a <> 0 Then
                  For i = 1 To a
                        If RS(0) = 0 Then X = X + 1
                        RS.MoveNext
                  Next i
                  RS.MoveFirst
                  
                  ReDim MATRIZ1(a - X): a = 1
                  Do While Not RS.EOF
                        If RS(0) <> 0 Then
                        MATRIZ1(a) = RS(0)
                        a = a + 1
                        End If
                        RS.MoveNext:
                  Loop
            End If
            RS.Close: Conecta_Financeiro.Close
      End Function
      Function dataduh_financeiro_relatorio()
            dataduh_relatorio_layout
            Application.ScreenUpdating = False
                                                  
                  With Cells
                            .Clear
                            .Font.Name = "COURIER NEW"
                            .Font.Name = "Swis721 BT"
                            .RowHeight = 15
                  End With
                  
                  a = Array("Vencimento", "Duplicata", "Nome", "Cobrança", "Valor", "Sub Total")
                  b = Array(12, 18, 127, 10.5, 13, 13)
                  For i = 1 To 6
                            Cells(4, i) = a(i)
                            Columns(i).ColumnWidth = b(i)
                  Next i
     
                  Columns("E:F").NumberFormat = "#,##0.00"
                  Columns("E:F").HorizontalAlignment = xlRight
                  Columns(1).HorizontalAlignment = xlLeft
                  Columns(2).HorizontalAlignment = xlCenter

                  
                  Rows("1:1").RowHeight = 25
                  Rows("2:3").RowHeight = 20.5
                  Rows("1:1").Font.Size = 22
                  Rows("1:1").Borders(xlEdgeBottom).Weight = xlThin
                  Rows("1:2").Font.Bold = True
                  
                  ReDim MATRIZ1(UBound(DATADUHLIST1), 6)
                  'Exit Function
                  For i = 1 To UBound(DATADUHLIST1)
                        MATRIZ1(i, 1) = VBA.Format(DATADUHLIST1(i, 2), "mm/dd/yyyy")
                        MATRIZ1(i, 2) = (DATADUHLIST1(i, 3))
                        MATRIZ1(i, 3) = (DATADUHLIST1(i, 4))
                        MATRIZ1(i, 4) = (DATADUHLIST1(i, 5))
                        MATRIZ1(i, 5) = CDbl(DATADUHLIST1(i, 6))
                        MATRIZ1(i, 6) = ""
                  Next i
                   
                  Set BASE = Range(Cells(5, 1), Cells(UBound(MATRIZ1) + 4, 6))
                  
                  BASE.Value = MATRIZ1
                  BASE.Borders(xlEdgeTop).Weight = xlThin
                  BASE.Borders(xlEdgeBottom).Weight = xlHairline
                  BASE.Borders(xlInsideHorizontal).Weight = xlHairline
                  
                  'calculos subtotal----------------------------------
                  Lin = Range("a4").CurrentRegion.Rows.Count
                  X = 0: SOMATUDO = 0
                  For i = 5 To Lin + 4
                         SOMATUDO = SOMATUDO + Cells(i, 5).Value
                         X = X + 1
                         
                         If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                               Cells(i - X + 1, 6).Value = SOMATUDO
                               SOMATUDO = 0: X = 0
                               Range(Cells(i, 1), Cells(i, 6)).Borders(xlEdgeBottom).Weight = xlThin
                         End If
                         If Cells(i, 4).Value = "DEPOSITO" Then
                               Cells(i, 4).Interior.Color = RGB(180, 255, 180)
                         End If
                  Next i
                  'calculos subtotal----------------------------------
                  
                  Cells(1, 4) = FormatCurrency(TOTAL, 2)
                  Set BASE = Range(Cells(1, 4), Cells(1, 6))
                  BASE.HorizontalAlignment = xlCenterAcrossSelection
                  BASE.Interior.Color = RGB(60, 60, 50)
                  BASE.Font.Color = RGB(255, 255, 255)
                                  
                  CODIGO = ""
                  For i = 1 To UBound(MEMFILTRO)
                        If MEMFILTRO(i, 4) <> "" Then
                              If MEMFILTRO(i, 5) <> 0 Then
                                    CODIGO = CODIGO & MEMFILTRO(i, 4) & ">" & MEMFILTRO(i, 5) & " | "
                              End If
                        End If
                  Next i
                  
                  If UBound(MATRIZ1) = 1 Then
                        CODIGO = CODIGO & UBound(MATRIZ1) & " REGISTRO."
                  Else
                        CODIGO = CODIGO & UBound(MATRIZ1) & " REGISTROS."
                  End If
                  
                  Cells(1, 1).Value = VIRTUALTAB(1, 10)
                  Cells(2, 1).Value = "Relatório Financeiro"
                  Cells(3, 1).Value = CODIGO

            Application.ScreenUpdating = True
      End Function
      Function dataduh_financeiro_assistente()
            SQL = "CAIXA": Banco_Financeiro Conecta_Financeiro
            RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
            
            If CDbl(RS.RecordCount) <> 0 Then
                  ReDim RECORDVBA(40, 2): ReDim BDTITULOS(40)
                  For i = 1 To 40
                        RECORDVBA(i, 1) = RS.Fields(i - 1).Name
                        RECORDVBA(i, 2) = 0
                  Next
            End If
            
            RECORDVBA(2, 2) = ROTULOS(3)
            RECORDVBA(3, 2) = CDate(Date)
            RECORDVBA(4, 2) = Format(Now, "hh:mm")
            
            RS.Close: Conecta_Financeiro.Close
      End Function
      Function dataduh_financeiro_carregaassistente()
            ReDim DIMENSOES(2)
            DIMENSOES(1) = (LISTAS(2).Width * 23) / 100
            DIMENSOES(2) = LISTAS(2).Width - DIMENSOES(1) - 2
            DIMENSOES(1) = DIMENSOES(1) & ";" & DIMENSOES(2)
      
            With LISTAS(2)
                  .Clear
                  .ColumnCount = 2
                  .ColumnWidths = DIMENSOES(1)
                  .List = RECORDVBA
            End With
      End Function
      Function dataduh_financeiro_salariosexe()
                If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
                If CAIXAS(2) <> "Pagar Colaboradores" Then Exit Function
                If CAIXAS(3) <> "Criar" Then Exit Function
               
                ReDim MATRIZ1(7, 2)
                MATRIZ1(1, 1) = "SALÁRIO": MATRIZ1(1, 2) = "401"
                MATRIZ1(2, 1) = "BÔNUS": MATRIZ1(2, 2) = "406"
                MATRIZ1(3, 1) = "13° PARCELA 1": MATRIZ1(3, 2) = "403"
                MATRIZ1(4, 1) = "13° PARCELA 2": MATRIZ1(4, 2) = "403"
                MATRIZ1(5, 1) = "13° TOTAL": MATRIZ1(5, 2) = "403"
                MATRIZ1(6, 1) = "RESCISÃO": MATRIZ1(6, 2) = "404"
                MATRIZ1(7, 1) = "FÉRIAS": MATRIZ1(7, 2) = "402"
                
                X = 0
                For i = 1 To UBound(MATRIZ1)
                        If CONTROLCX(1) = MATRIZ1(i, 1) Then
                        X = 1: ORIGEM = (MATRIZ1(i, 2))
                        End If
                Next i
                
                If Not VBA.IsDate(CONTROLCX(2).Text) = True Then CONTROLCX(2).Text = ""
                
                If X = 0 Then
                        CONTROLCX(1) = ""
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Escolha uma opção !"
                        Exit Function
                End If
                
                If CONTROLCX(2) = "" Or CONTROLCX(2) = 0 Then
                        CONTROLCX(2) = ""
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Escolha uma data !"
                        Exit Function
                End If
                
                DATA1 = CDate(CONTROLCX(2))
                If DATA1 < CDate(Date) Then DATA1 = CDate(Date)
                If DATA1 > CDate(Date) + 15 Then DATA1 = CDate(Date)
                CONTROLCX(2) = DATA1
                
                dataduh_ferramenta_listaselect2
                
                If MATRIZ2(1) = 0 Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione todos ou alguns !"
                        Exit Function
                End If
                
                Cells.Clear
                For k = 1 To UBound(MATRIZ2)
                      ReDim RECORDVBA(39)
                      
                      RECORDVBA(1) = ROTULOS(3): RECORDVBA(2) = CDate(Date): RECORDVBA(3) = Format(Now, "hh:mm")
                      RECORDVBA(4) = "DESPESA": RECORDVBA(5) = "CONTÁBIL": RECORDVBA(7) = "GERAL"
                      
                      CAMPO = "Status": CRITERIO = "PRESIDENTE"
                      dataduh_buscador_contato
                      
                      RECORDVBA(8) = MATRIZ1(1): RECORDVBA(9) = MATRIZ1(7): RECORDVBA(10) = MATRIZ1(6)
                      RECORDVBA(11) = MATRIZ1(17): RECORDVBA(12) = MATRIZ1(18): RECORDVBA(13) = MATRIZ1(20)

                      CRITERIO = MATRIZ2(k): CAMPO = "ID"
                      dataduh_buscador_contato
                      RECORDVBA(6) = MATRIZ1(32)
                      
                      RECORDVBA(14) = MATRIZ1(1): RECORDVBA(15) = MATRIZ1(7): RECORDVBA(16) = MATRIZ1(6)
                      RECORDVBA(17) = MATRIZ1(17): RECORDVBA(18) = MATRIZ1(18): RECORDVBA(19) = MATRIZ1(19)
                      RECORDVBA(20) = MATRIZ1(20): RECORDVBA(21) = MATRIZ1(21): RECORDVBA(22) = MATRIZ1(22)
                      If MATRIZ1(17) = "CHEQUE" Then
                      RECORDVBA(11) = "CHEQUE": RECORDVBA(12) = "SICOOB": RECORDVBA(13) = "0"
                      End If
                      
                      DATA1 = CONTROLCX(2)
                      dataduh_ferramenta_fimsemana
                      
                      RECORDVBA(23) = CDbl(DATA1): RECORDVBA(24) = (DATA1): RECORDVBA(25) = 0
                      RECORDVBA(26) = "ATIVO": RECORDVBA(27) = ORIGEM: RECORDVBA(28) = 0
                      RECORDVBA(29) = "FOLHA": RECORDVBA(30) = "PAGAMENTO DE " & CONTROLCX(1)
                      
                      SQL = "CAIXA"
                      Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
                      
                      RS.AddNew
                      For j = 1 To 39
                            If RECORDVBA(j) = "" Then RECORDVBA(j) = 0
                            RS(j) = RECORDVBA(j)
                      Next
                      RS.Update
                      
                      RS.Close: Conecta_Financeiro.Close
                      
                Next k
                
                dataduh_parametro_comandoson
                
                Nome = CAIXAS(1)
                ReDim MEMFILTRO(1, 5)
                
                MEMFILTRO(1, 1) = "Vencimento"
                MEMFILTRO(1, 2) = DATA1
                MEMFILTRO(1, 3) = DATA1
                MEMFILTRO(1, 4) = "Contabil"
                MEMFILTRO(1, 5) = ORIGEM
                
                CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
                dataduh_financeiro_filtrarexe
                dataduh_financeiro_filtrar2
      End Function
      Function dataduh_financeiro_salarios()
               If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
               If CAIXAS(2) <> "Pagar Colaboradores" Then Exit Function
            
               dataduh_formulario_colaborador
               
               For i = 3 To 6
                        CONTROLRT(i).Visible = False
                        CONTROLCX(i).Visible = False
               Next i
    
               With CONTROLCX(1)
                     .Clear
                     .AddItem "SALÁRIO"
                     .AddItem "BÔNUS"
                     .AddItem "FÉRIAS"
                     .AddItem "13° PARCELA 1"
                     .AddItem "13° PARCELA 2"
                     .AddItem "13° TOTAL"
                     .AddItem "RESCISÃO"
                     .Value = "SALÁRIO"
               End With
                  
               With CAIXAS(3)
                     .Clear
                     .AddItem "Criar"
                     .AddItem "Sair"
               End With
                                         
               CONTROLRT(1) = "Tipo"
               CONTROLRT(2) = "Data"
               
               ROTULOS(22) = "Escolha o tipo e digite a data, selecione todos ou alguns."
               ROTULOS(23) = "Selecione o comando 'Criar' e Execute."
      
      End Function
      Function dataduh_financeiro_excluir()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Excluir" Then Exit Function
      
            CODIGO = MsgBox("Deseja realmente excluir ?", vbYesNo, "DATADUH")
            If CODIGO = 7 Then Exit Function
            
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) = 0 Then Exit Function
            
            For i = 1 To UBound(MATRIZ2)
                  SQL = "SELECT * FROM CAIXA WHERE ID = " & MATRIZ2(i)
                  
                  Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
                  On Error Resume Next
                  RS.Update
                  RS.Delete
                  RS.Close: Conecta_Financeiro.Close
            Next i

            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            
      End Function
      Function dataduh_financeiro_clonarexe()
               If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
               If CAIXAS(2) <> "Clonar" Then Exit Function
               If CAIXAS(3) <> "Criar" Then Exit Function

               If CONTROLCX(3) = "" Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !"
                        If CONTROLRT(1) = "" Then ROTULOS(25) = "Digite a data e o intervalo !"
                        If CONTROLRT(1) <> "" Then ROTULOS(25) = "Digite a data !"
                        Exit Function
               End If
               
               dataduh_ferramenta_operacao
               
               If CONTROLRT(1) = "" Then
                     If CONTROLCX(4) = "" Then CONTROLCX(4) = 0
                     dataduh_ferramenta_listaselect1
                     CAMPO = "ID": CRITERIO = MATRIZ2(1)
                     dataduh_buscador_titulo
                     
                     ReDim RECORDVBA(CONTROLCX(2), 40)
                     
                     DATA1 = CONTROLCX(3)
                     d = Day(DATA1): m = Month(DATA1): a = Year(DATA1)
      
                     For i = 1 To UBound(RECORDVBA)
                           For j = 1 To 40
                                RECORDVBA(i, j) = MATRIZ1(j)
                           Next j
            
                           RECORDVBA(i, 1) = 0
                           RECORDVBA(i, 2) = ROTULOS(3)
                           RECORDVBA(i, 3) = CDate(Date)
                           RECORDVBA(i, 4) = Format(Now, "hh:mm")
            
                           If CONTROLCX(4) > 28 And CONTROLCX(4) < 32 Then
                                DATA1 = d & "/" & m & "/" & a
                                m = m + 1
                                If m = 13 Then
                                      m = 1: a = a + 1
                                End If
                                dataduh_ferramenta_fimsemana
                           Else
                                DATA1 = CDate(DATA1) + CDbl(CONTROLCX(4))
                                dataduh_ferramenta_fimsemana
                           End If
            
                           RECORDVBA(i, 24) = CONTROLCX(5)
                           RECORDVBA(i, 25) = DATA1
                           RECORDVBA(i, 29) = "PARC. " & i & "/" & CONTROLCX(2)
                           RECORDVBA(i, 31) = CONTROLCX(6)
                           RECORDVBA(i, 38) = SERIAL
                     Next i
               Else
                     DATA1 = CONTROLCX(3)
                     dataduh_ferramenta_fimsemana
                     dataduh_ferramenta_listaselect1
                 
                     ReDim RECORDVBA(UBound(MATRIZ2), 40)
   
                     For k = 1 To UBound(RECORDVBA)
                           CAMPO = "ID": CRITERIO = MATRIZ2(k)
                           dataduh_buscador_titulo
                           
                           For j = 1 To 40
                                RECORDVBA(k, j) = MATRIZ1(j)
                           Next j

                           RECORDVBA(k, 1) = 0
                           RECORDVBA(k, 2) = ROTULOS(3)
                           RECORDVBA(k, 3) = CDate(Date)
                           RECORDVBA(k, 4) = Format(Now, "hh:mm")
                           RECORDVBA(k, 25) = DATA1
                           RECORDVBA(k, 29) = 0
                           RECORDVBA(k, 38) = SERIAL
                     Next k
               End If
               
               For i = 1 To UBound(RECORDVBA)
               SQL = "CAIXA"
                      Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
                      
                      RS.AddNew
                      For j = 1 To 39
                            If RECORDVBA(i, j + 1) = "" Then RECORDVBA(i, j + 1) = 0
                            RS(j) = RECORDVBA(i, j + 1)
                      Next
                      RS.Update
                      
               RS.Close: Conecta_Financeiro.Close
               Next i
               
                               
                CAMADAS(3).Visible = False
                Nome = CAIXAS(1)
                ReDim MEMFILTRO(1, 5)

                MEMFILTRO(1, 1) = "Data"
                MEMFILTRO(1, 2) = CDate(Date)
                MEMFILTRO(1, 3) = CDate(Date)
                MEMFILTRO(1, 4) = "Operacao"
                MEMFILTRO(1, 5) = SERIAL
                
                CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
                dataduh_financeiro_filtrarexe
                dataduh_financeiro_filtrar2
               
               If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
               ROTULOS(24) = "Operação: " & SERIAL: ROTULOS(25) = "Clonagem bem sucedida !"
      End Function
      Function dataduh_financeiro_clonar()
               If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
               If CAIXAS(2) <> "Clonar" Then Exit Function
                              
               dataduh_ferramenta_listaselect1
               If MATRIZ2(1) = 0 Then
                    If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                    ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                    Exit Function
               End If
                      
               dataduh_formulario_filtros
               
               CAMPO = "ID": CRITERIO = MATRIZ2(1)
               dataduh_buscador_titulo
               CONTROLCX(5) = MATRIZ1(24)
               CONTROLCX(6) = MATRIZ1(31)

               CONTROLCX(1).Visible = False
               LISTAS(2).Visible = False
                  
               With CAIXAS(3)
                     .Clear
                     .AddItem "Criar"
                     .AddItem "Sair"
               End With
                          
               CONTROLRT(2) = "Repetições"
               CONTROLRT(3) = "Data"
               CONTROLRT(4) = "Intervalo"
               CONTROLRT(5) = "Duplicata"
               CONTROLRT(6) = "Histórico"
               
               dataduh_ferramenta_listaselect1
               If UBound(MATRIZ2) = 1 Then CONTROLRT(1) = ""

               If UBound(MATRIZ2) > 1 Then
               CONTROLRT(1) = "CLONAR SELEÇÃO"
               For i = 2 To 6
               If i <> 3 Then
               CONTROLRT(i).Visible = False
               CONTROLCX(i).Visible = False
               End If
               Next i
               End If
               
               CONTROLCX(2) = 1
     End Function
     Function dataduh_financeiro_editar()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            CAIXAS(2) = "Editar"
            
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) = 0 Then Exit Function
            dataduh_formulario_editar
                       
            dataduh_financeiro_assistente
            
            CONTROLCX(5).Clear
            For i = 1 To UBound(RECORDVBA)
                  CONTROLCX(5).AddItem RECORDVBA(i, 1)
            Next i

            CAMPO = "ID": CRITERIO = MATRIZ2(1)
            dataduh_buscador_titulo
            
            ROTULOS(11) = CAIXAS(1) & " - " & MATRIZ1(16)
            
            If MATRIZ1(1) = "VAZIO" Then
            LISTAS(2).Clear
            Exit Function
            End If
            
            For i = 1 To UBound(RECORDVBA)
                  RECORDVBA(i, 2) = MATRIZ1(i)
                  RECORDVBA(2, 2) = ROTULOS(3)
                  RECORDVBA(3, 2) = CDate(Date)
                  RECORDVBA(4, 2) = Format(Now, "hh:mm")
            Next i
            
            dataduh_financeiro_carregaassistente
                                    
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            dataduh_financeiro_editarmasc
      End Function
      Function dataduh_financeiro_editarmasc()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Editar" Then Exit Function
            If CAIXAS(3) = "Sair" Then Exit Function

            If COMPARADOR(5) <> CONTROLCX(5) Then
                  If CONTROLCX(5) <> "" Then
                        CODIGO = CONTROLCX(5): CONTROLCX(6).Clear
                        dataduh_financeiro_abastececombo
                        For i = 1 To UBound(MATRIZ1)
                              If MATRIZ1(i) <> 0 Then CONTROLCX(6).AddItem MATRIZ1(i)
                        Next i
                  End If
            End If

            If CONTROLCX(5) <> "" And CONTROLCX(6) <> "" Then
                  CAIXAS(3) = "Alterar"
            Else
                  CAIXAS(3) = ""
            End If
            
            If CONTROLCX(5) = "ID" Then CONTROLCX(5) = ""
            If CONTROLCX(5) = "Usuario" Then CONTROLCX(5) = ""
            If CONTROLCX(5) = "Data" Then CONTROLCX(5) = ""
            If CONTROLCX(5) = "Hora" Then CONTROLCX(5) = ""
            If CONTROLCX(5) = "Pessoa" Then CONTROLCX(5) = ""
            If CONTROLCX(5) = "CPFCNPJ" Then CONTROLCX(5) = ""
            
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
      End Function
      Function dataduh_financeiro_novoregistro1()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            dataduh_formulario_cadastro

            ROTULOS(13) = 1
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Sair"
            End With
            
            dataduh_financeiro_assistente
            dataduh_financeiro_carregaassistente
            
            For i = 1 To 6
                  CONTROLRT(i) = RECORDVBA(i + 4, 1)
                  CONTROLCX(i) = RECORDVBA(i + 4, 2)
                  If RECORDVBA(i + 4, 2) = 0 Then CONTROLCX(i) = ""
            Next i
            
            dataduh_menu_abastecerglobal
            
            ReDim COMPARADOR(6)
            
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
                  If i > 4 Then
                        CONTROLRT(i).Visible = False: CONTROLCX(i).Visible = False
                  End If
            Next
            
            ROTULOS(14) = "[ Classificação do Título ]"
            CONTROLCX(1) = "DESPESA": CONTROLCX(2) = "CONTÁBIL"
            CONTROLCX(4) = "GERAL"
            
            dataduh_parametro_comandosoff
    
      End Function
      Function dataduh_financeiro_novoregistro1masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> "1" Then Exit Function
  
            TRAVACOMANDO = 0
            For i = 1 To 4
                  RECORDVBA(i + 4, 2) = CONTROLCX(i)
                  If CONTROLCX(i) = "" Then RECORDVBA(i + 4, 2) = 0
                  If i < 5 Then
                  If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
                  End If
            Next i
 
            If TRAVACOMANDO = 4 Then
                  CAIXAS(3) = "Avançar"
            Else
                  CAIXAS(3) = ""
            End If
            
            For i = 1 To 40
                 If RECORDVBA(i, 2) = "" Then RECORDVBA(i, 2) = 0
            Next i
            
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro2()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
         
            dataduh_formulario_cadastro
            ROTULOS(13) = 2
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Voltar"
                  .AddItem "Sair"
            End With
            
            dataduh_financeiro_carregaassistente
                   
            For i = 1 To 6
                  CONTROLRT(i) = RECORDVBA(i + 8, 1)
                  CONTROLCX(i) = RECORDVBA(i + 8, 2)
                  If RECORDVBA(i + 8, 2) = 0 Then CONTROLCX(i) = ""
            Next i
            
            ROTULOS(14) = "[ Informações do Pagador ]"
            CONTROLRT(1) = "Id": CONTROLRT(2) = "Nome": CONTROLRT(3) = "CPF/CNPJ":
            CONTROLRT(4) = "Cobrança": CONTROLRT(5) = "Banco": CONTROLRT(6) = "Conta":

            dataduh_menu_abastecerglobal
            
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next

      End Function
      Function dataduh_financeiro_novoregistro2masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 2 Then Exit Function

            'Buscador universal--------------
            If CONTROLCX(1) <> COMPARADOR(1) Then
            CAMPO = "ID": CRITERIO = CONTROLCX(1)
            dataduh_buscador_contato
            End If
            If CONTROLCX(2) <> COMPARADOR(2) Then
            CAMPO = "Nome": CRITERIO = CONTROLCX(2)
            dataduh_buscador_contato
            End If
            If CONTROLCX(3) <> COMPARADOR(3) Then
            CAMPO = "CPFCNPJ": CRITERIO = CONTROLCX(3)
            dataduh_buscador_contato
            End If
            'Buscador universal--------------

            SOMATUDO = 0
            For i = 1 To 3
                  If CONTROLCX(i) <> COMPARADOR(i) Then SOMATUDO = SOMATUDO + 1
            Next i
            
            If SOMATUDO <> 0 Then
            CONTROLCX(1) = MATRIZ1(1): CONTROLCX(2) = MATRIZ1(7): CONTROLCX(3) = MATRIZ1(6)
            CONTROLCX(4) = MATRIZ1(17): CONTROLCX(5) = MATRIZ1(18): CONTROLCX(6) = MATRIZ1(20)
            End If
            
            TRAVACOMANDO = 0
            For i = 1 To 6
                  RECORDVBA(i + 8, 2) = CONTROLCX(i)
                  If CONTROLCX(i) = "" Then RECORDVBA(i + 8, 2) = 0
                  COMPARADOR(i) = CONTROLCX(i)
                  If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
            Next i
            
            If TRAVACOMANDO = 6 Then
                  CAIXAS(3) = "Avançar"
            Else
                  CAIXAS(3) = ""
            End If
            
            For i = 1 To 40
                 If RECORDVBA(i, 2) = "" Then RECORDVBA(i, 2) = 0
            Next i
            
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro3()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            dataduh_formulario_cadastro
            ROTULOS(13) = 3
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Voltar"
                  .AddItem "Sair"
            End With
            
            dataduh_financeiro_carregaassistente
            
            For i = 1 To 6
                  CONTROLRT(i) = RECORDVBA(i + 14, 1)
                  CONTROLCX(i) = RECORDVBA(i + 14, 2)
                  If RECORDVBA(i + 14, 2) = 0 Then CONTROLCX(i) = ""
            Next i
            
            ROTULOS(14) = "[ Identificação do Recebedor ]"
            CONTROLRT(1) = "Id": CONTROLRT(2) = "Nome": CONTROLRT(3) = "CPF/CNPJ":
            dataduh_menu_abastecerglobal
            
            ReDim COMPARADOR(6)
            
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
                  If i > 3 Then
                        CONTROLRT(i).Visible = False: CONTROLCX(i).Visible = False
                  End If
            Next
            
            
      End Function
       Function dataduh_financeiro_novoregistro3masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 3 Then Exit Function

            'Buscador universal--------------
            If CONTROLCX(1) <> COMPARADOR(1) Then
            CAMPO = "ID": CRITERIO = CONTROLCX(1)
            dataduh_buscador_contato
            End If
            If CONTROLCX(2) <> COMPARADOR(2) Then
            CAMPO = "Nome": CRITERIO = CONTROLCX(2)
            dataduh_buscador_contato
            End If
            If CONTROLCX(3) <> COMPARADOR(3) Then
            CAMPO = "CPFCNPJ": CRITERIO = CONTROLCX(3)
            dataduh_buscador_contato
            End If
            'Buscador universal--------------

            SOMATUDO = 0
            For i = 1 To 3
                  If CONTROLCX(i) <> COMPARADOR(i) Then SOMATUDO = SOMATUDO + 1
            Next i
            
            If SOMATUDO <> 0 Then
            CONTROLCX(1) = MATRIZ1(1): CONTROLCX(2) = MATRIZ1(7): CONTROLCX(3) = MATRIZ1(6)
            RECORDVBA(15, 2) = MATRIZ1(1): RECORDVBA(16, 2) = MATRIZ1(7): RECORDVBA(17, 2) = MATRIZ1(6)
            RECORDVBA(18, 2) = MATRIZ1(17): RECORDVBA(19, 2) = MATRIZ1(18): RECORDVBA(20, 2) = MATRIZ1(19)
            RECORDVBA(21, 2) = MATRIZ1(20): RECORDVBA(22, 2) = MATRIZ1(21): RECORDVBA(23, 2) = MATRIZ1(22)
            RECORDVBA(28, 2) = MATRIZ1(10): RECORDVBA(30, 2) = MATRIZ1(12): RECORDVBA(31, 2) = MATRIZ1(13)
            RECORDVBA(32, 2) = MATRIZ1(40): RECORDVBA(29, 2) = 1: RECORDVBA(27, 2) = "ATIVO"
            End If
            
            TRAVACOMANDO = 0
            For i = 1 To 3
                  COMPARADOR(i) = CONTROLCX(i)
                  If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
            Next i
            
            If TRAVACOMANDO = 3 Then
                  CAIXAS(3) = "Avançar"
            Else
                  CAIXAS(3) = ""
            End If
            
            For i = 1 To 40
                 If RECORDVBA(i, 2) = "" Then RECORDVBA(i, 2) = 0
            Next i
      
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro4()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            dataduh_formulario_cadastro
            ROTULOS(13) = 4
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Voltar"
                  .AddItem "Sair"
                  .Value = ""
            End With
            
            dataduh_financeiro_carregaassistente
            
            For i = 1 To 6
                  CONTROLRT(i) = RECORDVBA(i + 17, 1)
                  CONTROLRT(6) = "Chave Pix"
                  CONTROLCX(i) = RECORDVBA(i + 17, 2)
                  If CONTROLCX(i) = 0 Then CONTROLCX(i) = ""
            Next

            dataduh_menu_abastecerglobal
            
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            ROTULOS(14) = "[ Detalhes do Pagamento ]"
            
            dataduh_financeiro_novoregistro4masc
            
      End Function
      Function dataduh_financeiro_novoregistro4masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 4 Then Exit Function
                      
            For i = 1 To 6
                  RECORDVBA(i + 17, 2) = CONTROLCX(i)
            Next i
            
            If CONTROLCX(1) <> "DEPOSITO" And CONTROLCX(1) <> "PIX" Then
                  For i = 2 To 6
                        On Error Resume Next
                        CONTROLRT(i).Visible = False
                        On Error Resume Next
                        CONTROLCX(i).Visible = False
                  Next
            Else
                  For i = 2 To 6
                        CONTROLRT(i).Visible = True: CONTROLCX(i).Visible = True
                        CONTROLRT(6).Visible = False: CONTROLCX(6).Visible = False
                  Next
            End If
            
            If CONTROLCX(1) = "PIX" Then
                  For i = 2 To 5
                        On Error Resume Next
                        CONTROLRT(i).Visible = False
                        On Error Resume Next
                        CONTROLCX(i).Visible = False
                        CONTROLRT(6).Visible = True: CONTROLCX(6).Visible = True
                  Next
            End If
            
            
            '----------------------------------------------------
            If CONTROLCX(1) = "DEPÓSITO" Then CONTROLCX(1) = "DEPOSITO"
            If CONTROLCX(1) <> "DEPOSITO" And CONTROLCX(1) <> "PIX" Then CAIXAS(3) = "Avançar"
            
            If CONTROLCX(1) = "DEPOSITO" Then
                  TRAVACOMANDO = 0
                  For i = 2 To 5
                        If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
                        If TRAVACOMANDO = 4 Then CAIXAS(3) = "Avançar"
                        If TRAVACOMANDO <> 4 Then CAIXAS(3) = ""
                  Next i
            End If
            
            If CONTROLCX(1) = "PIX" And CONTROLCX(6) <> "" Then CAIXAS(3) = "Avançar"
            
            If CONTROLCX(5) <> "CORRENTE" And CONTROLCX(5) <> "POUPANÇA" Then
            RECORDVBA(22, 2) = "CORRENTE"
            Else
            RECORDVBA(22, 2) = CONTROLCX(5)
            End If
            '----------------------------------------------------
            
            For i = 1 To 40
                 If RECORDVBA(i, 2) = "" Then RECORDVBA(i, 2) = 0
            Next i
                  
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro5()
      If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
      If CAIXAS(2) <> "Novo Registro" Then Exit Function
      
            dataduh_formulario_cadastro
            ROTULOS(13) = 5
            
            dataduh_financeiro_carregaassistente
            
            For i = 1 To 6
                  If i < 4 Then
                  CONTROLRT(i) = RECORDVBA(i + 23, 1): CONTROLCX(i) = RECORDVBA(i + 23, 2)
                  If RECORDVBA(i + 23, 2) = 0 Then CONTROLCX(i) = ""
                  End If
                  If i = 4 Then
                  CONTROLRT(i) = RECORDVBA(i + 24, 1): CONTROLCX(i) = RECORDVBA(i + 24, 2)
                  If RECORDVBA(i + 24, 2) = 0 Then CONTROLCX(i) = ""
                  End If
                  If i = 6 Then
                  CONTROLRT(i) = RECORDVBA(i + 23, 1): CONTROLCX(i) = RECORDVBA(i + 23, 2)
                  If RECORDVBA(i + 23, 2) = 0 Then CONTROLCX(i) = ""
                  End If
                  CONTROLRT(5) = "Descrição"
            Next i
            
            CONTROLCX(6).Clear
            For i = 1 To 12
                  CONTROLCX(6).AddItem i
            Next i
            
            dataduh_menu_abastecerglobal
            
            ROTULOS(14) = "[ Detalhes da Duplicata ]"
            CAIXAS(3) = ""
            
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            CODIGO = CONTROLCX(4)
            dataduh_buscador_contabil
            CONTROLCX(5) = CODIGO
            
            dataduh_financeiro_novoregistro5masc
      End Function
       Function dataduh_financeiro_novoregistro5masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 5 Then Exit Function
            
            If CONTROLCX(2) <> "" Then
                  If CDate(CONTROLCX(2)) < CDate(Date) Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Data antiga !"
                  End If
            End If
            
            If CONTROLCX(3) <> "" Then
                  On Error Resume Next
                  If Not VBA.IsNumeric(CONTROLCX(3).Text) = True Then CONTROLCX(3).Text = ""
                  If CDbl(CONTROLCX(3)) < 1 Then CONTROLCX(3) = CDbl(CONTROLCX(3)) * -1
            End If
            
            'busca contabil-------------------------
            If CONTROLCX(4) <> COMPARADOR(4) Then
                  CODIGO = CONTROLCX(4)
                  dataduh_buscador_contabil
                  CONTROLCX(5) = CODIGO
            End If
            If CONTROLCX(5) <> COMPARADOR(5) Then
                  CODIGO = CONTROLCX(5)
                  dataduh_buscador_contabil
                  CONTROLCX(4) = CODIGO
            End If
            'busca contabil-------------------------
            
            If CONTROLCX(6) <> "" Then
                  If CONTROLCX(6) > 12 Then CONTROLCX(6) = 12
                  If CONTROLCX(6) < 1 Then CONTROLCX(6) = 1
            End If
            
            If CONTROLCX(2) <> "" Then
            DATA1 = CONTROLCX(2)
            dataduh_ferramenta_fimsemana
            CONTROLCX(2) = DATA1
            End If
                      
            RECORDVBA(24, 2) = CONTROLCX(1)
            RECORDVBA(25, 2) = CONTROLCX(2)
            RECORDVBA(26, 2) = CONTROLCX(3)
            RECORDVBA(28, 2) = CONTROLCX(4)
            RECORDVBA(29, 2) = CONTROLCX(6)
            
            If CONTROLCX(6) = "" Then
                  RECORDVBA(29, 2) = 1
            Else
                  RECORDVBA(29, 2) = CONTROLCX(6)
            End If
            
            TRAVACOMANDO = 0
            For i = 1 To 4
                  COMPARADOR(i) = CONTROLCX(i)
                  If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
            Next i
            
            If TRAVACOMANDO = 4 Then CAIXAS(3) = "Avançar"
            If TRAVACOMANDO <> 4 Then CAIXAS(3) = ""
            
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            For i = 1 To 40
                 If RECORDVBA(i, 2) = "" Then RECORDVBA(i, 2) = 0
            Next i
 
            ReDim PARCELAMENTO(RECORDVBA(29, 2), 2)
            DATA1 = RECORDVBA(25, 2)
            d = Day(CDate(DATA1)): m = Month(CDate(DATA1)): a = Year(CDate(DATA1))
            
            SOMATUDO = 0
            For i = 1 To UBound(PARCELAMENTO)
                  dataduh_ferramenta_fimsemana
                  PARCELAMENTO(i, 1) = VBA.Format(DATA1, "dd/mm/yyyy")
                  PARCELAMENTO(i, 2) = (Fix((CDbl(RECORDVBA(26, 2)) / CDbl(RECORDVBA(29, 2))) * 100)) / 100
                  m = m + 1
                  If m = 13 Then a = a + 1
                  If m = 13 Then m = 1
                  DATA1 = d & "/" & m & "/" & a
                  SOMATUDO = SOMATUDO + PARCELAMENTO(i, 2)
            Next i
            i = i - 1: PARCELAMENTO(i, 2) = (RECORDVBA(26, 2) - SOMATUDO) + PARCELAMENTO(i, 2)
            
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro6()
            dataduh_financeiro_novoregistro5masc
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            PARCELAS = "SIM": CALCULADORA = ""
            
            dataduh_formulario_design
            
            ROTULOS(11) = CAIXAS(1): ROTULOS(12) = CAIXAS(2)
            ROTULOS(13) = 6: ROTULOS(14) = ""
            ROTULOS(14) = "[ Revisão das Parcelas ]"
            
            If UBound(PARCELAMENTO) > 6 Then
            ROTULOS(22) = "Parcelas 6 de " & UBound(PARCELAMENTO)
            Else
            
            End If
            
            If CDbl(RECORDVBA(29, 2)) = 1 Then
            ROTULOS(23) = RECORDVBA(29, 2) & " Parcela"
            Else
            ROTULOS(23) = RECORDVBA(29, 2) & " Parcelas com intervalo de 30 dias"
            End If
            
            With CAIXAS(3)
            .Clear
            .AddItem "Voltar"
            .AddItem "Sair"
            End With

            For i = 1 To UBound(PARCELAMENTO)
                  If i < 7 Then
                  CONTROLCX(i) = CDate(PARCELAMENTO(i, 1))
                  CAIXAS(i + 9) = CDbl(PARCELAMENTO(i, 2))
                  End If
            Next i


            For i = 1 To 6
                  If i > CDbl(RECORDVBA(29, 2)) Then
                        CONTROLRT(i).Visible = False
                        CONTROLCX(i).Visible = False
                        CAIXAS(i + 9).Visible = False
                  End If
            Next i
            
            dataduh_financeiro_novoregistro6masc
            dataduh_financeiro_carregaassistente
            
      End Function
      Function dataduh_financeiro_novoregistro6masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 6 Then Exit Function
            
            ReDim COMPARADOR(1, 2)
            For i = 1 To 6
                  If i <= CDbl(RECORDVBA(29, 2)) Then
                  
                  DATA1 = CONTROLCX(i)
                  dataduh_ferramenta_fimsemana
                  CONTROLCX(i) = DATA1
                  
                  PARCELAMENTO(i, 1) = CONTROLCX(i)
                  PARCELAMENTO(i, 2) = Round(CAIXAS(i + 9), 2)
                  If CONTROLCX(i) = "" Then COMPARADOR(1, 1) = 1
                  If CAIXAS(i + 9) = "" Then COMPARADOR(1, 1) = 1
                  End If
                                       
                  On Error Resume Next
                  If CONTROLCX(i) <> "" Then
                        If CDate(CONTROLCX(i)) < CDate(Date) Then
                              If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                              ROTULOS(24) = "Atenção !": ROTULOS(25) = "Data antiga !"
                        End If
                  End If
            Next i
            
            If COMPARADOR(1, 1) = 1 Or COMPARADOR(1, 1) = 1 Then
                  CAIXAS(3) = ""
            Else
                  CAIXAS(3) = "Avançar"
            End If
            
            dataduh_menu_mascaraglobal
      End Function
      Function dataduh_financeiro_novoregistro7()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            PARCELAS = "SIM": CALCULADORA = ""
            
            dataduh_formulario_design
            
            ROTULOS(11) = CAIXAS(1): ROTULOS(12) = CAIXAS(2)
            ROTULOS(13) = 7: ROTULOS(14) = ""
            ROTULOS(14) = "[ Revisão das Parcelas ]"
            ROTULOS(22) = ""
            
            If CDbl(RECORDVBA(29, 2)) = 1 Then
            ROTULOS(23) = RECORDVBA(29, 2) & " Parcela"
            Else
            ROTULOS(23) = RECORDVBA(29, 2) & " Parcelas com intervalo de 30 dias"
            End If
            
            With CAIXAS(3)
            .Clear
            .AddItem "Voltar"
            .AddItem "Sair"
            End With
            
            If UBound(PARCELAMENTO) < 7 Then
            For i = 1 To 6
                        CONTROLRT(i).Visible = False
                        CONTROLCX(i).Visible = False
                        CAIXAS(i + 9).Visible = False
            Next i
            CAIXAS(3) = "Avançar"
            dataduh_financeiro_carregaassistente
            Exit Function
            End If

            For i = 1 To 6
                  CONTROLRT(i) = "Parcela - " & i + 6
                  If i > CDbl((UBound(PARCELAMENTO) - 6)) Then
                        CONTROLRT(i).Visible = False
                        CONTROLCX(i).Visible = False
                        CAIXAS(i + 9).Visible = False
                  End If
                  
                   If i + 6 <= CDbl(UBound(PARCELAMENTO)) Then
                        CONTROLCX(i) = PARCELAMENTO(i + 6, 1)
                        CAIXAS(i + 9) = PARCELAMENTO(i + 6, 2)
                  End If
            Next i
  
            dataduh_financeiro_novoregistro7masc
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro7masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 7 Then Exit Function
            
            ReDim COMPARADOR(1, 2)
            For i = 1 To 6
                  If i + 6 <= CDbl(RECORDVBA(29, 2)) Then
                  
                  DATA1 = CONTROLCX(i)
                  dataduh_ferramenta_fimsemana
                  CONTROLCX(i) = DATA1
                  
                  PARCELAMENTO(i + 6, 1) = CONTROLCX(i)
                  On Error Resume Next
                  PARCELAMENTO(i + 6, 2) = Round(CAIXAS(i + 9), 2)
                  If CONTROLCX(i) = "" Then COMPARADOR(1, 1) = 1
                  If CAIXAS(i + 9) = "" Then COMPARADOR(1, 1) = 1
                  End If
                                       
                  On Error Resume Next
                  If CONTROLCX(i) <> "" Then
                        If CDate(CONTROLCX(i)) < CDate(Date) Then
                              If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                              ROTULOS(24) = "Atenção !": ROTULOS(25) = "Data antiga !"
                        End If
                  End If
            Next i
            
            If COMPARADOR(1, 1) = 1 Or COMPARADOR(1, 1) = 1 Then
                  CAIXAS(3) = ""
            Else
                  CAIXAS(3) = "Avançar"
            End If

            dataduh_menu_mascaraglobal
      End Function
      Function dataduh_financeiro_novoregistro8()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            dataduh_formulario_cadastro
            ROTULOS(13) = 8
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Voltar"
                  .AddItem "Sair"
            End With
                               
            For i = 1 To 6
                  CONTROLRT(i) = RECORDVBA(i + 29, 1)
                  CONTROLCX(i) = RECORDVBA(i + 29, 2)
                  If RECORDVBA(i + 29, 2) = 0 Then CONTROLCX(i) = ""
            Next i
            
            ROTULOS(14) = "[ Observações Gerais ]"

            dataduh_menu_abastecerglobal
            
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
            
            CONTROLRT(6).Visible = False: CONTROLCX(6).Visible = False

            dataduh_financeiro_novoregistro8masc
      End Function
      Function dataduh_financeiro_novoregistro8masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 8 Then Exit Function
            
            For i = 1 To 5
                  RECORDVBA(i + 29, 2) = CONTROLCX(i)
                  If RECORDVBA(i + 29, 2) = "" Then RECORDVBA(i + 29, 2) = 0
                  
            Next i
            
            TRAVACOMANDO = 0
            For i = 1 To 3
                  COMPARADOR(i) = CONTROLCX(i)
                  If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
            Next i
            
            If TRAVACOMANDO = 3 Then
                  CAIXAS(3) = "Avançar"
            Else
                  CAIXAS(3) = ""
            End If

            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistro9()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            
            dataduh_formulario_cadastro
            ROTULOS(13) = 9
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Voltar"
                  .AddItem "Sair"
            End With
                               
            For i = 1 To 3
                  CONTROLRT(i) = RECORDVBA(i + 34, 1)
                  CONTROLCX(i) = RECORDVBA(i + 34, 2)
                  If RECORDVBA(i + 34, 2) = 0 Then CONTROLCX(i) = ""
                  CONTROLRT(i + 3).Visible = False: CONTROLCX(i + 3).Visible = False
                  CONTROLRT(1) = "Id": CONTROLRT(3) = "Observações"
            Next i
            
            X = 1: dataduh_ferramenta_operacao
            ROTULOS(14) = "[ Direcionamento do Título ]"
            ROTULOS(22) = "Operação: "
            ROTULOS(23) = SERIAL

            dataduh_financeiro_novoregistro9masc
            CONTROLCX(1) = 1
            
      End Function
      Function dataduh_financeiro_novoregistro9masc()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 9 Then Exit Function
            
            For i = 1 To 3
                  RECORDVBA(i + 34, 2) = CONTROLCX(i)
                  If RECORDVBA(i + 34, 2) = "" Then RECORDVBA(i + 34, 2) = 0
                  RECORDVBA(38, 2) = ROTULOS(23)
            Next i
            
            TRAVACOMANDO = 0
            For i = 1 To 3
                  COMPARADOR(i) = CONTROLCX(i)
                  If CONTROLCX(i) <> "" Then TRAVACOMANDO = TRAVACOMANDO + 1
            Next i

            If TRAVACOMANDO = 3 Then CAIXAS(3) = "Registrar"
            If TRAVACOMANDO <> 3 Then CAIXAS(3) = ""
            
            dataduh_financeiro_carregaassistente
      End Function
      Function dataduh_financeiro_novoregistrocomprovante()
            If ROTULOS(11) <> "FINANCEIRO" Then Exit Function
            If ROTULOS(12) <> "Novo Registro" Then Exit Function
            If ROTULOS(13) <> 9 Then Exit Function
            
            dataduh_relatorio_layout
            Application.ScreenUpdating = False
                        Cells.RowHeight = 12
                        Rows("1:3").RowHeight = 14
                        Rows("1:3").Font.Size = 15
                        Rows("1:3").Font.Bold = True
                        Cells.Font.Name = "Courier New"
            
                        Columns(1).ColumnWidth = 16
                        Columns(2).ColumnWidth = 60
                        Columns(3).ColumnWidth = 16
                        Columns(4).ColumnWidth = 16
                                 
                        Cells(1, 1) = UCase("Comprovante de Lançamento")
                        Cells(2, 1) = UCase("DATADUH - Módulo Financeiro")
                        Cells(4, 1) = "Campos"
                        Cells(4, 2) = "Critérios"
                        
                        Cells(44, 1) = "Vencimentos"
                        Cells(44, 2) = "Valores"
            
                        Range(Cells(44, 1), Cells(44, 2)).Font.Bold = True
                        Range(Cells(1, 1), Cells(1, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                        Range(Cells(4, 1), Cells(4, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                        Range(Cells(42, 1), Cells(42, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                        Range(Cells(44, 1), Cells(44, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                        Range(Cells(56, 1), Cells(56, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                        
                        Range(Cells(5, 1), Cells(42, 2)).Value = RECORDVBA
                        Range(Cells(45, 1), Cells(44 + UBound(PARCELAMENTO), 2)).Value = PARCELAMENTO
                        Range(Cells(45, 1), Cells(56, 1)).NumberFormat = "mm/dd/yyyy"
                        Range(Cells(45, 2), Cells(56, 2)).NumberFormat = "#,##0.00"
            Application.ScreenUpdating = False
            
            Set BASE = Range(Cells(1, 1), Cells(56, 2))
            
            v = "VERTICAL"
            l = 1: a = 1: TITULO = "$1:$4"
                
            ReDim MATRIZ1(9)
            For i = 1 To 9
            MATRIZ1(i) = 1
            Next
                
            dataduh_impressao_simples BASE, TITULO, v, l, a
            dataduh_ferramenta_pdf
      End Function
      Function dataduh_financeiro_novoregistroavancar()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            If CAIXAS(3) <> "Avançar" Then Exit Function
      
            Select Case ROTULOS(13)
            Case 1
                dataduh_financeiro_novoregistro2
            Case 2
                dataduh_financeiro_novoregistro3
            Case 3
                dataduh_financeiro_novoregistro4
            Case 4
                dataduh_financeiro_novoregistro5
            Case 5
                dataduh_financeiro_novoregistro6
            Case 6
                dataduh_financeiro_novoregistro7
            Case 7
                dataduh_financeiro_novoregistro8
            Case 8
                dataduh_financeiro_novoregistro9
            End Select
      End Function
      Function dataduh_financeiro_novoregistrovoltar()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            If CAIXAS(3) <> "Voltar" Then Exit Function
            
            Select Case ROTULOS(13)
            Case 2
                dataduh_financeiro_novoregistro1
            Case 3
                dataduh_financeiro_novoregistro2
            Case 4
                dataduh_financeiro_novoregistro3
            Case 5
                dataduh_financeiro_novoregistro4
            Case 6
                dataduh_financeiro_novoregistro5
            Case 7
                dataduh_financeiro_novoregistro6
            Case 8
                dataduh_financeiro_novoregistro7
            Case 9
                dataduh_financeiro_novoregistro8
            End Select
      End Function
      Function dataduh_financeiro_novoregistrogravar()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro" Then Exit Function
            If CAIXAS(3) <> "Registrar" Then Exit Function

            For k = 1 To UBound(PARCELAMENTO)
            
            SQL = "CAIXA"
            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
            RS.AddNew
            
            For i = 1 To 39
                  RS(i) = RECORDVBA(i + 1, 2)
                  RS(24) = PARCELAMENTO(k, 1)
                  RS(25) = PARCELAMENTO(k, 2)
                  If UBound(PARCELAMENTO) > 1 Then
                        RS(28) = "PARC. " & k & "/" & UBound(PARCELAMENTO)
                  Else
                        RS(28) = 0
                  End If
            Next
            
            RS.Update
            RS.Close: Conecta_Financeiro.Close
            Next k
  
            dataduh_financeiro_novoregistrocomprovante
            dataduh_parametro_comandoson
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            MEMFILTRO(1, 1) = "Data"
            MEMFILTRO(1, 2) = CDate(Date)
            MEMFILTRO(1, 3) = CDate(Date)
            MEMFILTRO(1, 4) = "Operacao"
            MEMFILTRO(1, 5) = SERIAL
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            
            If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
            ROTULOS(24) = "Operação: " & SERIAL: ROTULOS(25) = "Registrado !"
      End Function
      Function dataduh_financeiro_alterardados()
               If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
               If CAIXAS(2) <> "Alterar Dados" Then Exit Function
                              
               dataduh_ferramenta_listaselect1
               If MATRIZ2(1) = 0 Then
                    If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                    ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                    Exit Function
               End If
                      
               dataduh_formulario_filtros

               LISTAS(2).Visible = False
               
               With CAIXAS(3)
                     .Clear
                     .AddItem "Alterar"
                     .AddItem "Sair"
               End With
                                                      
               CONTROLRT(1) = "Campo"
               CONTROLRT(2) = "Novo Critério"
               
               For i = 2 To 6
                    CONTROLRT(i).Visible = False
                    CONTROLCX(i).Visible = False
               Next
               
               dataduh_financeiro_assistente
               
               CONTROLCX(1).Clear
               CONTROLCX(1).AddItem "Sequência"
               For i = 1 To 40
                    CONTROLCX(1).AddItem RECORDVBA(i, 1)
               Next i
                                    
               ReDim COMPARADOR(6)
               For i = 1 To 6
                     COMPARADOR(i) = CONTROLCX(i)
               Next
      End Function
      Function dataduh_financeiro_alterardadosmasc()
               If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
               If CAIXAS(2) <> "Alterar Dados" Then Exit Function
               
               Nome = Array("ID", "Usuario", "Data", "Hora", "Movimento", "Classe", "Fazenda", "Rateio", _
               "Pgid", "Pgnome", "Pgcpfcnpj", "Pgcobranca", "Pgbanco", "Pgconta", "Idforn", "Nome", "Cpfcnpj")
               
               For i = 1 To 17
               If UCase(CONTROLCX(1)) = UCase(Nome(i)) Then CONTROLCX(1) = ""
               Next i
                  
                  If COMPARADOR(1) <> CONTROLCX(1) Then
                        If CONTROLCX(1) <> "" Then
                              CONTROLRT(3).Visible = True
                              CONTROLCX(3).Visible = True
                              CONTROLRT(3) = CONTROLCX(1)
                              If CONTROLCX(1) = "Sequência" Then CONTROLRT(3) = "Duplicata"
                              CODIGO = CONTROLCX(1): CONTROLCX(3).Clear
                              dataduh_financeiro_abastececombo
                              For i = 1 To UBound(MATRIZ1)
                                    If MATRIZ1(i) <> 0 Then CONTROLCX(3).AddItem MATRIZ1(i)
                              Next i
                        Else
                              CONTROLRT(3).Visible = False
                              CONTROLCX(3).Visible = False
                        End If
                        
                  End If
                  
                  If CONTROLCX(1) = "Sequência" Then
                        If Not IsNumeric(CONTROLCX(3)) = True Then CONTROLCX(3) = ""
                        If IsNumeric(CONTROLCX(3)) = True Then CONTROLCX(3) = Format(CONTROLCX(3), "#")
                  End If
                  
                  For i = 1 To 6
                        COMPARADOR(i) = CONTROLCX(i)
                  Next
      End Function
      Function dataduh_financeiro_editarcriterio()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Editar" Then Exit Function

            dataduh_ferramenta_listaselect2
            If MATRIZ2(1) = 0 Then Exit Function
            
            CONTROLCX(5) = MATRIZ2(1): CONTROLCX(6) = "": CAIXAS(3) = ""
                       
            dataduh_financeiro_editarmasc
      End Function
      Function dataduh_financeiro_editaralterar()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Editar" Then Exit Function
            If CAIXAS(3) <> "Alterar" Then Exit Function
            If RECORDVBA(1, 2) = "" Then Exit Function
            If RECORDVBA(1, 2) = 0 Then Exit Function
            
            For i = 1 To 40
                  If RECORDVBA(i, 1) = CONTROLCX(5) Then RECORDVBA(i, 2) = CONTROLCX(6)
            Next i
            
            dataduh_financeiro_carregaassistente
    
            SQL = "SELECT * FROM CAIXA WHERE ID = " & RECORDVBA(1, 2)

            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
            
            For i = 2 To 40
                  If RECORDVBA(i, 2) = "" Then RECORDVBA(i, 2) = 0
                  RS(i - 1) = RECORDVBA(i, 2)
            Next
            
            RS.Update: RS.Close: Conecta_Financeiro.Close
            
            DATA1 = MEMFILTRO(1, 2): DATA2 = MEMFILTRO(1, 3)
            dataduh_ferramenta_invertedata
            MEMFILTRO(1, 2) = DATA1: MEMFILTRO(1, 3) = DATA2
                        
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            
            CAMADAS(3).Visible = False
            
            If CAMADAS(4).Visible = False Then
            If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
            End If
            ROTULOS(24) = "Aviso !": ROTULOS(25) = "Alterado !"
            
      End Function
      Function dataduh_financeiro_alterardadosexe()
               If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
               If CAIXAS(2) <> "Alterar Dados" Then Exit Function
               If CAIXAS(3) <> "Alterar" Then Exit Function

                If CONTROLCX(1) = "" And CONTROLCX(2) = "" Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Aviso !": ROTULOS(25) = "Nada a alterar !"
                        Exit Function
                End If
                              
                X = 1: dataduh_ferramenta_operacao
                dataduh_ferramenta_listaselect1
                
                For i = 1 To UBound(MATRIZ2)
                      Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET Usuario = '" & ROTULOS(3) & "' WHERE ID = " & MATRIZ2(i), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                      Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET Data = '" & CDate(Date) & "' WHERE ID = " & MATRIZ2(i), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                      Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET Hora = '" & Format(Now, "hh:mm") & "' WHERE ID = " & MATRIZ2(i), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                      Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET Operacao = '" & SERIAL & "' WHERE ID = " & MATRIZ2(i), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                      Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET " & CONTROLRT(3) & " = '" & CONTROLCX(3) & "' WHERE ID = " & MATRIZ2(i), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                      If CONTROLCX(1) = "Sequência" Then CONTROLCX(3) = CONTROLCX(3) + 1
                Next i
                
                CAMADAS(3).Visible = False
                Nome = CAIXAS(1)
                ReDim MEMFILTRO(1, 5)

                MEMFILTRO(1, 1) = "Data"
                MEMFILTRO(1, 2) = CDate(Date)
                MEMFILTRO(1, 3) = CDate(Date)
                MEMFILTRO(1, 4) = "Operacao"
                MEMFILTRO(1, 5) = SERIAL
                
                CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
                dataduh_financeiro_filtrarexe
                dataduh_financeiro_filtrar2
               
                If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                ROTULOS(24) = "Operação: " & SERIAL: ROTULOS(25) = "Alteração bem sucedida !"
      End Function
      Function dataduh_financeiro_resumo()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Resumir" Then Exit Function
  
           X = 1: dataduh_ferramenta_operacao
           dataduh_ferramenta_listaselect1
           
           If MATRIZ2(1) = 0 Then
                If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione todos ou alguns !"
                Exit Function
           End If

           For i = 1 To UBound(MATRIZ2)
                Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET Operacao = '" & SERIAL & "' WHERE ID = " & MATRIZ2(i), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
           Next i

           Nome = CAIXAS(1)
           ReDim MEMFILTRO(1, 5)

           MEMFILTRO(1, 1) = "Data"
           MEMFILTRO(1, 2) = CDate("01/01/2001")
           MEMFILTRO(1, 3) = CDate("01/01/2101")
           MEMFILTRO(1, 4) = "Operacao"
           MEMFILTRO(1, 5) = SERIAL
           
           CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
           dataduh_financeiro_filtrarexe
           dataduh_financeiro_filtrar2
               
      End Function
      Function dataduh_financeiro_detalhes()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Detalhes" Then Exit Function
            Application.ScreenUpdating = False
                  CAMADAS(3).Visible = False
                  
                  dataduh_ferramenta_listaselect1
                  If MATRIZ2(1) = 0 Then
                          If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                          ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um registro !"
                          Exit Function
                  End If

                  Planilha2.Select: Planilha2.Select
                  
                  With Cells
                        .Clear
                        .Font.Name = "Courier New"
                        .Font.Size = 13
                        .RowHeight = 16
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlLeft
                  End With
                              
                  Rows("1:3").RowHeight = 14
                  Rows("1:3").Font.Size = 15
                  Rows("1:3").Font.Bold = True
                  
                  Columns(1).ColumnWidth = 16
                  Columns(2).ColumnWidth = 60
                  Columns(3).ColumnWidth = 16
                  Columns(4).ColumnWidth = 16
                      
                  Cells(1, 1) = UCase("Detalhes do Registro")
                  Cells(2, 1) = UCase("DATADUH - Módulo Financeiro")
                  Cells(4, 1) = "Campos"
                  Cells(4, 2) = "Critérios"
                  
                  Range(Cells(1, 1), Cells(1, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                  Range(Cells(4, 1), Cells(4, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                  Range(Cells(42, 1), Cells(42, 4)).Borders(xlEdgeBottom).Weight = xlHairline
                                               
                  dataduh_financeiro_assistente
                  
                  dataduh_ferramenta_listaselect1
                  If MATRIZ2(1) = 0 Then Exit Function
                  
                  CAMPO = "ID": CRITERIO = MATRIZ2(1)
                  dataduh_buscador_titulo
                  For i = 1 To UBound(RECORDVBA) - 2
                        Cells(i + 4, 1).Value = RECORDVBA(i, 1)
                        Cells(i + 4, 2).Value = MATRIZ1(i)
                  Next i
                  
                  Set BASE = Range(Cells(5, 1), Cells(42, 4))
                  BASE.Font.Size = 12
                  Set BASE = Range(Cells(1, 1), Cells(42, 4))
                  
                  v = "VERTICAL": l = 1: a = 1: TITULO = ""
            Application.ScreenUpdating = True
            dataduh_impressao_simples BASE, TITULO, v, l, a
            dataduh_ferramenta_pdf
            Planilha1.Select: Planilha1.Select:
      End Function
      Function dataduh_financeiro_recibolayout()
            Application.ScreenUpdating = False
            
                  Planilha3.Select: Planilha3.Select
                  Cells.Select: Selection.UnMerge: Cells(1, 1).Select
                  Columns(1).ColumnWidth = 75: Columns(2).ColumnWidth = 35
                  
                  Planilha2.Select: Planilha2.Select
                  Cells.Select: Selection.UnMerge: Cells(1, 1).Select
                  Columns(1).ColumnWidth = 75: Columns(2).ColumnWidth = 35
                                    
                  With Cells
                        .Clear
                        .Font.Name = "Swis721 BT"
'                        .Font.Name = "verdana"
                        .Font.Size = 15
                        .RowHeight = 20
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlLeft
                  End With
                  Columns(2).HorizontalAlignment = xlRight
                  
                  a = Array(1, 10, 14)
                  For i = 1 To 3
                        Rows(a(i)).Font.Size = 19
                        Rows(a(i)).RowHeight = 28
                  Next i
                  
                  Rows(1).RowHeight = 38: Rows(1).Font.Size = 30: Rows(2).RowHeight = 60
                  Rows(2).VerticalAlignment = xlBottom: Rows(6).VerticalAlignment = xlBottom
                  
                  a = Array(1, 2, 6, 10, 14)
                  For i = 1 To 5
                        Rows(a(i)).Font.Bold = True
                  Next i
                  
                  a = Array(1, 20, 10, 14)
                  For i = 1 To 4
                        Rows(a(i)).Borders(xlEdgeBottom).Weight = xlThin
                  Next i
                  
                  a = Array(10, 14): b = 220
                  For i = 1 To 2
                        'Rows(a(i)).Interior.Color = RGB(b, b, b)
                        Rows(a(i)).NumberFormat = "$ #,##0.00"
                        Rows(a(i)).RowHeight = 18
                  Next i
                  Rows(22).NumberFormat = "dd/mm/yy;@"

                  With Cells(1, 2)
                        .Font.Size = 9
                  End With
               
                  Set BASE = Range("A11:B13")
                  With BASE
                        .VerticalAlignment = xlTop
                        .MergeCells = True
                        .Font.Name = "courier new"
                        .WrapText = True
                  End With
                  
                  Set BASE = Range("A15:B17")
                  With BASE
                        .VerticalAlignment = xlTop
                        .MergeCells = True
                        .Font.Name = "courier new"
                        .WrapText = True
                  End With
                  
            ActiveWindow.Zoom = 100
            Application.ScreenUpdating = False
      End Function
      Function dataduh_financeiro_documentos()
             If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
             If CAIXAS(2) <> "Documentos" Then Exit Function
                            
             dataduh_ferramenta_listaselect1
             If MATRIZ2(1) = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                  Exit Function
             End If
                    
             dataduh_formulario_filtros

             LISTAS(2).Visible = False
             
             With CONTROLCX(1)
                   .Clear
                   .AddItem "Recibos"
                   .AddItem "Cheques"
                   .AddItem "Canhoto dos Cheques"
                   .AddItem "E-mail Pagamento"
                   .AddItem "E-mail Cancelamento"
             End With
             
             With CAIXAS(3)
                   .Clear
                   .AddItem "Criar"
                   .AddItem "Sair"
             End With
                                    
             CONTROLRT(1) = "Tipo do documento"
   
             For i = 2 To 6
                  CONTROLRT(i).Visible = False
                  CONTROLCX(i).Visible = False
             Next
                   
             ReDim COMPARADOR(6)
             For i = 1 To 6
                   COMPARADOR(i) = CONTROLCX(i)
             Next
    End Function
    Function dataduh_financeiro_documentosmasc()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Documentos" Then Exit Function
            
            If CONTROLCX(1) = "Cheques" Then
                  If CONTROLCX(3) <> "" Then
                  If CDate(CONTROLCX(3)) < CDate(Date) Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Data antiga !"
                        CONTROLCX(3) = ""
                  End If
                  End If
            End If
            
            For i = 2 To 6
                CONTROLRT(i).Visible = False
                CONTROLCX(i).Visible = False
            Next
            
            If CONTROLCX(1) = "Cheques" Then
                CONTROLRT(2).Visible = True: CONTROLRT(2) = "Praça"
                CONTROLCX(2).Visible = True
                CONTROLRT(3).Visible = True: CONTROLRT(3) = "Data"
                CONTROLCX(3).Visible = True
            End If
      End Function
      Function dataduh_financeiro_recibo()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Documentos" Then Exit Function
            If CAIXAS(3) <> "Criar" Then Exit Function
            If CONTROLCX(1) <> "Recibos" Then Exit Function
            
            If CONTROLCX(1) = "" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Nada a criar !"
                  Exit Function
            End If
            
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                  Exit Function
             End If
             
            dataduh_financeiro_recibolayout
            
            'eliminando valores de 0 -----------------
            a = 0
            For j = 1 To UBound(MATRIZ2)
                  CAMPO = "ID": CRITERIO = MATRIZ2(j)
                  dataduh_buscador_titulo
                  If MATRIZ1(26) <> 0 Then a = a + 1
                  If MATRIZ1(26) = 0 Then MATRIZ2(j) = 0
            Next j
            
            ReDim MATRIZ3(a): a = 1
            For j = 1 To UBound(MATRIZ2)
                  If MATRIZ2(j) <> 0 Then
                        MATRIZ3(a) = MATRIZ2(j)
                        a = a + 1
                  End If
            Next j
            ReDim MATRIZ2(UBound(MATRIZ3)): MATRIZ2 = MATRIZ3
            'eliminando valores de 0 -----------------
            
            'eliminando boletos ----------------------
            a = 0
            For j = 1 To UBound(MATRIZ2)
                  CAMPO = "ID": CRITERIO = MATRIZ2(j)
                  dataduh_buscador_titulo
                  If MATRIZ1(18) <> "BOLETO" Then a = a + 1
                  If MATRIZ1(18) = "BOLETO" Then MATRIZ2(j) = 0
            Next j
                        
            If a = 0 Then
            If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
            ROTULOS(24) = "Aviso !": ROTULOS(25) = "Nada encontrado !"
            Exit Function
            End If

            ReDim MATRIZ3(a): a = 1
            For j = 1 To UBound(MATRIZ2)
                  If MATRIZ2(j) <> 0 Then
                        MATRIZ3(a) = MATRIZ2(j)
                        a = a + 1
                  End If
            Next j
            ReDim MATRIZ2(UBound(MATRIZ3)): MATRIZ2 = MATRIZ3
            'eliminando boletos ----------------------
           
            Application.ScreenUpdating = False
                  For j = 1 To UBound(MATRIZ2)
                        
                        Planilha2.Select: Planilha2.Select
                        Cells.ClearContents
                              
                        CAMPO = "ID": CRITERIO = MATRIZ2(j)
                        dataduh_buscador_titulo
                        
                        X = 1: dataduh_ferramenta_operacao
                        Banco_Financeiro Conecta_Financeiro: RS.Open "UPDATE CAIXA SET Operacao = '" & SERIAL & "' WHERE ID = " & MATRIZ2(j), Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                        '---------------------------------------------
                        Cells(1, 1) = "Recibo": Cells(2, 1) = "Pagador": Cells(6, 1) = "Recebedor"
                        Cells(10, 1) = "Total": Cells(14, 1) = "Motivo"
                        Cells(1, 2) = SERIAL: Cells(3, 1) = MATRIZ1(10)
                        Cells(7, 1) = MATRIZ1(16): Cells(10, 2) = MATRIZ1(26)
                        Cells(15, 1) = MATRIZ1(31): Cells(22, 2) = MATRIZ1(25)
                        Cells(21, 1) = MATRIZ1(16)
                        Cells(22, 1) = MATRIZ1(17)
                        '---------------------------------------------
                        Cells(11, 1).Select: ActiveCell.Formula2R1C1 = "=UPPER(VExtensoFree(R[-1]C[1]))"
                        '---------------------------------------------
     
                        b = 45
                        Rows("1:" & b).Copy
                        linha = (j * b) - (b - 1) & ":" & (j * b)
                        Planilha3.Select: Planilha3.Select
                        Rows(linha).Select: ActiveSheet.Paste
                        '---------------------------------------------
                  Next j
                  
                  j = j - 1: Set BASE = Range(Cells(1, 1), Cells(j * b, 2))
                  v = "VERTICAL": l = 1: a = j: TITULO = ""
                  
                  TRAVACOMANDO = "SIM"
                  ReDim MATRIZ1(9)
                  For i = 1 To 9
                        MATRIZ1(i) = 0.2
                  Next
                  
                  dataduh_impressao_simples BASE, TITULO, v, l, a
                  dataduh_ferramenta_pdf
                  
                  Planilha1.Select: Planilha1.Select
            Application.ScreenUpdating = True
      End Function
      Function dataduh_financeiro_chequelayout()
            Application.ScreenUpdating = False
                  Planilha2.Select: Planilha2.Select: Cells.Clear
                                       
                  With Cells
                        .Clear
                        .Font.Name = "Swis721 BT"
                        .Font.Size = 10
                        .RowHeight = 20
                        .ColumnWidth = 78
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlRight
                  End With
                  
                  n = Array(27.5, 28.5, 20.5, 19.5)
                  For i = 1 To 4
                        Rows(i).RowHeight = n(i)
                  Next i
                  
                  Rows(1).Font.Size = 13
                  Rows(1).NumberFormat = """#"" $ #,##0.00_.""#"";"
                  Rows(1).VerticalAlignment = xlTop
                  Rows(2).WrapText = True
                  Rows(2).HorizontalAlignment = xlLeft
                  Rows(2).VerticalAlignment = xlTop
                  Rows(3).HorizontalAlignment = xlLeft
            Application.ScreenUpdating = False
      End Function
      Function dataduh_financeiro_cheque()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Documentos" Then Exit Function
            If CAIXAS(3) <> "Criar" Then Exit Function
            If CONTROLCX(1) <> "Cheques" Then Exit Function
            
            For i = 1 To 3
                  If CONTROLCX(i) = "" Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Complete para continuar !"
                        Exit Function
                  End If
            Next i
            
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                  Exit Function
            End If
            
            dataduh_financeiro_chequelayout
            
            dataduh_ferramenta_listaselect1
            'eliminando valores de 0 -----------------
            a = 0
            For j = 1 To UBound(MATRIZ2)
                  CAMPO = "ID": CRITERIO = MATRIZ2(j)
                  dataduh_buscador_titulo
                  If MATRIZ1(26) <> 0 Then a = a + 1
                  If MATRIZ1(26) = 0 Then MATRIZ2(j) = 0
            Next j
            
            If a = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Registros sem valores !"
                  CAMADAS(3).Visible = False
                  Exit Function
            End If

            ReDim MATRIZ3(a): a = 1
            For j = 1 To UBound(MATRIZ2)
                  If MATRIZ2(j) <> 0 Then
                        MATRIZ3(a) = MATRIZ2(j)
                        a = a + 1
                  End If
            Next j
            ReDim MATRIZ2(UBound(MATRIZ3)): MATRIZ2 = MATRIZ3
            'eliminando <> de cheque -----------------
            a = 0
            For j = 1 To UBound(MATRIZ2)
                  CAMPO = "ID": CRITERIO = MATRIZ2(j)
                  dataduh_buscador_titulo
                  If MATRIZ1(18) = "CHEQUE" Then a = a + 1
                  If MATRIZ1(18) <> "CHEQUE" Then MATRIZ2(j) = 0
            Next j
            
            If a = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Nada encontrado !"
                  Exit Function
            End If
            
            ReDim MATRIZ3(a): a = 1
            For j = 1 To UBound(MATRIZ2)
                  If MATRIZ2(j) <> 0 Then
                        MATRIZ3(a) = MATRIZ2(j)
                        a = a + 1
                  End If
            Next j
            ReDim MATRIZ2(UBound(MATRIZ3)): MATRIZ2 = MATRIZ3
            'eliminando -------------------------------
            
            Application.ScreenUpdating = False
                  DATA1 = Format(CONTROLCX(3), "DD/MM/YYYY")
                  CIDADE = UCase(CONTROLCX(2) & ", " & Day(DATA1) & "             " & Format(CDate(DATA1), "MMMM") & "               " & Year(DATA1) & "       ")
                  
                  c = 0
                  For j = 1 To UBound(MATRIZ2)
                        CAMPO = "ID": CRITERIO = MATRIZ2(j)
                        dataduh_buscador_titulo
                        
                        Cells(1, j) = MATRIZ1(26)
                        Cells(3, j) = MATRIZ1(16)
                        Cells(4, j) = CIDADE
                        
                        Cells(2, j).Select: ActiveCell.Formula2R1C1 = "=UPPER(VExtensoFree(R[-1]C[0]))"
                        Cells(2, j) = "                       (" & Cells(2, j).Value & ")- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
                  Next j
                  j = j - 1: c = j
                  Set IMPRESS = Range(Cells(1, 1), Cells(4, j))
                  
                  If Cells(1, 1) <> 0 Then
                        dataduh_impressao_cheque
                        dataduh_ferramenta_pdf
                  End If
                  Planilha1.Select: Planilha1.Select
            Application.ScreenUpdating = True
      End Function
      Function dataduh_financeiro_canhoto()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Documentos" Then Exit Function
            If CAIXAS(3) <> "Criar" Then Exit Function
            If CONTROLCX(1) <> "Canhoto dos Cheques" Then Exit Function
            Application.ScreenUpdating = False
                  Planilha2.Select: Planilha2.Select
                  
                  With Cells
                        .Clear
                        .Font.Name = "Swis721 BT"
                        .Font.Size = 11
                        .RowHeight = 15
                        .ColumnWidth = 35
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlLeft
                  End With
                                    
                  dataduh_ferramenta_listaselect1
                  If MATRIZ2(1) = 0 Then
                        If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                        ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                        Exit Function
                  End If
                   
                  For j = 1 To UBound(MATRIZ2)
                        CAMPO = "ID": CRITERIO = MATRIZ2(j)
                        dataduh_buscador_titulo
                        If MATRIZ1(18) <> "CHEQUE" Then
                              If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                              ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione somente CHEQUES !"
                              Exit Function
                        End If
                  Next j

                  c = 1: l = 1
                  For j = 1 To UBound(MATRIZ2)
                        CAMPO = "ID": CRITERIO = MATRIZ2(j)
                        dataduh_buscador_titulo
                              Cells(l, c) = MATRIZ1(25) & " CHEQUE: " & MATRIZ1(24)
                              Cells(l + 1, c) = MATRIZ1(16)
                              Cells(l + 2, c) = "CPF:" & MATRIZ1(17)
                              Cells(l + 3, c) = MATRIZ1(26)
                              Cells(l + 3, c).Interior.Color = RGB(240, 240, 240)
                              Cells(l + 3, c).Font.Bold = True
                              Cells(l + 3, c).Font.Size = 12
                              
                              If MATRIZ1(28) = 406 Then Cells(l + 4, c) = Left(MATRIZ1(31), 19)
                              If MATRIZ1(28) <> 406 Then Cells(l + 4, c) = Left(MATRIZ1(31), 20)
                              Rows(l + 4).Borders(xlEdgeBottom).Weight = xlHairline
                               
                        c = c + 1
                        If c = 6 Then
                        l = l + 5: c = 1
                        End If
                  Next j

                  Set BASE = Range(Cells(1, 1), Cells(l + 5, 5))
                  v = "HORIZONTAL": l = 1: a = 10: TITULO = ""
                  
                  TRAVACOMANDO = "SIM"
                  ReDim MATRIZ1(9)
                  For i = 1 To 9
                        MATRIZ1(i) = 0.2
                  Next
                  
                  dataduh_impressao_simples BASE, TITULO, v, l, a
                  dataduh_ferramenta_pdf

                  Planilha1.Select: Planilha1.Select
            Application.ScreenUpdating = True
      End Function
      Function dataduh_financeiro_textos()
                  If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
                  If CAIXAS(2) <> "Documentos" Then Exit Function
                  If CAIXAS(3) <> "Criar" Then Exit Function
                  If Left(CONTROLCX(1), 6) <> "E-mail" Then Exit Function
                  
                  Application.ScreenUpdating = False
                        Planilha2.Select: Planilha2.Select
                        
                        With Cells
                              .Clear
                              .Font.Name = "Swis721 BT"
                              .Font.Name = "COURIER NEW"
                              .Font.Size = 12
                              .RowHeight = 15
                              .ColumnWidth = 100
                              .VerticalAlignment = xlCenter
                              .HorizontalAlignment = xlLeft
                        End With
                        
                        dataduh_ferramenta_listaselect1
                        If MATRIZ2(1) = 0 Then
                              If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                              ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um ou mais !"
                              Exit Function
                        End If
                        
                        For j = 1 To UBound(MATRIZ2)
                              CAMPO = "ID": CRITERIO = MATRIZ2(j)
                              dataduh_buscador_titulo
                              If MATRIZ1(18) = "CHEQUE" Then
                                    If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                                    ROTULOS(24) = "Atenção !": ROTULOS(25) = "Não Selecione CHEQUES !"
                                    Exit Function
                              End If
                        Next j
                
                        If Hour(Now) >= 12 Then Cells(2, 1).Value = "Boa tarde."
                        If Hour(Now) <= 12 Then Cells(2, 1).Value = "Bom dia."
                        Cells(3, 1) = ".": Cells(5, 1) = "." ': Cells(6, 1) = ".": Cells(7, 1) = ".":
                 
                        For j = 1 To UBound(MATRIZ2)
                              CAMPO = "ID": CRITERIO = MATRIZ2(j)
                              dataduh_buscador_titulo
                              
                              If j = 1 Then
                                    Cells(6, 1).Value = "COOPERADO/PAGADOR: " & MATRIZ1(10) & " - " & MATRIZ1(14)
                                    If UBound(MATRIZ2) > 1 Then
                                          If CONTROLCX(1) = "E-mail Pagamento" Then
                                                Cells(1, 1).Value = "AUTORIZAÇÕES - " & MATRIZ1(10)
                                                Cells(4, 1).Value = "Por gentileza, efetue os seguintes pagamentos:"
                                          End If
                                          If CONTROLCX(1) = "E-mail Cancelamento" Then
                                                Cells(1, 1).Value = "CANCELAMENTO DE AUTORIZAÇÕES - " & MATRIZ1(10)
                                                Cells(4, 1).Value = "Por gentileza, 'CANCELE' os seguintes pagamentos:"
                                          End If
                                          Cells(7, 1).Value = "----------------- FAVORECIDOS --------------------------------------------------------------------"
                                    Else
                                          If CONTROLCX(1) = "E-mail Pagamento" Then
                                                Cells(1, 1).Value = "AUTORIZAÇÃO - " & MATRIZ1(10)
                                                Cells(4, 1).Value = "Por gentileza, efetue o seguinte pagamento:"
                                          End If
                                          If CONTROLCX(1) = "E-mail Cancelamento" Then
                                                Cells(1, 1).Value = "CANCELAMENTO DE AUTORIZAÇÃO - " & MATRIZ1(10)
                                                Cells(4, 1).Value = "Por gentileza, 'CANCELE' o seguinte pagamento:"
                                          End If
                                          Cells(7, 1).Value = "----------------- FAVORECIDO ---------------------------------------------------------------------"
                                    End If
                              End If
                              
                              If MATRIZ1(18) = "BOLETO" Then
                              Cells(j + 7, 1) = MATRIZ1(25) & " | " & MATRIZ1(16) & " - " & MATRIZ1(18) & " - R$ " & VBA.Format(MATRIZ1(26), "#,##0.00")
                              Else
                              Cells(j + 7, 1) = MATRIZ1(25) & " | " & MATRIZ1(16) & " - " & MATRIZ1(18) & " - " & MATRIZ1(19) & " - R$ " & VBA.Format(MATRIZ1(26), "#,##0.00")
                              End If
                        Next j
                        j = j + 6
                        j = j + 1: Cells(j, 1).Value = "------------------------------------------------------------------------------------------------------"
                        j = j + 1: Cells(j, 1) = "."
                        j = j + 1: Cells(j, 1) = "."
                        j = j + 1: Cells(j, 1).Value = "POR FAVOR, CONFIRME O RECEBIMENTO!"
                        j = j + 1: Cells(j, 1) = "."
                        j = j + 1: Cells(j, 1) = "."
                        j = j + 1: Cells(j, 1).Value = "Atenciosamente..."
                        j = j + 1: Cells(j, 1).Value = ROTULOS(3)
                        Columns("A:A").EntireColumn.AutoFit
                        
                        
                        Set BASE = Range(Cells(1, 1), Cells(j, 1))
                        v = "VERTICAL": l = 1: a = 10: TITULO = ""
                        
                        TRAVACOMANDO = "SIM"
                        ReDim MATRIZ1(9)
                        For i = 1 To 9
                        MATRIZ1(i) = 0.2
                        Next
                        
                        dataduh_impressao_simples BASE, TITULO, v, l, a
                        dataduh_ferramenta_pdf
                  Application.ScreenUpdating = False
      Planilha1.Select: Planilha1.Select
      End Function
      Function dataduh_financeiro_autorizacoeslayout()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Autorizações" Then Exit Function
            'Layout---------------------------------
            Application.ScreenUpdating = False
                  '---------------------------------
                  Planilha3.Select: Planilha3.Select
                  With Cells
                        .Clear
                        .Font.Name = "Swis721 BT"
                        .Font.Size = 14
                        .RowHeight = 20
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlLeft
                  End With
                  a = Array(80, 30, 20)
                  For i = 1 To 3
                        Columns(i).ColumnWidth = a(i)
                  Next i
                  '---------------------------------
                  Planilha2.Select: Planilha2.Select
                    With Cells
                        .Clear
                        .Font.Name = "Swis721 BT"
                        .Font.Size = 14
                        .RowHeight = 20
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlLeft
                  End With
                  For i = 1 To 3
                        Columns(i).ColumnWidth = a(i)
                  Next i
                  
                  '---------------------------------
                  Columns(2).HorizontalAlignment = xlRight
                  Columns(3).HorizontalAlignment = xlRight
                  'Columns(3).Font.Size = 17
                  Columns(3).NumberFormat = "#,##0.00"
                  Rows(2).NumberFormat = "dd/mm/yy;@"
                  Rows(2).Font.Size = 20
                  Rows("2:4").RowHeight = 21
                  Rows(4).Font.Bold = False
                  Rows("2:5").Font.Bold = True
                  Rows("22:26").Font.Bold = True
                  
                  Rows(26).Font.Size = 11
                  Rows(26).VerticalAlignment = xlTop
                                     
                  a = Array(2, 5, 21, 25)
                  For i = 1 To 4
                  Rows(a(i)).Borders(xlEdgeBottom).Weight = xlThin
                  Next i
                                    
            Application.ScreenUpdating = False
      End Function
      Function dataduh_financeiro_autorizacoes()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Autorizações" Then Exit Function
            
            CAMADAS(3).Visible = False
            Banco_Financeiro Conecta_Financeiro: RS.Open "DELETE * FROM MEMORIA", Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
          
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) <> 0 Then
                  CAIXAS(2) = "Resumir"
                  dataduh_financeiro_resumo
                  CAIXAS(2) = "Autorizações"
            End If

            dataduh_ferramenta_listatotal1
            If UBound(MATRIZ2) > 200 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Muitos registros !"
                  Exit Function
            End If
            
            dataduh_financeiro_resumocomvalor
            
            If MATRIZ2(1) = "VAZIO" Then
            
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Dados Inválidos !"
            
            Exit Function
            End If
            
            For k = 1 To UBound(MATRIZ2)
                  CAMPO = "ID": CRITERIO = MATRIZ2(k)
                  dataduh_buscador_titulo

                  SQL = "MEMORIA"
                  Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
                  RS.AddNew

                  For i = 1 To 39
                        RS(i) = MATRIZ1(i + 1)
                  Next

                  RS.Update
                  RS.Close: Conecta_Financeiro.Close
            Next k

            dataduh_financeiro_autorizacoeslayout
                        
            CODIGO = "Vencimento"
            dataduh_financeiro_abastececombomemoria
            ReDim VENCIMENTO(UBound(MATRIZ1))
            For i = 1 To (UBound(MATRIZ1))
                 VENCIMENTO(i) = MATRIZ1(i)
            Next i
            
            CODIGO = "Pgid"
            dataduh_financeiro_abastececombomemoria
            ReDim PGID(UBound(MATRIZ1))
            For i = 1 To (UBound(MATRIZ1))
                 PGID(i) = MATRIZ1(i)
            Next i
            
            CODIGO = "Banco"
            dataduh_financeiro_abastececombomemoria
            On Error Resume Next
            ReDim BANCOS(UBound(MATRIZ1))
            For i = 1 To (UBound(MATRIZ1))
                 BANCOS(i) = MATRIZ1(i)
            Next i

            ReDim TIPO(2): TIPO(1) = "CORRENTE": TIPO(2) = "POUPANÇA":

            ReDim PAGINA(1): PAGINA(1) = 1
            'deposito>banco>tipo--------------------------
            For v = 1 To UBound(VENCIMENTO)
                DATA1 = VENCIMENTO(v): DATA2 = VENCIMENTO(v)
                dataduh_ferramenta_invertedata
                
                For p = 1 To UBound(PGID)
                        For b = 1 To UBound(BANCOS)
                          For t = 1 To UBound(TIPO)
                            SQL = "SELECT * FROM MEMORIA WHERE Vencimento BETWEEN #" & DATA1 & "# AND #" & DATA1 & "#"
                            SQL = SQL & " AND Classe LIKE 'CONTÁBIL%'"
                            SQL = SQL & " AND Pgid LIKE '" & PGID(p) & "%'"
                            SQL = SQL & " AND Cobranca LIKE 'DEPOSITO%'"
                            SQL = SQL & " AND Banco LIKE '" & BANCOS(b) & "%'"
                            SQL = SQL & " AND Tipo LIKE '" & TIPO(t) & "%'"
                            SQL = SQL & " ORDER BY Nome"
                            
                            Banco_Financeiro Conecta_Financeiro
                            RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
                            a = 0: a = CDbl(RS.RecordCount)
                            
                            If a <> 0 Then
                            
                              If a Mod 18 <> 0 Then
                                  a = Round((a / 16), 0) + 1
                              Else
                                  a = (a / 16)
                              End If
                              
                              For k = 1 To a
                                  Planilha2.Select: Planilha2.Select
                                  Cells.ClearContents
                                  Set BASE = Range(Cells(6, 1), Cells(21, 3))
                                  BASE.Borders(xlTop).LineStyle = xlNone
                                  BASE.Borders(xlBottom).LineStyle = xlNone
                                  
                                  '--------------------------------
                                  Cells(3, 1) = "AUTORIZAÇÃO DE DEPÓSITO EM CONTA " & RS(21) & " NO BANCO " & RS(18)
                                  If RS(18) = "BANCO DO BRASIL" Then Cells(3, 1) = "AUTORIZAÇÃO DE DEPÓSITO EM CONTA " & RS(21) & " NO " & RS(18)
                                  Cells(5, 1) = "Nome | CPF/CNPJ": Cells(5, 2) = "Agência | Conta"
                                  Cells(5, 3) = "Valor": Cells(22, 2).Value = "Total"
                                  Cells(26, 3).Value = "FINANCEIRO": Cells(2, 3).Value = RS(24)
                                  Cells(2, 1).Value = RS(9) & " - " & RS(13)
                                  Cells(26, 1).Value = RS(9) & " - " & RS(13)
                                  '--------------------------------
  
                                  For i = 1 To 16
                                      On Error Resume Next
                                      If RS(25) <> 0 Then
                                          On Error Resume Next
                                          If RS(0) = "" Then Exit For
                                          Cells(i + 5, 1) = Left(RS(15), 31) & " " & RS(16)
                                          On Error Resume Next
                                          Cells(i + 5, 2) = RS(19) & " / " & RS(20)
                                          On Error Resume Next
                                          Cells(i + 5, 3) = RS(25)
                                          Cells(22, 3) = CDbl(Cells(22, 3)) + CDbl(RS(25))
                                          If Cells(i + 5, 3) > 0 Then Rows(i + 5).Borders(xlEdgeBottom).Weight = xlHairline
                                      End If
                                      RS.MoveNext
                                  Next i
                                                                  
                                  If Cells(2, 1) <> "" Or Cells(6, 1) <> "" Then
                                  Rows("1:27").Copy: Rows("28:54").Select: ActiveSheet.Paste
                                  Rows("1:54").Copy:
                                  
                                  linha = (PAGINA(1) * 54) - 53 & ":" & (PAGINA(1) * 54)
                                  PAGINA(1) = PAGINA(1) + 1
                                  
                                  Planilha3.Select: Planilha3.Select
                                  Rows(linha).Select: ActiveSheet.Paste
                                  End If
                          Next k
                        End If
                        RS.Close: Conecta_Financeiro.Close
                      Next t
                    Next b
                Next p
            Next v
                      

            'pix--------------------------
            For v = 1 To UBound(VENCIMENTO)
                DATA1 = VENCIMENTO(v): DATA2 = VENCIMENTO(v)
                dataduh_ferramenta_invertedata
                
                For p = 1 To UBound(PGID)
                            SQL = "SELECT * FROM MEMORIA WHERE Vencimento BETWEEN #" & DATA1 & "# AND #" & DATA1 & "#"
                            SQL = SQL & " AND Classe LIKE 'CONTÁBIL%'"
                            SQL = SQL & " AND Pgid LIKE '" & PGID(p) & "%'"
                            SQL = SQL & " AND Cobranca LIKE 'PIX%'"
                            SQL = SQL & " ORDER BY Nome"
                            
                            Banco_Financeiro Conecta_Financeiro
                            RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
                            a = 0: a = CDbl(RS.RecordCount)
                            
                            If a <> 0 Then
                            
                              If a Mod 18 <> 0 Then
                                  a = Round((a / 16), 0) + 1
                              Else
                                  a = (a / 16)
                              End If
                              
                              For k = 1 To a
                                  Planilha2.Select: Planilha2.Select
                                  Cells.ClearContents
                                  Set BASE = Range(Cells(6, 1), Cells(21, 3))
                                  BASE.Borders(xlTop).LineStyle = xlNone
                                  BASE.Borders(xlBottom).LineStyle = xlNone
                                  
                                  '--------------------------------
                                  Cells(3, 1) = "AUTORIZAÇÃO DE PAGAMENTO VIA PIX"
                                  Cells(5, 1) = "Nome | CPF/CNPJ": Cells(5, 2) = "Chave Pix"
                                  Cells(5, 3) = "Valor": Cells(22, 2).Value = "Total"
                                  Cells(26, 3).Value = "FINANCEIRO": Cells(2, 3).Value = RS(24)
                                  Cells(2, 1).Value = RS(9) & " - " & RS(13)
                                  Cells(26, 1).Value = RS(9) & " - " & RS(13)
                                  '--------------------------------
  
                                  For i = 1 To 16
                                      On Error Resume Next
                                      If RS(25) <> 0 Then
                                          On Error Resume Next
                                          If RS(0) = "" Then Exit For
                                          Cells(i + 5, 1) = Left(RS(15), 31) & " " & RS(16)
                                          On Error Resume Next
                                          Cells(i + 5, 2) = RS(22)
                                          If Cells(i + 5, 2) = 0 Then Cells(i + 5, 2) = ""
                                          On Error Resume Next
                                          Cells(i + 5, 3) = RS(25)
                                          Cells(22, 3) = CDbl(Cells(22, 3)) + CDbl(RS(25))
                                          If Cells(i + 5, 3) > 0 Then Rows(i + 5).Borders(xlEdgeBottom).Weight = xlHairline
                                      End If
                                      RS.MoveNext
                                  Next i
                                                                  
                                  If Cells(2, 1) <> "" Or Cells(6, 1) <> "" Then
                                  Rows("1:27").Copy: Rows("28:54").Select: ActiveSheet.Paste
                                  Rows("1:54").Copy:
                                  
                                  linha = (PAGINA(1) * 54) - 53 & ":" & (PAGINA(1) * 54)
                                  PAGINA(1) = PAGINA(1) + 1
                                  
                                  Planilha3.Select: Planilha3.Select
                                  Rows(linha).Select: ActiveSheet.Paste
                                  End If
                          Next k
                        End If
                        RS.Close: Conecta_Financeiro.Close
                Next p
            Next v

            'boleto--------------------------
            For v = 1 To UBound(VENCIMENTO)
                DATA1 = VENCIMENTO(v): DATA2 = VENCIMENTO(v)
                dataduh_ferramenta_invertedata
                
                For p = 1 To UBound(PGID)
                            SQL = "SELECT * FROM MEMORIA WHERE Vencimento BETWEEN #" & DATA1 & "# AND #" & DATA1 & "#"
                            SQL = SQL & " AND Classe LIKE 'CONTÁBIL%'"
                            SQL = SQL & " AND Pgid LIKE '" & PGID(p) & "%'"
                            SQL = SQL & " AND Cobranca LIKE 'BOLETO%'"
                            SQL = SQL & " ORDER BY Nome"
                            
                            Banco_Financeiro Conecta_Financeiro
                            RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
                            a = 0: a = CDbl(RS.RecordCount)
                            
                            If a <> 0 Then
                            
                              If a Mod 18 <> 0 Then
                                  a = Round((a / 16), 0) + 1
                              Else
                                  a = (a / 16)
                              End If
                              
                              For k = 1 To a
                                  Planilha2.Select: Planilha2.Select
                                  Cells.ClearContents
                                  Set BASE = Range(Cells(6, 1), Cells(21, 3))
                                  BASE.Borders(xlTop).LineStyle = xlNone
                                  BASE.Borders(xlBottom).LineStyle = xlNone
                                  
                                  '--------------------------------
                                  Cells(3, 1) = "AUTORIZAÇÃO DE PAGAMENTO DE DUPLICATAS"
                                  Cells(5, 1) = "Nome | CPF/CNPJ":
                                  Cells(5, 3) = "Valor": Cells(22, 2).Value = "Total"
                                  Cells(26, 3).Value = "FINANCEIRO": Cells(2, 3).Value = RS(24)
                                  Cells(2, 1).Value = RS(9) & " - " & RS(13)
                                  Cells(26, 1).Value = RS(9) & " - " & RS(13)
                                  '--------------------------------
  
                                  For i = 1 To 16
                                      On Error Resume Next
                                      If RS(25) <> 0 Then
                                          On Error Resume Next
                                          If RS(0) = "" Then Exit For
                                          Cells(i + 5, 1) = RS(15)
                                          On Error Resume Next
                                          Cells(i + 5, 3) = RS(25)
                                          Cells(22, 3) = CDbl(Cells(22, 3)) + CDbl(RS(25))
                                          If Cells(i + 5, 3) > 0 Then Rows(i + 5).Borders(xlEdgeBottom).Weight = xlHairline
                                      End If
                                      RS.MoveNext
                                  Next i
                                                                  
                                  If Cells(2, 1) <> "" Or Cells(6, 1) <> "" Then
                                  Rows("1:27").Copy: Rows("28:54").Select: ActiveSheet.Paste
                                  Rows("1:54").Copy:
                                  
                                  linha = (PAGINA(1) * 54) - 53 & ":" & (PAGINA(1) * 54)
                                  PAGINA(1) = PAGINA(1) + 1
                                  
                                  Planilha3.Select: Planilha3.Select
                                  Rows(linha).Select: ActiveSheet.Paste
                                  End If
                          Next k
                        End If
                        RS.Close: Conecta_Financeiro.Close
                Next p
            Next v

            Planilha3.Select: Planilha3.Select
            If Cells(2, 1) = "" Or Cells(6, 1) = "" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Nada a visualizar !"
                  Exit Function
            End If
            
            PAGINA(1) = PAGINA(1) - 1
            Set BASE = Range(Cells(1, 1), Cells(PAGINA(1) * 54, 3))
            v = "VERTICAL": l = 1: a = PAGINA(1): TITULO = ""
            
            TRAVACOMANDO = "SIM"
            ReDim MATRIZ1(9)
            For i = 1 To 9
            MATRIZ1(i) = 0.5
            Next
            
            dataduh_impressao_simples BASE, TITULO, v, l, a
            dataduh_ferramenta_pdf
            Planilha1.Select: Planilha1.Select
      End Function
      Function dataduh_financeiro_fixarpagamento()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Classificar Pagamento" Then Exit Function
            
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um registro !"
                  Exit Function
            End If
            
            dataduh_formulario_filtros
                        
            LISTAS(2).Visible = False
            
            CAMPO = "ID": CRITERIO = MATRIZ2(1)
            dataduh_buscador_titulo
            
            CONTROLRT(2) = MATRIZ1(1)
            CONTROLRT(3) = "Classificar como ..."
            CONTROLRT(4) = "Dia (1 à 28)"
            
            CONTROLCX(3).AddItem "FIXA"
            CONTROLCX(3).AddItem "FAVORITA"
            
            CONTROLCX(4) = Day(MATRIZ1(25))
            If CONTROLCX(4) > 28 Then CONTROLCX(4) = 28
            
            
            With CAIXAS(3)
                .Clear
                .AddItem "Criar"
                .AddItem "Sair"
            End With
            
            For i = 1 To 6
                  If i < 2 Or i > 3 Then
                        CONTROLRT(i).Visible = False
                        CONTROLCX(i).Visible = False
                        CONTROLCX(2).Visible = False
                  End If
            Next i
      
      End Function
      Function dataduh_financeiro_fixarpagamentoexe()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Classificar Pagamento" Then Exit Function
            If CAIXAS(3) <> "Criar" Then Exit Function
            If CONTROLCX(3) <> "FIXA" Then Exit Function
            
            If CONTROLCX(4) <> "" Then
            On Error Resume Next
            If Not VBA.IsNumeric(CONTROLCX(4).Text) = True Then CONTROLCX(4).Text = ""
                  If IsNumeric(CONTROLCX(4)) = True Then CONTROLCX(4) = Format(CONTROLCX(4), "#")
                  If CONTROLCX(4) < 1 Then CONTROLCX(4) = ""
                  If CONTROLCX(4) > 28 Then CONTROLCX(4) = ""
            End If
            
            If CONTROLCX(4) = "" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Digite o dia do Vencimento !"
                  Exit Function
            End If
            
            X = 1: dataduh_ferramenta_operacao
            
            CAMPO = "ID": CRITERIO = CONTROLRT(2)
            dataduh_buscador_titulo
            
            SQL = "CAIXA"
            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
            RS.AddNew
            
            For i = 1 To 39
                  RS(i) = MATRIZ1(i + 1)
                  If i = 1 Then RS(i) = ROTULOS(3)
                  If i = 2 Then RS(i) = CDate(Date)
                  If i = 3 Then RS(i) = Format(Now, "hh:mm")
                  If i = 5 Then RS(i) = "FIXA"
                  If i = 23 Then RS(i) = "DIA " & CONTROLCX(4)
                  If i = 24 Then RS(i) = CDate(Date)
                  If i = 25 Then RS(i) = 0
                  If i = 28 Then RS(i) = 0
                  If i = 30 Then RS(i) = "FIXA-" & MATRIZ1(i + 1)
                  If i = 37 Then RS(i) = SERIAL
                  If i = 38 Then RS(i) = CONTROLCX(4)
            Next i
            
            RS.Update
            RS.Close: Conecta_Financeiro.Close
            
            CAMADAS(3).Visible = False
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            
            MEMFILTRO(1, 1) = "Data"
            MEMFILTRO(1, 2) = CDate(Date)
            MEMFILTRO(1, 3) = CDate(Date)
            MEMFILTRO(1, 4) = "Operacao"
            MEMFILTRO(1, 5) = SERIAL
            
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            ROTULOS(24) = "Operação: " & SERIAL: ROTULOS(25) = "Pagamento fixado para dia " & CONTROLCX(4)
      End Function
      Function dataduh_financeiro_favoritarpagamentoexe()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Classificar Pagamento" Then Exit Function
            If CAIXAS(3) <> "Criar" Then Exit Function
            If CONTROLCX(3) <> "FAVORITA" Then Exit Function
            
            X = 1: dataduh_ferramenta_operacao
            
            CAMPO = "ID": CRITERIO = CONTROLRT(2)
            dataduh_buscador_titulo
            
            SQL = "CAIXA"
            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
            RS.AddNew
            
            For i = 1 To 39
                  RS(i) = MATRIZ1(i + 1)
                  If i = 1 Then RS(i) = ROTULOS(3)
                  If i = 2 Then RS(i) = CDate(Date)
                  If i = 3 Then RS(i) = Format(Now, "hh:mm")
                  If i = 5 Then RS(i) = "FAVORITA"
                  If i = 23 Then RS(i) = "FAV "
                  If i = 24 Then RS(i) = CDate(Date)
                  If i = 25 Then RS(i) = 0
                  If i = 28 Then RS(i) = 0
                  If i = 30 Then RS(i) = "FAV-" & MATRIZ1(i + 1)
                  If i = 37 Then RS(i) = SERIAL
            Next i
            
            RS.Update
            RS.Close: Conecta_Financeiro.Close
            
            CAMADAS(3).Visible = False
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            
            MEMFILTRO(1, 1) = "Data"
            MEMFILTRO(1, 2) = CDate(Date)
            MEMFILTRO(1, 3) = CDate(Date)
            MEMFILTRO(1, 4) = "Operacao"
            MEMFILTRO(1, 5) = SERIAL
            
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            ROTULOS(24) = "Operação: " & SERIAL: ROTULOS(25) = "Classificada como Favorita !"
      End Function
      Function dataduh_financeiro_novoregistrofavorito()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro Favorito" Then Exit Function
                        
            X = 1: dataduh_ferramenta_operacao
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            MEMFILTRO(1, 1) = "Data"
            MEMFILTRO(1, 2) = CDate("01/01/2000")
            MEMFILTRO(1, 3) = CDate("01/01/2100")
            MEMFILTRO(1, 4) = "Classe"
            MEMFILTRO(1, 5) = "FAVORITA"
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
            
            CAMADAS(4).Visible = False
            
            CAIXAS(2) = "Novo Registro Favorito"
            
            dataduh_formulario_filtros
            
            LISTAS(2).Visible = False
            
            With CAIXAS(3)
                .Clear
                .AddItem "Criar"
                .AddItem "Visualizar"
                .AddItem "Sair"
            End With
            
            CONTROLRT(1) = "Duplicata"
            CONTROLRT(2) = "Vencimento"
            CONTROLRT(3) = "Valor"
            CONTROLRT(4) = "Histórico"
            
            For i = 5 To 6
                  CONTROLRT(i).Visible = False
                  CONTROLCX(i).Visible = False
            Next i
 
            X = 1: dataduh_ferramenta_operacao
      End Function
      Function dataduh_financeiro_novoregistrofavoritoexe()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro Favorito" Then Exit Function
            If CAIXAS(3) <> "Criar" Then Exit Function
            
            dataduh_ferramenta_listaselect1
            If MATRIZ2(1) = 0 Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Selecione um registro !"
                  Exit Function
            End If
            
            If CONTROLCX(2) = "" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Digite o vencimento !"
                  Exit Function
            End If
                        
            CAMPO = "ID": CRITERIO = MATRIZ2(1)
            dataduh_buscador_titulo
            
            SQL = "CAIXA"
            Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
            RS.AddNew
            
            For i = 1 To 39
                  RS(i) = MATRIZ1(i + 1)
                  If i = 1 Then RS(i) = ROTULOS(3)
                  If i = 2 Then RS(i) = CDate(Date)
                  If i = 3 Then RS(i) = Format(Now, "hh:mm")
                  If i = 5 Then RS(i) = "CONTÁBIL"
                  
                  If CONTROLCX(1) <> "" Then
                        If i = 23 Then RS(i) = CONTROLCX(1)
                  End If
                  
                  If i = 24 Then RS(i) = CONTROLCX(2)
                  
                  If CONTROLCX(3) <> "" Then
                        If i = 25 Then RS(i) = CONTROLCX(3)
                  End If
                  
                  If i = 28 Then RS(i) = 0
                  
                  If CONTROLCX(4) <> "" Then
                        If i = 30 Then RS(i) = CONTROLCX(4)
                  End If
                  
                  If i = 37 Then RS(i) = SERIAL
            Next
            RS.Update
            RS.Close: Conecta_Financeiro.Close
            
            For i = 1 To 4
                  CONTROLCX(i) = ""
            Next i
            
            If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
            ROTULOS(24) = "Aviso !": ROTULOS(25) = "Registro criado !"
      End Function
      Function dataduh_financeiro_novoregistrofavoritover()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Novo Registro Favorito" Then Exit Function
            If CAIXAS(3) <> "Visualizar" Then Exit Function
            
            CAMADAS(3).Visible = False
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            
            MEMFILTRO(1, 1) = "Data"
            MEMFILTRO(1, 2) = CDate("01/01/2000")
            MEMFILTRO(1, 3) = CDate("01/01/2100")
            MEMFILTRO(1, 4) = "Operacao"
            MEMFILTRO(1, 5) = SERIAL
            
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
      End Function
      Function dataduh_financeiro_resumosemana()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Resumir Semana" Then Exit Function
            
            X = 2: dataduh_ferramenta_semanadata
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            
            MEMFILTRO(1, 1) = "Vencimento"
            MEMFILTRO(1, 2) = DATA1
            MEMFILTRO(1, 3) = DATA2
            MEMFILTRO(1, 4) = ""
            MEMFILTRO(1, 5) = ""
            
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
      End Function
      Function dataduh_financeiro_resumomes()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Resumir Mês" Then Exit Function
            
            X = 2: dataduh_ferramenta_mesinteiro
            
            Nome = CAIXAS(1)
            ReDim MEMFILTRO(1, 5)
            
            MEMFILTRO(1, 1) = "Vencimento"
            MEMFILTRO(1, 2) = DATA1
            MEMFILTRO(1, 3) = DATA2
            MEMFILTRO(1, 4) = ""
            MEMFILTRO(1, 5) = ""
            
            CAIXAS(1) = Nome: CAIXAS(2) = "Filtro"
            dataduh_financeiro_filtrarexe
            dataduh_financeiro_filtrar2
      End Function
      Function dataduh_financeiro_resumocomvalor()
                  ReDim MATRIZ3(DATADUH.ListBox1.ListCount)
                  a = 1
                  For i = 0 To DATADUH.ListBox1.ListCount - 1
                        If CDbl(DATADUH.ListBox1.List(i, 5)) <> 0 Then
                              MATRIZ3(a) = DATADUH.ListBox1.List(i, 0)
                              a = a + 1
                        End If
                  Next i
                    
                  If a > 1 Then
                        ReDim MATRIZ2(a - 1)
                        For i = 1 To UBound(MATRIZ2)
                              MATRIZ2(i) = MATRIZ3(i)
                        Next i
                  Else
                        ReDim MATRIZ2(1): MATRIZ2(1) = "VAZIO"
                  End If
      End Function


      Function dataduh_financeiro_calculos()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Calculos" Then Exit Function
            
            dataduh_formulario_filtros
                        
            CONTROLRT(1) = "Tipo de Cálculo"
            
            With CONTROLCX(1)
                  .Clear
                  .AddItem "FOLHA DE SALÁRIOS"
            End With
            
            With CAIXAS(3)
                  .Clear
                  .AddItem "Calcular"
                  .AddItem "Sair"
            End With
            
            For i = 2 To 6
                  CONTROLRT(i).Visible = False
                  CONTROLCX(i).Visible = False
            Next
            
            LISTAS(2).Visible = False
            
            ReDim COMPARADOR(6)
            For i = 1 To 6
                  COMPARADOR(i) = CONTROLCX(i)
            Next
      End Function
      Function dataduh_financeiro_calculosmasc()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Calculos" Then Exit Function
            
            For i = 2 To 6
                  CONTROLRT(i).Visible = False
                  CONTROLCX(i).Visible = False
            Next
            
            If CONTROLCX(1) = "FOLHA DE SALÁRIOS" Then
                  CONTROLRT(2) = "Data"
                  CONTROLRT(2).Visible = True
                  CONTROLCX(2).Visible = True
            End If
            

      End Function
      Function dataduh_financeiro_gerarcontafixa()
            ReDim MEMFILTRO(1, 5)
            MEMFILTRO(1, 1) = "Vencimento"
            MEMFILTRO(1, 2) = "01/01/1950"
            MEMFILTRO(1, 3) = "01/01/2100"
            MEMFILTRO(1, 4) = "CLASSE"
            MEMFILTRO(1, 5) = "FIXA"
            
            dataduh_financeiro_filtrarexe
                     
            CODIGO = CDate("01/" & Month(CDate(Date)) & "/" & Year(CDate(Date))) + 33
            CODIGO = CDate("01/" & Month(CDate(CODIGO)) & "/" & Year(CDate(CODIGO)))
      
            'Gravando novo registro---------------------------
            For i = 1 To UBound(VIRTUALTAB)
                  If VIRTUALTAB(i, 39) <> CODIGO Then
                        SQL = "CAIXA": Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic
                        RS.AddNew
                        For j = 1 To 39
                              If VIRTUALTAB(i, j + 1) = "" Then VIRTUALTAB(i, j + 1) = 0
                              RS(j) = VIRTUALTAB(i, j + 1)
                              If j + 1 = 2 Then RS(j) = "EDUARDO" 'ROTULOS(3)
                              If j + 1 = 3 Then RS(j) = CDate(Date)
                              If j + 1 = 4 Then RS(j) = Format(Now, "hh:mm")
                              If j + 1 = 6 Then RS(j) = "CONTÁBIL"
                              If j + 1 = 24 Then RS(j) = VIRTUALTAB(i, 24)
                              If j + 1 = 25 Then RS(j) = CDate(VIRTUALTAB(i, 33) & "/" & Month(CDate(CODIGO)) & "/" & Year(CDate(CODIGO)))
                              
                              If j = 24 Then
                              If Weekday(RS(24)) = 7 Then RS(24) = RS(24) - 1
                              If Weekday(RS(24)) = 1 Then RS(24) = RS(24) - 2
                              End If
                              
                              If j + 1 = 39 Then RS(j) = "0"
                        Next
                        RS.Update: RS.Close: Conecta_Financeiro.Close
                                                
                        SQL = "UPDATE CAIXA SET Campo_38 = '" & CODIGO & "' WHERE ID = " & VIRTUALTAB(i, 1)
                        Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                  End If
            Next i
      End Function
      Function dataduh_financeiro_diagnostico()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Relatório de Erros" Then Exit Function
            CAMADAS(3).Visible = False
            dataduh_relatorio_layout
            Application.ScreenUpdating = False
                  Cells.Font.Name = "Courier New"
                  Cells.HorizontalAlignment = xlLeft
                  
                  a = Array(8, 13, 80, 13.5, 13.5)
                  For i = 1 To 5
                        Columns(i).ColumnWidth = a(i)
                  Next i
                  
                  Cells(1, 1) = "Relatório de Erros"
                  Cells(2, 1) = "DATADUH - Módulo Financeiro"
                  Cells(3, 1) = "."
                  Cells(4, 1) = "Id"
                  Cells(4, 2) = "Vencimento"
                  Cells(4, 3) = "Nome"
                  Cells(4, 4) = "Cobrança"
                  Cells(4, 5) = "Diagnóstico"
                       
                  ReDim MEMFILTRO(3, 5)
                  MEMFILTRO(1, 1) = "Vencimento"
                  
                  MEMFILTRO(1, 4) = "CLASSE"
                  MEMFILTRO(1, 5) = "CONTÁBIL"
                  MEMFILTRO(2, 4) = "Cobranca"
                  MEMFILTRO(2, 5) = "DEPOSITO"
                  MEMFILTRO(3, 4) = "Tipo"
                  MEMFILTRO(3, 5) = "FATURA"
                                                      
                  For s = 1 To 2
                        MEMFILTRO(1, 2) = CDate(Date)
                        MEMFILTRO(1, 3) = "01/01/2100"
                        dataduh_financeiro_filtrarexe
                        CODIGO = "CORRENTE"
                        If VIRTUALTAB(1, 1) <> "VAZIO" Then
                              For i = 1 To UBound(VIRTUALTAB)
                                    SQL = "UPDATE CAIXA SET Tipo = '" & CODIGO & "' WHERE ID = " & VIRTUALTAB(i, 1)
                                    Banco_Financeiro Conecta_Financeiro: RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockPessimistic: Conecta_Financeiro.Close
                              Next i
                        End If
                        MEMFILTRO(3, 5) = 0
                  Next s
                  
                  
                  l = 4
                  t = Array("Banco", "Agencia", "Conta", "Pix")
                  For s = 1 To 4
                  MEMFILTRO(3, 4) = t(s)
                  MEMFILTRO(1, 2) = CDate(Date)
                  MEMFILTRO(1, 3) = "01/01/2100"
                  dataduh_financeiro_filtrarexe
                        If VIRTUALTAB(1, 1) <> "VAZIO" Then
                                    v = 0
                                    For i = 1 To UBound(VIRTUALTAB)
                                          If s < 4 Then
                                                If Len(VIRTUALTAB(i, s + 18)) = 1 Then
                                                      l = l + 1
                                                      Cells(l, 1) = VIRTUALTAB(i, 1)
                                                      Cells(l, 2) = VIRTUALTAB(i, 25)
                                                      Cells(l, 3) = VIRTUALTAB(i, 16)
                                                      Cells(l, 4) = MEMFILTRO(2, 5)
                                                      Cells(l, 5) = MEMFILTRO(3, 4) & " = 0"
                                                      v = v + 1
                                                End If
                                          Else
                                                If Len(VIRTUALTAB(i, s + 19)) = 1 Then
                                                      l = l + 1
                                                      Cells(l, 1) = VIRTUALTAB(i, 1)
                                                      Cells(l, 2) = VIRTUALTAB(i, 25)
                                                      Cells(l, 3) = VIRTUALTAB(i, 16)
                                                      Cells(l, 4) = MEMFILTRO(2, 5)
                                                      Cells(l, 5) = MEMFILTRO(3, 4) & " = 0"
                                                      v = v + 1
                                                End If
                                          End If
                                    Next i
                                    If v > 0 Then l = l + 1
                                    For i = 1 To 120
                                                       Cells(l, 1) = Cells(l, 1).Value & "-"
                                    Next i
                        End If
                  If s = 3 Then MEMFILTRO(2, 5) = "PIX"
                  Next s
            Application.ScreenUpdating = False
                        
            If Cells(6, 1) = "" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Aviso !": ROTULOS(25) = "Não há erros !"
                  Exit Function
            End If

            Set BASE = Range("A1").CurrentRegion
            Cells(3, 1) = ""
            
            v = "VERTICAL"
            l = 1: a = 100: TITULO = "$1:$4"
                
            ReDim MATRIZ1(9)
            For i = 1 To 9
                  MATRIZ1(i) = 0.5
            Next
                
            dataduh_impressao_simples BASE, TITULO, v, l, a
            dataduh_ferramenta_pdf

      End Function
      Function dataduh_financeiro_calculofolha()
            If CAIXAS(1) <> "FINANCEIRO" Then Exit Function
            If CAIXAS(2) <> "Calculos" Then Exit Function
            If CAIXAS(3) <> "Calcular" Then Exit Function
            If CONTROLCX(1) <> "FOLHA DE SALÁRIOS" Then Exit Function
                     
            If CONTROLCX(2) = "" Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Digite uma data !"
                  Exit Function
            End If
                              
            If CDbl(CDate(CONTROLCX(2))) > CDbl(Date) Then
                  If CAMADAS(4).Visible = False Then dataduh_parametro_aviso
                  ROTULOS(24) = "Atenção !": ROTULOS(25) = "Use uma data antiga !"
                  Exit Function
            End If
            
            DATA1 = CDate("01/" & Month(CONTROLCX(2)) & "/" & Year(CONTROLCX(2)))
            DATA2 = CDate(DATA1) + 32
            DATA2 = CDate("01/" & Month(DATA2) & "/" & Year(DATA2)) - 1
            
            dataduh_relatorio_layout
            Application.ScreenUpdating = False
                  With Cells
                        .Font.Size = 14
                        .RowHeight = 30
                        .ColumnWidth = 22
                        .HorizontalAlignment = xlLeft
                  End With
                  
                  Rows("1:1").Font.Size = 20
                  Rows("6:6").Font.Bold = True
                  Rows("12:12").Font.Bold = True
                  Rows("1:2").RowHeight = 22
                  
                  For i = 3 To 5
                        Columns(i).HorizontalAlignment = xlCenter
                  Next i
                  
                  Columns(2).HorizontalAlignment = xlRight
                  Columns(6).HorizontalAlignment = xlRight
                  Columns(2).Font.Bold = True
                  Columns(5).Font.Bold = True
                  
                  CAMPO = "Status": CRITERIO = "PRESIDENTE"
                  dataduh_buscador_contato
                  CRITERIO = MATRIZ1(1)
                  
                  Cells(1, 1) = MATRIZ1(7)
                  Cells(2, 1) = "RESUMO DA FOLHA DE PAGAMENTOS"
                  Cells(1, 6) = UCase(MonthName(Month(DATA1)))
                  Cells(2, 6) = Year(DATA1)
                  
                  ReDim VIRTUALTAB(7, 4)
                  VIRTUALTAB(1, 1) = "" '"Resumo"
                  VIRTUALTAB(2, 1) = "Salário"
                  VIRTUALTAB(3, 1) = "Férias"
                  VIRTUALTAB(4, 1) = "13º Salário"
                  VIRTUALTAB(5, 1) = "Rescisão"
                  VIRTUALTAB(6, 1) = "Bonus"
                  VIRTUALTAB(7, 1) = "Total"
                  VIRTUALTAB(1, 2) = "Depósito"
                  VIRTUALTAB(1, 3) = "Cheque"
                  VIRTUALTAB(1, 4) = "Total"
                  
                  dataduh_ferramenta_invertedata
                  For i = 1 To 2
                        SQL = "SELECT Contabil, SUM(VALOR) FROM CAIXA "
                        SQL = SQL & "WHERE Vencimento BETWEEN #" & DATA1 & "# AND #" & DATA2 & "#"
                        SQL = SQL & " AND PGID LIKE '" & CRITERIO & "%'"
                        If i = 1 Then SQL = SQL & " AND COBRANCA LIKE 'DEPOSITO%'"
                        If i = 2 Then SQL = SQL & " AND COBRANCA LIKE 'CHEQUE%'"
                        SQL = SQL & "GROUP BY Contabil"
                        
                        Banco_Financeiro Conecta_Financeiro
                        RS.Open SQL, Conecta_Financeiro, adOpenKeyset, adLockReadOnly
                        a = 0: a = CDbl(RS.RecordCount)
                        
                        If a <> 0 Then
                              For j = 1 To a
                                    If RS(0) = 401 Then VIRTUALTAB(2, i + 1) = RS(1)
                                    If RS(0) = 402 Then VIRTUALTAB(3, i + 1) = RS(1)
                                    If RS(0) = 403 Then VIRTUALTAB(4, i + 1) = RS(1)
                                    If RS(0) = 404 Then VIRTUALTAB(5, i + 1) = RS(1)
                                    If RS(0) = 406 Then VIRTUALTAB(6, i + 1) = RS(1)
                                    RS.MoveNext
                              Next j
                        End If
                        RS.Close: Conecta_Financeiro.Close
                  Next i
                  
                  For i = 2 To 6
                        For j = 2 To 3
                              VIRTUALTAB(i, 4) = VIRTUALTAB(i, 4) + VIRTUALTAB(i, j)
                        Next j
                        VIRTUALTAB(7, 2) = VIRTUALTAB(7, 2) + VIRTUALTAB(i, 2)
                        VIRTUALTAB(7, 3) = VIRTUALTAB(7, 3) + VIRTUALTAB(i, 3)
                        VIRTUALTAB(7, 4) = VIRTUALTAB(7, 4) + VIRTUALTAB(i, 4)
                  Next i
                  
                  For i = 2 To 6
                        For j = 2 To 4
                              If VIRTUALTAB(i, j) = 0 Then VIRTUALTAB(i, j) = ""
                        Next j
                  Next i
                  
                  Set BASE = Range(Cells(6, 2), Cells(12, 5))
                  BASE.Interior.Color = RGB(250, 254, 255)
                  Range(Cells(6, 2), Cells(6, 5)).Interior.Color = RGB(240, 244, 244)
                  Range(Cells(12, 2), Cells(12, 5)).Interior.Color = RGB(230, 234, 235)
                  Range(Cells(7, 2), Cells(11, 2)).Interior.Color = RGB(240, 240, 250)
                  Range(Cells(7, 5), Cells(11, 5)).Interior.Color = RGB(240, 240, 250)
                  Cells(12, 5).Interior.Color = RGB(255, 255, 50)
                  Cells(12, 5).Font.Size = 18
                  BASE.Value = VIRTUALTAB
                  
                  Set BASE = Range(Cells(1, 1), Cells(12, 6))
                  If Cells(1, 1) = "" Then Exit Function
                  
                  v = "VERTICAL"
                  l = 1: a = 1: TITULO = ""
                  
                  dataduh_impressao_simples BASE, TITULO, v, l, a
                  dataduh_ferramenta_pdf
            Application.ScreenUpdating = True
      End Function
