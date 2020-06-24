Public Class frmBegin
    Private Sub frmBegin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call txtCLS()
        Call fmtListView()
        Call showData()

        Call showDataList()

    End Sub

    Function chkMixQty(strDocNo As String) As Integer
        txtSQL = "Select  (BOM_RM_Values)as BOM_RM_Values "

        txtSQL = txtSQL & "From  BOMmastH "
        txtSQL = txtSQL & "Left Join BOMmastD "
        txtSQL = txtSQL & "On BOMmastH.BOM_No=BOMmastD.BOM_No "
        txtSQL = txtSQL & "Left Join BOM_RMmast "
        txtSQL = txtSQL & "On BOMMastD.BOM_RM_Code=BOM_RMmast.RM_Code "

        txtSQL = txtSQL & "Where "
        txtSQL = txtSQL & "( BOMmastD.BOM_RM_Code='08001' "
        txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08002' "
        txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08003' "
        txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08004' "
        txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08005') "
        txtSQL = txtSQL & "And BOMmastH.BOM_No='" & strDocNo & "' "

        txtSQL = txtSQL & "and BOM_RM_Values > 0 "

        Dim subds As New DataSet
        Dim subDa As SqlClient.SqlDataAdapter


        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subds, "Master")
        If subds.Tables("Master").Rows.Count > 0 Then
            Return subds.Tables("Master").Rows(0).Item("BOM_RM_Values")
        Else
            Return 0
        End If
    End Function
    Sub showData()

        txtSQL = "Select BOMmastH.BOM_Date,Right( BOMmastH.BOM_No,1)as KeyType, BOMmastH.BOM_No,isnull(Dtl_Num,0)as dtl_Num,"
        'txtSQL = txtSQL & "sum(case when (BOMmastD.BOM_RM_Code='08001' "
        'txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08002'"
        'txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08003'"
        'txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08004'"
        'txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='08005'"
        'txtSQL = txtSQL & ") then isnull((BOMmastD.BOM_RM_Values),0) else 0 end) as sum08005,"
        txtSQL = txtSQL & "sum(case when (BOMmastD.BOM_RM_Code='07001') then isnull((BOMmastD.BOM_RM_Values),0) else 0 end) as sum07001,"
        txtSQL = txtSQL & "sum(case when (BOMmastD.BOM_RM_Code='07002') then isnull((BOMmastD.BOM_RM_Values),0) else 0 end) as sum07002,"
        txtSQL = txtSQL & "sum(case when (BOMmastD.BOM_RM_Code='07003') then isnull((BOMmastD.BOM_RM_Values),0) else 0 end) as sum07003 "

        txtSQL = txtSQL & "From  BOMmastH "
        txtSQL = txtSQL & "Left Join BOMmastD "
        txtSQL = txtSQL & "On BOMmastH.BOM_No=BOMmastD.BOM_No "
        txtSQL = txtSQL & "Left Join BOM_RMmast "
        txtSQL = txtSQL & "On BOMMastD.BOM_RM_Code=BOM_RMmast.RM_Code "
        txtSQL = txtSQL & "Left Join TranDataD_E "
        txtSQL = txtSQL & "ON BOMmastH.BOM_No=TranDataD_E.Dtl_No "

        txtSQL = txtSQL & "Where (BOMmastD.BOM_RM_Code='07001' "
        txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='07002' "
        txtSQL = txtSQL & "or BOMmastD.BOM_RM_Code='07003' "

        txtSQL = txtSQL & ") "

        Dim strDateNow1 As String = Year(Now()) - 543 & "/" & Format(Month(Now()), "0#") & "/" & Format(Microsoft.VisualBasic.Day(Now()), "0#") ' & " 00:00:00"
        lbDate01.Text = Format(CDate(strDateNow1), "dd/MM/yyyy")

        txtSQL = txtSQL & "And (BOMmastH.BOM_Date='" & strDateNow1 & "' "
        Dim iDatediff As Integer = 0
        If CDate(Now).DayOfWeek = 6 Then
            idateDiff = 2
        Else
            iDatediff = 1
            'txtSQL = txtSQL & "or TranDataH_e.Trh_Date='" & DateAdd(DateInterval.Day, 1, CDate(strDateNow1)) & "' "
        End If
        Dim date01 As Date = DateAdd(DateInterval.Day, iDatediff, Now)
        Dim strDateNow2 As String = Year(date01) - 543 & "/" & Format(Month(date01), "0#") & "/" & Format(Microsoft.VisualBasic.Day(date01), "0#") ' & " 00:00:00"

        lbDate02.Text = Format(CDate(strDateNow2), "dd/MM/yyyy")
        txtSQL = txtSQL & "or BOMmastH.BOM_Date='" & strDateNow2 & "') "
        'txtSQL = txtSQL & "and BOM_RM_Values>0 "
        txtSQL = txtSQL & "Group by BOMmastH.BOM_Date,Right( BOMmastH.BOM_No,1), BOMmastH.BOM_No,Dtl_Num "
        txtSQL = txtSQL & "Order by Right( BOMmastH.BOM_No,1), BOMmastH.BOM_No "
        Dim subds As New DataSet
        Dim subDa As SqlClient.SqlDataAdapter


        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subds, "Master")
        Call txtCLS()
        With subds.Tables("Master")
            Dim qtyNo As Integer = 0
            'If .Rows.Count > 0 Then
            Dim strDate1 As String
            ' Dim strDate2 As String
            Dim anydata() As String
            Dim lvi As ListViewItem
            Dim strdocDate As String
            Dim strDocNo As String
            Dim dblMix10 As Double
            Dim dblMix22 As Double
            Dim dblMix60 As Double
            Dim dblDtlQty As Integer
            lsvDateList.Items.Clear()

            For i = 0 To subds.Tables("Master").Rows.Count - 1
                qtyNo = 0
                'strDocNo = ""
                dblMix10 = 0
                dblMix22 = 0
                dblMix60 = 0
                dblDtlQty = 0

                dblDtlQty = subds.Tables("Master").Rows(i).Item("dtl_Num")
                strDocNo = subds.Tables("Master").Rows(i).Item("BOM_No")
                strdocDate = subds.Tables("Master").Rows(i).Item("BOM_Date")
                qtyNo = chkMixQty(strDocNo)
                ' MsgBox(subds.Tables("Master").Rows(i).Item("BOM_No") & "-" & " qtyNo = " & qtyNo)
                If qtyNo = 0 Then
                    qtyNo = 1
                End If
                'If strDocNo = "27020FB" Then

                '    MsgBox("")

                'End If

                lbQty1.Text = qtyNo
                strDate1 = Format(DateAdd(DateInterval.Year, -543, (subds.Tables("Master").Rows(i).Item("BOM_Date"))), "yyyy/MM/dd")
                'strDate2 = Format(DateAdd(DateInterval.Year, -543, (subds.Tables("Master").Rows(i).Item("Trh_Date"))), "yyyy/MM/dd")
                If dblDtlQty > 0 Then
                    If strDateNow1 = strDate1 Then
                        If .Rows(i).Item("KeyType") = "A" Then

                            If .Rows(i).Item("sum07001") > 0 Then
                                dblMix10 = (.Rows(i).Item("sum07001") * qtyNo)
                                lb10sA_1.Text = Format(dblMix10 + CDbl(lb10sA_1.Text), "#,##0.00")
                            End If

                            If .Rows(i).Item("sum07002") > 0 Then

                                dblMix22 = (.Rows(i).Item("sum07002") * qtyNo)
                                lb22sA_1.Text = Format(dblMix22 + CDbl(lb22sA_1.Text), "#,##0.00")
                            End If

                            If .Rows(i).Item("sum07003") > 0 Then
                                dblMix60 = (.Rows(i).Item("sum07003") * qtyNo)
                                lb60sA_1.Text = Format(dblMix60 + CDbl(lb60sA_1.Text), "#,##0.00")
                            End If
                        Else
                            If .Rows(i).Item("sum07001") > 0 Then
                                dblMix10 = (.Rows(i).Item("sum07001") * qtyNo)
                                lb10sB_1.Text = Format(dblMix10 + CDbl(lb10sB_1.Text), "#,##0.00")
                            End If

                            If .Rows(i).Item("sum07002") > 0 Then
                                    dblMix22 = (.Rows(i).Item("sum07002") * qtyNo)
                                    lb22sB_1.Text = Format(dblMix22 + CDbl(lb22sB_1.Text), "#,##0.00")
                            End If
                            If .Rows(i).Item("sum07003") > 0 Then
                                dblMix60 = (.Rows(i).Item("sum07003") * qtyNo)
                                lb60sB_1.Text = Format(dblMix60 + CDbl(lb60sB_1.Text), "#,##0.00.00")
                            End If

                        End If

                    ElseIf strDateNow2 = strDate1 Then ' subds.Tables("Master").Rows(i).Item("Trh_Date") Then


                        If .Rows(i).Item("KeyType") = "A" Then

                            If .Rows(i).Item("sum07001") > 0 Then
                                dblMix10 = (.Rows(i).Item("sum07001") * qtyNo)
                                lb10sA_2.Text = Format(dblMix10 + CDbl(lb10sA_2.Text), "#,##0.00")
                            End If
                            If .Rows(i).Item("sum07002") > 0 Then
                                dblMix22 = (.Rows(i).Item("sum07002") * qtyNo)
                                lb22sA_2.Text = Format(dblMix22 + CDbl(lb22sA_2.Text), "#,##0.00")
                            End If
                            If .Rows(i).Item("sum07003") > 0 Then
                                dblMix60 = (.Rows(i).Item("sum07003") * qtyNo)
                                lb60sA_2.Text = Format(dblMix60 + CDbl(lb60sA_2.Text), "#,##0.00")
                            End If

                        Else

                            If .Rows(i).Item("sum07001") > 0 Then
                                dblMix10 = (.Rows(i).Item("sum07001") * qtyNo)
                                lb10sB_2.Text = Format(dblMix10 + CDbl(lb10sB_2.Text), "#,##0.00")
                            End If
                            If .Rows(i).Item("sum07002") > 0 Then
                                dblMix22 = (.Rows(i).Item("sum07002") * qtyNo)
                                lb22sB_2.Text = Format(dblMix22 + CDbl(lb22sB_2.Text), "#,##0.00")
                            End If
                            If .Rows(i).Item("sum07003") > 0 Then
                                dblMix60 = (.Rows(i).Item("sum07003") * qtyNo)
                                lb60sB_2.Text = Format(dblMix60 + CDbl(lb60sB_2.Text), "#,##0.00")
                            End If

                        End If

                    End If

                End If

                anydata = New String() {strdocDate, strDocNo, qtyNo.ToString("#0"), dblMix10.ToString("#,##0"), dblMix22.ToString("#,##0"), dblMix60.ToString("#,##0")}
                lvi = New ListViewItem(anydata)
                lsvDateList.Items.Add(lvi)

            Next
            ' End If
            lb10sTotal_1.Text = Format(CDbl(lb10sA_1.Text) + CDbl(lb10sB_1.Text), "#,##0")
            lb22sTotal_1.Text = Format(CDbl(lb22sA_1.Text) + CDbl(lb22sB_1.Text), "#,##0")
            lb60sTotal_1.Text = Format(CDbl(lb60sA_1.Text) + CDbl(lb60sB_1.Text), "#,##0")

            lb10sTotal_2.Text = Format(CDbl(lb10sA_2.Text) + CDbl(lb10sB_2.Text), "#,##0")
            lb22sTotal_2.Text = Format(CDbl(lb22sA_2.Text) + CDbl(lb22sB_2.Text), "#,##0")
            lb60sTotal_2.Text = Format(CDbl(lb60sA_2.Text) + CDbl(lb60sB_2.Text), "#,##0")

        End With



    End Sub
    Sub fmtListView()

        With lsvData
            '.Columns.Add("#", 40, HorizontalAlignment.Center) '1
            .Columns.Add("ชุดงาน", 230, HorizontalAlignment.Left) '1

            .Columns.Add("ชื่อชุดงาน", 570, HorizontalAlignment.Left) '0

            .Columns.Add("เวลาเริ่ม", 410, HorizontalAlignment.Right) '0
            .Columns.Add("รอ", 170, HorizontalAlignment.Right) '0
            .Columns.Add("รอ", 150, HorizontalAlignment.Right) '0
            .Columns.Add("Scale", 100, HorizontalAlignment.Center) '0

            .View = View.Details
            .GridLines = True
        End With
        With lsvDateList
            .Columns.Add("วันที่", 200, HorizontalAlignment.Center) '1
            .Columns.Add("ชุดงาน", 200, HorizontalAlignment.Left) '1

            .Columns.Add("จำนวนถัง", 200, HorizontalAlignment.Right) '0

            .Columns.Add("น้ำยา 10วิ", 200, HorizontalAlignment.Right) '0
            .Columns.Add("น้ำยา 22วิ", 200, HorizontalAlignment.Right) '0
            .Columns.Add("น้ำยา 60วิ", 200, HorizontalAlignment.Right) '0
            .Columns.Add("น้ำหนักรวม", 200, HorizontalAlignment.Center) '0

            .View = View.Details
            .GridLines = True
        End With
    End Sub
    Sub showDataList()

        Dim subds As New DataSet
        Dim subDa As SqlClient.SqlDataAdapter

        Dim strDateNow1 As String = Year(Now()) - 543 & "/" & Format(Month(Now()), "0#") & "/" & Format(Microsoft.VisualBasic.Day(Now()), "0#") ' & " 00:00:00"
        'lbDate01.Text = Format(CDate(strDateNow1), "dd/MM/yyyy")

        txtSQL = "Select * "
        txtSQL = txtSQL & "From cmMixMast "
        txtSQL = txtSQL & "Left Join BOMmastH "
        txtSQL = txtSQL & "On CM_Mix_Doc=BOM_No "
        txtSQL = txtSQL & "Left Join TranDataH_PM "
        txtSQL = txtSQL & "On CM_Mix_Doc=Trh_No "

        txtSQL = txtSQL & "where CM_Mix_Date='" & strDateNow1 & "' "
        txtSQL = txtSQL & "And Not(cm_Mix_Scale='5') "
        txtSQL = txtSQL & "And Trh_type ='M' "
        txtSQL = txtSQL & "And Trh_Chk_print='0' "
        txtSQL = txtSQL & "And CM_Mix_Chk=0 "
        txtSQL = txtSQL & "And trh_date >='2020-01-01 00:00:00' "

        txtSQL = txtSQL & "Order by CM_Mix_DateTime asc"

        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subds, "Master")
        Dim anydata() As String
        Dim lvi As ListViewItem
        Dim strDocNo As String
        Dim strStkName As String
        Dim strScale As String
        Dim strTime As String
        Dim dblTime As Integer
        Dim strConVTime As String
        lsvData.Items.Clear()

        For i = 0 To subds.Tables("Master").Rows.Count - 1
            With subds.Tables("Master").Rows(i)
                strDocNo = .Item("cm_Mix_Doc")
                strStkName = .Item("BOM_PC_Name")
                strTime = .Item("CM_Mix_DateTime")
                strScale = .Item("CM_Mix_Scale")

                dblTime = DateDiff(DateInterval.Minute, CDate(strToDate3(Now(), 1)), CDate(strToDate3(strTime, 1)))
                strTime = Mid(strTime, 1, 16)
                If IsNumeric(dblTime) Then
                    strConVTime = changeInt2Time(dblTime)
                Else
                    strConVTime = ""
                End If

                anydata = New String() {strDocNo, strStkName, strTime, strConVTime, dblTime.ToString("#,##0"), strScale}
                lvi = New ListViewItem(anydata)
                lsvData.Items.Add(lvi)

            End With
        Next

    End Sub

    Sub txtCLS()

        lb22sA_1.Text = 0
        lb22sB_1.Text = 0
        lb10sA_1.Text = 0
        lb10sB_1.Text = 0
        lb60sA_1.Text = 0
        lb60sB_1.Text = 0


        lb22sA_2.Text = 0
        lb22sB_2.Text = 0
        lb10sA_2.Text = 0
        lb10sB_2.Text = 0
        lb60sA_2.Text = 0
        lb60sB_2.Text = 0

        lb10sTotal_1.Text = 0
        lb60sTotal_1.Text = 0
        lb22sTotal_1.Text = 0


    End Sub
    Private Sub lb60sA_Click(sender As Object, e As EventArgs) Handles lb60sA_1.Click

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lbDate.Text = Now

    End Sub

    Private Sub GroupBox5_Enter(sender As Object, e As EventArgs) Handles GroupBox5.Enter

    End Sub

    Private Sub cmbExit_Click(sender As Object, e As EventArgs) Handles cmbExit.Click
        End
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        showData()
        showDataList()
    End Sub
End Class
