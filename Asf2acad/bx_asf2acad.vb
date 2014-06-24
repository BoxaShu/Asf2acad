'' bx_asf2acad.vb
'' © Шульжицкий Владимир, 2014 (boxa.shu@gmail.com)
'' Назначение: Плагин для AutoCAD, предназначен импортирования в автокад, результатов расчетов сохраненных в asf формате .
'' Команда: bx_asf2acad

Imports System.IO
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.Colors
Public Class acad__boxashu
    Const CrLf As String = ControlChars.CrLf 'Environment.NewLine

    <CommandMethod("bx_asf2acad")> _
    Public Sub bx_asf2acad()


        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor


        Dim openFileDialog1 As New OpenFileDialog("Выберите asf файл",
                                                  "*.asf",
                                                  "asf",
                                                  "Выбор файла",
                                                  OpenFileDialog.OpenFileDialogFlags.NoUrls)
        openFileDialog1.ShowDialog()
        Dim PATH As String = openFileDialog1.Filename
        Dim fn1 As String

        If PATH.ToUpper Like "?:\*.asf".ToUpper Then
            fn1 = Split(PATH, "\")(Split(PATH, "\").Rank) & "_" & Date.Now
            fn1 = Replace(fn1, ":", "-")
        Else
            Exit Sub
        End If

        Dim pnt_txt As String = "Точки" ' False - точки ; True - текст
        Dim pnt_txt_opt As PromptKeywordOptions = New PromptKeywordOptions(CrLf & "Способ вывода:")
        pnt_txt_opt.Keywords.Add("Точки")
        pnt_txt_opt.Keywords.Add("тЕкст")
        pnt_txt_opt.Keywords.Add("Солид")
        pnt_txt_opt.Keywords.Default = "Точки"
        Dim pnt_txt_res As PromptResult = acEd.GetKeywords(pnt_txt_opt)
        If pnt_txt_res.Status <> PromptStatus.OK Then
            Exit Sub
        Else
            If pnt_txt_res.StringResult = "Точки" Then
                pnt_txt = "Точки"
            End If
            If pnt_txt_res.StringResult = "тЕкст" Then
                pnt_txt = "тЕкст"
            End If
            If pnt_txt_res.StringResult = "Солид" Then
                pnt_txt = "Солид"
            End If
        End If


        Dim s_arm As PromptDoubleOptions = New PromptDoubleOptions(CrLf &
                                "Введите площадь фоновой арматуры в см2:")
        s_arm.AllowNone = False
        s_arm.AllowZero = True
        s_arm.AllowNegative = False

        Dim s_res As PromptDoubleResult = acEd.GetDouble(s_arm)
        If s_res.Status <> PromptStatus.OK Then
            acDoc.Editor.WriteMessage(CrLf & "Ошибка ввода площади.")
            Exit Sub
        End If

        Dim GL_POLYs As New Collection
        Dim GL_KNOTs As New Collection
        Dim GL_ELEMs As New Collection
        Dim GL_ARMs As New Collection

        Dim filestr As String = Trim(PATH)

        Dim text_file As String
        Try
            ' Открываем файл D:\test1.txt в стандартной кодировке 1251
            Dim F_R As New StreamReader(filestr, System.Text.Encoding.GetEncoding(1251))
            text_file = F_R.ReadToEnd
            F_R.Close() ' Закрываем файл

            'GL POLY
            Try
                'Разбиваем файл на разделы
                Dim t1() As String = Split(text_file, "GL POLY")
                Dim t2() As String = Split(t1(1), "GP KNOT")
                Dim t3() As String = Split(t2(0), CrLf)
                For Each i In t3
                    If i Like "*.??? *.??? *.???" Then

                        i = i.Replace("      ", " ")
                        i = i.Replace("     ", " ")
                        i = i.Replace("    ", " ")
                        i = i.Replace("   ", " ")
                        i = i.Replace("  ", " ")


                        Dim t4() As String = Split(Trim(i), " ")
                        Dim gl_poly As New GL_POLY
                        gl_poly.x = CType(Val(Trim(t4(0))), Double)
                        gl_poly.y = CType(Val(Trim(t4(1))), Double)
                        gl_poly.z = CType(Val(Trim(t4(2))), Double)

                        GL_POLYs.Add(gl_poly)
                    End If
                Next
            Catch ex As Exception
                acDoc.Editor.WriteMessage(CrLf & "Проблема с файлом. Он не прочитан.")
                acDoc.Editor.WriteMessage(CrLf & ex.Message)
                Exit Sub
            End Try

            'GL KNOT
            Try
                'Разбиваем файл на разделы
                Dim t1() As String = Split(text_file, "GP KNOT")
                Dim t2() As String = Split(t1(1), "GF ELEM")
                Dim t3() As String = Split(t2(0), CrLf)
                For Each i In t3
                    If i Like "* * * *" Then

                        i = i.Replace("      ", " ")
                        i = i.Replace("     ", " ")
                        i = i.Replace("    ", " ")
                        i = i.Replace("   ", " ")
                        i = i.Replace("  ", " ")

                        Dim t4() As String = Split(Trim(i), " ")
                        Dim GL_KNOT As New GL_KNOT

                        GL_KNOT.N = CType(Val(Trim(t4(0))), Integer)
                        GL_KNOT.x = CType(Val(Trim(t4(1))), Double)
                        GL_KNOT.y = CType(Val(Trim(t4(2))), Double)
                        GL_KNOT.z = CType(Val(Trim(t4(3))), Double)

                        GL_KNOTs.Add(GL_KNOT, GL_KNOT.N)
                    End If
                Next
            Catch ex As Exception
                acDoc.Editor.WriteMessage(CrLf & ex.Message)
            End Try

            'GF ARM
            Try
                'Разбиваем файл на разделы
                Dim t3() As String = Split(text_file, CrLf)
                For Each i In t3
                    If i Like "QM *" Then

                        i = i.Substring(2)
                        i = i.Replace("      ", " ")
                        i = i.Replace("     ", " ")
                        i = i.Replace("    ", " ")
                        i = i.Replace("   ", " ")
                        i = i.Replace("  ", " ")

                        Dim t4() As String = Split(Trim(i), " ")
                        Dim GL_ARM As New GL_ARM

                        GL_ARM.d1 = CType(Val(Trim(t4(0))), Integer) 'Dim d1 As Integer
                        GL_ARM.d2 = CType(Val(Trim(t4(1))), Integer) 'Dim d2 As Integer

                        GL_ARM.x = CType(Val(Trim(t4(2))), Double) 'Dim x As Double
                        GL_ARM.y = CType(Val(Trim(t4(3))), Double) 'Dim y As Double
                        GL_ARM.z = CType(Val(Trim(t4(4))), Double) 'Dim z As Double

                        GL_ARM.Asv_X = CType(Val(Trim(t4(6))), Double) 'Dim Asn_X As Double
                        GL_ARM.Asv_Y = CType(Val(Trim(t4(5))), Double) 'Dim Asn_Y As Double
                        GL_ARM.Asv_Z = CType(Val(Trim(t4(7))), Double) 'Dim Asn_Z As Double

                        GL_ARM.Asn_X = CType(Val(Trim(t4(8))), Double) 'Dim Asv_X As Double
                        GL_ARM.Asn_Y = CType(Val(Trim(t4(9))), Double) 'Dim Asv_Y As Double
                        GL_ARM.Asn_Z = CType(Val(Trim(t4(10))), Double) 'Dim Asv_Z As Double

                        GL_ARMs.Add(GL_ARM)
                    End If
                Next
            Catch ex As Exception
                acDoc.Editor.WriteMessage(CrLf & ex.Message)
            End Try
        Catch ex As Exception
            acDoc.Editor.WriteMessage(CrLf & ex.Message)
            text_file = "0"
        End Try



        Dim pm As ProgressMeter = New ProgressMeter()
        pm.Start("Заполняю элементы:")
        pm.SetLimit(GL_KNOTs.Count)

        'GF ELEM
        Try
            'Разбиваем файл на разделы
            Dim t1() As String = Split(text_file, "GF ELEM")
            Dim t2() As String = Split(t1(1), "QR")
            Dim t3() As String = Split(t2(0), CrLf)
            For Each i In t3
                If i Like "* * * * *" Then

                    'i = i.Replace("      ", " ")
                    'i = i.Replace("     ", " ")
                    'i = i.Replace("    ", " ")
                    'i = i.Replace("   ", " ")
                    'i = i.Replace("  ", " ")
                    'i = i.Trim
                    'Dim t4() As String = Split(Trim(i), " ")
                    'Dim GL_ELEM As New GL_ELEM
                    'GL_ELEM.N = CType(Val(Trim(t4(0))), Integer)
                    'GL_ELEM.p1 = CType(Val(Trim(t4(1))), Integer)
                    'GL_ELEM.p2 = CType(Val(Trim(t4(2))), Integer)
                    'GL_ELEM.p3 = CType(Val(Trim(t4(3))), Integer)
                    'GL_ELEM.p4 = CType(Val(Trim(t4(4))), Integer)

                    Dim GL_ELEM As New GL_ELEM
                    GL_ELEM.N = Integer.Parse(i.Substring(0, 5).Trim)
                    GL_ELEM.p1 = Integer.Parse(i.Substring(5, 5).Trim)
                    GL_ELEM.p2 = Integer.Parse(i.Substring(10, 5).Trim)
                    GL_ELEM.p3 = Integer.Parse(i.Substring(15, 5).Trim)
                    GL_ELEM.p4 = Integer.Parse(i.Substring(20, 5).Trim)



                    Dim p1 As GL_KNOT = GL_KNOTs.Item(GL_ELEM.p1)
                    Dim p2 As GL_KNOT = GL_KNOTs.Item(GL_ELEM.p2)
                    Dim p3 As GL_KNOT = GL_KNOTs.Item(GL_ELEM.p3)
                    Dim p4 As GL_KNOT

                    If GL_KNOTs.Contains(GL_ELEM.p4) = True Then
                        p4 = GL_KNOTs.Item(GL_ELEM.p4)
                    Else
                        p4 = New GL_KNOT
                        p4.N = 0
                        p4.x = p1.x
                        p4.y = p1.y
                        p4.z = p1.z
                    End If

                    Dim x_min As Double = {p1.x, p2.x, p3.x, p4.x}.Min
                    Dim y_min As Double = {p1.y, p2.y, p3.y, p4.y}.Min

                    Dim x_max As Double = {p1.x, p2.x, p3.x, p4.x}.Max
                    Dim y_max As Double = {p1.y, p2.y, p3.y, p4.y}.Max

                    For Each e As GL_ARM In GL_ARMs
                        If e.x >= x_min And e.y >= y_min And e.x <= x_max And e.y <= y_max Then
                            GL_ELEM.Asn_X = e.Asn_X
                            GL_ELEM.Asn_Y = e.Asn_Y
                            GL_ELEM.Asn_Z = e.Asn_Z

                            GL_ELEM.Asv_X = e.Asv_X
                            GL_ELEM.Asv_Y = e.Asv_Y
                            GL_ELEM.Asv_Z = e.Asv_Z
                            Exit For
                        End If
                    Next
                    GL_ELEMs.Add(GL_ELEM)
                End If

                pm.MeterProgress()
                System.Windows.Forms.Application.DoEvents()
            Next

        Catch ex As Exception
            acDoc.Editor.WriteMessage(CrLf & ex.Message)
        End Try
        pm.Stop()

        Dim p1_opt As PromptPointOptions = New PromptPointOptions(CrLf & "Укажите точку вставки Asn_X:")
        p1_opt.AllowNone = False
        Dim p1_res As PromptPointResult = acEd.GetPoint(p1_opt)
        If p1_res.Status <> PromptStatus.OK Then
            Exit Sub
        End If

        If pnt_txt = "Точки" Then
            Call AddPLine(GL_POLYs, p1_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asn_X > s_res.Value Then
                    Call AddPointAndSetPointStyle(New Point3d(i.x + p1_res.Value.X, i.y + p1_res.Value.Y, i.z),
                                                  i.Asn_X,
                                                  34
                                                  )
                End If
            Next
        End If
        If pnt_txt = "тЕкст" Then
            Call AddPLine(GL_POLYs, p1_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asn_X > s_res.Value Then
                    Call AddText(New Point3d(i.x + p1_res.Value.X, i.y + p1_res.Value.Y, i.z), i.Asn_X)
                End If
            Next
        End If

        If pnt_txt = "Солид" Then
            'Вывод элемента
            Dim Block_name As String = fn1 & "_Asn_X"
            Dim ObjID As ObjectId = function_lib.CreateBlockDefinition(Block_name)

            For Each i As GL_ELEM In GL_ELEMs

                If i.Asn_X > s_res.Value Then
                    Dim p1 As GL_KNOT = GL_KNOTs.Item(i.p1)
                    'p1.x = p1.x + p1_res.Value.X
                    'p1.y = p1.y + p1_res.Value.Y

                    Dim p2 As GL_KNOT = GL_KNOTs.Item(i.p2)
                    'p2.x = p2.x + p1_res.Value.X
                    'p2.y = p2.y + p1_res.Value.Y

                    Dim p3 As GL_KNOT = GL_KNOTs.Item(i.p3)
                    'p3.x = p3.x + p1_res.Value.X
                    'p3.y = p3.y + p1_res.Value.Y

                    Dim p4 As GL_KNOT
                    If GL_KNOTs.Contains(i.p4) = True Then
                        p4 = GL_KNOTs.Item(i.p4)
                    Else
                        p4 = New GL_KNOT
                        p4.N = 0
                        p4.x = p1.x
                        p4.y = p1.y
                        p4.z = p1.z
                    End If
                    'p4.x = p4.x + p1_res.Value.X
                    'p4.y = p4.y + p1_res.Value.Y

                    Call Add2DSolid(Block_name, p1, p2, p3, p4, i.Asn_X)
                End If
            Next

            Call AddPLine2block(Block_name, GL_POLYs)
            Call Add2Block(Block_name, p1_res.Value)
        End If


        Dim p2_opt As PromptPointOptions = New PromptPointOptions(CrLf & "Укажите точку вставки Asn_Y:")
        p2_opt.AllowNone = False
        Dim p2_res As PromptPointResult = acEd.GetPoint(p2_opt)
        If p2_res.Status <> PromptStatus.OK Then
            Exit Sub
        End If

        If pnt_txt = "Точки" Then
            Call AddPLine(GL_POLYs, p2_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asn_Y > s_res.Value Then
                    Call AddPointAndSetPointStyle(New Point3d(i.x + p2_res.Value.X, i.y + p2_res.Value.Y, i.z), i.Asn_Y, 34)
                End If
            Next
        End If
        If pnt_txt = "тЕкст" Then
            Call AddPLine(GL_POLYs, p2_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asn_Y > s_res.Value Then
                    Call AddText(New Point3d(i.x + p2_res.Value.X, i.y + p2_res.Value.Y, i.z), i.Asn_Y)
                End If
            Next
        End If

        If pnt_txt = "Солид" Then
            Dim Block_name As String = fn1 & "_Asn_Y"
            Dim ObjID As ObjectId = CreateBlockDefinition(Block_name)

            'Вывод элемента
            For Each i As GL_ELEM In GL_ELEMs

                If i.Asn_Y > s_res.Value Then
                    Dim p1 As GL_KNOT = GL_KNOTs.Item(i.p1)
                    Dim p2 As GL_KNOT = GL_KNOTs.Item(i.p2)
                    Dim p3 As GL_KNOT = GL_KNOTs.Item(i.p3)
                    Dim p4 As GL_KNOT
                    If GL_KNOTs.Contains(i.p4) = True Then
                        p4 = GL_KNOTs.Item(i.p4)
                    Else
                        p4 = New GL_KNOT
                        p4.N = 0
                        p4.x = p1.x
                        p4.y = p1.y
                        p4.z = p1.z
                    End If
                    Call Add2DSolid(Block_name, p1, p2, p3, p4, i.Asn_Y)
                End If
            Next
            Call AddPLine2block(Block_name, GL_POLYs)
            Call Add2Block(Block_name, p2_res.Value)
        End If

        Dim p3_opt As PromptPointOptions = New PromptPointOptions(CrLf & "Укажите точку вставки Asv_X:")
        p3_opt.AllowNone = False
        Dim p3_res As PromptPointResult = acEd.GetPoint(p3_opt)
        If p3_res.Status <> PromptStatus.OK Then
            Exit Sub
        End If

        If pnt_txt = "Точки" Then
            Call AddPLine(GL_POLYs, p3_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asv_X > s_res.Value Then
                    Call AddPointAndSetPointStyle(New Point3d(i.x + p3_res.Value.X, i.y + p3_res.Value.Y, i.z), i.Asv_X, 34)
                End If
            Next
        End If
        If pnt_txt = "тЕкст" Then
            Call AddPLine(GL_POLYs, p3_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asv_X > s_res.Value Then
                    Call AddText(New Point3d(i.x + p3_res.Value.X, i.y + p3_res.Value.Y, i.z), i.Asv_X)
                End If
            Next
        End If


        If pnt_txt = "Солид" Then

            Dim Block_name As String = fn1 & "_Asv_X"
            Dim ObjID As ObjectId = CreateBlockDefinition(Block_name)
            'Вывод элемента
            For Each i As GL_ELEM In GL_ELEMs
                If i.Asv_X > s_res.Value Then
                    Dim p1 As GL_KNOT = GL_KNOTs.Item(i.p1)
                    Dim p2 As GL_KNOT = GL_KNOTs.Item(i.p2)
                    Dim p3 As GL_KNOT = GL_KNOTs.Item(i.p3)
                    Dim p4 As GL_KNOT
                    If GL_KNOTs.Contains(i.p4) = True Then
                        p4 = GL_KNOTs.Item(i.p4)
                    Else
                        p4 = New GL_KNOT
                        p4.N = 0
                        p4.x = p1.x
                        p4.y = p1.y
                        p4.z = p1.z
                    End If
                    Call Add2DSolid(Block_name, p1, p2, p3, p4, i.Asv_X)
                End If
            Next
            Call AddPLine2block(Block_name, GL_POLYs)
            Call Add2Block(Block_name, p3_res.Value)
        End If


        Dim p4_opt As PromptPointOptions = New PromptPointOptions(CrLf & "Укажите точку вставки Asv_Y:")
        p4_opt.AllowNone = False
        Dim p4_res As PromptPointResult = acEd.GetPoint(p4_opt)
        If p4_res.Status <> PromptStatus.OK Then
            Exit Sub
        End If

        If pnt_txt = "Точки" Then
            Call AddPLine(GL_POLYs, p4_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asv_Y > s_res.Value Then
                    Call AddPointAndSetPointStyle(New Point3d(i.x + p4_res.Value.X, i.y + p4_res.Value.Y, i.z), i.Asv_Y, 34)
                End If
            Next
        End If
        If pnt_txt = "тЕкст" Then
            Call AddPLine(GL_POLYs, p4_res.Value)
            For Each i As GL_ARM In GL_ARMs
                If i.Asv_Y > s_res.Value Then
                    Call AddText(New Point3d(i.x + p4_res.Value.X, i.y + p4_res.Value.Y, i.z), i.Asv_Y)
                End If
            Next
        End If


        If pnt_txt = "Солид" Then
            Dim Block_name As String = fn1 & "Asv_Y"
            Dim ObjID As ObjectId = CreateBlockDefinition(Block_name)

            'Вывод элемента
            For Each i As GL_ELEM In GL_ELEMs
                If i.Asv_Y > s_res.Value Then
                    Dim p1 As GL_KNOT = GL_KNOTs.Item(i.p1)
                    Dim p2 As GL_KNOT = GL_KNOTs.Item(i.p2)
                    Dim p3 As GL_KNOT = GL_KNOTs.Item(i.p3)
                    Dim p4 As GL_KNOT
                    If GL_KNOTs.Contains(i.p4) = True Then
                        p4 = GL_KNOTs.Item(i.p4)
                    Else
                        p4 = New GL_KNOT
                        p4.N = 0
                        p4.x = p1.x
                        p4.y = p1.y
                        p4.z = p1.z
                    End If
                    Call Add2DSolid(Block_name, p1, p2, p3, p4, i.Asv_Y)
                End If
            Next
            Call AddPLine2block(Block_name, GL_POLYs)
            Call Add2Block(Block_name, p4_res.Value)
        End If



    End Sub


    <CommandMethod("bx_GetXData")> _
    Public Sub bx_GetXData()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Dim opt As PromptEntityOptions = New PromptEntityOptions(CrLf & "Select entity: ")
        opt.AllowNone = False

        Dim res As PromptEntityResult = acEd.GetEntity(opt)
        If res.Status <> PromptStatus.OK Then
            Exit Sub
        End If

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim obj As DBObject = acTrans.GetObject(res.ObjectId, OpenMode.ForRead)
            Dim rb As ResultBuffer = obj.XData
            If rb = Nothing Then
                acEd.WriteMessage(CrLf & "Entity does not have XData attached.")
            Else
                Dim n As Integer = 0
                For Each tv As TypedValue In rb
                    acEd.WriteMessage(CrLf & "TypedValue {0} - type: {1}, value: {2}",
                                n,
                                tv.TypeCode,
                                tv.Value
                                )
                    n += n
                Next
                rb.Dispose()
            End If
            acTrans.Commit()
        End Using
    End Sub
End Class





