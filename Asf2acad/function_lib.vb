'' function_lib.vb
'' © Шульжицкий Владимир, 2014 (boxa.shu@gmail.com)

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors


Module function_lib
    Const CrLf As String = ControlChars.CrLf

    Function getcolor(ByRef color As Integer) As Integer
        'acPoint.ColorIndex = 10 + CInt(colorIndex) ^ 2
        getcolor = 10 + color
    End Function

#Region "вывод примитивов"
    Public Sub AddPointAndSetPointStyle(ByVal p1 As Point3d, ByVal colorIndex As Integer, ByVal type As Integer)

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)
            '' Создание точки с координатой (4, 3, 0) в пространстве Модели
            Dim acPoint As DBPoint = New DBPoint(p1)

            acPoint.ColorIndex = getcolor(colorIndex)
            acPoint.SetDatabaseDefaults()

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acPoint)
            acTrans.AddNewlyCreatedDBObject(acPoint, True)

            '' Установка стиля для всех объектов точек в чертеже
            acCurDb.Pdmode = type
            acCurDb.Pdsize = 0.3

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub


    Sub AddPLine(ByVal GL_POLYs As Collection, ByVal p3d As Point3d)

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков пространства Модели для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)

            '' Создание полилинии с двумя сегментами (3 точки)
            Dim acPoly As Polyline = New Polyline()
            acPoly.SetDatabaseDefaults()

            Dim q As Integer = 0
    For Each i As GL_POLY In GL_POLYs
                acPoly.AddVertexAt(q, New Point2d(i.x + p3d.X, i.y + p3d.Y), 0, 0, 0) '37.000  21.950   0.000
                q = q + 1
            Next

            acPoly.Closed = True

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub


    Sub AddLine(ByVal p1 As Point3d, ByVal p2 As Point3d, ByVal colorIndex As Integer)

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков пространства Модели для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)

            '' Создание отрезка начинающегося в 5,5 и заканчивающегося в 12,3
            Dim acLine As Line = New Line(p1, p2)
            acLine.ColorIndex = getcolor(colorIndex)
            acLine.SetDatabaseDefaults()

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub

    Sub AddText(ByVal p1 As Point3d, ByVal colorIndex As Double)

        '' Устанавливаем текущий документ и базу данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Начинаем транзакцию
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открываем таблицу Блока для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                    OpenMode.ForRead)

            '' Открываем запись таблицы Блока пространство Модели (Model space) для записи
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                       OpenMode.ForWrite)

            '' Создаем однострочный текстовый объект
            Dim acText As DBText = New DBText()
            acText.SetDatabaseDefaults()
            acText.Position = p1 'New Point3d(2, 2, 0)
            acText.Height = 0.1
            acText.TextString = String.Format("{0:n2}", Math.Round(colorIndex, 2)) 'A_s.ToString '"Hello, World."
            acText.ColorIndex = getcolor(colorIndex)
            'acText.AlignmentPoint = p1
            acText.HorizontalMode = TextHorizontalMode.TextCenter
            acText.VerticalMode = TextVerticalMode.TextVerticalMid
            acText.AlignmentPoint = p1
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            '' Сохраняем изменения и закрываем транзакцию
            acTrans.Commit()
        End Using
    End Sub


    Sub Add2DSolid(ByRef ObjID As String,
                   ByRef p1 As GL_KNOT,
                   ByVal p2 As GL_KNOT,
                   ByVal p3 As GL_KNOT,
                   ByVal p4 As GL_KNOT,
                   ByVal colorIndex As Double)

        'Dim x_min As Double = {p1.x, p2.x, p3.x, p4.x}.Min
        'Dim y_min As Double = {p1.y, p2.y, p3.y, p4.y}.Min

        'Dim x_max As Double = {p1.x, p2.x, p3.x, p4.x}.Max
        'Dim y_max As Double = {p1.y, p2.y, p3.y, p4.y}.Max

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков для записи
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(ObjID), OpenMode.ForWrite)

            Dim ac2DSolidSqr As Solid
            If p4.N <> 0 Then

                ac2DSolidSqr = New Solid(New Point3d(p1.x, p1.y, p1.z), _
                                                 New Point3d(p2.x, p2.y, p2.z), _
                                                 New Point3d(p4.x, p4.y, p4.z), _
                                                 New Point3d(p3.x, p3.y, p3.z))

            Else

                ac2DSolidSqr = New Solid(New Point3d(p1.x, p1.y, p1.z), _
                                         New Point3d(p2.x, p2.y, p2.z), _
                                         New Point3d(p3.x, p3.y, p3.z))
            End If

            AddRegAppTableRecord("bx_asf2acad")
            Dim rb As ResultBuffer = New ResultBuffer(
                                     New TypedValue(1001, "bx_asf2acad"),
                                     New TypedValue(1000, colorIndex.ToString)
                                     )
            ac2DSolidSqr.XData = rb
            rb.Dispose()

            'ac2DSolidSqr.ColorIndex = 10 + CInt(colorIndex) ^ 2
            ac2DSolidSqr.ColorIndex = getcolor(colorIndex)
            'ac2DSolidSqr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 0)
            ac2DSolidSqr.SetDatabaseDefaults()

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(ac2DSolidSqr)
            acTrans.AddNewlyCreatedDBObject(ac2DSolidSqr, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub


    Function CreateBlockDefinition(ByVal block_name As String) As ObjectId

        Dim newBtrId As ObjectId

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()


            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)

            Dim newBtr As BlockTableRecord = New BlockTableRecord()
            newBtr.Name = block_name
            newBtrId = acBlkTbl.Add(newBtr)

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acTrans.AddNewlyCreatedDBObject(newBtr, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using

    End Function


    Sub Add2Block(ByRef ObjID As String, ByRef p1 As Point3d)

        '' Получение текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        'Данные базовой отметки
        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim bt As BlockTable = TryCast(acCurDb.BlockTableId.GetObject(OpenMode.ForRead), BlockTable)
            Dim blockDef As BlockTableRecord = TryCast(bt(ObjID).GetObject(OpenMode.ForRead), BlockTableRecord)
            'Also open modelspace - we'll be adding our BlockReference to it
            Dim ms As BlockTableRecord = TryCast(bt(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite), BlockTableRecord)
            'Create new BlockReference, and link it to our block definition

            'Горизонтальное направление
            Using blockRef As New BlockReference(p1, blockDef.ObjectId)
                'Add the block reference to modelspace
                ms.AppendEntity(blockRef)
                acTrans.AddNewlyCreatedDBObject(blockRef, True)
                'Iterate block definition to find all non-constant 
                ' AttributeDefinitions
                For Each id As ObjectId In blockDef
                    Dim obj As DBObject = id.GetObject(OpenMode.ForRead)
                    Dim attDef As AttributeDefinition = TryCast(obj, AttributeDefinition)
                    If (attDef IsNot Nothing) AndAlso (Not attDef.Constant) Then
                        'This is a non-constant AttributeDefinition 
                        'Create a new AttributeReference
                        Using attRef As New AttributeReference()
                            attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform)
                            'Add the AttributeReference to the BlockReference
                            blockRef.AttributeCollection.AppendAttribute(attRef)
                            acTrans.AddNewlyCreatedDBObject(attRef, True)
                        End Using
                    End If
                Next
                'blockRef.Layer = block_layer
                'blockRef.Rotation = alfa - Math.PI / 2 'Math.PI / 2
                'blockRef.LineWeight = LineWeight.ByLayer
                blockRef.Position = p1
                blockRef.ScaleFactors = New Scale3d(1000)
            End Using
            ' Сохранение нового объекта в базе данных
            acTrans.Commit()
            ' Очистка транзакции
        End Using
    End Sub

    Sub AddPLine2block(ByRef ObjID As String, ByVal GL_POLYs As Collection)

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Открытие таблицы Блоков для чтения
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Открытие записи таблицы Блоков пространства Модели для записи
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(ObjID), OpenMode.ForWrite)

            '' Создание полилинии с двумя сегментами (3 точки)
            Dim acPoly As Polyline = New Polyline()
            acPoly.SetDatabaseDefaults()

            Dim q As Integer = 0
            For Each i As GL_POLY In GL_POLYs
                acPoly.AddVertexAt(q, New Point2d(i.x, i.y), 0, 0, 0) '37.000  21.950   0.000
                q = q + 1
            Next
            acPoly.Closed = True

            '' Добавление нового объекта в запись таблицы блоков и в транзакцию
            acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)

            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub

    Sub AddRegAppTableRecord(ByRef regAppName As String)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim rat As RegAppTable = CType(acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead, False), RegAppTable)

            If rat.Has(regAppName) = False Then
                rat.UpgradeOpen()
                Dim ratr As RegAppTableRecord = New RegAppTableRecord()
                ratr.Name = regAppName
                rat.Add(ratr)
                acTrans.AddNewlyCreatedDBObject(ratr, True)
            End If
            '' Сохранение нового объекта в базе данных
            acTrans.Commit()
        End Using
    End Sub

#End Region


End Module
