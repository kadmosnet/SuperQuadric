Imports DXFReaderNET
Imports Microsoft.Win32
Public Class Form1

    Dim rxg, ryg, rzg, a1, a2, a3, eps1, eps2, eta, omega As Double
    Dim prec, num, r, p, j, k, v1, v2, va1, va2 As Integer
    Dim color As Integer

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        DxfReaderNET1.Rendering = RenderingType.WireFrame
        DxfReaderNET1.Refresh()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        DxfReaderNET1.Rendering = RenderingType.Shaded
        DxfReaderNET1.Refresh()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        DxfReaderNET1.Rendering = RenderingType.ShadedEdges
        DxfReaderNET1.Refresh()

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        With SaveFileDialog1
            .DefaultExt = "jpg"
            .Filter = "JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|BMP (*.bmp)|*.bmp"




            .FilterIndex = 1
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Select Case .FilterIndex
                    Case 1
                        DxfReaderNET1.Image.Save(.FileName, Imaging.ImageFormat.Jpeg)
                    Case 2
                        DxfReaderNET1.Image.Save(.FileName, Imaging.ImageFormat.Png)
                    Case 3
                        DxfReaderNET1.Image.Save(.FileName, Imaging.ImageFormat.Bmp)

                End Select


            End If



        End With
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        With DxfReaderNET1

            .NewDrawing()
            color = 1
            r = 16
            p = 32
            rxg = 0
            ryg = 0
            rzg = 0
            eps1 = 1
            eps2 = 1
            Center.X = 0
            Center.Y = 0
            Center.Z = 0
            prec = 4
            a1 = 20
            a2 = 20
            a3 = 20

            Dim k As Integer
            Dim j As Integer
            color = 10
            eps1 = 0.5
            eps2 = 0.5

            For j = 0 To 3
                For k = 0 To 5
                    Center.Y = j * 60
                    Center.X = k * 60
                    GenerateSuperQuadric()
                    color += 10
                    eps1 += 0.06
                    eps2 += 0.06
                Next k
            Next j

            DxfReaderNET1.DisplayPredefinedView(PredefinedViewType.SW_Isometric)
            DxfReaderNET1.Refresh()
            DxfReaderNET1.ZoomExtents()
        End With
    End Sub

    Dim Center As Vector3

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        r = Val(TextBoxRings.Text)
        p = Val(TextBoxPointsXring.Text)
        rxg = Val(TextBoxRotationX.Text)
        ryg = Val(TextBoxRotationY.Text)
        rzg = Val(TextBoxRotationZ.Text)
        eps1 = Val(TextBoxEps1.Text)
        eps2 = Val(TextBoxEps2.Text)
        Center.X = Val(TextBoxCenterX.Text)
        Center.Y = Val(TextBoxCenterY.Text)
        Center.Z = Val(TextBoxCenterZ.Text)
        prec = 4
        a1 = Val(TextBoxRadiusX.Text)
        a2 = Val(TextBoxRadiusY.Text)
        a3 = Val(TextBoxRadiusZ.Text)
        color = Val(TextBoxColor.Text)
        GenerateSuperQuadric()
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = New Vector3(-Math.Sqrt(1 / 3), -Math.Sqrt(1 / 3), Math.Sqrt(1 / 3))
        DxfReaderNET1.Refresh()
        DxfReaderNET1.ZoomExtents()

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SaveRegistry()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DxfReaderNET1.NewDrawing()
        LoadRegistry()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = Vector3.UnitZ

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = -Vector3.UnitZ

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = -Vector3.UnitX

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = Vector3.UnitX

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = -Vector3.UnitY

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = Vector3.UnitY

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = New Vector3(-Math.Sqrt(1 / 3), -Math.Sqrt(1 / 3), Math.Sqrt(1 / 3))

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = New Vector3(Math.Sqrt(1 / 3), Math.Sqrt(1 / 3), Math.Sqrt(1 / 3))

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = New Vector3(Math.Sqrt(1 / 3), Math.Sqrt(1 / 3), Math.Sqrt(1 / 3))

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        DxfReaderNET1.DXF.VPorts("*Active").ViewDirection = New Vector3(-Math.Sqrt(1 / 3), Math.Sqrt(1 / 3), Math.Sqrt(1 / 3))

        DxfReaderNET1.ZoomExtents()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DxfReaderNET1.NewDrawing()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With SaveFileDialog1
            .DefaultExt = "dxf"
            .Filter = "AutoCAD R10 DXF|*.dxf" '1
            .Filter += "|AutoCAD R11 and R12 DXF|*.dxf" '2
            .Filter += "|AutoCAD R13 DXF|*.dxf" '3
            .Filter += "|AutoCAD R14 DXF|*.dxf" '4
            .Filter += "|AutoCAD 2000 DXF|*.dxf" '5
            .Filter += "|AutoCAD 2004 DXF|*.dxf" '6
            .Filter += "|AutoCAD 2007 DXF|*.dxf" '7
            .Filter += "|AutoCAD 2010 DXF|*.dxf" '8
            .Filter += "|AutoCAD 2013 DXF|*.dxf" '9
            .Filter += "|AutoCAD 2018 DXF|*.dxf" '10

            .Filter += "|AutoCAD R10 binary DXF|*.dxf" '11
            .Filter += "|AutoCAD R11 and R12 binary DXF|*.dxf" '12
            .Filter += "|AutoCAD R13 binary DXF|*.dxf" '13
            .Filter += "|AutoCAD R14 binary DXF|*.dxf" '14
            .Filter += "|AutoCAD 2000 binary DXF|*.dxf" '15
            .Filter += "|AutoCAD 2004 binary DXF|*.dxf" '16
            .Filter += "|AutoCAD 2007 binary DXF|*.dxf" '17
            .Filter += "|AutoCAD 2010 binary DXF|*.dxf" '18
            .Filter += "|AutoCAD 2013 binary DXF|*.dxf" '19
            .Filter += "|AutoCAD 2018 binary DXF|*.dxf" '20


            Select Case DxfReaderNET1.DXF.DrawingVariables.AcadVer
                Case Header.DxfVersion.AutoCad10
                    .FilterIndex = 1
                Case Header.DxfVersion.AutoCad12

                    .FilterIndex = 2

                Case Header.DxfVersion.AutoCad13

                    .FilterIndex = 3
                Case Header.DxfVersion.AutoCad14

                    .FilterIndex = 4
                Case Header.DxfVersion.AutoCad2000

                    .FilterIndex = 5
                Case Header.DxfVersion.AutoCad2004

                    .FilterIndex = 6
                Case Header.DxfVersion.AutoCad2007

                    .FilterIndex = 7
                Case Header.DxfVersion.AutoCad2010


                    .FilterIndex = 8
                Case Header.DxfVersion.AutoCad2013

                    .FilterIndex = 9
                Case Header.DxfVersion.AutoCad2018
                    .FilterIndex = 10

            End Select

            If Not DxfReaderNET1.FileName Is Nothing Then
                Dim FileInfo As New IO.FileInfo(DxfReaderNET1.FileName())

                .FileName = FileInfo.Name
            Else
                .FileName = "SuperQuadric.dxf"
            End If


            If DxfReaderNET1.DXF.IsBinary Then .FilterIndex += 10

            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                Dim dxfver As DXFReaderNET.Header.DxfVersion = Header.DxfVersion.AutoCad2013
                Select Case .FilterIndex
                    Case 1, 11
                        dxfver = Header.DxfVersion.AutoCad10
                    Case 2, 12
                        dxfver = Header.DxfVersion.AutoCad12
                    Case 3, 13
                        dxfver = Header.DxfVersion.AutoCad13

                    Case 4, 14
                        dxfver = Header.DxfVersion.AutoCad14

                    Case 5, 15
                        dxfver = Header.DxfVersion.AutoCad2000

                    Case 6, 16
                        dxfver = Header.DxfVersion.AutoCad2004


                    Case 7, 17
                        dxfver = Header.DxfVersion.AutoCad2007

                    Case 8, 16
                        dxfver = Header.DxfVersion.AutoCad2010


                    Case 9, 19
                        dxfver = Header.DxfVersion.AutoCad2013

                    Case 10, 20
                        dxfver = Header.DxfVersion.AutoCad2018


                End Select

                If .FilterIndex > 10 Then
                    DxfReaderNET1.DXF.IsBinary = True
                End If

                DxfReaderNET1.WriteDXF(.FileName, dxfver, DxfReaderNET1.DXF.IsBinary)

            End If

        End With
    End Sub


    Private Sub GenerateSuperQuadric()


        Dim alpha As Double = rxg * Math.PI / 180
        Dim beta As Double = ryg * Math.PI / 180
        Dim gamma As Double = rzg * Math.PI / 180

        Dim matr() As Vector3 = Nothing

        num = 0
        eta = Math.PI / 2
        For j = 1 To r
            omega = 0
            For k = 1 To p
                ReDim Preserve matr(num + 1)
                matr(num + 1) = Superellipsoid(eta, omega)
                num += 1
                omega += 2 * Math.PI / p

            Next k
            eta -= Math.PI / (r - 1)

        Next j

        Dim m As New Matrix
        For k = 1 To matr.Length - 1
            matr(k) = Matrix.RotationX(alpha) * Matrix.RotationX(beta) * Matrix.RotationX(gamma) * matr(k)

            matr(k) += Center


        Next k

        'dxf generation
        For k = 0 To r * p - p - 1
            v1 = k + p + 1
            v2 = k + 1
            va1 = v1
            va2 = v2


            If v1 Mod p = 0 Then
                va1 = v1 - p
                va2 = v2 - p
            End If
            Dim FirstVertex As Vector3
            Dim SecondVertex As Vector3
            Dim ThirdVertex As Vector3
            Dim FourthVertex As Vector3

            FirstVertex = matr(k + 1)
            SecondVertex = matr(k + p + 1)
            ThirdVertex = matr(va1 + 1)
            FourthVertex = matr(va2 + 1)

            DxfReaderNET1.AddFace3D(FirstVertex, SecondVertex, ThirdVertex, FourthVertex, color)

        Next k
        'DxfReaderNET1.Refresh()
        'DxfReaderNET1.ZoomExtents()
    End Sub

    Private Function Superellipsoid(eta As Double, omega As Double) As Vector3
        Dim v As Vector3 = Vector3.Zero
        v.X = a1 * elev(Math.Cos(eta), eps1) * elev(Math.Cos(omega), eps2)
        v.Y = a2 * elev(Math.Cos(eta), eps1) * elev(Math.Sin(omega), eps2)
        v.Z = a3 * elev(Math.Sin(eta), eps1)
        Return v
    End Function
    Private Function elev(x As Double, y As Double) As Double
        Dim buf As Double = 0
        Dim e As Double = 0
        Dim test As Boolean = True

        If x <> 0 Then
            buf = y * Math.Log(Math.Abs(x))
        Else
            test = False
        End If
        If buf > -88 Then
            e = Math.Exp(buf) * Math.Sign(x)
        Else
            test = False
        End If
        If Not test Then
            e = 0
        End If
        Return e
    End Function


    Private Sub LoadRegistry()

        Me.Width = Registry.GetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wWidth", Screen.PrimaryScreen.Bounds.Width - 40)
        Me.Height = Registry.GetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wHeight", Screen.PrimaryScreen.Bounds.Height - 60)
        Me.Left = Registry.GetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wLeft", 20)
        Me.Top = Registry.GetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wTop", 20)


    End Sub

    Private Sub SaveRegistry()

        Registry.SetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wWidth", Me.Width)
        Registry.SetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wHeight", Me.Height)
        Registry.SetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wLeft", Me.Left)
        Registry.SetValue("HKEY_CURRENT_USER\Software\SuperQuadric", "m_wTop", Me.Top)

    End Sub
End Class
