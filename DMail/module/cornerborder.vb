Module cornerborder

    Public Sub ctrlCornerBorder(ctrl As Control, CurveSize As Integer)

        Try

            Dim p As New System.Drawing.Drawing2D.GraphicsPath

            p.StartFigure()
            p.AddArc(New Rectangle(0, 0, CurveSize, CurveSize), 180, 90)
            'p.AddLine(CurveSize, 0, ctrl.Width - CurveSize, 0)

            p.AddArc(New Rectangle(ctrl.Width - CurveSize, 0, CurveSize, CurveSize), -90, 90)
            'p.AddLine(ctrl.Width, CurveSize, ctrl.Width, ctrl.Height - CurveSize)

            p.AddArc(New Rectangle(ctrl.Width - CurveSize, ctrl.Height - CurveSize, CurveSize, CurveSize), 0, 90)
            'p.AddLine(ctrl.Width - 40, ctrl.Height, 40, ctrl.Height)

            p.AddArc(New Rectangle(0, ctrl.Height - CurveSize, CurveSize, CurveSize), 90, 90)
            p.CloseFigure()

            ctrl.Region = New Region(p)
            p.Dispose()

        Catch ex As Exception
            MsgBox(Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Information)
        End Try

    End Sub

End Module
