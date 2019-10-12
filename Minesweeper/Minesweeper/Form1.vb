

Public Class Form1

    Dim Count As New MyPoint(30, 16)
    Dim StartLoc As New MyPoint(5, 5)
    Dim EachSize As New MyPoint(50, 50)
    Dim MyBtn(,) As MyButton
    Dim Clicked As Boolean = False
    Dim BLOCK As Boolean = False
    Dim LblLeft As Label
    Dim BombCount As Integer = 99
    Dim MyDefaultBackColor As Color = Color.FromArgb(100, 100, 250)
    Function CountToLoc(ByRef C As MyPoint) As MyPoint
        Return New MyPoint(Me.StartLoc.X + C.X * EachSize.X, StartLoc.Y + C.Y * EachSize.Y)
    End Function
    Private Sub Fresh()
        Me.Controls.Clear()
        Clicked = False
        Me.Size = New Size(EachSize.X * Count.X + 2 * StartLoc.X + 20, EachSize.Y * Count.Y + 2 * StartLoc.Y + 40 + 50)
        ReDim MyBtn(Count.X - 1, Count.Y - 1)

        '畫圖
        For i = 0 To Count.X - 1
            For j = 0 To Count.Y - 1
                MyBtn(i, j) = New MyButton
                MyBtn(i, j).Size = New Size(EachSize.X, EachSize.Y)
                MyBtn(i, j).Count = New MyPoint(i, j)
                MyBtn(i, j).Location = CountToLoc(New MyPoint(i, j))
                MyBtn(i, j).Text = ""
                MyBtn(i, j).value = 0
                MyBtn(i, j).BackColor = MyDefaultBackColor
                MyBtn(i, j).Font = Font
                AddHandler MyBtn(i, j).MyClick, AddressOf MyBtn_MyClick
                AddHandler MyBtn(i, j).UnCover, AddressOf MyBtn_UnCover
                AddHandler MyBtn(i, j).RightClick, AddressOf MyBtn_RightClick
                AddHandler MyBtn(i, j).DualClick, AddressOf MyBtn_DualClick
                AddHandler MyBtn(i, j).DualPreClick, AddressOf MyBtn_DualPreClick
                Controls.Add(MyBtn(i, j))
                MyBtn(i, j).Show()
            Next j
        Next i

        LblLeft = New Label
        LblLeft.Text = Str(BombCount)
        LblLeft.Size = New Size(100, 50)
        LblLeft.Font = Font
        LblLeft.Location = New MyPoint(Me.StartLoc.X + Count.X * EachSize.X - 100, StartLoc.Y + Count.Y * EachSize.Y)
        Me.Controls.Add(LblLeft)
        LblLeft.Show()


    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MyDefaultBackColor
        Me.Fresh()
    End Sub

#Region "Click"

    Sub FirstClick(c As MyPoint)
        '埋雷
        Dim Str As New Starter(Count.X * Count.Y - Sides(c).Count - 1, BombCount)
        Dim l As List(Of MyPoint) = Sides(c)
        For Each i In MyBtn
            For Each j In l
                If i.Count = j Then GoTo NextI
            Next j
            If i.Count = c Or Sides(c).Contains(i.Count) Then Continue For
            If Str.Read Then i.value = -1
NextI:
        Next i

        '算雷

        For Each i In MyBtn
            If i.value = -1 Then Continue For
            For Each s In Sides(i.Count)
                If MyBtn(s.X, s.Y).value = -1 Then i.value += 1
            Next s
        Next i

    End Sub
    Sub MyBtn_MyClick(sender As Object, e As EventArgs)
        Dim s As MyButton = sender
        Dim c As MyPoint = CType(e, MyButtonClickEventArgs).Count
        If Clicked = False Then
            Clicked = True
            FirstClick(c)
        End If
        If s.BackColor <> Color.Chocolate Then s.UCover()
    End Sub
    Sub MyBtn_RightClick(sender As Object, e As EventArgs)
        Dim s As MyButton = sender
        Dim c As MyPoint = CType(e, MyButtonClickEventArgs).Count
        If s.BackColor = MyDefaultBackColor Then
            s.BackColor = Color.Chocolate
            LblLeft.Text = Str(Val(LblLeft.Text) - 1)
        ElseIf s.BackColor = Color.Chocolate Then
            s.BackColor = MyDefaultBackColor
            LblLeft.Text = Str(Val(LblLeft.Text) + 1)
        End If
    End Sub
    Sub MyBtn_DualClick(sender As Object, e As EventArgs)
        Dim s As MyButton = sender
        Dim c As MyPoint = CType(e, MyButtonClickEventArgs).Count
        For Each i In Sides(s.Count)
            If MyBtn(i.X, i.Y).BackColor = Color.LightBlue Then MyBtn(i.X, i.Y).BackColor = MyDefaultBackColor
        Next i
        If s.BackColor <> Color.White Then

            Exit Sub
        Else
            Dim x As Integer = 0
            For Each i In Sides(s.Count)
                If MyBtn(i.X, i.Y).BackColor = Color.Chocolate Then x += 1
            Next i
            If x = s.value Then
                For Each i In Sides(s.Count)
                    If MyBtn(i.X, i.Y).BackColor <> Color.Chocolate Then MyBtn(i.X, i.Y).UCover()
                Next i
            End If
        End If

    End Sub
    Sub MyBtn_DualPreClick(sender As Object, e As EventArgs)
        Dim s As MyButton = sender
        Dim c As MyPoint = CType(e, MyButtonClickEventArgs).Count
        For Each i In Sides(s.Count)
            If MyBtn(i.X, i.Y).BackColor = MyDefaultBackColor Then MyBtn(i.X, i.Y).BackColor = Color.LightBlue
        Next i
    End Sub

#End Region
    Sub MyBtn_UnCover(sender As Object, e As EventArgs)
        If BLOCK Then Exit Sub
        Dim s As MyButton = sender
        Dim c As MyPoint = CType(e, MyButtonUnCoverEventArgs).Count
        If s.Uncovered Then Exit Sub
        s.Uncovered = True
        Select Case s.value
            Case 0
                s.BackColor = Color.White
                For Each i In Sides(s.Count)
                    MyBtn(i.X, i.Y).UCover()
                Next i
            Case -1
                Defeat(s)
                Exit Sub
            Case Else
                s.Text = Str(s.value)
                s.BackColor = Color.White
        End Select
        For Each i In MyBtn
            If i.value = -1 Or i.BackColor = Color.White Then Continue For Else Exit Sub
        Next
        Succeed()
    End Sub
    Sub Defeat(s As MyButton)
        For Each i In MyBtn
            If i.value = -1 Then
                i.Text = "*"
            Else
                If i.BackColor = Color.Chocolate Then
                    i.Text = "X"
                    i.BackColor = Color.Red
                End If
            End If

        Next i
        s.BackColor = Color.Red
        Dim a As MsgBoxResult = MsgBox("您已失敗, 是否再來一局?", MsgBoxStyle.YesNo, "")
        If a = MsgBoxResult.Yes Then Fresh() Else BLOCK = True

    End Sub
    Sub Succeed()
        For Each i In MyBtn
            If i.value = -1 Then
                i.Text = "*"
            Else

            End If
        Next i
        Dim a As MsgBoxResult = MsgBox("您已完成, 是否再來一局?", MsgBoxStyle.YesNo, "")
        If a = MsgBoxResult.Yes Then Fresh() Else BLOCK = True

    End Sub

    Private Function BSides(ByVal P As MyPoint) As List(Of MyPoint)
        Dim ret As New List(Of MyPoint)
        ret.Clear()
        ret.Add(P.Left())
        ret.Add(P.Right())
        ret.Add(P.Up())
        ret.Add(P.Down())
        ret.Add(P.LeftUp())
        ret.Add(P.RightUp())
        ret.Add(P.LeftDown())
        ret.Add(P.RightDown())
        Return ret
    End Function
    Private Function Sides(ByVal P As MyPoint) As List(Of MyPoint)
        Dim ret As New List(Of MyPoint)
        ret.Clear()
        For Each i In BSides(P)
            If i.Valid(New MyPoint(Count)) Then ret.Add(i)
        Next
        Return ret
    End Function


End Class
Public Class MyPoint
    Public X As Integer, Y As Integer
    Private Sub New()

    End Sub
    Sub New(p As Point)
        Me.X = p.X
        Me.Y = p.Y
    End Sub
    Shared Widening Operator CType(p As MyPoint) As Point
        Return New Point(p.X, p.Y)
    End Operator

    Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    Function Valid(Ref As MyPoint)
        If X >= 0 And X < Ref.X And Y >= 0 And Y < Ref.Y Then Return True Else Return False
    End Function
    Function Left() As MyPoint
        Return New MyPoint(X - 1, Y)
    End Function
    Function Right() As MyPoint
        Return New MyPoint(X + 1, Y)
    End Function
    Function Up() As MyPoint
        Return New MyPoint(X, Y - 1)
    End Function
    Function Down() As MyPoint
        Return New MyPoint(X, Y + 1)
    End Function
    Function LeftUp()
        Return Me.Left().Up()
    End Function
    Function LeftDown()
        Return Me.Left().Down()
    End Function
    Function RightUp()
        Return Me.Right().Up()
    End Function
    Function RightDown()
        Return Me.Right().Down()
    End Function
    Public Shared Operator =(l As MyPoint, r As MyPoint)
        Return l.X = r.X And l.Y = r.Y
    End Operator
    Public Shared Operator <>(l As MyPoint, r As MyPoint)
        Return l.X <> r.X Or l.Y <> r.Y
    End Operator
End Class
Public Class MyButtonEventArgs
    Inherits EventArgs
    Sub New(c As MyPoint)
        Count = c
    End Sub
    Public Count As MyPoint
End Class
Public Class MyButtonClickEventArgs
    Inherits MyButtonEventArgs
    Sub New(c As MyPoint)
        MyBase.New(c)
    End Sub
End Class
Public Class MyButtonUnCoverEventArgs
    Inherits MyButtonEventArgs
    Sub New(c As MyPoint)
        MyBase.New(c)
    End Sub
End Class

Public Class MyButton
    Inherits Button
    Public value As Integer
    Public Count As MyPoint
    Public MouseLeft As Boolean = False, MouseRight As Boolean = False, DualClicked As Boolean = False
    Public Uncovered As Boolean = False
    Event MyClick(sender As MyButton, e As MyButtonClickEventArgs)
    Event UnCover(sender As MyButton, e As MyButtonUnCoverEventArgs)
    Event RightClick(sender As MyButton, e As MyButtonClickEventArgs)
    Event DualClick(sender As MyButton, e As MyButtonClickEventArgs)
    Event DualPreClick(sender As MyButton, e As MyButtonClickEventArgs)
    Public Sub MyBtn_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        Select Case e.Button
            Case MouseButtons.Left
                MouseLeft = False
                If MouseRight Then
                    RaiseEvent DualClick(Me, New MyButtonClickEventArgs(Me.Count))
                Else
                    If Not DualClicked Then RaiseEvent MyClick(Me, New MyButtonClickEventArgs(Me.Count))
                End If

            Case MouseButtons.Right
                MouseRight = False
                If MouseLeft Then
                    RaiseEvent DualClick(Me, New MyButtonClickEventArgs(Me.Count))
                Else

                End If
        End Select
    End Sub
    Public Sub MyBtn_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        Select Case e.Button
            Case MouseButtons.Left
                MouseLeft = True
                If MouseRight Then
                    DualClicked = True
                    RaiseEvent DualPreClick(Me, New MyButtonClickEventArgs(Me.Count))
                Else
                    DualClicked = False
                End If

            Case MouseButtons.Right
                MouseRight = True
                If MouseLeft Then
                    DualClicked = True
                    RaiseEvent DualPreClick(Me, New MyButtonClickEventArgs(Me.Count))
                Else
                    RaiseEvent RightClick(Me, New MyButtonClickEventArgs(Me.Count))
                End If
        End Select
    End Sub
    Public Sub MClick()
        RaiseEvent MyClick(Me, New MyButtonClickEventArgs(Me.Count))
    End Sub
    Public Sub UCover()
        RaiseEvent UnCover(Me, New MyButtonUnCoverEventArgs(Me.Count))
    End Sub
    Public Sub RClick()
        RaiseEvent RightClick(Me, New MyButtonClickEventArgs(Me.Count))
    End Sub
    Public Sub DClick()
        RaiseEvent DualClick(Me, New MyButtonClickEventArgs(Me.Count))
    End Sub
End Class


Class Starter
    Dim LeftCount As Integer
    Dim Seed As Integer
    Dim LeftReal As Integer
    Dim r As Random
    Sub New(all As Integer, real As Integer, Optional s As Integer = 0)
        Randomize()
        If s = 0 Then s = Now.Ticks Mod Integer.MaxValue
        Me.LeftCount = all
        Me.LeftReal = real
        r = New Random(s)
    End Sub
    Function Read() As Boolean
        If LeftReal = 0 Then Return False
        Dim a As Double = r.NextDouble
        If a <= CDbl(LeftReal) / LeftCount Then
            LeftCount -= 1
            LeftReal -= 1
            Return True
        Else
            LeftCount -= 1
            Return False
        End If
    End Function
End Class