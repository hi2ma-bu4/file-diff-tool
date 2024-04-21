Const WIDTH = 500
Const HEIGHT = 150
Const BAR_WIDTH = 350
Const BAR_HEIGHT = 16
Const BAR_BG = "#C0C0C0"
Const BAR_FG = "#0066FF"

Class ProgressBar
    Private strTitle1
    Private nCurrent1
    Private nStartTime, nCurrentTime
    Private objIE
    Private div1, div2, div3

    Private Sub Class_Initialize
        Set objIE = CreateObject("InternetExplorer.Application")
        objIE.Visible = False
        objIE.Navigate2 "about:blank"
        objIE.Document.Title = "�i����"
        objIE.AddressBar = False
        objIE.MenuBar = False
        objIE.ToolBar = False
        objIE.Resizable = False
        objIE.Width = WIDTH
        objIE.Height = HEIGHT
        objIE.Top = 0
        objIE.Left = 0
        Set div1 = objIE.Document.CreateElement("div")
        div1.Id = "div1"
        div1.style.position = "absolute"
        div1.style.top = "10px"
        div1.style.left = "10px"
        div1.style.backgroundColor = BAR_BG
        div1.style.width = BAR_WIDTH & "px"
        div1.style.height = BAR_HEIGHT & "px"
        div1.style.border = "1px solid"
        div1.style.overflow = "hidden"
        Set div2 = objIE.Document.CreateElement("div")
        div2.Id = "div2"
        div2.style.position = "relative"
        div2.style.top = "1px"
        div2.style.left = "1px"
        div2.style.backgroundColor = BAR_FG
        div2.style.width = "0px"
        div2.style.height = (BAR_HEIGHT - 2) & "px"
        div2.style.overflow = "hidden"
        Set div3 = objIE.Document.CreateElement("div")
        div3.Id = "div3"
        div3.style.position = "absolute"
        div3.style.top = "45px"
        div3.style.left = "10px"

        objIE.Document.Body.AppendChild(div1)
        div1.AppendChild(div2)
        objIE.Document.Body.AppendChild(div3)

        nStartTime = Timer()
    End Sub

    Private Sub Class_Terminate
        on error resume next
        objIE.Quit
        Set objIE = Nothing
        on error goto 0
    End Sub

    Public Sub SetTitle (t)
        strTitle1 = t
        objIE.Document.Title = t & String(40, "�@")
        objIE.Visible = True
    End Sub

    Public Sub SetProgress(n1)
        nCurrent1 = n1
        Repaint
    End Sub

    Private Sub Repaint

    Dim nAverage
    Dim nElapsedTime
    Dim nRemain
    Dim strRemain
    Dim strPercent

    Dim w1
    Dim style1, style2

        nCurrentTime = Timer()
        nElapsedTime = nCurrentTime - nStartTime

        w1 = BAR_WIDTH * (nCurrent1)
        strRemain = "�s��"
        If nElapsedTime <> 0 Then
            nAverage = nCurrent1 / nElapsedTime
            If nAverage <> 0 Then
                nRemain = Round((1 - nCurrent1) / nAverage, 1)
            End If
            If nRemain > 60 Then
                strRemain = "��" & CStr(Round(nRemain / 60, 0)) & "��"
            Else
                strRemain = FormatNumber(nRemain, 1) & "�b"
            End If
        End If

        strPercent = FormatNumber(nCurrent1 * 100, 1)

        on error resume next
        div2.style.width = (w1 -1) & "px"
        div3.innerText = strPercent & "%�I���@--�@�c�萄��F" & strRemain
        objIE.Visible = True
        objIE.Document.all(0).Click
        on error goto 0
    End Sub
End Class