Public Class About

    Private Sub lblLink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblLink.LinkClicked
        Process.Start(String.Format("mailto:{0}", "greglalonde@conservative.ca?subject=CIMS/AD Account Simplifier Bug"))
    End Sub
End Class