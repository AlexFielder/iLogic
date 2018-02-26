Attribute VB_Name = "WebBrowserDialog"

Private oWebBrowserDialog As WebBrowserDialog

Sub DismissWebBrowser()
    Set oWebBrowserDialog = Nothing
End Sub

Sub WebBrowserDialogSample()
    ' Create a WebBrowserDialog
    Set oWebBrowserDialog = ThisApplication.WebBrowserDialogs.Add("MyBrowser", False)
    oWebBrowserDialog.WindowState = kNormalWindow
    
    ' Nagigate to a web site
    Call oWebBrowserDialog.Navigate("http://www.autodesk.com")

    ' Play a tutorial video if you have the Interactive Tutorial installed
    ' Call oWebBrowserDialog.Navigate("C:\Users\Public\Documents\Autodesk\Inventor 2017\Interactive Tutorial\en-US\Fundamentals\Video\Drawings.webm")

    ' Delete it - commenteted
    ' oWebBrowserDialog.Delete
End Sub

