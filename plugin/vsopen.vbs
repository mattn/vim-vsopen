Set dte = GetObject("", "VisualStudio.DTE")
dte.MainWindow.Activate
dte.MainWindow.Visible = True
dte.UserControl = True
For i = 0 To WScript.Arguments.Count - 1
    dte.ItemOperations.OpenFile(WScript.Arguments(i))
Next

