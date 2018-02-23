# UW
VBA code
Sub Run()
    'Set ie = CreateObject("internetExplorer.application")
    Dim ie As New SHDocVw.InternetExplorer
    ie.Visible = True
    ie.navigate "en.wikipedia.org/wiki/Main_Page"
    Do While ie.readyState <> READYSTATE_COMPLETE
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
    Loop
    Debug.Print ie.LocationName, ie.LocationURL
    
    ie.document.forms("searchform").elements("search").Value = "Document object model"
    ie.document.forms("searchform").elements("go").Click
    
End Sub

