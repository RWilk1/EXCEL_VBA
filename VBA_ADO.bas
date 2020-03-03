Attribute VB_Name = "ADO"

' ActiveX Data Objects (AD0) Library
Sub CopyDataFromDatabase()

    ' Tworzenie zmiennych umo¿liwiaj¹cych po³¹cznie z baz¹ danych SQL
    Dim MoviesConn As ADODB.connection
    Dim MoviesData As ADODB.Recordset
       
    Set MoviesConn = New ADODB.connection
    Set MoviesData = New ADODB.Recordset
    
    MoviesConn.ConnectionString = _
    "Provider=SQLNCLI11;Server=RAFALWILK\RAFALWILKSQL;Database=Movies;Trusted_Connection=yes;"
    
    ' Otwieranie po³¹czenia
    MoviesConn.Open
    
    On Error GoTo CloseConnection
    
    With MoviesData
        .ActiveConnection = MoviesConn
'        .Source = "tblActor"
        .Source = "SELECT ActorName, ActorDOB, ActorGender FROM tblActor WHERE YEAR(ActorDOB) = '1980';"
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
        .Open
    End With
    
    On Error GoTo CloseRecordset
    
'    Worksheets.Add
'    Range("A2").CopyFromRecordset MoviesData
    
    ' Inny sposób pobierania danych z tabeli SQL
    Worksheets("Arkusz3").Range("A1").CopyFromRecordset MoviesData
    
CloseRecordset:
    MoviesData.Close
    
CloseConnection:
    MoviesConn.Close
    
' https://www.youtube.com/watch?v=-c2QvyPpkAM&list=PLNIs-AWhQzckr8Dgmgb3akx_gFMnpxTN5&index=48
 
End Sub

Sub CopyDataFromDatabase2()

    ' Tworzenie zmiennych umo¿liwiaj¹cych po³¹cznie z baz¹ danych SQL
    Dim MoviesConn As ADODB.connection
    Dim MoviesData As ADODB.Recordset
       
    Set MoviesConn = New ADODB.connection
    Set MoviesData = New ADODB.Recordset
    
    MoviesConn.ConnectionString = _
        "Provider=SQLOLEDB;" & _
        "Data Source=RAFALWILK\RAFALWILKSQL;" & _
        "Database=Movies;" & _
        "User ID=uid; Password=pwd;" & _
        "Trusted_Connection=yes;"
    ' "Provider=SQLNCLI11;" & _


    ' Otwieranie po³¹czenia
    MoviesConn.Open
    
    On Error GoTo CloseConnection
    
    With MoviesData
        .ActiveConnection = MoviesConn
'        .Source = "tblActor"
        .Source = "SELECT ActorName, ActorDOB, ActorGender FROM tblActor WHERE YEAR(ActorDOB) = '1980';"
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
        .Open
    End With
    
    On Error GoTo CloseRecordset
    
'    Worksheets.Add
'    Range("A2").CopyFromRecordset MoviesData
    
    ' Inny sposób pobierania danych z tabeli SQL
    Worksheets("Arkusz3").Range("A1").CopyFromRecordset MoviesData
    
CloseRecordset:
    MoviesData.Close
    
CloseConnection:
    MoviesConn.Close
    
' https://www.youtube.com/watch?v=-c2QvyPpkAM&list=PLNIs-AWhQzckr8Dgmgb3akx_gFMnpxTN5&index=48
 
End Sub


