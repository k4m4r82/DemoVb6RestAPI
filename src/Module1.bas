Attribute VB_Name = "Module1"
Option Explicit

' untuk mendapatkan api key daftar di http://rajaongkir.com/akun/daftar
Private Const RAJA_ONGKIR_API_KEY As String = ""

Public Sub Main()
    GetProvinsi
    GetKabupaten
    Call GetCost
End Sub

Private Sub GetCost()
    Dim item        As Scripting.Dictionary
    Dim cost        As Scripting.Dictionary
    Dim costDetail  As Scripting.Dictionary
    Dim objJson     As Object
    Dim apiURL      As String
    Dim postData    As String
    Dim postResult  As String
    
    apiURL = "http://api.rajaongkir.com/starter/cost"
    
    postData = "origin=501&destination=114&weight=1700&courier=jne"
    postResult = PostRequest(apiURL, postData, RAJA_ONGKIR_API_KEY)

    Debug.Print "postResult : " & postResult & vbCrLf
    
    Set objJson = ModJSON.parse(postResult)
    
    Debug.Print "Info kota pengirim: Prov ID: " & objJson.item("rajaongkir").item("origin_details").item("province_id") & _
                ", Prov Name: " & objJson.item("rajaongkir").item("origin_details").item("province") & ", " & _
                "Kab ID: " & objJson.item("rajaongkir").item("origin_details").item("city_id") & ", Kab Name: " & _
                objJson.item("rajaongkir").item("origin_details").item("city_name") & vbCrLf
    
    Debug.Print "Info kota tujuan: Prov ID: " & objJson.item("rajaongkir").item("destination_details").item("province_id") & _
                ", Prov Name: " & objJson.item("rajaongkir").item("destination_details").item("province") & ", " & _
                "Kab ID: " & objJson.item("rajaongkir").item("destination_details").item("city_id") & ", Kab Name: " & _
                objJson.item("rajaongkir").item("destination_details").item("city_name") & vbCrLf
    
    Debug.Print "Info layanan dan biaya:"
    For Each item In objJson.item("rajaongkir").item("results")
        For Each cost In item.item("costs")
            Set costDetail = cost.item("cost")(1)
            Debug.Print "--> service: " & cost.item("service") & ", description: " & cost.item("description") & _
                        ", biaya: " & costDetail.item("value") & ", estimasi (hari): " & costDetail.item("etd")
        Next cost
    Next item
    
End Sub

Private Sub GetKabupaten()
    Dim kabupaten   As Scripting.Dictionary
    Dim objJson     As Object
    Dim apiURL      As String
    Dim jsonResult  As String
    
    Dim provID      As String
    provID = "1" ' bali
    
    apiURL = "http://api.rajaongkir.com/starter/city?id=&province=" & provID
    jsonResult = GetRequest(apiURL, RAJA_ONGKIR_API_KEY)
    
    Debug.Print "jsonResult : " & jsonResult & vbCrLf
    
    Debug.Print "Daftar Kabupaten Bali"
    Debug.Print "================================================"
    Set objJson = ModJSON.parse(jsonResult)
    For Each kabupaten In objJson.item("rajaongkir").item("results")
        Debug.Print "Kab ID: " & kabupaten.item("city_id") & ", Kab Name: " & kabupaten.item("city_name") & ", Postal Code: " & kabupaten.item("postal_code")
    Next kabupaten
End Sub

Private Sub GetProvinsi()
    Dim prov        As Scripting.Dictionary
    Dim objJson     As Object
    Dim apiURL      As String
    Dim jsonResult  As String
    
    apiURL = "http://api.rajaongkir.com/starter/province?id="
    jsonResult = GetRequest(apiURL, RAJA_ONGKIR_API_KEY)
    
    Debug.Print "jsonResult : " & jsonResult & vbCrLf
    
    Debug.Print "Daftar Provinsi"
    Debug.Print "================================================"
    Set objJson = ModJSON.parse(jsonResult)
    For Each prov In objJson.item("rajaongkir").item("results")
        Debug.Print "Prov ID: " & prov.item("province_id") & ", Prov Name: " & prov.item("province")
    Next prov
End Sub

Public Function GetRequest(url As String, ByVal Key As String) As String
    Dim http As MSXML2.XMLHTTP
    
    On Error GoTo errHandler
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    http.Open "GET", url, False
    http.setRequestHeader "key", Key
    http.send

    GetRequest = http.responseText
    Set http = Nothing
    
    Exit Function
errHandler:
    
End Function

Public Function PostRequest(ByVal url As String, ByVal postData As String, ByVal Key As String) As String
    Dim http As MSXML2.XMLHTTP
    
    On Error GoTo errHandler
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    http.Open "POST", url, False
    http.setRequestHeader "key", Key
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    http.send postData

    PostRequest = http.responseText
    Set http = Nothing
    
    Exit Function
errHandler:
End Function
