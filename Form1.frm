VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rcs1 As Recordset
Dim dbs1 As Database
Dim i As Integer
Dim wrks As Workspace
Dim nom As String, sreq As String, nom1 As String
Dim tdf As TableDef
 '     Set wrks = DBEngine.Workspaces(0)
   
    nom1 = "Frm_bv2"
    nom = "c:\hydraulique\bo_v4\defchamps.mdb"
     Set dbs1 = OpenDatabase(nom)
'        For Each tdf In dbs1.TableDefs
'            If tdf.Name = "tempo" Then
'                dbs1.TableDefs.Delete "temp"
'            End If
'        Next
    sreq = "select *    from Feuil1    where form = '" & nom1 & "' ;"
     Set rcs1 = dbs1.OpenRecordset(sreq)
    If rcs1.RecordCount > 0 Then
        With rcs1
    rcs1.MoveFirst
    While Not rcs1.EOF
         Debug.Print .Fields(1).Value

        rcs1.MoveNext
    Wend


'    For i = 1 To .Fields.Count
'    Debug.Print .Fields(i - 1).Name
''        Debug.Print .Fields(i - 1).Name, .Fields(i - 1).Type, .Fields(i - 1).Size
'    Next
        End With
    End If
rcs1.Close
dbs1.Close
End Sub


