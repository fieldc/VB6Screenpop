Module VBComtestConsole
    Public Class Callback
        Implements SSFIntegration.IRecordRequestCallback
        Public Sub Callback(clientId As String, taxId As String, contractId As String,contractCount As Int32) Implements SSFIntegration.IRecordRequestCallback.Callback
            Console.WriteLine(vbTab & "[Callback.Callback] got event: clientId=" & clientId & " taxId=" & taxId & " contractId=" & contractId + " contractId=" + contractCount.ToString())
        End Sub
    End Class

    Sub Main()
        Dim obj_h As SSFIntegration.IRecordRequest

        Try
            obj_h = CreateObject("SSFIntegration.SSFComMessageProxy")
        Catch ex As Exception
            obj_h = Nothing
        End Try

        If obj_h Is Nothing Then
            System.Console.WriteLine("Error Creating Object")
            System.Console.ReadKey()
            Exit Sub
        End If

        obj_h.RegisterCallBack(New Callback())

        Dim docontinue As Boolean
        docontinue = True

        Dim clientId As String
        Dim taxId As String
        Dim contractId As String

        Do While docontinue
            System.Console.WriteLine("Press 'q' to quit")
            System.Console.Write("ClientId: ")
            clientId = System.Console.ReadLine().Trim()
            If clientId.ToLower().Equals("q") Then
                docontinue = False
                Exit Do
            End If
            System.Console.Write("TaxId: ")
            taxId = System.Console.ReadLine().Trim()
            If taxId.ToLower().Equals("q") Then
                docontinue = False
                Exit Do
            End If

            System.Console.Write("ContractId: ")
            contractId = System.Console.ReadLine().Trim()
            If contractId.ToLower().Equals("q") Then
                docontinue = False
                Exit Do
            End If
            System.Console.WriteLine("Raising events clientId=" & clientId & " taxId=" & taxId & " contractId=" & contractId & " count=1")
            obj_h.RaiseRecordRequested(clientId, taxId, contractId, "1")
        Loop

    End Sub
End Module
