Public Class SboTables
    Inherits CollectionBase
    ' Restricts to TableNames types, items that can be added to the collection.
    Public Sub Add(ByVal oTableName As enTableNamesType)
        ' Invokes Add method of the List object to add a widget.
        List.Add(oTableName)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        ' Check to see if there is a widget at the supplied index.
        If index > Count - 1 Or index < 0 Then
            ' If no widget exists, a messagebox is shown and the operation is 
            ' cancelled.
            ' System.Windows.Forms.MessageBox.Show("Index not valid!")
            Dim ex As New Exception("Index invalid")
            Throw ex
        Else
            ' Invokes the RemoveAt method of the List object.
            List.RemoveAt(index)
        End If
    End Sub

    ' This line declares the Item property as ReadOnly, and 
    ' declares that it will return a enTableNamesType object.
    Public ReadOnly Property Item(ByVal index As Integer) As enTableNamesType
        Get
            ' The appropriate item is retrieved from the List object and 
            ' explicitly cast to the enTableNamesType type, then returned to the 
            ' caller.
            Return CType(List.Item(index), enTableNamesType)
        End Get
    End Property

End Class
Public Class SboFields
    Inherits CollectionBase
    ' Restricts to TableNames types, items that can be added to the collection.
    Public Sub Add(ByVal oFieldName As enFieldNamesType)
        ' Invokes Add method of the List object to add a widget.
        List.Add(oFieldName)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        ' Check to see if there is a widget at the supplied index.
        If index > Count - 1 Or index < 0 Then
            ' If no widget exists, a messagebox is shown and the operation is 
            ' cancelled.
            ' System.Windows.Forms.MessageBox.Show("Index not valid!")
            Dim ex As New Exception("Index invalid")
            Throw ex
        Else
            ' Invokes the RemoveAt method of the List object.
            List.RemoveAt(index)
        End If
    End Sub

    ' This line declares the Item property as ReadOnly, and 
    ' declares that it will return a enTableNamesType object.
    Public ReadOnly Property Item(ByVal index As Integer) As enFieldNamesType
        Get
            ' The appropriate item is retrieved from the List object and 
            ' explicitly cast to the enTableNamesType type, then returned to the 
            ' caller.
            Return CType(List.Item(index), enFieldNamesType)
        End Get
    End Property

End Class

Public Class sboConditions
    Inherits CollectionBase
    ' Restricts to TableNames types, items that can be added to the collection.
    Public Sub Add(ByVal Condition0 As SAPbouiCOM.Condition)
        ' Invokes Add method of the List object to add a widget.
        List.Add(Condition0)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        ' Check to see if there is a widget at the supplied index.
        If index > Count - 1 Or index < 0 Then
            ' If no widget exists, a messagebox is shown and the operation is 
            ' cancelled.
            ' System.Windows.Forms.MessageBox.Show("Index not valid!")
            Dim ex As New Exception("Index invalid")
            Throw ex
        Else
            ' Invokes the RemoveAt method of the List object.
            List.RemoveAt(index)
        End If
    End Sub

    ' This line declares the Item property as ReadOnly, and 
    ' declares that it will return a enTableNamesType object.
    Public ReadOnly Property Item(ByVal index As Integer) As SAPbouiCOM.Condition
        Get
            ' The appropriate item is retrieved from the List object and 
            ' explicitly cast to the enTableNamesType type, then returned to the 
            ' caller.
            Return CType(List.Item(index), SAPbouiCOM.Condition)
        End Get
    End Property
End Class

Public Class SboUDOs
    Inherits CollectionBase
    ' Restricts to TableNames types, items that can be added to the collection.
    Public Sub Add(ByVal oUDOName As enUDONamesType)
        ' Invokes Add method of the List object to add a widget.
        List.Add(oUDOName)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        ' Check to see if there is a widget at the supplied index.
        If index > Count - 1 Or index < 0 Then
            ' If no widget exists, a messagebox is shown and the operation is 
            ' cancelled.
            ' System.Windows.Forms.MessageBox.Show("Index not valid!")
            Dim ex As New Exception("Index invalid")
            Throw ex
        Else
            ' Invokes the RemoveAt method of the List object.
            List.RemoveAt(index)
        End If
    End Sub

    ' This line declares the Item property as ReadOnly, and 
    ' declares that it will return a enTableNamesType object.
    Public ReadOnly Property Item(ByVal index As Integer) As enUDONamesType
        Get
            ' The appropriate item is retrieved from the List object and 
            ' explicitly cast to the enTableNamesType type, then returned to the 
            ' caller.
            Return CType(List.Item(index), enUDONamesType)
        End Get
    End Property

End Class
