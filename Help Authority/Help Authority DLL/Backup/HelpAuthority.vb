Namespace DynamicComponents
    Public Class HelpAuthority
        Inherits System.ComponentModel.Component

#Region " Component Designer generated code "

        Public Sub New(ByVal Container As System.ComponentModel.IContainer)
            MyClass.New()

            'Required for Windows.Forms Class Composition Designer support
            Container.Add(Me)
        End Sub

        Public Sub New()
            MyBase.New()

            'This call is required by the Component Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

        End Sub

        'Component overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Component Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Component Designer
        'It can be modified using the Component Designer.
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container()
        End Sub

#End Region

        Private Col_ControlName As New Collection()
        Private Col_ControlIndex As New Collection()
        Private Col_KeyValue As New Collection()
        Private Col_RequiredFields As New Collection()
        Private Col_GridKeyValue As New Collection()
        Private Col_FieldsType(30, 1) As String
        Private KeyLeavePos As Byte
        Private HasGrid As Boolean = False
        Private MyForm As New System.Windows.Forms.Form()
        Private MyGrid As New AxMSDataGridLib.AxDataGrid()
        Private HelpIdSender As String



        Private Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
            If str_String <> "" And int_Count <> 0 Then
                Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
            ElseIf int_Count = 0 Then
                Return str_String
            End If
        End Function


        Public Sub PrepareHelp(ByRef dm_DSN As ADODB.Connection, ByRef dm_Form As System.Windows.Forms.Form, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing)
            Dim MyProtect As New MyProtection()
            Dim ProductName As String


            
            ProductName = "DC Help Authority"
            MyProtect.SetInformation(ProductName)
            MyProtect.SetAlgorithms(1971, 18, 10, "maaat08")
            MyProtect.SetLicense(30)
            MyProtect.ShowAuthor()
            If MyProtect.NotLicensed Then
                MsgBox("Trial vesrion expired")
                Exit Sub
            End If



            Dim TxtCtrl As New Control()
            Dim X As Byte
            Dim Num As Byte
            Dim Num2 As Byte


            If Not (dm_Grid Is Nothing) Then
                HasGrid = True
            End If

            MyGrid = dm_Grid
            MyForm = dm_Form

            AddHandler dm_Form.Paint, AddressOf MyForm_Paint

            InitForm()
            ReadInitialValues()

            CN = dm_DSN
            X = 0

            For Each TxtCtrl In dm_Form.Controls
                If TypeName(TxtCtrl) = "TextBox" Then
                    Col_ControlName.Add(TxtCtrl)
                    Col_ControlIndex.Add(X)
                End If
                X += 1
            Next TxtCtrl


        End Sub


        Private Sub MyForm_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
            Static Flag As Byte = 0

            Flag += 1
            If Flag = 1 Then
                If Col_KeyValue.Count = 8 Then
                    Dim Num As Byte
                    For Num = 1 To 16
                        Col_KeyValue.Add(KeyLeavePos)
                    Next Num
                ElseIf Col_KeyValue.Count = 16 Then
                    Dim Num As Byte
                    For Num = 1 To 8
                        Col_KeyValue.Add(KeyLeavePos)
                    Next Num
                End If
                If Col_GridKeyValue.Count = 7 Then
                    Dim Num As Byte
                    For Num = 1 To 14
                        Col_GridKeyValue.Add(KeyLeavePos)
                    Next Num
                ElseIf Col_GridKeyValue.Count = 14 Then
                    Dim Num As Byte
                    For Num = 1 To 7
                        Col_GridKeyValue.Add(KeyLeavePos)
                    Next Num
                End If

            End If

            If HelpRtnID <> "" Then
                If UCase(sender.Controls(Col_KeyValue(8)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(8)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(1)).Text = HelpRtnName
                ElseIf UCase(sender.Controls(Col_KeyValue(16)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(16)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(9)).Text = HelpRtnName
                ElseIf UCase(sender.Controls(Col_KeyValue(24)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(24)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(17)).Text = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(6)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(6) + 2 <= MyGrid.Columns.Count Then
                        MyGrid.Col = Col_GridKeyValue(6) + 2
                    Else
                        MyGrid.Col = Col_GridKeyValue(6) - 2
                    End If
                    MyGrid.Columns(Col_GridKeyValue(6)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(6) + 1).Value = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(13)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(13) + 2 <= MyGrid.Columns.Count Then
                        MyGrid.Col = Col_GridKeyValue(13) + 2
                    Else
                        MyGrid.Col = Col_GridKeyValue(13) - 2
                    End If
                    MyGrid.Columns(Col_GridKeyValue(13)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(13) + 1).Value = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(20)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(20) + 2 <= MyGrid.Columns.Count Then
                        MyGrid.Col = Col_GridKeyValue(20) + 2
                    Else
                        MyGrid.Col = Col_GridKeyValue(20) - 2
                    End If
                    MyGrid.Columns(Col_GridKeyValue(20)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(20) + 1).Value = HelpRtnName
                End If
            End If

        End Sub

        Private Sub ControlHelp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F12 Then
                Dim oHelpFileForm As New HelpForm()
                Dim Num As Byte
                HelpFieldProperty(0) = sender.FindForm.Tag
                HelpFieldProperty(1) = sender.Name
                HelpFieldProperty(2) = ""
                For Num = 0 To Col_FieldsType.GetLength(0) - 1
                    If UCase(Col_FieldsType(Num, 0)) = UCase(sender.Name) Then
                        HelpFieldProperty(2) = Col_FieldsType(Num, 1)
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(3) = ""
                If TypeName(sender) = "TextBox" Then
                    Dim MyTextBox = CType(sender, TextBox)
                    HelpFieldProperty(3) = IIf(MyTextBox.MaxLength <> 0, MyTextBox.MaxLength, "")
                End If
                HelpFieldProperty(4) = ""
                For Num = 1 To Col_RequiredFields.Count
                    If UCase(MyForm.Controls(Col_RequiredFields(Num)).Name) = UCase(sender.Name) Then
                        HelpFieldProperty(4) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(5) = ""
                For Num = 5 To Col_KeyValue.Count Step 8
                    If UCase(Col_KeyValue(Num)) = UCase(sender.Name) Then
                        HelpFieldProperty(5) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(6) = TypeName(sender)
                oHelpFileForm.Show()
            End If
        End Sub

        Private Sub ColumnHelp_KeyDown(ByVal sender As Object, ByVal e As AxMSDataGridLib.DDataGridEvents_KeyDownEvent)
            If e.keyCode = Keys.F12 Then
                Dim oHelpFileForm As New HelpForm()
                Dim Num As Byte
                HelpFieldProperty(0) = sender.FindForm.Tag
                HelpFieldProperty(1) = sender.Name + "_" + sender.Columns(sender.Col).DataField
                HelpFieldProperty(2) = ""
                For Num = 0 To Col_FieldsType.GetLength(0) - 1
                    If UCase(Col_FieldsType(Num, 0)) = UCase(sender.Name) Then
                        HelpFieldProperty(2) = Col_FieldsType(Num, 1)
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(3) = ""
                If TypeName(sender) = "TextBox" Then
                    Dim MyTextBox = CType(sender, TextBox)
                    HelpFieldProperty(3) = IIf(MyTextBox.MaxLength <> 0, MyTextBox.MaxLength, "")
                End If
                HelpFieldProperty(4) = ""
                For Num = 1 To Col_RequiredFields.Count
                    If UCase(MyForm.Controls(Col_RequiredFields(Num)).Name) = UCase(sender.Name) Then
                        HelpFieldProperty(4) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(5) = ""
                For Num = 5 To Col_KeyValue.Count Step 8
                    If UCase(Col_KeyValue(Num)) = UCase(sender.Name) Then
                        HelpFieldProperty(5) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(6) = TypeName(sender)

                oHelpFileForm.Show()
            End If
        End Sub


        Private Sub InitForm()
            Dim Num As Integer = 0
            Dim MyControl As System.Windows.Forms.Control

            For Each MyControl In MyForm.Controls
                If TypeName(MyControl) <> "Label" Then
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf ControlHelp_KeyDown
                End If
                Num += 1
            Next MyControl

            AddHandler MyGrid.KeyDownEvent, AddressOf ColumnHelp_KeyDown

        End Sub

        Private Sub ReadInitialValues()
            Dim MyFile As String
            Dim MyFullPathFile As String
            Dim oInitValues As New InitValues()
            Dim FileNum As Byte

            Dim WinSys As String
            Dim WshShell As New Object()


            aInitValues(1) = ""
            aInitValues(2) = ""
            aInitValues(3) = ""
            aInitValues(4) = ""
            aInitValues(5) = ""
            aInitValues(6) = ""
            aInitValues(7) = ""
            aInitValues(8) = ""
            aInitValues(9) = ""
            aInitValues(10) = ""


            'cProperties = aInitValues(1) + IIf(HelpFieldProperty(2) <> "", HelpFieldProperty(2), aInitValues(5)) + Chr(13)
            'cProperties += aInitValues(2) + IIf(HelpFieldProperty(3) <> "", HelpFieldProperty(3).ToString + " " + aInitValues(6), aInitValues(5)) + Chr(13)
            'cProperties += aInitValues(3) + IIf(HelpFieldProperty(4) <> "", aInitValues(8), aInitValues(9)) + Chr(13)
            'cProperties += aInitValues(4) + IIf(HelpFieldProperty(5) <> "", aInitValues(8), aInitValues(9)) + Chr(13) + Chr(13)
            'cProperties += aInitValues(7) + Chr(13)


            WshShell = CreateObject("WScript.Shell")
            WinSys = WshShell.SpecialFolders("Fonts")
            WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
            WinSys += "System32\"
            FileNum = FreeFile()
            MyFullPathFile = WinSys + "DCHA30_Lang.dll"
            MyFile = Dir(WinSys + "DCHA30_Lang.dll")

            FileOpen(1, MyFullPathFile, OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared, 1000)
            If UCase(MyFile) = "DCHA30_LANG.DLL" Then
                FileGet(FileNum, oInitValues, 1)
                aInitValues(0) = oInitValues.Help_Caption.Trim + " "
                aInitValues(1) = oInitValues.Help_DataEntryType.Trim + " "
                aInitValues(2) = oInitValues.Help_MaxLenght.Trim + " "
                aInitValues(3) = oInitValues.Help_Required.Trim + " "
                aInitValues(4) = oInitValues.Help_HasDataHelp.Trim + " "
                aInitValues(5) = oInitValues.Help_Const_NotDefined.Trim + " "
                aInitValues(6) = oInitValues.Help_Const_Characters.Trim + " "
                aInitValues(7) = oInitValues.Help_Const_Description.Trim + " "
                aInitValues(8) = oInitValues.Help_Const_Yes.Trim + " "
                aInitValues(9) = oInitValues.Help_Const_No.Trim + " "
                ' Keep 5 room for Future add
                ' Keep 2 room for Future add
            End If
        End Sub

    End Class
End Namespace