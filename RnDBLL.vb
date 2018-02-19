''******************************************************************************************************
''* RnDBLL.vb
''* This Business Logic Layer was developed to bind data to a gridview and/or drop down list options 
''* based on business rules or user's criteria. Available options are Select, Insert, Update & Delete.
''*
''* Called From: Priorities.aspx - gvPriorities
''* Called From: TestingClassification.aspx - gvTestClass
''* Called From: TestIssuanceDetail.aspx - gvTMAssignments
''* Author  : LRey 05/18/2009
''* Added Supporting Documents  : Lalitha 02/18/2018
''******************************************************************************************************
Imports RnDDALTableAdapters

Public Class RnDBLL
#Region "Adapters"
    Private priorityAdapter As TestIssuance_PrioritiesTableAdapter = Nothing
    Protected ReadOnly Property Adapter1() As TestIssuance_PrioritiesTableAdapter
        Get
            If priorityAdapter Is Nothing Then
                priorityAdapter = New TestIssuance_PrioritiesTableAdapter()
            End If
            Return priorityAdapter
        End Get
    End Property


    Private classAdapter As Testing_ClassificationTableAdapter = Nothing
    Protected ReadOnly Property Adapter2() As Testing_ClassificationTableAdapter
        Get
            If classAdapter Is Nothing Then
                classAdapter = New Testing_ClassificationTableAdapter()
            End If
            Return classAdapter
        End Get
    End Property

    Private assignAdapter As TestIssuance_AssignmentsTableAdapter = Nothing
    Protected ReadOnly Property Adapter3() As TestIssuance_AssignmentsTableAdapter
        Get
            If assignAdapter Is Nothing Then
                assignAdapter = New TestIssuance_AssignmentsTableAdapter()
            End If
            Return assignAdapter
        End Get
    End Property

    Private pscpAdapter As TestIssuance_CustomerPartTableAdapter = Nothing
    Protected ReadOnly Property Adapter4() As TestIssuance_CustomerPartTableAdapter
        Get
            If pscpAdapter Is Nothing Then
                pscpAdapter = New TestIssuance_CustomerPartTableAdapter()
            End If
            Return pscpAdapter
        End Get
    End Property

    Private tcAdapter As TestIssuance_TestReportTableAdapter = Nothing
    Protected ReadOnly Property Adapter5() As TestIssuance_TestReportTableAdapter
        Get
            If tcAdapter Is Nothing Then
                tcAdapter = New TestIssuance_TestReportTableAdapter()
            End If
            Return tcAdapter
        End Get
    End Property
    Private pAdapter6 As Get_TestIssuance_Supporting_Doc_List1TableAdapter = Nothing
    Protected ReadOnly Property Adapter6() As Get_TestIssuance_Supporting_Doc_List1TableAdapter
        Get
            If pAdapter6 Is Nothing Then
                pAdapter6 = New Get_TestIssuance_Supporting_Doc_List1TableAdapter
            End If
            Return pAdapter6
        End Get
    End Property
#End Region 'EOF "Adapters"

#Region "Priorities"
    ''*****
    ''* Select TestIssuance_Priorities_Maint returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Select, True)> _
    Public Function GetPriorities(ByVal PriorityDescription As String) As RnDDAL.TestIssuance_PrioritiesDataTable

        Try
            If PriorityDescription = Nothing Then PriorityDescription = ""

            Return Adapter1.Get_TestIssuance_Priorities(PriorityDescription)

        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "PriorityDescription: " & PriorityDescription & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "GetPriorities : " & commonFunctions.convertSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/Prorities_Maint.aspx"

            UGNErrorTrapping.InsertErrorLog("GetPriorities : " & commonFunctions.convertSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try
    End Function
    ''*****
    ''* Insert New TestIssuance_Priorities_Maint
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Insert, True)> _
    Public Function InsertPriorities(ByVal ColorCode As String, ByVal PriorityDescription As String) As Boolean
        Try
            ' Create a new TestIssuance_Priorities_MaintRow instance
            Dim User As String = HttpContext.Current.Request.Cookies("UGNDB_User").Value

            ' Logical Rule - Cannot insert a record without a null  column
            If ColorCode = Nothing Then
                Throw New ApplicationException("Insert Cancelled: Color Code - is a required field.")
            End If
            If PriorityDescription = Nothing Then
                Throw New ApplicationException("Insert Cancelled: Prioritiy Description - is a required field.")
            End If

            ' Insert the new TestIssuance_Assignments row
            Dim rowsAffected As Integer = Adapter1.Insert_TestIssuance_Priorities(ColorCode, PriorityDescription, User)

            ' Return true if precisely one row was inserted, otherwise false
            Return rowsAffected = 1

        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "PriorityDescription: " & PriorityDescription & ", ColorCode:" & ColorCode & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "InsertPriorities : " & commonFunctions.convertSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/Prorities_Maint.aspx"

            UGNErrorTrapping.InsertErrorLog("InsertPriorities : " & commonFunctions.convertSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try
    End Function
    ''*****
    ''* Update TestIssuance_Priorities_Maint
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Insert, True)> _
    Public Function UpdatePriorities(ByVal PriorityDescription As String, ByVal Obsolete As Boolean, ByVal original_PID As Integer, ByVal ColorCode As String) As Boolean

        Try
            ' Create a new TestIssuance_Priorities_MaintRow instance
            Dim User As String = HttpContext.Current.Request.Cookies("UGNDB_User").Value

            ' Logical Rule - Cannot insert a record without a null column
            If ColorCode = Nothing Then
                Throw New ApplicationException("Insert Cancelled: Color Code - is a required field.")
            End If
            If PriorityDescription = Nothing Then
                Throw New ApplicationException("Insert Cancelled: Prioritiy Description - is a required field.")
            End If

            ' Insert the new TestIssuance_Assignments row
            Dim rowsAffected As Integer = Adapter1.Update_TestIssuance_Priorities(original_PID, ColorCode, PriorityDescription, Obsolete, User)

            ' Return true if precisely one row was inserted, otherwise false
            Return rowsAffected = 1

        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "PriorityDescription: " & PriorityDescription & ", ColorCode:" & ColorCode & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "UpdatePriorities : " & commonFunctions.convertSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/Prorities_Maint.aspx"

            UGNErrorTrapping.InsertErrorLog("UpdatePriorities : " & commonFunctions.convertSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try
    End Function

#End Region 'EOF "Priorities"

#Region "Testing Classification"
    ''*****
    ''* Select TestingClassification_Maint returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Select, True)> _
    Public Function GetTestingClassification(ByVal TestClassName As String) As RnDDAL.Testing_ClassificationDataTable

        Try
            If TestClassName = Nothing Then TestClassName = ""

            Return Adapter2.Get_Testing_Classification(TestClassName)
        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "TestClassName: " & TestClassName & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "GetTestingClassification : " & commonFunctions.replaceSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/TestingClass_Maint.aspx"

            UGNErrorTrapping.InsertErrorLog("GetTestingClassification : " & commonFunctions.replaceSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try

    End Function
    ''*****
    ''* Insert New TestingClassification_Maint
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Insert, True)> _
    Public Function InsertTestingClassification(ByVal TestClassName As String) As Boolean
        Try
            ' Create a new TestingClassification_MaintRow instance
            Dim User As String = HttpContext.Current.Request.Cookies("UGNDB_User").Value

            ' Logical Rule - Cannot insert a record without a null Subscriptions column
            If TestClassName = Nothing Then
                Throw New ApplicationException("Insert Cancelled: Testing Classification Name - is a required field.")
            End If

            ' Insert the new TestIssuance_Assignments row
            Dim rowsAffected As Integer = Adapter2.Insert_Testing_Classification(TestClassName, User)

            ' Return true if precisely one row was inserted, otherwise false
            Return rowsAffected = 1

        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "TestClassName: " & TestClassName & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "InsertTestingClassification : " & commonFunctions.replaceSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/TestingClass_Maint.aspx"

            UGNErrorTrapping.InsertErrorLog("InsertTestingClassification : " & commonFunctions.replaceSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try
    End Function
    ''*****
    ''* Update TestingClassification_Maint
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Insert, True)> _
    Public Function UpdateTestingClassification(ByVal TestClassName As String, ByVal Obsolete As Boolean, ByVal Original_TestClassID As Integer) As Boolean
        Try
            ' Create a new TestingClassification_MaintRow instance
            Dim User As String = HttpContext.Current.Request.Cookies("UGNDB_User").Value

            ' Logical Rule - Cannot insert a record without a null Subscriptions column
            If TestClassName = Nothing Then
                Throw New ApplicationException("Update Cancelled: Testing Classification Name - is a required field.")
            End If

            ' Insert the new TestIssuance_Assignments row
            Dim rowsAffected As Integer = Adapter2.Update_Testing_Classification(Original_TestClassID, TestClassName, Obsolete, User)

            ' Return true if precisely one row was inserted, otherwise false
            Return rowsAffected = 1


        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "TestClassName: " & TestClassName & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "UpdateTestingClassification : " & commonFunctions.replaceSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/TestingClass_Maint.aspx"

            UGNErrorTrapping.InsertErrorLog("UpdateTestingClassification : " & commonFunctions.replaceSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try

    End Function
#End Region 'EOF "Testing Classification"

#Region "Assignments"
    ''*****
    ''* Select TestIssuance_Assignments returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Select, True)> _
    Public Function GetTestIssuanceAssignments(ByVal RequestID As Integer) As RnDDAL.TestIssuance_AssignmentsDataTable

        Return Adapter3.GetData_TestIssuanceAssignments(RequestID)

    End Function

    ''*****
    ''* Insert a New row to TestIssuance_Assignments table
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Insert, True)> _
    Public Function InsertTestIssuanceAssignments(ByVal RequestID As Integer, ByVal TeamMemberID As Integer) As Boolean

        ' Create a new pscpRow instance
        Dim User As String = HttpContext.Current.Request.Cookies("UGNDB_User").Value

        ' Logical Rule - Cannot insert a record without null columns
        If TeamMemberID = Nothing Then
            Throw New ApplicationException("Insert Cancelled: Team Member is a required field.")
        End If

        ' Insert the new TestIssuance_Assignments row
        Dim rowsAffected As Integer = Adapter3.Insert_Test_Issuance_Assignments(RequestID, TeamMemberID, User)

        ' Return true if precisely one row was inserted, otherwise false
        Return rowsAffected = 1
    End Function

    ''*****
    ''* Delete TestIssuance_Assignments
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Delete, True)> _
    Public Function DeleteTestIssuanceAssignments(ByVal RequestID As Integer, ByVal TeamMemberID As Integer, ByVal original_RequestID As Integer, ByVal original_TeamMemberID As Integer) As Boolean

        Dim rowsAffected As Integer = Adapter3.Delete_Test_Issuance_Assignments(original_RequestID, original_TeamMemberID)

        ' Return true if precisely one row was deleted, otherwise false
        Return rowsAffected = 1

    End Function

#End Region 'EOF "Assignments"

#Region "Test Issuance Customer Part"
    ''*****
    ''* Select TestIssuance_CustomerPartNo returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Select, True)> _
    Public Function GetTestIssuanceCustomerPart(ByVal RequestID As Integer) As RnDDAL.TestIssuance_CustomerPartDataTable

        Try

            Return Adapter4.Get_TestIssuance_CustomerPart(RequestID)

        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "RequestID: " & RequestID & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value

            HttpContext.Current.Session("BLLerror") = "GetTestIssuanceCustomerPart : " & commonFunctions.convertSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData

            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RND/TestIssuanceList.aspx"

            UGNErrorTrapping.InsertErrorLog("GetTestIssuanceCustomerPart : " & commonFunctions.convertSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try
    End Function

    ''*****
    ''* Delete TestIssuance_CustomerPartNo
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Delete, True)> _
    Public Function DeleteTestIssuanceCustomerPart(ByVal RequestID As Integer, ByVal RowID As Integer, ByVal original_RequestID As Integer, ByVal original_RowID As Integer) As Boolean

        Dim rowsAffected As Integer = Adapter4.Delete_TestIssuance_CustomerPart(original_RequestID, original_RowID)

        ' Return true if precisely one row was deleted, otherwise false
        Return rowsAffected = 1

    End Function
#End Region 'EOF "Test Issuance Customer Part"

#Region "Test Report"
    ''*****
    ''* Select TestIssuance_TestReport returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Select, True)> _
    Public Function GetTestIssuanceTestReport(ByVal RequestID As Integer, ByVal TestReportID As Integer) As RnDDAL.TestIssuance_TestReportDataTable

        Return Adapter5.Get_TestIssuance_TestReport(RequestID, 0)
    End Function

    ''*****
    ''* Delete TestIssuance_TestReport
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Delete, True)> _
    Public Function DeleteTestIssuanceTestReport(ByVal TestReportID As Integer, ByVal RequestID As Integer, ByVal original_TestReportID As Integer, ByVal original_RequestID As Integer) As Boolean

        Dim rowsAffected As Integer = Adapter5.Delete_TestIssuance_TestReport(original_TestReportID, original_RequestID)

        ' Return true if precisely one row was deleted, otherwise false
        Return rowsAffected = 1

    End Function
#End Region 'EOF "Test Report"

#Region "TestIssuance Supporting Document"
    ''*****
    ''* Select TestIssuanceSupportingDoc returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Select, True)> _
    Public Function GetTestIssuanceSupportingDoc(ByVal RequestID As Integer) As RnDDAL.Get_TestIssuance_Supporting_Doc_ListDataTable

        Try
            If RequestID = Nothing Then RequestID = 0

            Return Adapter6.GetTestIssuanceSupportingDoc(RequestID)


        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "RequestID: " & RequestID & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value
            HttpContext.Current.Session("BLLerror") = "GetTestIssuanceSupportingDoc : " & commonFunctions.convertSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData
            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/TestIssuanceList.aspx"
            UGNErrorTrapping.InsertErrorLog("GetTestIssuanceSupportingDoc : " & commonFunctions.convertSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try

    End Function

    ''*****
    ''* Delete TestIssuanceSupportingDoc returning all rows
    ''*****
    <System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Delete, True)> _
    Public Function DeleteTestIssuanceSupportingDoc(ByVal RequestID As Integer, ByVal original_RowID As Integer) As Boolean

        Try


            Return Adapter6.sp_Delete_TestIssuance_Supporting_Doc(original_RowID, RequestID)

        Catch ex As Exception
            'on error, collect function data, error, and last page, then redirect to error page
            Dim strUserEditedData As String = "RowID: " & original_RowID & ", RequestID: " & RequestID & ", User: " & HttpContext.Current.Request.Cookies("UGNDB_User").Value
            HttpContext.Current.Session("BLLerror") = "DeleteTestIssuanceSupportingDoc : " & commonFunctions.convertSpecialChar(ex.Message, False) & " :<br/> RnDBLL.vb :<br/> " & strUserEditedData
            HttpContext.Current.Session("UGNErrorLastWebPage") = "~/RnD/TestIssuanceList.aspx"
            UGNErrorTrapping.InsertErrorLog("DeleteTestIssuanceSupportingDoc : " & commonFunctions.convertSpecialChar(ex.Message, False), "RnDBLL.vb", strUserEditedData)

            HttpContext.Current.Response.Redirect("~/UGNError.aspx", False)
            Return Nothing
        End Try

    End Function
#End Region 'EOF "Test Issuance Supporting Document"

End Class

