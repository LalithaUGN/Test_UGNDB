' ************************************************************************************************
' Name:	        TestIssuance_Supporting_Doc_Viewer.vb
' Purpose:	    This code is used to show all PDF Files inside popup windows
' Called From : TestIssuanceDetail.aspx
'
'' Date		       Author	    
'' 02/16/2018      Lalitha Jampana			Created .Net application
' ************************************************************************************************
Partial Class RnD_TestIssuance_Document
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            If Not Page.IsPostBack Then
                Dim strSupportingFileName As String = ""

                If HttpContext.Current.Request.QueryString("RowID") <> "" Then
                    ViewState("RowID") = CType(HttpContext.Current.Request.QueryString("RowID"), Integer)
                End If

                If HttpContext.Current.Request.QueryString("RequestID") <> "" Then
                    ViewState("RequestID") = CType(HttpContext.Current.Request.QueryString("RequestID"), Integer)
                End If

                If ViewState("RequestID") > 0 And ViewState("RowID") > 0 Then
                    Dim ds As DataSet = RnDModule.GetTestIssuanceSupportingDoc(ViewState("RowID"), ViewState("RequestID"))
                    If commonFunctions.CheckDataSet(ds) = True Then

                        If ds.Tables(0).Rows(0).Item("SupportingDocBinary") IsNot System.DBNull.Value Then

                            strSupportingFileName = ds.Tables(0).Rows(0).Item("SupportingDocName").ToString

                            If strSupportingFileName.Trim = "" Then
                                strSupportingFileName = "TestIssuance-SupportingDoc.pdf"
                            End If

                            Dim imagecontent As Byte() = DirectCast(ds.Tables(0).Rows(0).Item("SupportingDocBinary"), Byte())
                            Response.Clear()
                            Response.Buffer = True
                            Response.ContentType = ds.Tables(0).Rows(0).Item("SupportingDocEncodeType").ToString()

                            'avoid the prompt if PDF of JPEF
                            If ds.Tables(0).Rows(0).Item("SupportingDocEncodeType").ToString() = "application/pdf" _
                                Or ds.Tables(0).Rows(0).Item("SupportingDocEncodeType").ToString() = "image/pjpeg" Then
                                Response.AddHeader("Content-Disposition", "inline;filename=" & strSupportingFileName)
                            Else
                                Response.AddHeader("Content-Disposition", "attachment;filename=" & strSupportingFileName)
                            End If

                            Response.OutputStream.Write(imagecontent, 0, imagecontent.Length - 1)
                            Response.Flush()
                            Response.Close()
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

            'update error on web page
            lblErrors.Text = ex.Message

            'get current event name
            Dim mb As Reflection.MethodBase = Reflection.MethodBase.GetCurrentMethod

            'log and email error
            UGNErrorTrapping.UpdateUGNErrorLog(mb.Name & ":" & ex.Message, System.Web.HttpContext.Current.Request.Url.AbsolutePath)
        End Try

    End Sub

End Class
