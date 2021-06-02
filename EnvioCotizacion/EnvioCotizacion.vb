Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net.Mail

Public Class EnvioCotizacion

    Friend WithEvents SBOApplication As SAPbouiCOM.Application
    Public SBOCompany As SAPbobsCOM.Company
    Dim pdf, xml, pdfSAP, xmlSAP As String

    Public Function Cotizacion(ByVal DocNum As String, ByVal Tipo As String, ByVal psDirectory As String)

        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String
        Dim DocEntry, CardCode, CardName, DocDate, EmailC As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Call EnvioCorreo_SemiAutomatico('" & DocNum & "','" & Tipo & "')"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                DocEntry = oRecSettxb.Fields.Item("DocEntry").Value
                CardCode = oRecSettxb.Fields.Item("CardCode").Value
                CardName = oRecSettxb.Fields.Item("CardName").Value
                DocDate = oRecSettxb.Fields.Item("CreateDate").Value
                EmailC = oRecSettxb.Fields.Item("E_Mail").Value

                ExportarPDF(DocEntry, Tipo, DocNum, psDirectory)
                ValidarDoc(DocNum, Tipo, DocDate, CardCode, EmailC)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en la funcion Cotizacion. " & ex.Message)

        End Try

    End Function


    Public Function ExportarPDF(ByVal DocEntry As String, ByVal Tipo As String, ByVal DocNum As String, ByVal psDirectory As String)

        'MsgBox("Consulta de Documentos exitosa")
        Dim reportDocument As ReportDocument
        Dim diskFileDestinationOption As DiskFileDestinationOptions

        Try

            reportDocument = New ReportDocument

            reportDocument.Load(psDirectory + "\Informes\" + Tipo + ".rpt")

            Dim count As Integer = reportDocument.DataSourceConnections.Count
            reportDocument.DataSourceConnections(0).SetLogon(My.Settings.DbUserName, My.Settings.DbPassword)

            reportDocument.SetParameterValue(0, DocEntry)

            reportDocument.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            reportDocument.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
            diskFileDestinationOption = New DiskFileDestinationOptions

            diskFileDestinationOption.DiskFileName = My.Settings.Ruta & "\" & Tipo & "\" & DocNum & ".pdf"

            reportDocument.ExportOptions.ExportDestinationOptions = diskFileDestinationOption
            reportDocument.ExportOptions.ExportFormatOptions = New PdfRtfWordFormatOptions

            reportDocument.Export()
            'MsgBox("Exportacion de Documento Exitosa")
            reportDocument.Close()
            reportDocument.Dispose()
            GC.SuppressFinalize(reportDocument)

        Catch ex As Exception

            SBOApplication.MessageBox("Error en la funcion ExportarPDF. " & ex.Message)

        End Try

    End Function


    Public Function ValidarDoc(ByVal DocNum As String, ByVal Tipo As String, ByVal DocDate As Date, ByVal CardCode As String, ByVal EmailC As String)

        'MsgBox("Exportar Documento Exitoso")
        Dim Ruta As String
        pdf = Nothing
        xml = Nothing
        pdfSAP = Nothing
        xmlSAP = Nothing

        Try

            Ruta = My.Settings.Ruta & "\" & Tipo

            Dim dir As New System.IO.DirectoryInfo(Ruta)

            Dim fileList = dir.GetFiles("*.pdf", System.IO.SearchOption.TopDirectoryOnly)

            Dim FileQuery = From file In fileList
                            Where file.Extension = ".pdf" And file.Name.Trim.ToString.EndsWith(DocNum & ".pdf") And file.Name.Trim.ToString.StartsWith(DocNum & ".pdf")
                            Order By file.CreationTime
                            Select file

            pdf = Ruta & "\" & DocNum & ".pdf"

            If FileQuery.Count > 0 Then

                If EmailC <> "" Then

                    EnviarCorreo(DocNum, EmailC, pdf, Tipo, CardCode)

                Else

                    Dim stError As String
                    stError = "El socio de negocios no tiene asignado un correo electronico"
                    'Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en la funcion ValidarDoc. " & ex.Message)

        End Try

    End Function


    Public Function EnviarCorreo(ByVal DocNum As String, ByVal EmailC As String, ByVal pdf As String, ByVal Tipo As String, ByVal CardCode As String)

        'MsgBox("Validacion de Documentos exitosa")
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim oRecSettxb, oRecSettxb1, oRecSettxb2 As SAPbobsCOM.Recordset
        Dim stQuerytxb, stQuerytxb1, stQuerytxb2 As String
        Dim EmailU, Pass, EmailCC, Subject, Body, smtpService, Puerto, SegSSL As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Select ""U_Email"",""U_Password"",""U_EmailCC"",""U_Subject"",""U_Body"",""U_SMTP"",""U_Puerto"",""U_SeguridadSSL"" from ""@CORREOTEKNO"" where ""Name""='" & Tipo & "'"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                EmailU = oRecSettxb.Fields.Item("U_Email").Value
                Pass = oRecSettxb.Fields.Item("U_Password").Value
                EmailCC = oRecSettxb.Fields.Item("U_EmailCC").Value

                Subject = oRecSettxb.Fields.Item("U_Subject").Value
                Body = oRecSettxb.Fields.Item("U_Body").Value
                smtpService = oRecSettxb.Fields.Item("U_SMTP").Value
                Puerto = oRecSettxb.Fields.Item("U_Puerto").Value
                SegSSL = oRecSettxb.Fields.Item("U_SeguridadSSL").Value

                'Limpiamos correo destinatario, correo copia y archivos adjuntos
                message.To.Clear()
                message.CC.Clear()
                message.Attachments.Clear()

                'Llenamos encabezado de correo
                message.From = New MailAddress(EmailU)
                EmailC = ArreglarTexto(EmailC, ";", ",")
                message.To.Add(EmailC)
                If EmailCC.Count > 0 Then
                    message.CC.Add(EmailCC)
                End If
                message.Subject = Subject & DocNum

                'Llenamos el cuerpo del correo y prioridad
                message.Body = Body
                message.Priority = MailPriority.Normal

                'Adjuntamos archivos pdf
                Dim attpdf As New Net.Mail.Attachment(pdf)
                message.Attachments.Add(attpdf)

                'Llenamos datos de smtp
                smtp.Host = smtpService
                smtp.Credentials = New Net.NetworkCredential(EmailU, Pass)
                smtp.Port = Puerto
                smtp.EnableSsl = SegSSL

                'Enviamos Correo
                smtp.Send(message)

                If Tipo = "ORDR" Then

                    oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    stQuerytxb2 = "Update OINV set ""U_PDF""='" & pdf & "' where ""DocNum""=" & DocNum
                    oRecSettxb2.DoQuery(stQuerytxb2)

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en la funcion EnviarCorreo. " & ex.Message)

        End Try

    End Function


    Public Function ArreglarTexto(ByVal TextoOriginal As String, ByVal QuitarCaracter As String, ByVal PonerCaracter As String)

        TextoOriginal = TextoOriginal.Replace(QuitarCaracter, PonerCaracter)
        Return TextoOriginal

    End Function


End Class
