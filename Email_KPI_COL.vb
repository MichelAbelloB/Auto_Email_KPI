Sub KPI_COL()
    Dim objOutlook As Object
    Dim objMail As Object
    Dim strRecipients As String
    Dim strCC As String
    
    'Crea una instancia de Outlook
    Set objOutlook = CreateObject("Outlook.Application")
    
    'Crea un nuevo correo electrónico
    Set objMail = objOutlook.CreateItem(olMailItem)
    
    'Agrega los destinatarios separados por punto y coma
    strRecipients = Range("E3").Value
    objMail.To = strRecipients
    
    'Agrega los destinatarios en copia separados por punto y coma
    strCC = Range("E4").Value
    objMail.CC = strCC
    
    'Calcula la fecha del día anterior
    Dim fechaAnterior As Date
    Dim fechaHoy As Date
    fechaAnterior = DateAdd("d", -1, Date)
    fechaHoy = DateAdd("d", 0, Date)

    'Agrega el asunto y el cuerpo del correo
    objMail.Subject = "Reporting | Campaña | " & Format(fechaHoy, "dd.MM.yyyy") & " | Campaña| Dashboard Campaña | Insights |"
    
   'KPI a SCORE
    PP = Format(Range("E8").Value * 100, "0.00") + "%"
    PT = Format(Range("E7").Value * 100, "0.00") + "%"
    CR = Format(Range("E6").Value * 100, "0.00") + "%"
    RPC = Format(Range("E9").Value * 100, "0.00") + "%"
    LINK = Range("E10").Value
    
    'KPI a NOMBRE
    PP_N = Range("F8").Value
    PT_N = Range("F7").Value
    CR_N = Range("F6").Value
    RPC_N = Range("F9").Value
    
    'KPI a INDICADOR
    PP_T = Format(Range("G8").Value * 100, "0.00") + "%"
    PT_T = Format(Range("G7").Value * 100, "0.00") + "%"
    CR_T = Format(Range("G6").Value * 100, "0.00") + "%"
    RPC_T = Format(Range("G9").Value * 100, "0.00") + "%"
    
    'KPI a UNION
    PP_U = PP & " " & PP_N & " " & PP_T
    PT_U = PT & " " & PT_N & " " & PT_T
    CR_U = CR & " " & CR_N & " " & CR_T
    RPC_U = RPC & " " & RPC_N & " " & RPC_T
    
    Dim table As String
    table = "<table width='100%' border='1' style='border-collapse:collapse;'><tr><td colspan='2' style='background-color:#0066c4;color:white;text-align:center;' style='font-weight: bold;'>KPI COL</td></tr><tr><td style='font-weight: bold;'> ID </td><td> Campaña </tr><tr><td style='font-weight: bold;'> Campaña </td><td> Campaña </td></tr><tr><td style='font-weight: bold;'> LOB </td><td> KPI COL </td></tr><tr><td style='font-weight: bold;'> Fecha de Insight</td><td> " & Format(fechaAnterior, "dd/MM/yyyy") & "</td></tr></tr><tr><td style='font-weight: bold;'> Indicadores</td><td> PT, PP, RPC, Total Conversion</td></tr><tr><td style='font-weight: bold;'> Insight</td><td> Para el dia " & Format(fechaAnterior, "dd mmmm yyyy") & "; Promise to Pay es de " & PP_U & ", Payment Taken de " & PT_U & " , Conversion Rate de " & CR_U & " y RPC " & RPC_U & ".</td></tr><tr><td style='font-weight: bold;'> Link de informe</td><td> " & LINK & "</td></tr></table>"

    
    'Crea el cuerpo del correo en formato HTML
    Dim strBody As String
    strBody = "<html><body>" _
            & "<img src='C:\Users\report.png' alt='Banner'>" _
            & "<br>" _
            & table _
            & "<br>" _
            & "<p>Reporting | Campaña | " & Format(fechaAnterior, "dd.MM.yyyy") & " | Campaña | Dashboard Campaña | Insights</p>" _
            & "<img src='C:\Users\End_Insight.jpeg' alt='Banner'>" _
            & "<br>" _
            & "<b>Greetings,</b>" _
            & "<br>" _
            & "<p>Michel Enrique Abello Betancourt</p>" _
            & "<b>Reporting Analyst | Bogotá D.C, Colombia </b>" _
            & "<img src='C:\Users\banner.jpg' alt='Firma'>" _
            & "<p style='font-size: 8pt'>The information contained in this communication is privileged and confidential. The content is intended only for the use of the individual or entity named above. If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, distribution or copying of this communication is strictly prohibited. If you have received this communication in error, please notify me immediately by telephone or e-mail, and delete this message from your systems. Please consider the environmental impact of needlessly printing this e-mail. </p>" _
            & "</body></html>"
    
    'Agrega el cuerpo del correo en formato HTML
    objMail.HTMLBody = strBody

    'Envía el correo
    objMail.Send
    
    MsgBox "Correo Enviado a: " & strRecipients & strCC
    
    
    'Libera la memoria
    Set objMail = Nothing
    Set objOutlook = Nothing
End Sub