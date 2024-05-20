Imports System.Xml
Public Class Form1
    ' This program writes one record to an XML output file
    Dim strName As String
    Dim strCountry As String
    Dim strBirthdate As String
    Dim strMessage As String
    Dim strCommon As String
    Dim strBotanical As String
    Dim strZone As String
    Dim strPrice As String
    Dim strLight As String
    Dim strAvailability As String
    Dim strAuthor As String
    Dim strTitle As String
    Dim strGenre As String
    Dim strPriceB As String
    Dim strPublish As String
    Dim strDescription As String
    '-- Declare arrays
    Dim Author(11) As String
    Dim Title(11) As String
    Dim Genre(11) As String
    Dim PriceB(11) As Double
    Dim Publish(11) As String
    Dim Description(11) As String
    Dim totalValue As Double = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        strName = txtName.Text
        strCountry = txtCountry.Text
        strBirthdate = txtBirthdate.Text
        'Dim writer As New XmlTextWriter("C:\Users\02161730\Desktop\Players.xml", System.Text.Encoding.UTF8)
        Dim writer As New XmlTextWriter("Players.xml", System.Text.Encoding.UTF8)
        writer.Formatting = Formatting.Indented
        writer.indentation = 2
        writer.writeStartElement("Records")
        CreateXMLRec(strName, strCountry, strBirthdate, writer)

        writer.writeEndElement()

        writer.close()
        MsgBox("Player record saved successfully")
    End Sub
    Private Function CreateXMLRec(ByVal IncomingName As String, ByVal IncomingCountry As String, ByVal IncomingBirthdate As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Player")

        writer.WriteStartElement("Name")
        writer.WriteString(IncomingName)
        writer.WriteEndElement()

        writer.WriteStartElement("Country")
        writer.WriteString(IncomingCountry)
        writer.WriteEndElement()

        writer.WriteStartElement("Birthdate")
        writer.WriteString(IncomingBirthdate)
        writer.WriteEndElement()

        writer.WriteEndElement()

    End Function

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        m_xmld = New XmlDocument
        'm_xmld.Load("C:\Users\02161730\Desktop\Players.xml")
        m_xmld.Load("Players.xml")
        m_nodelist = m_xmld.GetElementsByTagName("Player")
        For Each m_node In m_nodelist
            strName = m_node.Item("Name").InnerText
            strCountry = m_node.Item("Country").InnerText
            strBirthdate = m_node.Item("Birthdate").InnerText

            strMessage = "Name:      " & strName & vbCrLf & "Country:      " & strCountry & vbCrLf & "Birthdate:     " & strBirthdate
            MsgBox(strMessage)
        Next


    End Sub

    Private Sub Show2_Click(sender As Object, e As EventArgs) Handles Show2.Click
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        m_xmld = New XmlDocument
        'm_xmld.Load("C:\Users\02161730\Desktop\Players.xml")
        m_xmld.Load("Players.xml")
        m_nodelist = m_xmld.GetElementsByTagName("Player")
        For Each m_node In m_nodelist
            strName = m_node.Item("Name").InnerText
            strCountry = m_node.Item("Country").InnerText
            strBirthdate = m_node.Item("Birthdate").InnerText
            lsbNames.Items.Add(strName)
            lsbCountires.Items.Add(strCountry)
            lsbBirthdates.Items.Add(strBirthdate)

        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        m_xmld = New XmlDocument
        'm_xmld.Load("C:\Users\02161730\Desktop\2020 Classes\2020 SD\XML_PlantCatalog.xml")
        m_xmld.Load("XML_PlantCatalog.xml")
        m_nodelist = m_xmld.GetElementsByTagName("PLANT")
        For Each m_node In m_nodelist
            strCommon = m_node.Item("COMMON").InnerText
            strBotanical = m_node.Item("BOTANICAL").InnerText
            strZone = m_node.Item("ZONE").InnerText
            strLight = m_node.Item("LIGHT").InnerText
            strPrice = m_node.Item("PRICE").InnerText
            strAvailability = m_node.Item("AVAILABILITY").InnerText
            lsbCommon.Items.Add(strCommon)
            lsbBotanical.Items.Add(strBotanical)
            lsbZone.Items.Add(strZone)
            lsbLight.Items.Add(strLight)
            lsbPrice.Items.Add(strPrice)
            lsbAvail.Items.Add(strAvailability)

        Next
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        lsbNames.Items.Clear()
        lsbCountires.Items.Clear()
        lsbBirthdates.Items.Clear()
        lsbCommon.Items.Clear()
        lsbBotanical.Items.Clear()
        lsbZone.Items.Clear()
        lsbLight.Items.Clear()
        lsbPrice.Items.Clear()
        lsbAvail.Items.Clear()
        lsbAuthor.Items.Clear()
        lsbTitle.Items.Clear()
        lsbGenre.Items.Clear()
        lsbPriceB.Items.Clear()
        lsbPublish.Items.Clear()
        lsbDescription.Items.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        m_xmld = New XmlDocument
        'm_xmld.Load("C:\Users\02161730\Desktop\2020 Classes\2020 SD\XML_PlantCatalog.xml")
        m_xmld.Load("XML_BookCatalog.xml")
        m_nodelist = m_xmld.GetElementsByTagName("book")
        Dim i As Integer = 0
        For Each m_node In m_nodelist
            strAuthor = m_node.Item("author").InnerText
            strTitle = m_node.Item("title").InnerText
            strGenre = m_node.Item("genre").InnerText
            strPriceB = m_node.Item("price").InnerText
            strPublish = m_node.Item("publish_date").InnerText
            strDescription = m_node.Item("description").InnerText
            lsbAuthor.Items.Add(strAuthor)
            lsbTitle.Items.Add(strTitle)
            lsbGenre.Items.Add(strGenre)
            lsbPriceB.Items.Add(strPriceB)
            lsbPublish.Items.Add(strPublish)
            lsbDescription.Items.Add(strDescription)
            '--read items into arrays
            Author(i) = strAuthor
            Title(i) = strTitle
            Genre(i) = strGenre
            'PriceB(i) = CDbl(strPriceB)
            PriceB(i) = Val(strPriceB)
            Publish(i) = strPublish
            Description(i) = strDescription
            totalValue = totalValue + PriceB(i)
            i = i + 1
        Next
        MsgBox("The total value is: " & totalValue)
        Dim lowestPrice As Double = 999
        For i = 0 To 11
            If PriceB(i) < lowestPrice Then lowestPrice = PriceB(i)
        Next
        MsgBox("The lowest price  is: " & lowestPrice)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        strName = txtName.Text
        strCountry = txtCountry.Text
        strBirthdate = txtBirthdate.Text


        Dim cr As String = Environment.NewLine
        Dim newPlayer As String = _
         "<Player>" & cr & _
    "    <Name>" & strName & "</Name>" & cr & _
    "    <Country>" & strCountry & "</Country>" & cr & _
    "    <Birthdate>" & strBirthdate & "</Birthdate>" & cr & _
    "  </Player>"
        Dim xd As New XmlDocument()
        xd.Load("Players.xml")

        ' Create a new XmlDocumentFragment for our document.
        Dim docFrag As XmlDocumentFragment = xd.CreateDocumentFragment()
        ' The Xml for this fragment is our newPerson string.
        docFrag.InnerXml = newPlayer
        ' The root element of our file is found using
        ' the DocumentElement property of the XmlDocument.
        Dim root As XmlNode = xd.DocumentElement
        ' Append our new Person to the root element.
        root.AppendChild(docFrag)

        'Save the Xml.
        xd.Save("Players.xml")


        MsgBox("Player record saved successfully")
    End Sub

    Private Sub btnDeleteNode_Click(sender As Object, e As EventArgs) Handles btnDeleteNode.Click

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        m_xmld = New XmlDocument
        m_xmld.Load("Players.xml")
        m_nodelist = m_xmld.GetElementsByTagName("Player")
        Dim p As Integer = 0
        For Each m_node In m_nodelist
            strName = m_node.Item("Name").InnerText
            strCountry = m_node.Item("Country").InnerText
            strBirthdate = m_node.Item("Birthdate").InnerText
            lsbNames.Items.Add(strName)
            lsbCountires.Items.Add(strCountry)
            lsbBirthdates.Items.Add(strBirthdate)
        Next

        Dim itemNumberToDelete As Integer = InputBox("Enter item number to delete fromxml file", "Delete Item")
        Dim t As Integer = 0
        For Each m_node In m_nodelist
            If t = itemNumberToDelete Then
                m_node.ParentNode.RemoveChild(m_node)
                Exit For
            End If
            t = t + 1
        Next
        
        ' Save the Xml.
        m_xmld.Save("Players.xml")
    End Sub
End Class
