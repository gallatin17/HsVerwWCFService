﻿' HINWEIS: Mit dem Befehl "Umbenennen" im Kontextmenü können Sie den Klassennamen "Service1" sowohl im Code als auch in der SVC-Datei und der Konfigurationsdatei ändern.
Public Class Service1
    Implements IService1

    Public Sub New()
    End Sub

    Public Function GetVerbrauch() As IEnumerable(Of IService1.Verbrauch) Implements IService1.GetVerbrauch
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_gesamtverbrauch As New List(Of IService1.Verbrauch)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()
        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND (Haushaltskategorie_ID = 1 OR Haushaltskategorie_ID = 4) ORDER BY Haushaltsunterkategorie_ID;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_verbrauch As New IService1.Verbrauch
            vlo_verbrauch.ID = vlo_row.Item("ID_Werte")
            vlo_verbrauch.Wert = vlo_row.Item("Anzahl")
            vlo_verbrauch.Datum = vlo_row.Item("Datum")
            vlo_verbrauch.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_verbrauch.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_verbrauch.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_verbrauch.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_verbrauch.Einheit = vlo_row.Item("Einheit")
            vlo_verbrauch.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_verbrauch.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_verbrauch.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_verbrauch.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")
            vlo_gesamtverbrauch.Add(vlo_verbrauch)
        Next
        Conn.Close()
        Return vlo_gesamtverbrauch

    End Function
    Public Function GetAusgabe() As IEnumerable(Of IService1.Ausgabe) Implements IService1.GetAusgabe
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_gesamtausgaben As New List(Of IService1.Ausgabe)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND (Haushaltskategorie_ID = 2 OR Haushaltskategorie_ID = 5) ORDER BY Haushaltsunterkategorie_ID;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_ausgabe As New IService1.Ausgabe
            vlo_ausgabe.ID = vlo_row.Item("ID_Werte")
            vlo_ausgabe.Wert = vlo_row.Item("Anzahl")
            vlo_ausgabe.Datum = vlo_row.Item("Datum")
            vlo_ausgabe.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_ausgabe.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_ausgabe.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_ausgabe.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_ausgabe.Einheit = vlo_row.Item("Einheit")
            vlo_ausgabe.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_ausgabe.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_ausgabe.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_ausgabe.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")
            vlo_gesamtausgaben.Add(vlo_ausgabe)
        Next
        Conn.Close()
        Return vlo_gesamtausgaben
    End Function

    Public Function GetVerbrauchbyTyp(ByVal verbrauchstyp As Long) As IEnumerable(Of IService1.Verbrauch) Implements IService1.GetVerbrauchbyTyp
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_gesamtverbrauch As New List(Of IService1.Verbrauch)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND Haushaltskategorie_ID = " & verbrauchstyp & " ORDER BY Haushaltsunterkategorie_ID, Datum;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_verbrauch As New IService1.Verbrauch
            vlo_verbrauch.ID = vlo_row.Item("ID_Werte")
            vlo_verbrauch.Wert = vlo_row.Item("Anzahl")
            vlo_verbrauch.Datum = vlo_row.Item("Datum")
            vlo_verbrauch.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_verbrauch.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_verbrauch.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_verbrauch.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_verbrauch.Einheit = vlo_row.Item("Einheit")
            vlo_verbrauch.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_verbrauch.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_verbrauch.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_verbrauch.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")
            vlo_gesamtverbrauch.Add(vlo_verbrauch)
        Next
        Conn.Close()
        Return vlo_gesamtverbrauch
    End Function

    Public Function GetVerbrauchbyUnterTyp(ByVal verbrauchsuntertyp As Long) As IEnumerable(Of IService1.Verbrauch) Implements IService1.GetVerbrauchbyUnterTyp
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_gesamtverbrauch As New List(Of IService1.Verbrauch)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND Haushaltsunterkategorie_ID = " & verbrauchsuntertyp & " ORDER BY Datum;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_verbrauch As New IService1.Verbrauch
            vlo_verbrauch.ID = vlo_row.Item("ID_Werte")
            vlo_verbrauch.Wert = vlo_row.Item("Anzahl")
            vlo_verbrauch.Datum = vlo_row.Item("Datum")
            vlo_verbrauch.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_verbrauch.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_verbrauch.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_verbrauch.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_verbrauch.Einheit = vlo_row.Item("Einheit")
            vlo_verbrauch.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_verbrauch.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_verbrauch.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_verbrauch.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")
            vlo_gesamtverbrauch.Add(vlo_verbrauch)
        Next
        Conn.Close()
        Return vlo_gesamtverbrauch
    End Function

    Public Function GetVerbrauchsTyp(ByVal Haushatskategorie_ID As Long) As IEnumerable(Of IService1.Verbrauchstyp) Implements IService1.GetVerbrauchsTyp
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_verbrauchstypen As New List(Of IService1.Verbrauchstyp)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Haushaltsunterkategorie, Haushaltsunterkategorie,Haushaltskategorie_ID FROM tbl_haushaltsunterkategorie WHERE Haushaltskategorie_ID = " & Haushatskategorie_ID & " ORDER BY Haushaltsunterkategorie;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_verbrauchstyp As New IService1.Verbrauchstyp
            vlo_verbrauchstyp.ID = vlo_row.Item("ID_Haushaltsunterkategorie")
            vlo_verbrauchstyp.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_verbrauchstyp.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_verbrauchstypen.Add(vlo_verbrauchstyp)

        Next
        Conn.Close()
        Return vlo_verbrauchstypen
    End Function

    Public Function GetAusgabebyTyp(ByVal ausgabentyp As Long) As IEnumerable(Of IService1.Ausgabe) Implements IService1.GetAusgabebyTyp
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_gesamtausgabe As New List(Of IService1.Ausgabe)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND Haushaltskategorie_ID = " & ausgabentyp & " ORDER BY Haushaltsunterkategorie_ID;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_ausgabe As New IService1.Ausgabe
            vlo_ausgabe.ID = vlo_row.Item("ID_Werte")
            vlo_ausgabe.Wert = vlo_row.Item("Anzahl")
            vlo_ausgabe.Datum = vlo_row.Item("Datum")
            vlo_ausgabe.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_ausgabe.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_ausgabe.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_ausgabe.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_ausgabe.Einheit = vlo_row.Item("Einheit")
            vlo_ausgabe.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_ausgabe.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_ausgabe.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_ausgabe.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")
            vlo_gesamtausgabe.Add(vlo_ausgabe)
        Next
        Conn.Close()
        Return vlo_gesamtausgabe
    End Function
    Public Function GetVerbrauchbyID(ByVal ID As Long) As IService1.Verbrauch Implements IService1.GetVerbrauchbyID
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_verbrauch As New IService1.Verbrauch
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND ID_Werte = " & ID & " ORDER BY Haushaltsunterkategorie_ID;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows

            vlo_verbrauch.ID = vlo_row.Item("ID_Werte")
            vlo_verbrauch.Wert = vlo_row.Item("Anzahl")
            vlo_verbrauch.Datum = vlo_row.Item("Datum")
            vlo_verbrauch.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_verbrauch.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_verbrauch.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_verbrauch.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_verbrauch.Einheit = vlo_row.Item("Einheit")
            vlo_verbrauch.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_verbrauch.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_verbrauch.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_verbrauch.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")
        Next
        Conn.Close()
        Return vlo_verbrauch
    End Function
    Public Function GetAusgabebyID(ByVal ID As Long) As IService1.Ausgabe Implements IService1.GetAusgabebyID
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_ausgabe As New IService1.Ausgabe
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie,Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor, ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, tbl_einheit WHERE Einheit_ID = ID_Einheit AND Zahlungsrythmus_ID = ID_Zahlungsrythmus AND Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie AND ID_Haushaltskategorie = Haushaltskategorie_ID AND ID_Werte = " & ID & " ORDER BY Haushaltsunterkategorie_ID;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows

            vlo_ausgabe.ID = vlo_row.Item("ID_Werte")
            vlo_ausgabe.Wert = vlo_row.Item("Anzahl")
            vlo_ausgabe.Datum = vlo_row.Item("Datum")
            vlo_ausgabe.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_ausgabe.HaushaltskategorieID = vlo_row.Item("Haushaltskategorie_ID")
            vlo_ausgabe.Haushaltsunterkategorie = vlo_row.Item("Haushaltsunterkategorie")
            vlo_ausgabe.HaushaltsunterkategorieID = vlo_row.Item("Haushaltsunterkategorie_ID")
            vlo_ausgabe.Einheit = vlo_row.Item("Einheit")
            vlo_ausgabe.EinheitID = vlo_row.Item("ID_Einheit")
            vlo_ausgabe.Zahlungsrythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_ausgabe.Zahlungsrythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_ausgabe.ZahlungsrythmusID = vlo_row.Item("ID_Zahlungsrythmus")

        Next
        Conn.Close()
        Return vlo_ausgabe
    End Function
    Public Function GetVerbrauchbyDate(ByVal datum As Date) As IService1.Verbrauch Implements IService1.GetVerbrauchbyDate
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Conn = Nothing
        Return Nothing
    End Function
    Public Function GetAusgabebyDate(ByVal datum As Date) As IService1.Ausgabe Implements IService1.GetAusgabebyDate
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Conn = Nothing
        Return Nothing
    End Function

    Public Function GetZahlungsrythmen() As IEnumerable(Of IService1.Zahlungsrythmus) Implements IService1.GetZahlungsrythmen
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_zahlungsrythmen As New List(Of IService1.Zahlungsrythmus)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()
        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Zahlungsrythmus, Zahlungsrythmus, Rythmusfaktor FROM tbl_zahlungsrythmus ORDER BY ID_Zahlungsrythmus;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_zahlungsrythmus As New IService1.Zahlungsrythmus
            vlo_zahlungsrythmus.Rythmus = vlo_row.Item("Zahlungsrythmus")
            vlo_zahlungsrythmus.Rythmusfaktor = vlo_row.Item("Rythmusfaktor")
            vlo_zahlungsrythmus.ID = vlo_row.Item("ID_Zahlungsrythmus")
            vlo_zahlungsrythmen.Add(vlo_zahlungsrythmus)
        Next
        Conn.Close()
        Return vlo_zahlungsrythmen

    End Function

    Public Function GetVarVerbrauchKat() As IEnumerable(Of IService1.VarVerbrauchKat) Implements IService1.GetVarVerbrauchKat
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_varverbrauchkats As New List(Of IService1.VarVerbrauchKat)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()
        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Haushaltsunterkategorie, Haushaltsunterkategorie FROM tbl_haushaltsunterkategorie WHERE Haushaltskategorie_ID = 1 ORDER BY Haushaltsunterkategorie;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_varverbrauchkat As New IService1.VarVerbrauchKat
            vlo_varverbrauchkat.VarVerbrauchKat = vlo_row.Item("Haushaltsunterkategorie")
            vlo_varverbrauchkat.ID = vlo_row.Item("ID_Haushaltsunterkategorie")
            vlo_varverbrauchkats.Add(vlo_varverbrauchkat)
        Next
        Conn.Close()
        Return vlo_varverbrauchkats

    End Function

    Public Function SetVerbrauch(ByVal vlo_verbrauch As IService1.Verbrauch) As Boolean Implements IService1.SetVerbrauch
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim myconnstring As String = ""

        SetVerbrauch = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter

        Try
            SetVerbrauch = "UPDATE tbl_werte SET Haushaltsunterkategorie_ID = " & vlo_verbrauch.HaushaltsunterkategorieID & ", Anzahl = " & vlo_verbrauch.Wert & ", Datum = '" & vlo_verbrauch.Datum & "', Bemerkung = '' WHERE ID_Werte =  " & vlo_verbrauch.ID & ";"
            adp_KVI_mysql.UpdateCommand = New MySql.Data.MySqlClient.MySqlCommand("UPDATE tbl_werte SET Haushaltsunterkategorie_ID = " & vlo_verbrauch.HaushaltsunterkategorieID & ", Anzahl = " & vlo_verbrauch.Wert & ", Datum = '" & vlo_verbrauch.Datum & "', Bemerkung = '' WHERE ID_Werte =  " & vlo_verbrauch.ID & ";", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
            adp_KVI_mysql.UpdateCommand.ExecuteNonQuery()
        Catch ex As Exception
            SetVerbrauch = "FEHLER " & ex.Message
        End Try

        adp_KVI_mysql.Dispose()
        Conn.Close()
        Return True
    End Function
    Public Function SetAusgabe(ByVal vlo_ausgabe As IService1.Ausgabe) As String Implements IService1.SetAusgabe
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim myconnstring As String = ""

        SetAusgabe = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter

        Try
            SetAusgabe = "UPDATE tbl_werte SET Haushaltsunterkategorie_ID = " & vlo_ausgabe.HaushaltsunterkategorieID & ", Anzahl = " & vlo_ausgabe.Wert & ", Datum = '" & vlo_ausgabe.Datum.ToString("yyy-MM-dd") & "', Bemerkung = '' WHERE ID_Werte =  " & vlo_ausgabe.ID & ";"
            adp_KVI_mysql.UpdateCommand = New MySql.Data.MySqlClient.MySqlCommand("UPDATE tbl_werte SET Haushaltsunterkategorie_ID = " & vlo_ausgabe.HaushaltsunterkategorieID & ", Anzahl = " & vlo_ausgabe.Wert & ", Datum = '" & vlo_ausgabe.Datum.ToString("yyy-MM-dd") & "', Bemerkung = '' WHERE ID_Werte =  " & vlo_ausgabe.ID & ";", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
            adp_KVI_mysql.UpdateCommand.ExecuteNonQuery()
        Catch ex As Exception
            SetAusgabe = "FEHLER " & ex.Message
        End Try

        'Nur dann Zahlungsrythmus wegschreiben wenn notwendig (0 = variable Ausgaben ohne Rythmus Einkäufe usw.)
        If vlo_ausgabe.ZahlungsrythmusID <> 0 Then

            Try
                SetAusgabe = "UPDATE tbl_haushaltsunterkategorie SET Zahlungsrythmus_ID = " & vlo_ausgabe.ZahlungsrythmusID & " WHERE ID_Haushaltsunterkategorie =  " & vlo_ausgabe.HaushaltsunterkategorieID & ";"
                adp_KVI_mysql.UpdateCommand = New MySql.Data.MySqlClient.MySqlCommand("UPDATE tbl_haushaltsunterkategorie SET Zahlungsrythmus_ID = " & vlo_ausgabe.ZahlungsrythmusID & " WHERE ID_Haushaltsunterkategorie =  " & vlo_ausgabe.HaushaltsunterkategorieID & ";", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
                adp_KVI_mysql.UpdateCommand.ExecuteNonQuery()
            Catch ex As Exception
                SetAusgabe = "FEHLER " & ex.Message
            End Try

        End If


        adp_KVI_mysql.Dispose()
        Conn.Close()
        Return SetAusgabe

    End Function
    Public Function SetVerbrauchNew(ByVal vlo_verbrauch As IService1.Verbrauch) As Boolean Implements IService1.SetVerbrauchNew
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim myconnstring As String = ""
        Dim vlo_id As Long = 0

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet

        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT MAX(ID_Werte) AS MAXID FROM tbl_werte;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        vlo_id = get_daten.Tables(0).Rows(0).Item("MAXID") + 1
        adp_KVI_mysql.InsertCommand = New MySql.Data.MySqlClient.MySqlCommand("INSERT INTO tbl_werte (ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Bemerkung) VALUES(" & vlo_id & "," & vlo_verbrauch.HaushaltsunterkategorieID & "," & vlo_verbrauch.Wert & ",'" & vlo_verbrauch.Datum.ToString("yyy-MM-dd") & "','');", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.InsertCommand.ExecuteNonQuery()

        adp_KVI_mysql.Dispose()
        Conn.Close()
        Return True
    End Function
    Public Function SetAusgabeNew(ByVal vlo_ausgabe As IService1.Ausgabe) As String Implements IService1.SetAusgabeNew
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim myconnstring As String = ""
        Dim vlo_id1 As Long = 0
        Dim vlo_id2 As Long = 0

        SetAusgabeNew = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet

        SetAusgabeNew = "SELECT MAX(ID_Haushaltsunterkategorie) AS MAXID FROM tbl_haushaltsunterkategorie;"

        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand(SetAusgabeNew, CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        'Neue Unterkategorie (wg. Zahlungsrythmus)
        vlo_id1 = get_daten.Tables(0).Rows(0).Item("MAXID") + 1

        Try
            SetAusgabeNew = "INSERT INTO tbl_haushaltsunterkategorie (ID_Haushaltsunterkategorie, Haushaltsunterkategorie, Haushaltskategorie_ID, Einheit_ID, Zahlungsrythmus_ID,Bemerkung) VALUES(" & vlo_id1 & ",'" & vlo_ausgabe.Haushaltsunterkategorie & "'," & vlo_ausgabe.HaushaltskategorieID & "," & vlo_ausgabe.EinheitID & "," & vlo_ausgabe.ZahlungsrythmusID & ",'');"
            adp_KVI_mysql.InsertCommand = New MySql.Data.MySqlClient.MySqlCommand(SetAusgabeNew, CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
            adp_KVI_mysql.InsertCommand.ExecuteNonQuery()
        Catch ex As Exception
            SetAusgabeNew = "FEHLER " & ex.Message
        End Try
        

        get_daten.Clear()

        Try
            SetAusgabeNew = "SELECT MAX(ID_Werte) AS MAXID FROM tbl_werte;"
            adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand(SetAusgabeNew, CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
            adp_KVI_mysql.Fill(get_daten)
        Catch ex As Exception
            SetAusgabeNew = "FEHLER " & ex.Message
        End Try
       

        'Neuen Wert eintragen
        vlo_id2 = get_daten.Tables(0).Rows(0).Item("MAXID") + 1

        Try
            SetAusgabeNew = "INSERT INTO tbl_werte (ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Bemerkung) VALUES(" & vlo_id2 & "," & vlo_id1 & "," & vlo_ausgabe.Wert & ",'" & vlo_ausgabe.Datum.ToString("yyy-MM-dd") & "','');"
            adp_KVI_mysql.InsertCommand = New MySql.Data.MySqlClient.MySqlCommand(SetAusgabeNew, CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
            adp_KVI_mysql.InsertCommand.ExecuteNonQuery()
        Catch ex As Exception
            SetAusgabeNew = "FEHLER " & ex.Message
        End Try
       
        adp_KVI_mysql.Dispose()
        Conn.Close()

        Return SetAusgabeNew
    End Function
    Public Function DeleteVerbrauch(ByVal vlo_verbrauch As IService1.Verbrauch) As Boolean Implements IService1.DeleteVerbrauch
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter

        adp_KVI_mysql.DeleteCommand = New MySql.Data.MySqlClient.MySqlCommand("DELETE FROM tbl_werte WHERE ID_Werte =  " & vlo_verbrauch.ID & ";", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.DeleteCommand.ExecuteNonQuery()

        adp_KVI_mysql.Dispose()
        Conn.Close()
        Return True
    End Function
    Public Function DeleteAusgabe(ByVal vlo_ausgabe As IService1.Ausgabe) As Boolean Implements IService1.DeleteAusgabe
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter

        'Wert löschen
        adp_KVI_mysql.DeleteCommand = New MySql.Data.MySqlClient.MySqlCommand("DELETE FROM tbl_werte WHERE ID_Werte =  " & vlo_ausgabe.ID & ";", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.DeleteCommand.ExecuteNonQuery()

        'Unterkategorie löschen
        adp_KVI_mysql.DeleteCommand = New MySql.Data.MySqlClient.MySqlCommand("DELETE FROM tbl_haushaltsunterkategorie WHERE ID_Haushaltsunterkategorie =  " & vlo_ausgabe.HaushaltsunterkategorieID & ";", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.DeleteCommand.ExecuteNonQuery()

        adp_KVI_mysql.Dispose()
        Conn.Close()
        Return True
    End Function
    Public Function GetHaushaltskategorien() As IEnumerable(Of IService1.Haushaltskategorie) Implements IService1.GetHaushaltskategorien
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_haushaltskategorien As New List(Of IService1.Haushaltskategorie)
        Dim myconnstring As String = ""

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()

        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT ID_Haushaltskategorie, Haushaltskategorie FROM tbl_haushaltskategorie;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
        adp_KVI_mysql.Fill(get_daten)

        adp_KVI_mysql.Dispose()

        For Each vlo_row As DataRow In get_daten.Tables(0).Rows
            Dim vlo_haushaltskategorie As New IService1.Haushaltskategorie
            vlo_haushaltskategorie.ID = vlo_row.Item("ID_Haushaltskategorie")
            vlo_haushaltskategorie.Haushaltskategorie = vlo_row.Item("Haushaltskategorie")
            vlo_haushaltskategorien.Add(vlo_haushaltskategorie)

        Next
        Conn.Close()
        Return vlo_haushaltskategorien
    End Function

    Public Function GetAuswertung() As IService1.Auswertung Implements IService1.GetAuswertung
        Dim Conn As MySql.Data.MySqlClient.MySqlConnection
        Dim vlo_auswertung As New IService1.Auswertung
        Dim myconnstring As String = ""
        Dim wertprojahr As Decimal = 0
        Dim vlo_rowcount As Integer = 0
        Dim vlo_alterwert As Integer = 0
        Dim vlo_neuerwert As Integer = 0
        Dim vlo_anzahl As Integer = 0
        Dim vlo_preis As Decimal = 0

        myconnstring = "Data Source=localhost;Database=db1145925-hausverwaltung;Password = kieran68;User ID = dbu1145925;pooling=false;Connection Timeout = 10;Default Command Timeout = 60"
        Conn = New MySql.Data.MySqlClient.MySqlConnection(myconnstring)
        Conn.Open()
        Dim adp_KVI_mysql As New MySql.Data.MySqlClient.MySqlDataAdapter
        Dim get_daten As New Data.DataSet
        Dim get_preis As New Data.DataSet
        Dim vlo_haushaltsunterkategorieid As Integer = 0

        For i As Integer = 1 To 5

            'Unterscheidung variabel oder fix -> variabel für 1 Jahr (letztes Jahr)

            Select Case i
                Case 1, 5 'variabel -> nur vergangenes Jahr 
                    adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT " _
                       & "ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie," _
                       & "Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor," _
                       & " ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit " _
                       & "FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, " _
                       & "tbl_einheit WHERE Einheit_ID = ID_Einheit And Zahlungsrythmus_ID = ID_Zahlungsrythmus " _
                       & "And Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie " _
                       & "And ID_Haushaltskategorie = Haushaltskategorie_ID And Haushaltskategorie_ID = " _
                       & i & " AND Datum BETWEEN '" & Year(Now) - 1 & "-01-01' AND '" & Year(Now) & "-01-01' ORDER BY Haushaltsunterkategorie_ID,Datum ASC;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
                Case Else
                    adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT " _
                        & "ID_Werte, Haushaltsunterkategorie_ID, Anzahl, Datum, Haushaltsunterkategorie," _
                        & "Haushaltsunterkategorie_ID, Haushaltskategorie, Haushaltskategorie_ID, Rythmusfaktor," _
                        & " ID_Zahlungsrythmus, Zahlungsrythmus, Einheit, ID_Einheit " _
                        & "FROM tbl_werte, tbl_haushaltskategorie, tbl_haushaltsunterkategorie, tbl_zahlungsrythmus, " _
                        & "tbl_einheit WHERE Einheit_ID = ID_Einheit And Zahlungsrythmus_ID = ID_Zahlungsrythmus " _
                        & "And Haushaltsunterkategorie_ID = ID_Haushaltsunterkategorie " _
                        & "And ID_Haushaltskategorie = Haushaltskategorie_ID And Haushaltskategorie_ID = " _
                        & i & " ORDER BY Haushaltsunterkategorie_ID;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
            End Select


            adp_KVI_mysql.Fill(get_daten)

            adp_KVI_mysql.Dispose()


            For Each vlo_row As DataRow In get_daten.Tables(0).Rows

                Select Case vlo_row.Item("ID_Zahlungsrythmus")
                    Case 1
                        wertprojahr = wertprojahr + vlo_row.Item("Anzahl")
                    Case 2
                        wertprojahr = wertprojahr + (vlo_row.Item("Anzahl") * 2)
                    Case 3
                        wertprojahr = wertprojahr + (vlo_row.Item("Anzahl") * 4)
                    Case 4
                        wertprojahr = wertprojahr + (vlo_row.Item("Anzahl") * 12)
                    Case Else
                        Select Case i
                            Case 1
                                'bei Wechsel Haushaltsunterkategorien (Verbrauchsarten) neu initialisieren
                                If vlo_haushaltsunterkategorieid <> vlo_row.Item("Haushaltsunterkategorie_ID") Then
                                    vlo_alterwert = 0
                                    vlo_neuerwert = 0
                                    vlo_anzahl = 0
                                    vlo_haushaltsunterkategorieid = vlo_row.Item("Haushaltsunterkategorie_ID")
                                End If

                                If vlo_rowcount = 0 Then
                                    vlo_rowcount = vlo_rowcount + 1
                                    vlo_alterwert = vlo_neuerwert
                                    vlo_neuerwert = vlo_row.Item("Anzahl")
                                Else
                                    vlo_rowcount = vlo_rowcount + 1
                                    vlo_alterwert = vlo_neuerwert
                                    vlo_neuerwert = vlo_row.Item("Anzahl")

                                    vlo_anzahl = vlo_neuerwert - vlo_alterwert

                                    If vlo_rowcount Mod 2 <> 0 Then
                                        adp_KVI_mysql.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand("SELECT Preis FROM tbl_verbrauchspreis WHERE Beginn <= '" & CDate(vlo_row.Item("Datum")).ToString("yyy-M-d HH:mm:ss") & "' AND Haushaltsunterkategorie_ID = " & vlo_row.Item("Haushaltsunterkategorie_ID") & " ORDER BY Beginn DESC LIMIT 1;", CType(Conn, MySql.Data.MySqlClient.MySqlConnection))
                                        adp_KVI_mysql.Fill(get_preis)

                                        For Each vlo_row_preis In get_preis.Tables(0).Rows
                                            vlo_preis = vlo_row_preis.Item("Preis")
                                        Next
                                    End If
                                    wertprojahr = wertprojahr + (vlo_anzahl * vlo_preis)
                                End If
                            Case 5
                                wertprojahr = wertprojahr + vlo_row.Item("Anzahl")
                        End Select

                End Select
            Next

            Select Case i
                Case 1 'Verbrauch variabel
                    vlo_auswertung.VerbrauchVarproJahr = wertprojahr.ToString & " €"
                    vlo_auswertung.VerbrauchVarproMonat = CommercialRound((wertprojahr / 12), 2).ToString & " €"
                Case 2 'Ausgabe fix
                    vlo_auswertung.AusgabenFixproJahr = wertprojahr.ToString & " €"
                    vlo_auswertung.AusgabenFixproMonat = CommercialRound((wertprojahr / 12), 2).ToString & " €"
                Case 3 'Einnahmen
                    vlo_auswertung.EinnahmenproJahr = wertprojahr.ToString & " €"
                    vlo_auswertung.EinnahmenproMonat = CommercialRound((wertprojahr / 12), 2).ToString & " €"
                Case 4 'Verbrauch fix
                    vlo_auswertung.VerbrauchFixproJahr = wertprojahr.ToString & " €"
                    vlo_auswertung.VerbrauchFixproMonat = CommercialRound((wertprojahr / 12), 2).ToString & " €"
                Case 5 'Ausgabe variabel
                    vlo_auswertung.AusgabenVarproJahr = wertprojahr.ToString & " €"
                    vlo_auswertung.AusgabenVarproMonat = CommercialRound((wertprojahr / 12), 2).ToString & " €"
            End Select

            get_daten.Tables(0).Clear()
            get_daten.Tables.Clear()
            wertprojahr = 0

        Next i

        Conn.Close()

        vlo_auswertung.AuswertungproJahr = "xxx"
        vlo_auswertung.AuswertungproMonat = "xxx"

        Return vlo_auswertung

    End Function

    Public Shared Function CommercialRound(value As Decimal, dec As Integer) As Decimal
        ' um die Anzahl der Dezimalstellen nach links verschieben
        Dim x As Decimal = value * Convert.ToDecimal(Math.Pow(10, dec))

        ' Dezimalstellen abtrennen
        Dim y As Decimal = Math.Floor(x)

        ' ist die Differenz größer oder gleich 0.5 soll aufgerundet werden
        If (x - y) >= 0.5D Then
            y += 1
        End If

        ' um die Anzahl der Dezimalstellen nach rechts verschieben 
        Return y / Convert.ToDecimal(Math.Pow(10, dec))
    End Function


End Class

