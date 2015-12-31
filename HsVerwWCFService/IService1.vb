' HINWEIS: Mit dem Befehl "Umbenennen" im Kontextmenü können Sie den Schnittstellennamen "IService1" sowohl im Code als auch in der Konfigurationsdatei ändern.
Imports System.Runtime.Serialization

<ServiceContract()>
Public Interface IService1

    <OperationContract()>
    Function GetVerbrauch() As IEnumerable(Of Verbrauch)
    <OperationContract()>
    Function GetAusgabe() As IEnumerable(Of Ausgabe)
    <OperationContract()>
    Function GetVerbrauchbyTyp(ByVal verbrauchstyp As Long) As IEnumerable(Of Verbrauch)
    <OperationContract()>
    Function GetVerbrauchsTyp(ByVal Haushatskategorie_ID As Long) As IEnumerable(Of Verbrauchstyp)
    <OperationContract()>
    Function GetVerbrauchbyUnterTyp(ByVal verbrauchsuntertyp As Long) As IEnumerable(Of Verbrauch)
    <OperationContract()>
    Function GetAusgabebyTyp(ByVal ausgabentyp As Long) As IEnumerable(Of Ausgabe)
    <OperationContract()>
    Function GetVerbrauchbyID(ByVal ID As Long) As Verbrauch
    <OperationContract()>
    Function GetAusgabebyID(ByVal ID As Long) As Ausgabe
    <OperationContract()>
    Function GetVerbrauchbyDate(ByVal datum As Date) As Verbrauch
    <OperationContract()>
    Function GetAusgabebyDate(ByVal datum As Date) As Ausgabe
    <OperationContract()>
    Function GetZahlungsrythmen() As IEnumerable(Of Zahlungsrythmus)
    <OperationContract()>
    Function GetVarVerbrauchKat() As IEnumerable(Of VarVerbrauchKat)
    <OperationContract()>
    Function SetVerbrauch(ByVal vlo_verbrauch As Verbrauch) As Boolean
    <OperationContract()>
    Function SetAusgabe(ByVal vlo_ausgabe As Ausgabe) As String
    <OperationContract()>
    Function SetVerbrauchNew(ByVal vlo_verbrauch As Verbrauch) As Boolean
    <OperationContract()>
    Function SetAusgabeNew(ByVal vlo_ausgabe As Ausgabe) As String
    <OperationContract()>
    Function DeleteVerbrauch(ByVal vlo_verbrauch As Verbrauch) As Boolean
    <OperationContract()>
    Function DeleteAusgabe(ByVal vlo_ausgabe As Ausgabe) As Boolean
    <OperationContract()>
    Function GetHaushaltskategorien() As IEnumerable(Of Haushaltskategorie)
    <OperationContract()>
    Function GetAuswertung() As Auswertung

    <DataContract()>
    Class Basis
        Private _id As Long
        <DataMember> Public Property ID As Long
            Get
                Return _id
            End Get
            Set(value As Long)
                _id = value
            End Set
        End Property

        Private _haushaltsunterkategorie As String
        <DataMember> Public Property Haushaltsunterkategorie As String
            Get
                Return _haushaltsunterkategorie
            End Get
            Set(value As String)
                _haushaltsunterkategorie = value
            End Set
        End Property

        Private _haushaltsunterkategorieid As Long
        <DataMember> Public Property HaushaltsunterkategorieID As Long
            Get
                Return _haushaltsunterkategorieid
            End Get
            Set(value As Long)
                _haushaltsunterkategorieid = value
            End Set
        End Property

        Private _haushaltskategorie As String
        <DataMember> Public Property Haushaltskategorie As String
            Get
                Return _haushaltskategorie
            End Get
            Set(value As String)
                _haushaltskategorie = value
            End Set
        End Property

        Private _haushaltskategorieid As Long
        <DataMember> Public Property HaushaltskategorieID As Long
            Get
                Return _haushaltskategorieid
            End Get
            Set(value As Long)
                _haushaltskategorieid = value
            End Set
        End Property

        Private _zahlungsrythmus As String
        <DataMember> Public Property Zahlungsrythmus As String
            Get
                Return _zahlungsrythmus
            End Get
            Set(value As String)
                _zahlungsrythmus = value
            End Set
        End Property

        Private _zahlungsrythmusid As Long
        <DataMember> Public Property ZahlungsrythmusID As Long
            Get
                Return _zahlungsrythmusid
            End Get
            Set(value As Long)
                _zahlungsrythmusid = value
            End Set
        End Property


        Private _zahlungsrythmusfaktor As Long
        <DataMember> Public Property Zahlungsrythmusfaktor As Long
            Get
                Return _zahlungsrythmusfaktor
            End Get
            Set(value As Long)
                _zahlungsrythmusfaktor = value
            End Set
        End Property

        Private _einheit As String
        <DataMember> Public Property Einheit As String
            Get
                Return _einheit
            End Get
            Set(value As String)
                _einheit = value
            End Set
        End Property

        Private _einheitid As Long
        <DataMember> Public Property EinheitID As Long
            Get
                Return _einheitid
            End Get
            Set(value As Long)
                _einheitid = value
            End Set
        End Property

        Private _wert As Long
        <DataMember> Public Property Wert As Long
            Get
                Return _wert
            End Get
            Set(value As Long)
                _wert = value
            End Set
        End Property

        Private _datum As Date
        <DataMember> Public Property Datum As Date
            Get
                Return _datum.ToShortDateString
            End Get
            Set(value As Date)
                _datum = value
            End Set
        End Property

    End Class

    <DataContract(Name:="Verbrauch")> Class Verbrauch
        Inherits Basis

        Private _verbrauchstyp As String
        <DataMember> Public Property Verbrauchstyp As String
            Get
                Return _verbrauchstyp
            End Get
            Set(value As String)
                _verbrauchstyp = value
            End Set
        End Property

    End Class

    <DataContract(Name:="Ausgabe")> Class Ausgabe
        Inherits Basis

        Private _ausgabentyp As String
        <DataMember> Public Property Ausgabentyp As String
            Get
                Return _ausgabentyp
            End Get
            Set(value As String)
                _ausgabentyp = value
            End Set
        End Property

    End Class

    <DataContract()>
    Class Verbrauchstyp
        Private _id As Long
        <DataMember> Public Property ID As Long
            Get
                Return _id
            End Get
            Set(value As Long)
                _id = value
            End Set
        End Property

        Private _haushaltsunterkategorie As String
        <DataMember> Public Property Haushaltsunterkategorie As String
            Get
                Return _haushaltsunterkategorie
            End Get
            Set(value As String)
                _haushaltsunterkategorie = value
            End Set
        End Property

        Private _haushaltskategorie As String
        <DataMember> Public Property Haushaltskategorie As String
            Get
                Return _haushaltskategorie
            End Get
            Set(value As String)
                _haushaltskategorie = value
            End Set
        End Property

        Private _haushaltskategorieid As Long
        <DataMember> Public Property HaushaltskategorieID As Long
            Get
                Return _haushaltskategorieid
            End Get
            Set(value As Long)
                _haushaltskategorieid = value
            End Set
        End Property

    End Class

    <DataContract()>
    Class Verbrauchspreis
        Private _id As Long
        <DataMember> Public Property ID As Long
            Get
                Return _id
            End Get
            Set(value As Long)
                _id = value
            End Set
        End Property

        Private _preis As Long
        <DataMember> Public Property Preis As Long
            Get
                Return _preis
            End Get
            Set(value As Long)
                _preis = value
            End Set
        End Property

        Private _haushaltsunterkategorieid As Long
        <DataMember> Public Property HaushaltsunterkategorieID As Long
            Get
                Return _haushaltsunterkategorieid
            End Get
            Set(value As Long)
                _haushaltsunterkategorieid = value
            End Set
        End Property

        Private _beginn As Date
        <DataMember> Public Property Beginn As Date
            Get
                Return _beginn
            End Get
            Set(value As Date)
                _beginn = value
            End Set
        End Property


    End Class

    <DataContract(Name:="Zahlungsrythmus")>
    Class Zahlungsrythmus
        Private _id As Long
        <DataMember> Public Property ID As Long
            Get
                Return _id
            End Get
            Set(value As Long)
                _id = value
            End Set
        End Property

        Private _rythmus As String
        <DataMember> Public Property Rythmus As String
            Get
                Return _rythmus
            End Get
            Set(value As String)
                _rythmus = value
            End Set
        End Property

        Private _rythmusfaktor As Integer
        <DataMember> Public Property Rythmusfaktor As Integer
            Get
                Return _rythmusfaktor
            End Get
            Set(value As Integer)
                _rythmusfaktor = value
            End Set
        End Property

        Private _beginn As Date
        <DataMember> Public Property Beginn As Date
            Get
                Return _beginn
            End Get
            Set(value As Date)
                _beginn = value
            End Set
        End Property


    End Class

    <DataContract(Name:="VarVerbrauchKat")>
    Class VarVerbrauchKat
        Private _id As Long
        <DataMember> Public Property ID As Long
            Get
                Return _id
            End Get
            Set(value As Long)
                _id = value
            End Set
        End Property

        Private _varverbrauchkat As String
        <DataMember> Public Property VarVerbrauchKat As String
            Get
                Return _varverbrauchkat
            End Get
            Set(value As String)
                _varverbrauchkat = value
            End Set
        End Property

    End Class

    <DataContract(Name:="Haushaltskategorie")>
    Class Haushaltskategorie
        Private _id As Long
        <DataMember> Public Property ID As Long
            Get
                Return _id
            End Get
            Set(value As Long)
                _id = value
            End Set
        End Property

        Private _haushaltskategorie As String
        <DataMember> Public Property Haushaltskategorie As String
            Get
                Return _haushaltskategorie
            End Get
            Set(value As String)
                _haushaltskategorie = value
            End Set
        End Property

    End Class

    <DataContract(Name:="Auswertung")>
    Class Auswertung

        Private _einnahmenpromonat As String
        <DataMember> Public Property EinnahmenproMonat As String
            Get
                Return _einnahmenpromonat
            End Get
            Set(value As String)
                _einnahmenpromonat = value
            End Set
        End Property

        Private _einnahmenprojahr As String
        <DataMember> Public Property EinnahmenproJahr As String
            Get
                Return _einnahmenprojahr
            End Get
            Set(value As String)
                _einnahmenprojahr = value
            End Set
        End Property

        Private _verbrauchpromonat As String
        <DataMember> Public Property VerbrauchproMonat As String
            Get
                Return _verbrauchpromonat
            End Get
            Set(value As String)
                _verbrauchpromonat = value
            End Set
        End Property

        Private _verbrauchprojahr As String
        <DataMember> Public Property VerbrauchproJahr As String
            Get
                Return _verbrauchprojahr
            End Get
            Set(value As String)
                _verbrauchprojahr = value
            End Set
        End Property

        Private _ausgabenpromonat As String
        <DataMember> Public Property AusgabenproMonat As String
            Get
                Return _ausgabenpromonat
            End Get
            Set(value As String)
                _ausgabenpromonat = value
            End Set
        End Property

        Private _ausgabenprojahr As String
        <DataMember> Public Property AusgabenproJahr As String
            Get
                Return _ausgabenprojahr
            End Get
            Set(value As String)
                _ausgabenprojahr = value
            End Set
        End Property
        Private _auswertungpromonat As String
        <DataMember> Public Property AuswertungproMonat As String
            Get
                Return _auswertungpromonat
            End Get
            Set(value As String)
                _auswertungpromonat = value
            End Set
        End Property

        Private _auswertungprojahr As String
        <DataMember> Public Property AuswertungproJahr As String
            Get
                Return _auswertungprojahr
            End Get
            Set(value As String)
                _auswertungprojahr = value
            End Set
        End Property
    End Class

End Interface

