Imports System.Runtime.InteropServices

<ComClass(ComSrvTest.ClassId, ComSrvTest.InterfaceId, ComSrvTest.EventsId)>
Public Class ComSrvTest

#Region "COM GUIDs"
    ' この GUID は、このクラス、およびその COM インターフェイスの
    ' COM の ID を提供します。これらを変更すると、
    ' 既存のクライアントがクラスにアクセスできなくなります。
    Public Const ClassId As String = "045C1BD4-569C-4d9e-A1B2-DDF07783364B"
    Public Const InterfaceId As String = "B73B6948-4CC6-4b91-BAC4-80D286C526D9"
    Public Const EventsId As String = "521AC157-DD94-4f90-B87C-1CBC13C17CD7"
#End Region

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure Book
        <MarshalAs(UnmanagedType.BStr)> _
        Public title As String
        Public cost As Integer
    End Structure

    Public Function getBook() As Book
        Dim book As New Book
        book.title = "The Ruby Book"
        book.cost = 20
        Return book
    End Function

    Public Function getBooks() As Book()
        Dim book() As Book = {New Book, New Book}
        book(0).title = "The CRuby Book"
        book(0).cost = 30
        book(1).title = "The JRuby Book"
        book(1).cost = 40
        Return book
    End Function

    Public Sub getBookByRefObject(ByRef obj As Object)
        Dim book As New Book
        book.title = "The Ruby Reference Book"
        book.cost = 50
        obj = book
    End Sub

    Public Function getVer2BookByValBook(<MarshalAs(UnmanagedType.Struct)> ByVal book As Book) As Book
        Dim ret As New Book
        ret.title = book.title + " ver2"
        ret.cost = book.cost * 1.1
        Return ret
    End Function

    Public Sub getBookByRefBook(<MarshalAs(UnmanagedType.LPStruct)> ByRef book As Book)
        book.title = "The Ruby Reference Book2"
        book.cost = 44
    End Sub

    Public Sub getVer3BookByRefBook(<MarshalAs(UnmanagedType.LPStruct)> ByRef book As Book)
        book.title += " ver3"
        book.cost *= 1.2
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub
End Class
