Attribute VB_Name = "VariableType_CompSetup"
Option Explicit



'For Country and Currency
Public Type Country_and_Currency
    CountryName     As String
    CurrencySymbol  As String
End Type
'For User Information
Public Type User_Info
    FullName As String
    UserName As String
    Password As String
End Type
'For Company Inforamation
Public Type Company_Info
    CompanyName  As String
    ContactName  As String
    StreetAdd    As String
    City         As String
    ZipCode      As String
    Phone        As String
    Fax          As String
    EAdd         As String
    WebSite      As String
    BusinessType As String
End Type
