VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim GoXML As New CGoXML
    
    Dim strFileName As String
    strFileName = App.Path & "\" & "Input.xml"
    
' Use XPath to access XML elements...
'
' XPath QUICK START:
' "/alphabet/letter[0]" - the first (!) 'letter' subnode (absolute addressation)
' "letter[@position]" - 'letter' subnode(s) with a 'position' attribute (relative addressation)
' "letter[@position=5]" - 'letter' subnode(s) with a 'position' attribute of 5 (relative addressation)
'
' For more info on XPath, take a look at the MS XML
' Parser SDK (<= 3.0) or MS XML Core Services (4.0)
' available from http://www.microsoft.com
'
' --- Insert Shameless Self Promotion Here ---
'
' If you need some hands on experience in XPath,
' drop by at http://www.write4food.de
' I have prepared an online XPath Trainer.
' Work on three different XML documents (simple,
' medium, complex) or bring your own!
'
' ---------------------------------------------
'
' The following code includes *all* the public methods and properties
' currently supported by GoXML.

'-----------------------------------------------------------------------
    txtOutput = txtOutput & _
        "Initialize: " & _
        GoXML.Initialize(pavAuto)
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "OpenFromFile: " & _
        GoXML.OpenFromFile(strFileName, True)
        
'    txtOutput = txtOutput & vbCrLf & _
        "OpenFromString: " & _
        GoXML.OpenFromString("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><alphabet><letter position=""1"">A</letter><letter position=""2"">B</letter><letter position=""3"">C</letter><letter position=""4"">D</letter><letter position=""5"">E</letter></alphabet>", True)
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "WriteNode: " & _
        GoXML.WriteNode("/alphabet/letter[0]", "The letter A")
    
    txtOutput = txtOutput & vbCrLf & _
        "WriteAttribute: " & _
        GoXML.WriteAttribute("/alphabet/letter[0]", "comment", "My favorite letter")

'    txtOutput = txtOutput & vbCrLf & _
        "WriteCDataSection: " & _
        GoXML.WriteCDataSection("/alphabet/letter[4]", "<html><head></head><body>This is a CData Section</body></html>")
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "NodeCount: " & _
        GoXML.NodeCount("/alphabet/letter")
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "ReadNode: " & _
        GoXML.ReadNode("/alphabet/letter[0]")
    
    txtOutput = txtOutput & vbCrLf & _
        "ReadNodeXML: " & _
        GoXML.ReadNodeXML("/alphabet/letter[0]")

'    txtOutput = txtOutput & vbCrLf & _
        "ReadAttribute: " & _
        GoXML.ReadAttribute("/alphabet/letter[0]", "comment")
'-----------------------------------------------------------------------
'    txtOutput = txtOutput & vbCrLf & _
        "DeleteNode: " & _
        GoXML.DeleteNode("/alphabet/letter[0]")

'    txtOutput = txtOutput & vbCrLf & _
        "DeleteAttribute: " & _
        GoXML.DeleteAttribute("/alphabet/letter[0]", "comment")
'-----------------------------------------------------------------------
'    txtOutput = txtOutput & vbCrLf & _
        "InsertNode: " & _
        GoXML.InsertNode("/alphabet/letter[0]", "lower_case_version", "a", , , norCHILD)
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "XMLDocumentVersion: " & _
        GoXML.XMLDocumentInfo(xdiVERSION)
    
    txtOutput = txtOutput & vbCrLf & _
        "XMLDocumentEncoding: " & _
        GoXML.XMLDocumentInfo(xdiENCODING)
    
    txtOutput = txtOutput & vbCrLf & _
        "XMLDocumentStandalone: " & _
        GoXML.XMLDocumentInfo(xdiSTANDALONE)
    
    txtOutput = txtOutput & vbCrLf & _
        "XML: " & _
        GoXML.XML
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "Reparse: " & _
        GoXML.Reparse
'-----------------------------------------------------------------------
    txtOutput = txtOutput & vbCrLf & _
        "XMLParserVersion: " & _
        GoXML.XMLParserVersion
'-----------------------------------------------------------------------
'    txtOutput = txtOutput & vbCrLf & _
        "Save: " & _
        GoXML.Save(App.Path & "\" & "Output.xml")
'-----------------------------------------------------------------------
    
    Set GoXML = Nothing
End Sub

