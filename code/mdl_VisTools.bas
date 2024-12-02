Attribute VB_Name = "mdl_VisTools"
Option Explicit
Option Private Module

' taken from https://johnvisiomvp.ca/2021/03/08/the-adventures-of-visio-section-shapes-continues/

Public MsgDict As Dictionary

Sub LoadDictionary()
    Set MsgDict = CreateObject("Scripting.Dictionary")
    MsgDict.Add Key:="G0", Item:="Default"
    MsgDict.Add Key:="G136", Item:="Tab0"
    MsgDict.Add Key:="G137", Item:="Geometry"
    MsgDict.Add Key:="G138", Item:="MoveTo"
    MsgDict.Add Key:="G139", Item:="LineTo"
    MsgDict.Add Key:="G140", Item:="ArcTo"
    MsgDict.Add Key:="G141", Item:="InfiniteLine"
    MsgDict.Add Key:="G143", Item:="Ellipse"
    MsgDict.Add Key:="G144", Item:="EllipticalArcTo"
    MsgDict.Add Key:="G150", Item:="Tab2"
    MsgDict.Add Key:="G151", Item:="Tab10"
    MsgDict.Add Key:="G153", Item:="CnnctPt"
    MsgDict.Add Key:="G162", Item:="CtlPt"
    MsgDict.Add Key:="G165", Item:="SplineBeg"
    MsgDict.Add Key:="G166", Item:="SplineSpan"
    MsgDict.Add Key:="G170", Item:="CtlPtTip"
    MsgDict.Add Key:="G181", Item:="Tab60"
    MsgDict.Add Key:="G185", Item:="CnnctNamed"
    MsgDict.Add Key:="G186", Item:="CnnctPtABCD"
    MsgDict.Add Key:="G187", Item:="CnnctNamedABCD"
    MsgDict.Add Key:="G193", Item:="PolylineTo"
    MsgDict.Add Key:="G195", Item:="NURBSTo"
    MsgDict.Add Key:="G236", Item:="RelCubBezTo"
    MsgDict.Add Key:="G237", Item:="RelQuadBezTo"
    MsgDict.Add Key:="G238", Item:="RelMoveTo"
    MsgDict.Add Key:="G239", Item:="RelLineTo"
End Sub

Function GetMsgName(msgCode As String, MsgNo As Integer) As String
    Dim ky As String
    ky = msgCode & Trim(Str(MsgNo))
    If MsgDict.Exists(ky) Then
        GetMsgName = ky & " " & MsgDict(ky)
    Else
      GetMsgName = "**<<<Unknown " & ky & ">>**"
      Debug.Print "**<<<Unknown " & ky & ">>**"
    End If
End Function

Function ResolveEnum(msgCode As String, MsgNo As Integer) As String
    Dim ky As String
    ky = msgCode & Trim(Str(MsgNo))
    If MsgDict.Exists(ky) Then
        ResolveEnum = MsgDict(ky)
    Else
      ResolveEnum = "**<<<Unknown " & ky & ">>**"
      Debug.Print "**<<<Unknown " & ky & ">>**"
    End If
End Function

Sub test()
    Set MsgDict = CreateObject("Scripting.Dictionary")
    Call LoadDictionary
    Debug.Print ResolveEnum("G", 138)
End Sub
