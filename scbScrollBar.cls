VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "scbScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'* Name                 :   scbScrollBar Object Class Module                            *
'* Author               :   Joseph Solomon, CPA                                         *
'* Purpose              :   Class object referencing a table that will allow the scroll *
'*                              bar to control a dynamic number of options listed in    *
'*                              tblControls                                             *
'* Main Call(s)         :   BindScrollBar, WriteFrameOptionControls                     *
'* Rules                :   1.  Must have a reference to the Microsoft Forms 2.0 library*
'*                          2.  Must a table named tblControls with fields: BoundForm   *
'*                                  BoundControl, ControlType, ParentControl,           *
'*                                  ControlValue, and Caption.                          *
'*                                                                                      *
'* Usage Example:                                                                       *
'* ~~~~~~~~~~~~~~~~~~~                                                                  *
'* Dim mclsScrollBar As scbScrollBar                                                    *
'* Set mclsScrollBar = New scbScrollBar                                                 *
'* mclsScrollBar.BindScrollBar _                                                        *
'*     probjParentControl:=Me.fmeExceptionType, _                                       *
'*     probjBoundForm:=Me, _                                                            *
'*     probjScrollBar:=vsbExceptionType.Object                                          *
'* mclsScrollBar.WriteFrameOptionControls                                               *
'*                                                                                      *
'* Revision History:                                                                    *
'* Rev#     Revision Date   Description                                                 *
'* ~~~~     ~~~~~~~~~~~~~   ~~~~~~~~~~~                                                 *
'*  1       2012/12/19      Initial Release                                             *
'*  2       2012/12/20      Debugged and tested.  Wrote documentation.                  *
'****************************************************************************************

Option Compare Database
Option Explicit
Private WithEvents mobjScrollBar As MSForms.ScrollBar
Attribute mobjScrollBar.VB_VarHelpID = -1
Private WithEvents mrstScrollBar As ADODB.Recordset
Attribute mrstScrollBar.VB_VarHelpID = -1
Private mobjBoundForm As Form
Private mobjParentControl As Control
Private mblnIsSelected As Boolean

'****************************************************************************************
'*  <Methods>                                                                           *
'*                                                                                      *
'*  Name                        Scope       Description                                 *
'*  ~~~~~~~~~~~~~               ~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~       *
'*  BindScrollBar               Public      Sets parameters probjParentControl,         *
'*                                              probjBoundForm, probjScrollBar as Module*
'*                                              level variables, opens Control          *
'*                                              Properties recordset, formats the       *
'*                                              scrollbar, and writes the first set of  *
'*                                              options.                                *
'*  WriteFrameOptionControls    Public      Changes the properties of the Option Boxes  *
'*                                              to correspond with a value from the     *
'*                                              scrollbar.                              *
'*  OpenRecordSet               Private     Opens the recordset of holding the control's*
'*                                              properties.                             *
'*  Find; Move; MoveFirst;      Public      ADODB recordset commands that are used to   *
'*  MoveLast; MoveNext;                         control the recordset outside of the    *
'*  Requery; AddNew; Update;                    class module                            *
'*  Delete                                                                              *
'*  glngCountChildObjects       Private     Count child objects in the parent object.   *
'*  mlngMaxOptionValue          Public      Finds the largest ControlValue Field in the *
'*                                              recordset. Assumes OptionValues are     *
'*                                              stored in ControlValue field.           *
'****************************************************************************************

Public Sub BindScrollBar( _
    ByRef probjParentControl As Object, _
    ByRef probjBoundForm As Form, _
    ByRef probjScrollBar As MSForms.ScrollBar)
    '                                Leaving this blank will show hidden controls at the
    '                                end of the list being scrolled.
    'probjParentControl         'Parent control holding the child objects that the
    '                               Scrollbar will be controlling
    'probjBoundForm             'Form the Scrollbar will be attached to
    'probjScrollBar             'Object name of the Scrollbar being manipulated
    
    Set mobjScrollBar = probjScrollBar
    Set mobjBoundForm = probjBoundForm
    Set mobjParentControl = probjParentControl
    OpenRecordset
    WriteFrameOptionControls
End Sub

Public Sub WriteFrameOptionControls()
    Dim objControl As Object
    Dim lngNumControls As Long  'Number of static controls the scrollbar will control
    On Error GoTo WriteFrameOptionControls_Err
   
    'Verify scrollbar has been bound to this class.
    If mobjScrollBar Is Nothing Then
        Err.Raise 23101, "Active X scrollbar object not " & _
            "yet bound to a form or control.  Run BindScrollBar method first."
    End If
    
    'Skip to the record corresponding to the ScrollBar's value
    Move Me.gobjScrollBar.Value - 1, adBookmarkFirst
    
    'Format Scroll Bar
    With mobjScrollBar
        .Enabled = True
        lngNumControls = glngCountChildObjects(acOptionButton)
        If mrstScrollBar.RecordCount <= lngNumControls Then
            .Enabled = False
        End If
        .Min = 1
        .Max = mrstScrollBar.RecordCount - lngNumControls + 1
        If .Max < 1 Then .Max = 1
    End With 'mobjScrollBar
    
    'Walk each Frame OptionControl
    For Each objControl In mobjParentControl.Controls
        With objControl
            Debug.Print .Name

            If .ControlType = acOptionButton Then
                If Not mrstScrollBar.EOF Then
                    .Visible = True
                    .OptionValue = CLng(gstrControlValue)
                    .Controls.Item(0).Caption = gstrCaption
                    MoveNext
                Else
                    .Visible = False
                End If
            End If
        End With
    Next objControl

WriteFrameOptionControls_Exit:
    'Move to begining of recordset for subsequent scrollbar changes
    MoveFirst
    Exit Sub
    
WriteFrameOptionControls_Err:
        Select Case MsgBox( _
            Prompt:=Err.Description, _
            Buttons:=vbAbortRetryIgnore Or vbExclamation, _
            Title:="Error")
            Case Is = vbIgnore: Resume Next
            Case Is = vbRetry: Resume
            Case Is = vbAbort: Resume WriteFrameOptionControls_Exit
        End Select
        Resume WriteFrameOptionControls_Exit
End Sub

Public Sub AddNewOptionButton(ByVal pvstrCaption As String)
    With mrstScrollBar
        .AddNew
        !ControlValue = mlngMaxOptionValue
        !Caption = pvstrCaption
        !BoundForm = mobjBoundForm.Name
        !BoundControl = mobjParentControl
        !ControlType = "acObjectButton"
        .Update
        .Requery
    End With
    WriteFrameOptionControls
    Set mobjParentControl = Null
End Sub

Private Function mlngMaxOptionValue()
    With mrstScrollBar
        .MoveFirst
        Do Until .EOF
            If mlngMaxOptionValue < !ControlValue Then
                mlngMaxOptionValue = !ControlValue
            End If
        Loop
    End With
End Function

Private Sub OpenRecordset()
    With mrstScrollBar
        .ActiveConnection = CurrentProject.Connection
        .Source = _
            "SELECT * " & _
            "FROM tblControls " & _
            "WHERE BoundForm = '" & mobjBoundForm.Name & "' AND " & _
            "BoundControl = '" & mobjParentControl.Name & "' " & _
            "ORDER BY ControlID ASC"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function glngCountChildObjects( _
    Optional ByVal pvenmControlType As AcControlType) As Long
    'pvenmControlType               'Optional control type to count.
    Dim objControl As Control       'Child object being walked.
    
    glngCountChildObjects = 0
    For Each objControl In mobjParentControl.Controls
        If (pvenmControlType = 0) Or (objControl.ControlType = pvenmControlType) Then
            glngCountChildObjects = glngCountChildObjects + 1
        End If
    Next objControl
End Function

'ADODB.Recordset Methods
Public Sub Move(ByVal pvNumRecords As Long, Optional pvStart As BookmarkEnum)
    mrstScrollBar.Move pvNumRecords, pvStart
End Sub
Public Sub Find( _
pvstrCriteria As String, _
Optional pvSkipRecords As ADO_LONGPTR, _
Optional pvSearchdirection As SearchDirectionEnum = adSearchForward)
    mrstScrollBar.Find pvstrCriteria, pvSkipRecords, pvSearchdirection
End Sub
Public Sub MoveFirst()
    mrstScrollBar.MoveFirst
End Sub
Public Sub MoveNext()
    mrstScrollBar.MoveNext
End Sub
Public Sub MoveLast()
    mrstScrollBar.MoveLast
End Sub
Public Sub Requery()
    mrstScrollBar.Requery
End Sub
Public Sub Delete()
    mrstScrollBar.Delete
End Sub
Public Sub Update()
    mrstScrollBar.Update
End Sub
Public Sub AddNew()
    mrstScrollBar.AddNew
End Sub

'****************************************************************************************
'*  <Properties>                                                                        *
'*                                                                                      *
'*  Name                        Scope       Description                                 *
'*  ~~~~~~~~~~~~~               ~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~       *
'*  gblnEOF                     Public      ADODB Property to check if recordset cursor *
'*                                              is at the end of the file.              *
'*  gobjScrollBar               Public      Scrollbar object this class is wrapped to   *
'*  gstrCaption                 Private     Value to be written to the associated       *
'*                                              child's label caption                   *
'*  gstrControlValue            Public      The Value property assigned to the child    *
'*                                              object.                                 *
'*  gblnIsSelected              Private     Boolean to determine whether a selection was*
'*                                              made on a child object.                 *
'****************************************************************************************
Public Property Get gblnEOF() As Boolean
    gblnEOF = mrstScrollBar.EOF
End Property
Public Property Get gobjScrollBar() As Object
    Set gobjScrollBar = mobjScrollBar
End Property
Public Property Get gstrCaption() As String
    gstrCaption = mrstScrollBar!Caption
End Property
Public Property Let gstrCaption(pvstrCaption As String)
    With mrstScrollBar
        !ControlValue = pvstrCaption
        .Update
    End With
    gstrCaption = pvstrCaption
End Property
Public Property Get gstrControlValue() As String
    gstrControlValue = CStr(mrstScrollBar!ControlValue)
End Property
Public Property Let gstrControlValue(pvControlValue As String)
    With mrstScrollBar
        !ControlValue = pvControlValue
        .Update
    End With
    gstrControlValue = pvControlValue
End Property
Public Property Get gblnIsSelected() As Boolean
    gblnIsSelected = mblnIsSelected
End Property
Public Property Let gblnIsSelected(ByVal mblnIsSelected As Boolean)
    If Not IsNull(mobjParentControl) Then mblnIsSelected = mblnIsSelected
End Property

Private Sub Class_Initialize()
    Set mrstScrollBar = New ADODB.Recordset
End Sub
Private Sub Class_Terminate()
    mrstScrollBar.Close
    Set mrstScrollBar = Nothing
    Set mobjScrollBar = Nothing
    Set mobjBoundForm = Nothing
    Set mobjParentControl = Nothing
End Sub
Private Sub mobjScrollBar_Change()
    WriteFrameOptionControls
    mobjParentControl.SetFocus
End Sub

