Attribute VB_Name = "Ws_Public_Functions"
Option Explicit

'These two functions (with the accompanying constant and variable) are
'used to call the WindowProc function below. These declarations
'can be used to tell the OS to call the function, when certain
'Winsock Events occur.
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public OldWndProc As Long

'These two variables are used to raise events in the user control
Private UC As New Collection
Private MaxUCCount As Integer

'These four variables are used to allow us to use encryption on send and
'receive
Public CryptionObject As Object
Public ICanUseCryptionObject As Boolean
Public IShouldUseCryptionObject() As Boolean
Public CryptionKey() As String

'These two variables are used for tracking current states
Public WinsockStates(9) As String
Public CurrentState() As Integer

'These three variables are used to track sockets
Public m_lngSocks() As Long
Public m_intSocketAsync() As Integer
Public m_intMaxSockCount As Integer
Public m_intConnectionsAlert As Integer
'

'Returns the socket stack index for the specified socketID.
'If the socket ID does not exist in the stack, -1 is returned.
Public Function GetIndexFromsID(SocketID As Long) As Integer

  Dim x As Integer
  
  For x = 1 To m_intMaxSockCount
    If m_lngSocks(x) = SocketID Then
      GetIndexFromsID = x
      Exit Function
    End If
  Next x
  
  GetIndexFromsID = -1

End Function

'An internal function for delays. The WaitTime should be specified in
'seconds. If waittime is not passed, a value of 1 second is used.
Public Function WaitJustOneSecond(Optional WaitTime As Single = 1) As Boolean

  Dim sTimer As Variant
  
  sTimer = Timer
  
  Do Until Timer > sTimer + WaitTime
    DoEvents
  Loop
  
  WaitJustOneSecond = True

End Function

'This sub is used at start up, to reference the user control
'from within this function.
Public Function SetControlHost(ByVal ControlInstance As TTOSock) As String
  
  Dim objTTOSock As TTOSock
  Dim NewKey As String
  
  'This will ensure a unique key
  NewKey = "a" & UC.Count + 1
  
  Set objTTOSock = ControlInstance
  UC.Add objTTOSock, NewKey
  
  'If the count is larger than the Maximum Count, we need to
  'increase the maximum count so that we are sure that we will
  'be able to raise events to each instance.
  If UC.Count > MaxUCCount Then MaxUCCount = UC.Count
  
  Set objTTOSock = Nothing
  Set ControlInstance = Nothing
  
  SetControlHost = NewKey
      
End Function

'This function is called by the OS, when Winsock events are
'raised. lParam contains the event or error code, wParam contains
'the Socket ID.
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
  'Under Windows NT Server 4.0 with SP6 I found that when we were trying
  'to close the control, a value continued to loop through WindowProc until
  'a stack over flow occured. After testing, I found that in all instances,
  'no matter the the value of uMsg, lParam was 0. I am assuming that
  'this is a garbage message that is floating around looking to be destroyed
  'but I can't prove it. Everything still appears to be working fine, but
  'you should watch out for error caused because of this.
  If lParam = 0 Then Exit Function
  
  'We only want to look at messages that are addressed to us,
  'so we check to see if the message number matches our designated
  'number. We designated the number 4025 when we set up the
  'messaging system.
    
  If uMsg > 4025 And uMsg < 4026 + MaxUCCount Then
      
    Dim WSAEvent As Long
    Dim WSAError As Long
    Dim TempUC As TTOSock
        
    'We need to create a useable instance of TTOSock
    Set TempUC = UC.Item("a" & uMsg - 4025)
    
    'Checks for errors and events
    WSAEvent = WSAGetSelectEvent(lParam)
    WSAError = WSAGetAsyncError(lParam)
      
    'Deals with each event
    Select Case WSAEvent
      Case FD_ACCEPT
                
        TempUC.RaiseConnectionRequest wParam
              
      Case FD_READ
          
        ReceiveDataNew wParam, "a" & uMsg - 4025
        
      Case FD_CONNECT
        
         TempUC.RaiseConnected wParam
        
      Case FD_CLOSE
         
        TempUC.RaisePeerClosing wParam
           
      Case FD_WRITE
        
  
        
      Case FD_OOB
     
    End Select
    
  Else
    
    'Passes on the event.
    WindowProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, ByVal lParam)

  End If
  
  Set TempUC = Nothing

End Function

'This function checks for new data on the specified socket
Private Function ReceiveDataNew(SocketID As Long, UCKey As String)

  Dim RecvBuffer As String
  Dim fixstr As String * 1024
  Dim RetByteErr As Integer
  fixstr = ""
  RecvBuffer = ""
   
  'Attempts to receive data from the socket
  RetByteErr = recv(SocketID, fixstr, 1024, 0)
   
  'Pick the info out of the junk
  If RetByteErr < 0 Then
    'HandleError
    Exit Function
  ElseIf RetByteErr = 0 Then
    'Connection was closed
    Exit Function
  Else
    RecvBuffer = Left$(fixstr, RetByteErr)
  End If
  
  If RecvBuffer <> "" Then
    'Raises the new data arrival event
    Dim TempUC As TTOSock
    Set TempUC = UC.Item(UCKey)
    If ICanUseCryptionObject And IShouldUseCryptionObject(GetIndexFromsID(SocketID)) Then RecvBuffer = CryptionObject.Decrypt(RecvBuffer, CryptionKey(GetIndexFromsID(SocketID)))
    TempUC.RaiseDataArrival SocketID, RecvBuffer
    Set TempUC = Nothing
  End If
    
End Function

'This function destroys the UC object that was created in order to
'access controls in the usercontrol. This method must be called
'whenever we are attempting to destroy the user control. This can be
'a deadly circular reference.
Public Sub CleanUp(UCKey As String)
  
  On Error Resume Next
  UC.Remove UCKey
      
End Sub

  
Public Sub CleanUpAll()

  Dim x As Integer
  
  On Error Resume Next
  
  For x = UC.Count To 0 Step -1
    UC.Remove x
  Next x

End Sub

Public Function ResolveIPtoNBO(IP As String) As Long

  Dim NBO As Long
  
  NBO = inet_addr(IP)
  
  If NBO = -1 Then NBO = GetHostByNameAlias(IP)
   
  ResolveIPtoNBO = NBO
  
End Function
  
  
