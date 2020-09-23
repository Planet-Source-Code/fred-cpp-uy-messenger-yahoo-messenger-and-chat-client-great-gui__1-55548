Attribute VB_Name = "ws_Release_Notes"
'Copyright 2000, TotalTakeOut.com
'Casey Goodhew - goodhewc@hotmail.com
'
'http://ttosock.2y.net
'http://www.TotalTakeOut.com
'
'You can see the latest updates to the documentation at:
'http://ttosock.2y.net
'
'If you would like to be notified of future updates to the control,
'please email me at goodhewc@hotmail.com
'
'Version Stamp - 1.2.1 - November 3, 2000
'
'If you find a bug in this code, or if you know of a better way
'to implement a function, PLEASE LET ME KNOW! When I receive bug
'reports, comments, or recommendations, I will post an updated
'version of the control as soon as possible.
'
'I think that we should all share this code freely, in order to
'develop the ultimate Winsock API user control. If anyone finds
'a better way of doing something, please let me know, so that we
'can all have the advantage, and no one is left behind.
'
'This souce code (both in complied and non-compiled verions)
'is free for you to use and distribute, under the condition
'that TotalTakeOut.com is acknoledged, and credit is
'given where credit is due.
'
'Speaking of which, credit is due to:
'         Zack Lantz - Lantz@1st.net - ICQ 17255556
'         Jay Freeman (saurik) - saurik@saurik.com - www.saurik.com
'         Daniel - sigsegv@mail.utexas.edu - Originator of Winsock_API.bas
'                                            (Originally named Wsksock.bas)
'Thanks for the help guys!
'
'YOU MAY NOT REDISTIBUTE THIS SOURCE CODE IN ANY WAY, SHAPE OR
'FORM, FOR ANY SUM OF MONEY.
'
'This module must be distributed with the uncompiled project.
'
'I have worked for several months to create this control. I came to
'the conclusion that there are not enough resouces on the internet
'in refernce to the Winsock API. I decided to freely distribute this
'control as both a learning tool and a pratical user control for the
'Winsock API.
'
'
'
'November 3, 2000 Release (Version 1.2.1)
'
'- A bug was discovered due to the way that messaging is delivered in
'    Windows NT Server 4. Details on this bug can be found @
'    ttosock.2y.net under Bug #0004.
'
'
'
'
'October 24, 2000 Release (Version 1.2.0)
'
'- Connect Method added
'    Now we call call out for connections instead of just listening.
'
'- Memory leak discovered
'    Various minor changes made to clean this up.
'    YOU WILL NEED TO INVOKE A NEW FUNCTION TO FIX THIS!!!
'    Check it out the Resolved Bugs (Bug Resolution #0003) section @
'    http://ttosock.2y.net
'
'- ListenNow Method now returns the key of the ListeningOnSocket
'    It the ListeningSocket property has been removed. This is
'    because if we had more than one listening socket, only the
'    newest socket was being returned.
'
'
'
'
'August 24, 2000 Release (Version 1.1.1)
'
'- Changes made to WindowProc Function Message Handle (uMsg):
'    The base address was increased from 1025 to 4025 in order to
'    avoid possible capture of an internal Windows message.
'    Thanks to Jeremy Stein for helping out on this. This will
'    be posted as Bug Resolution #0002.
'
'- Changes made to FD_CLOSE in WindowProc Function:
'    I've added "TempUC.Disconnect wParam" before raising the
'    PeerClosing Event, simply because if the remote end is closing
'    the connect, we might as well too. This just takes a little bit
'    control away from the user, in the respect that you will NOT NEED TO
'    specify disconnect in your applications now, it is taken care of
'    for you.
'
'    If anyone can see any problems with this, please let me know.
'
'    I should note that you will not need to change your existing code
'    because of this. Although the Disconnect method can still be used
'    in your PeerClosing Event, it is not needed.
'
'- Threading Changed:
'    The threading for the control was updated to an Apartment Threaded
'    model from a Single Threaded model, after an email from
'    Stephan Strittmatter. His email was actually a possible bug report,
'    regarding using the control in an ActiveX dll.
'    This possible bug is labelled #0003.
'
'
'
'
'July 27, 2000 Release (Version 1.1.0)
'
'- Internal Cryption added. Check out the control's home page for detail.
'
'
'
'
'
'
'
'
