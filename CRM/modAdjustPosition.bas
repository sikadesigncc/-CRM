Attribute VB_Name = "modAjustPosition"
Option Explicit

Public Sub ap()
    frm.Top = (Screen.Height - frm * .Height) / 2 * 0.9
    frm*.Left = (Screen.Width - .Width) / 2

End Sub
