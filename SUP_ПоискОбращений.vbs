'Fedotov_PV �������� 2012. ��� ������ ��������� DIRECTUM
'�������� ��������� ����� ��������� � ������� ���������
'VBScript.


  '��������� ��� �������� "�� ���������"
  SYSCODE = "XXXXXXX"									'��� ������� ������
  ScriptName = "�������� ����������� ��� � �����������" '��� ��������
  
  '������� �����������
  DIM Login, App, Scripts, Script
  SET Login = CreateObject("SBLogon.LoginPoint")
  SET App = Login.GetApplication("SystemCode=" & SYSCODE)
  '�������� ������� ���������
  SET Scripts = App.ScriptFactory
  SET Script = Scripts.GetObjectByName(ScriptName)
  '���������� ������� ����������, �.�. �������� "������" � ��������� IS-Builder ���������� ����������, ������� ������ ����� �������� �����������.
  On Error resume next
  Script.Execute
  
  ' �������������� �������, ����� ���������
  '�������� ������� ����� ����������
  'ApplicationFolder = App.Connection.SystemInfo.CorePath 
  '������ ������� �� ��������� ������
  '"C:\Program Files\NPO Computer\IS-Builder 7.10.1\SBLauncher.exe" -SYS=XXXXXX -CT=Script -F="�������� ����������� ��� � �����������"
  'SET Shell = CreateObject("WScript.Shell")
  ' ������������� ������� �������������� �������� �������. ������ & """" ��������, ��� � ������ ����������� ���� �������
  'CMD = """" & ApplicationFolder & "SBLauncher.exe" & """" & "-SYS=" & SYSCODE & " -CT=Script -F=""" & ScriptName & """"
  'Shell.Run CMD
  
  '����� ��������
  'MsgBox("End")
