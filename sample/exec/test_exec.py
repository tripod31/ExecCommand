import win32com.client

instance = win32com.client.Dispatch("ExecCommand.ExecCommand")

instance.ExeFile="ping.exe"
instance.arg = "localhost -n 1"

ret = instance.Exec
print (instance.stdout+instance.stderr)
