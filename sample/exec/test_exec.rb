require 'win32ole'
instance = WIN32OLE.new('ExecCommand.ExecCommand')

instance.ExeFile="ping.exe"
instance.arg = "localhost -n 1"

i = instance.Exec
print instance.stdout
