Dim SavePsw 
On Error resume next
Set SavePsw = CreateObject("WScript.Shell")
SavePsw.Run "cmdkey.exe chcp 1251 /add:192.168.0.100 /user:moscow\�������������_ /pass:begin"
SavePsw.Run "cmdkey.exe chcp 1251 /add:192.168.0.204 /user:moscow\�������������_ /pass:begin"
SavePsw.Run "cmdkey.exe chcp 1251 /add:192.168.0.205 /user:moscow\�������������_ /pass:begin"
SavePsw.Run "cmdkey.exe chcp 1251 /add:192.168.0.210 /user:moscow\�������������_ /pass:begin"
SavePsw.Run "cmdkey.exe chcp 1251 /add:192.168.1.100 /user:moscow\�������������_ /pass:begin"
SavePsw.Run "cmdkey.exe chcp 1251 /add:192.168.3.100 /user:moscow\�������������_ /pass:begin"
SavePsw.Run "cmdkey.exe chcp 1251 /add:10.1.2.254 /user:moscow\�������������_ /pass:begin"
if Err.Number>0 then
SavePsw.Run "net use \\bak-server /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\192.168.0.204 /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\lab-server /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\192.168.0.205 /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\bicepz /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\192.168.0.210 /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\win2kserv1 /user:moscow\�������������_ begin /persistent:yes"
SavePsw.Run "net use \\192.168.0.100 /user:moscow\�������������_ begin /persistent:yes"
end if
On Error goto 0