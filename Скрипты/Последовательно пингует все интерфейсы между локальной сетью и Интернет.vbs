::Successively pings each interface between LAN and internet, then gateway, 
::both dns servers, and a known good internet address. 
:: 
::Outputs to text file %OutPutFile%.   
::Repeats indefinitely until stopped with Ctrl-C. 
::The output file grows by ~14MB/day. 
:: 
::You must first edit the set variables commands immediately below these remarks. 
::If a variable is left blank, related commands will be skipped,  
::e.g., if you have no firewall, leave 
::"SET LanSideFirewall=" blank  
::and also leave 
::"SET WanSideFirewall=" blank. 
::**************************************************************************** 
SET OutPutFile=c:\test\Connectivity_Test_Log.txt 
SET LanSideSBS= 
SET WanSideSBS= 
SET LanSideFirewall= 
SET WanSideFirewall= 
SET LanSideModem= 
SET WanSideModem= 
SET Gateway= 
SET DNS1= 
SET DNS2= 
SET KnownGoodInternet=winfiles.com 
::**************************************************************************** 
:: 
:runagain 
echo. >> %OutPutFile% 
echo ********************************************************** >> %OutPutFile% 
date /t >> %OutPutFile% 
echo. | time >> %OutPutFile% 
if defined LanSideSBS echo. >> %OutPutFile% 
if defined LanSideSBS echo ........................................ >> %OutPutFile% 
if defined LanSideSBS echo LAN side of SBS >> %OutPutFile% 
if defined LanSideSBS ping %LanSideSBS% >> %OutPutFile% 
if defined WanSideSBS echo. >> %OutPutFile% 
if defined WanSideSBS echo ........................................ >> %OutPutFile% 
if defined WanSideSBS echo WAN side of SBS >> %OutPutFile% 
if defined WanSideSBS ping %WanSideSBS% >> %OutPutFile% 
if defined LanSideFirewall echo. >> %OutPutFile% 
if defined LanSideFirewall echo ................................... >> %OutPutFile% 
if defined LanSideFirewall echo LAN side of firewall >> %OutPutFile% 
if defined LanSideFirewall ping %LanSideFirewall% >> %OutPutFile% 
if defined WanSideFirewall echo. >> %OutPutFile% 
if defined WanSideFirewall echo ................................... >> %OutPutFile% 
if defined WanSideFirewall echo WAN side of firewall >> %OutPutFile% 
if defined WanSideFirewall ping %WanSideFirewall% >> %OutPutFile% 
if defined LanSideModem echo. >> %OutPutFile% 
if defined LanSideModem echo ...................................... >> %OutPutFile% 
if defined LanSideModem echo LAN side of cable or dsl modem >> %OutPutFile% 
if defined LanSideModem ping %LanSideModem% >> %OutPutFile% 
if defined WanSideModem echo. >> %OutPutFile% 
if defined WanSideModem echo ...................................... >> %OutPutFile% 
if defined WanSideModem echo WAN side of cable or dsl modem >> %OutPutFile% 
if defined WanSideModem ping %WanSideModem% >> %OutPutFile% 
if defined Gateway echo. >> %OutPutFile% 
if defined Gateway echo ............................................ >> %OutPutFile% 
if defined Gateway echo Gateway address >> %OutPutFile% 
if defined Gateway ping %Gateway% >> %OutPutFile% 
if defined DNS1 echo. >> %OutPutFile% 
if defined DNS1 echo ............................................... >> %OutPutFile% 
if defined DNS1 echo DNS1 >> %OutPutFile% 
if defined DNS1 ping %DNS1% >> %OutPutFile% 
if defined DNS2 echo. >> %OutPutFile% 
if defined DNS2 echo ............................................... >> %OutPutFile% 
if defined DNS2 echo DNS2 >> %OutPutFile% 
if defined DNS2 ping %DNS2% >> %OutPutFile% 
if defined KnownGoodInternet echo. >> %OutPutFile% 
if defined KnownGoodInternet echo .................................. >> %OutPutFile% 
if defined KnownGoodInternet echo Known good internet address >> %OutPutFile% 
if defined KnownGoodInternet ping %KnownGoodInternet% >> %OutPutFile% 
goto runagain 
