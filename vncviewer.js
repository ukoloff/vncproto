//
// vnc:... URL handler (See <http://net.ekb.ru/help/VNC/>)
//

WScript.Interactive=false;

if(WScript.Arguments.length!=1) Help();

var Cmd=WScript.Arguments(0);

if('-install'==Cmd) Install();
else if('-uninstall'==Cmd) unInstall();
else if(/^vnc:(.+)/.test(Cmd)) Run(RegExp.$1);
else Help();

function Help()
{
 WScript.Echo('Usage:\n'
    +'\t'+WScript.ScriptName+' -install\n'
    +'\t'+WScript.ScriptName+' -uninstall\n'
    +'\t'+WScript.ScriptName+' vnc:HOST/IP\n');
 WScript.Quit();
}

function findBin()
{
 var F=WScript.ScriptFullName.replace(/\.[^.\/\\]*$/, '')+'.exe';
 var fso=new ActiveXObject("Scripting.FileSystemObject");
 if(fso.FileExists(F)) return F;
 WScript.Echo('File not found: '+F+'!');
 WScript.Quit();
}

function WSh()
{
 return WScript.CreateObject("WScript.Shell");
}

function Install()
{
 findBin();
 var X=WSh();
 X.RegWrite('HKCR\\vnc\\', 'VNC: Protocol');
 X.RegWrite('HKCR\\vnc\\URL Protocol', '');
 X.RegWrite('HKCR\\vnc\\shell\\open\\command\\',
    'wscript "'+WScript.ScriptFullName+'" "%1"');
}

function unInstall()
{
 var X=WSh();
 X.RegDelete('HKCR\\vnc\\shell\\open\\command\\');
 X.RegDelete('HKCR\\vnc\\shell\\open\\');
 X.RegDelete('HKCR\\vnc\\shell\\');
 X.RegDelete('HKCR\\vnc\\');
}

function Run(Host)
{
 var X=WSh();
 X.Run('"'+findBin()+'" "'+unescape(Host)+'"'
// +' -password=***'
 );
}

//--[EOF]------------------------------------------------------------
