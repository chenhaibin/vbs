<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">
<title>mp3 Jukebox</title>
</head>

<body bgcolor="#FFFFFF">
<%
''''''''''''''''''''''''
'Constants to used later
''''''''''''''''''''''''

Dim strOutput
commaspace = ", "
carriageReturn = vbCrLf
%>
<form action="<%= request.servervariables("script_name") %>"
method="POST">
    <div align="center"><center><table border="1" cellspacing="0"
    width="80%">
        <tr>
            <td><a name="rol"></a><p align="center"><img
            src="ray%20of%20light.gif" width="199" height="201"> </p>
            <p align="center"><input type="checkbox" name="opt"><font
            size="2" face="Tahoma">Madonna - Ray Of Light</font></p>
            </td>
            <td><table border="0" cellspacing="0">
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\Drowned%20world%20-%20Substitute%20for%20Love.mp3">1.
                    Drowned World / Substitute for Love</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\swim.mp3">2.
                    Swim</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\ray%20of%20light">3.
                    Ray Of Light</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\candy%20perfume%20girl.mp3">4.
                    Candy Perfume Girl</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\skin.mp3">5.
                    Skin</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\Nothing%20Really%20matters.mp3">6.
                    Nothing Really Matters</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\sky%20fits%20heaven.mp3">7.
                    Sky Fits Heaven</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\Shanti%20-%20Ashtangi.mp3">8.
                    Shanti / Ashtangi</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\frozen.mp3">9.
                    Frozen</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\the%20power%20of%20good-bye.mp3">10.
                    The Power Of Goodbye</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\to%20have%20and%20not%20to%20hold.mp3">11.
                    To Have and Not To Hold</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\Little%20Star.mp3">12.
                    Little Star</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\ray%20of%20light\Mer%20girl.mp3">13.
                    Mer Girl</font></td>
                </tr>
            </table>
            </td>
        </tr>
        <tr>
            <td align="center"><a name="ds"></a><img
            src="due%20south.gif" width="199" height="201"><p
            align="center"><input type="checkbox" name="C32"><font
            size="2" face="Tahoma">Soundtrack - dueSouth</font> </p>
            </td>
            <td><div align="left"><table border="0">
                <tr>
                    <td><input type="checkbox" name="C15"><font
                    size="1" face="Tahoma">1. dueSouth Theme</font></td>
                </tr>
                <tr>
                    <td><input type="checkbox" name="C16"><font
                    size="1" face="Tahoma">2. Bone Of Contention</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C17">3. Cabin Music</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C18">4. Possession
                    (Piano Version)</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C19">5. Horses</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C20">6. Akua Tuta</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C21">7. American Woman</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C22">8. Henry Martin</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C23">9. Ride Forever</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C24">10. Flying</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C25">11. dueSouth Theme</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C26">12. Neon Blue</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C27">13. Victoria's
                    Secret</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C28">14. Calling
                    Occupants...</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C29">15. Eia, Mater</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C30">16. Fraser / Inuit
                    Soliloquy</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C31">17. Dief's In Love</font></td>
                </tr>
            </table>
            </div></td>
        </tr>
        <tr>
            <td align="center"><a name="surr"></a><img
            src="surrender.gif" width="199" height="201"><p><input
            type="checkbox" name="C33"><font size="2"
            face="Tahoma">Chemical Brothers - Surrender</font> </p>
            </td>
            <td><table border="0">
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C34">1. Music: Response</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C35">2. Under the
                    Influence</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C36">3. Out of Control</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C37">4. Orange Wedge</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C39">5. Let Forever Be</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C38">6. The Sunshine
                    Underground</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C40">7.Asleep from Day</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C41">8. Got Glint?</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C42">9. Hey Boy Hey
                    Girl</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C43">10.Surrender</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C44">11. Dream On</font></td>
                </tr>
            </table>
            </td>
        </tr>
        <tr>
            <td align="center"><a name="sc"></a><img
            src="sheryl%20crow.gif" width="199" height="201"><p><input
            type="checkbox" name="C45"><font size="2"
            face="Tahoma">Sheryl Crow - Sheryl Crow</font></p>
            </td>
            <td><table border="0">
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C46">1. My Favourite
                    Mistake</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C47">2. There Goes The
                    Neighbourhood</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C48">3. Riverwide</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C49">4. It Don't Hurt</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C50">5. Maybe That's
                    Something</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C51">6. Am I Getting
                    Through (parts I and II)</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C52">7. Anything But
                    Down</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C53">8. The Different
                    Kind</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C54">9. Mississippi</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C55">10. Members Only</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C56">11. Crash and Burn</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C57">12. Resuscitation</font></td>
                </tr>
            </table>
            </td>
        </tr>
        <tr>
            <td align="center"><img src="hackers.gif" width="199"
            height="201"><p><a name="hack"></a><input
            type="checkbox" name="C58"><font size="2"
            face="Tahoma">Soundtrack - Hackers</font></p>
            </td>
            <td><table border="0">
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C59">1. Original
                    Bedroom Rockers</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C60">2. Cowgirl</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C61">3. Voodoo People</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C62">4. Open Up</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C63">5. Phoebus Apollo</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C64">6. The Joker</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C65">7. Halcyon &amp;
                    On &amp; On</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C66">8. Communicate</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C67">9. One Love</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C68">10. Connected</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C69">11. Eyes, Lips,
                    Body</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C70">12 . Good Grief</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C71">13. Richest Junkie
                    Still Alive</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="C73">14. Heaven Knows</font></td>
                </tr>
            </table>
            </td>
        </tr>
        <tr>
            <td>&nbsp;</td>
            <td><table border="0">
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\irrest.mp3">1.
                    Irresistable U Are</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\intense.mp3">2.
                    Intense</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\Iamifee.mp3">3.
                    I Am I Feel</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\alisha.mp3">4.
                    Alisha Rules the World</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\whiteroo.mp3">5.
                    White Room</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\stonein.mp3">6.
                    Stone In My Shoe</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\persline.mp3">7.
                    Personality Lines</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\indestru.mp3">8.
                    Indestructible</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\wontmiss.mp3">9.
                    I Won't Miss You</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\goldrule.mp3">10.
                    The Golden Rule</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\justlike.mp3">11.
                    Just the Way You Like It</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\airbrea.mp3">12.
                    Air We Breathe</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\aattic\adoreu.mp3">13.
                    Adore U</font></td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\ironic.mp3">1.
                    Ironic</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\frontrow.mp3">2.
                    Front Row</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\handinmy.mp3">3.
                    Hand In My Pocket</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\alliwant.mp3">4.
                    All I Really Want</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\notdoc.mp3">5.
                    Not The Doctor</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\oughtakn.mp3">6.
                    You Oughta Know</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\thanku.mp3">7.
                    Thank U</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\alanis\youlearn.mp3">8.
                    You Learn</font></td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\touchmy.mp3">1.
                    I Touch Myself</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\austheme.mp3">2.
                    Austin's Theme</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\bbc.mp3">3. BBC</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\bossanov.mp3">4.
                    Soul Bossanova</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\Carnival.mp3">5.
                    Carnival</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\female.mp3">6.
                    Female of the Species</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\looklove.mp3">7.
                    The Look Of Love</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\medley.mp3">8.
                    Shag-A-Delic Medley</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\theseday.mp3">9.
                    These Days</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\apowers.mp3">10.
                    Austin Powers</font></td>
                </tr>
                <tr>
                    <td><font size="1" face="Tahoma"><input
                    type="checkbox" name="opt"
                    value="\\10.0.0.1\mp3s\austin\youshow.mp3">11.
                    You Showed Me</font></td>
                </tr>
            </table>
            </td>
        </tr>
    </table>
    </center></div><p align="center"><input type="submit"
    name="B1" value="Submit"><input type="reset" name="B2"
    value="Reset"></p>
</form>
<%
''''''''''''''''''''''''''''''''
'User input into Variables
''''''''''''''''''''''''''''''''

opt=request.form("opt")

''''''''''''''''''''''''''''''''
'Creates file system object
''''''''''''''''''''''''''''''''

set fso = createobject("scripting.filesystemobject")

''''''''''''''''''''''''''''''''
'Creates a text file of name playlist.m3u on the web server
''''''''''''''''''''''''''''''''

set pllst = fso.createtextfile(server.mappath("/playlist.m3u"), true)

''''''''''''''''''''''''''''''''
'write users selection into text file
''''''''''''''''''''''''''''''''

if opt <> "" then
        stroutput=opt
                stroutput=replace(stroutput, commaspace, carriagereturn)
        pllst.writeline stroutput
end if

''''''''''''''''''''''''''''''''
'close the text file
''''''''''''''''''''''''''''''''

pllst.close

%></body>
</html>
