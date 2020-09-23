<div align="center">

## Make your own Mp3 Player\!


</div>

### Description

Every wanted to make your own Mp3 Player instead of using WinAmp etc...This is your chance. Read this tuturial and find out how you can make your own Mp3 Player.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin.md)
**Level**          |Intermediate
**User Rating**    |4.7 (93 globes from 20 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-make-your-own-mp3-player__1-26022/archive/master.zip)





### Source Code

<html>
<head>
</head>
<body>
<p><b><font size="2" face="Verdana"><font color="#FF0000">Welcome!</font><br>
<font color="#FF0000">
In the tutorial I will try to explain how to make your own mp3 player using the
Media Player control.<br>
It will be a fully featured audio player with a playlist and options like:
&quot;Repeat&quot; and &quot;Random Play&quot;.<br>
</font><font color="#808080"><br>
Gray Text = Things to do!<br>
</font><font color="#0000FF">Blue Text = Source Code<br>
</font><font color="#FF0000">Red Text&nbsp; = Information about the code
etc...&nbsp;</font></font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#808080">1) Start VB<br>
2) Press CTRL + T<br>
3) Insert: Windows Media Player Control, Microsoft Common Dialog Control and Microsoft
Windows Common Controls.<br>
4) Put on your form: 6 Command buttons (Playback), A Label (Time Label), A
Textbox, 2 Sliders (Volume and seekbar), A Listbox (PLS)<br>
5) Set the Media Player Control Invisible.</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">The first thing we'll have
to do is making a &quot;Open File&quot; button. This way a user is able to
choose an audio file.<br>
</font><font size="2" face="Verdana" color="#808080">Select one of the 6 command
buttons you've created earlier in this tutorial and put the following code on
it.</font></b></p>
<p><b><font size="2" color="#0000FF" face="Verdana"><i>On Error Resume Next</i></font></b><font size="2" face="Verdana" color="#0000FF"><b><i><br>
CommonDialog1.Filter = "Audio Files|*.wav;*.mid;*.mp3;mp2;*.mod|"<br>
CommonDialog1.Flags = cdlOFNHideReadOnly<br>
CommonDialog1.CancelError = True<br>
CommonDialog1.DialogTitle = "Choose an mediafile to open"<br>
CommonDialog1.FileName = ""<br>
CommonDialog1.ShowOpen<br>
<br>
List1.AddItem CommonDialog1.FileName<br>
List1.ListIndex = List1.ListIndex + 1<br>
MediaPlayer1.FileName = CommonDialog1.FileName<br>
Text1.Text = CommonDialog1.File</i></b></font><font size="2" face="Verdana" color="#0000FF"><b><i>name</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors<br>
Line2: It will display only mp3,
wav, mid files etc...<br>
Line3: This will remove the &quot;Read Only&quot; checkbox at the end of the open
dialog<br>
Line4: This will handle the error you get if you click the cancel button<br>
Line5: This will set the text between the &quot;...&quot; on the titlebar of the
open dialog<br>
Line5: This will show the open dialog<br>
Line6: An Empty Line!<br>
Line7: This will put the file
you've selected in the listbox (Used as playlist Control)<br>
Line8: This will select the file you've chosen in the PlayList<br>
Line9: This tells the Media Player Control which file it needs to play<br>
Line10: </b></font><b><font size="2" face="Verdana" color="#FF0000"> Put the name of the file we're going to play in the &quot;Filename&quot;
textbox.</font></b></p>
<hr noshade color="#000000">
<p><font size="2" face="Verdana" color="#FF0000"><b>Now that we're playing a
file we can put other code in our player such as &quot;Play Selected
Track&quot;...<br>
</b></font><b><font size="2" face="Verdana" color="#808080">Select one of the 5
command buttons you've created earlier in this tutorial and put the following
code on it.</font></b></p>
<p><b><font size="2" color="#0000FF" face="Verdana"><i>On Error Resume Next</i></font></b><font color="#0000FF" size="2" face="Verdana"><b><i><br>
MediaPlayer1.FileName = List1.Text<br>
Mediaplayer1.Play<br>
</i></b></font><font size="2" face="Verdana" color="#0000FF"><b><i>text1.text =
mediaplayer1.filename</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors</b></font><b><font size="2" face="Verdana" color="#FF0000"><br>
Line2: The first line tells the Media Player Control the filename. In this case
the selected item in the PlayList.<br>
Line3: The second line tells the control that it must play the filename which
was set above<br>
Line4: Put the name of the file we're going to play in the &quot;Filename&quot;
textbox.</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">I don't know what you think,
But I'd like to be able to Pause the playing track :-)<br>
</font><font size="2" face="Verdana" color="#808080">Select one of the 4 command
buttons you've created earlier in this tutorial and put the following code on
it.</font></b></p>
<p><b><font size="2" color="#0000FF" face="Verdana"><i>On Error Resume Next</i></font></b><font color="#0000FF" size="2" face="Verdana"><b><i><br>
If MediaPlayer1.PlayState = mpPlaying Then<br>
MediaPlayer1.Pause<br>
Else<br>
MediaPlayer1.Play<br>
End If</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors</b></font><br>
<b><font size="2" face="Verdana" color="#FF0000">Line2: If the Media Control is
playing a file then...<br>
Line3: Pause the playing file!<br>
Line4: Else. If it's not playing a file. So it's either stoped or already
paused...<br>
Line5: Play the file which is still in the memory of the Media Player Control.</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Every audio player contains
a stop button. Well here's the code for it...<br>
</font><font size="2" face="Verdana" color="#808080">Select one of the 3 command
buttons you've created earlier in this tutorial and put the following code on
it.</font></b></p>
<p><b><font size="2" color="#0000FF" face="Verdana"><i>On Error Resume Next<br>
Mediaplayer1.stop</i></font></b></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors</b></font><br>
<b><font size="2" face="Verdana" color="#FF0000">Line2: Stop playing the current
file.</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Well, If a player contains a
PlayList I'd like to be able to switch between my tracks. Here's the
&quot;Previous Track&quot; code.<br>
</font><font size="2" face="Verdana" color="#808080">Select one of the 2 command
buttons you've created earlier in this tutorial and put the following code on
it.</font></b></p>
<p><b><font size="2" color="#0000FF" face="Verdana"><i>On Error Resume Next</i></font></b><font color="#0000FF" size="2" face="Verdana"><b><i><br>
List1.ListIndex = List1.ListIndex - 1<br>
MediaPlayer1.FileName = List1.Text<br>
MediaPlayer1.Play<br>
text1.text =
mediaplayer1.filename
</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors</b></font><b><font size="2" face="Verdana" color="#FF0000"><br>
Line2: This will go one item back from the selected item.<br>
Line3: </font></b><font size="2" face="Verdana" color="#FF0000"><b>
This tells the Media Player Control which file it needs to play</b></font><b><font size="2" face="Verdana" color="#FF0000"><br>
Line4: Say to the Media Player control : Play the file I've set above!<br>
Line5: Put the name of the file we're going to play in the &quot;Filename&quot;
textbox.</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Well, If a player contains a
PlayList I'd like to be able to switch between my tracks. Here's the &quot;Next
Track&quot; code.<br>
</font><font size="2" face="Verdana" color="#808080">Select the last
commandbutton that's left and put the following code on it.</font></b></p>
<p><b><font size="2" color="#0000FF" face="Verdana"><i>On Error Resume Next</i></font></b><font color="#0000FF" size="2" face="Verdana"><b><i><br>
List1.ListIndex = List1.ListIndex + 1<br>
MediaPlayer1.FileName = List1.Text<br>
MediaPlayer1.Play<br>
</i></b></font><font size="2" face="Verdana" color="#0000FF"><b><i>text1.text =
mediaplayer1.filename</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors</b></font><b><font size="2" face="Verdana" color="#FF0000"><br>
Line2: This will go one item further than the selected item.<br>
Line3: Tell the Media Player control which file it needs to load. In this case the
selected item in the PlayList.<br>
Line4: Play the file!<br>
Line5: Put the name of the file we're going to play in the &quot;Filename&quot;
textbox.</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Now that the PlayBack
controls are done we can add some more code to our player.<br>
</font><font size="2" face="Verdana" color="#808080">Put the following code in your form. It's a function which is called in the next Sub.</font></b></p>
<p><b><font size="2" face="Verdana" color="#0000FF"><i>Function ConvertTime(i As Integer)<br>
Secs = i Mod 60<br>
Mins = Int(i / 60) Mod 60<br>
Hours = Int(i / 3600)<br>
If Secs &lt; 10 Then Secs = "0" &amp; Secs<br>
If Mins &lt; 10 Then Mins = "0" &amp; Mins<br>
ConvertTime = Hours &amp; ":" &amp; Mins &amp; ":" &amp; Secs<br>
End Function</i></font></b></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: This is the name of
the function and the statement which tells VB it's a function.<br>
Line2: We make a variable named secs which converts I (which is specified when
the function is called) to seconds<br>
Line3: We make a variable named mins which converts I to minutes<br>
Line4: We make a variable which hours which converts&nbsp; I to hours<br>
Line5: If the number of seconds is less the then we put a 0 before the seconds
like this: 01,02,03 etc...<br>
Line6: The same as above but now with minutes<br>
Line7: Now we update the sub with the output format which you can get in the
next sub.<br>
Line8: The end of this function :-(</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Now that you've placed the
above function in your code we're ready to call it.<br>
</font><font color="#808080"><font size="2" face="Verdana">Insert a Timer
Control in your project and put the following code on it. Don't forget to set
it's interval to: &quot;1000&quot;.</font></font><font size="2" face="Verdana" color="#808080"><br>
Interval=&quot;1000&quot;: This means that the timer will update the Timer
Window every second.</font></b></p>
<p><b><font color="#0000FF" size="2" face="Verdana"><i>If MediaPlayer1.PlayState = mpPlaying Then<br>
Label1.Caption = ConvertTime(Round(MediaPlayer1.CurrentPosition, 0)) &amp; " / " &amp; ConvertTime(Round(MediaPlayer1.Duration, 0))<br>
Else<br>
Label1.Caption = &quot;00:00:00 / 0:00:00"<br>
End If</i></font></b></p>
<p><font color="#FF0000"><b><font size="2" face="Verdana">Line1: Are we playing
a file ?<br>
Line2: If we are playing a file update the Timer Window every second with the
time of the file you're playing<br>
Line3: If we are not playing a file...<br>
Line4: Put the following text in your Timer Window: &quot;00:00:00 /
00:00:00&quot;<br>
Line5: Stop our check function</font></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Note: The time of the file
will be showed like this: &quot;03:46:13 / 04:12:34&quot;<br>
The &quot;03:46:13&quot; indicates the current playing time and
&quot;04:12:34&quot; indicates the total time of the file.</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">We have playback controls, a
PlayList and a Timer Window but we don't have a seekbar yet!<br>
</font><font size="2" face="Verdana" color="#808080">Insert a Timer Control in
your project and put the following code on it. Don't forget to set it's interval
to: &quot;1000&quot;.<br>
Interval=&quot;1000&quot;: This means that the timer will update the Seekbar
control's position every second.</font></b></p>
<p><b><i><font color="#0000FF" size="2" face="Verdana">On Error Resume Next<br>
Slider1.Max = MediaPlayer1.Duration<br>
Slider1.Value = MediaPlayer1.CurrentPosition</font></i></b></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors</b></font><font color="#FF0000"><b><font size="2" face="Verdana"><br>
Line2: The Maximum value of the slider is the duration of the file we're playing<br>
Line3: Update the position of the slider to the current position of the Media
Player Control</font></b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">We're not ready with our
slider yet because we want to change the position of the track when moving
the slider.<br>
</font><font size="2" face="Verdana" color="#808080">Put the following code in
the Slider1_Scroll() function.</font></b></p>
<p><font color="#0000FF" size="2" face="Verdana"><b><i>On Error Resume next<br>
MediaPlayer1.CurrentPosition = Slider1.Value</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from ginving errors<br>
Line2: Update the position of the file we're playing to the slider's position.
Easy huh?</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Earlier in this tutorial I
said&nbsp; you had to put 2 slider control's on your form.<br>
Well, We've already used 1 slider so now we're going to use the second one. This
one is for the Volume Control.<br>
</font><font size="2" face="Verdana" color="#808080">Put the following code in
the Slider2_Scroll() function and set the Max Value of the slider to: &quot;2500&quot;</font></b></p>
<p><font color="#0000FF" size="2" face="Verdana"><b><i>Dim a As Integer, b As Integer<br>
Dim d, c<br>
c = Slider2.Value - 2500<br>
MediaPlayer1.Volume = c<br>
b = Slider2.Min<br>
a = Slider2.Value</i></b></font></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: Make 2 variables
called &quot;a&quot; and &quot;b&quot; and say to VB they are an Integer.<br>
Line2: Set variable &quot;D&quot; and &quot;C&quot;<br>
Line3: Update variable c with Slider2's value - 2500<br>
Line4: Set the volume of the Media Player control from Slider2's value<br>
Line5: Variable &quot;b&quot; is The Minium Value of Slider 2<br>
Line6: Variable &quot;a&quot; is the Maximum Value of Slider 2</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Well, Just to make sure a
few things you need to place this code in the Form itself.<br>
</font><font size="2" face="Verdana" color="#808080">Put the following code in
the Form_Load function.</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><i><b>timer1.interval = 1000<br>
timer2.interval = 1000<br>
Slider2.Max = 2500</b></i></font></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: Set the interval of
Timer1 to &quot;1000&quot; in case you forgot :-)<br>
Line2: Set the interval of Timer2 to &quot;1000&quot; in case you forgot :-)<br>
Line3: Set the maximum value of the volume slider to &quot;2500&quot; in case
you forgot :-)</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Congratulations! You're
Audio Player is now ready!<br>
You will find some more useful code for it below...</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">If you want to have an
&quot;Mute&quot; option in your Audio Player use the following code.</font></b></p>
<p><font color="#0000FF" size="2" face="Verdana"><b><i>If MediaPlayer1.Mute = True Then<br>
MediaPlayer1.Mute = False<br>
Else<br>
MediaPlayer1.Mute = True<br>
End If</i></b></font></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: If the sound is muted
then...<br>
Line2: UnMute the sound!<br>
Line3: Mute the sound because it's not muted yet!<br>
Line4: Stop the check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Your audio player contains a
PlayList but why would you need a PlayList if it cannot contain more than 1
file.<br>
Here's the code to add files to your PlayList.</font></b></p>
<p><font color="#0000FF" size="2" face="Verdana"><i><b>On Error Resume Next<br>
CommonDialog1.Filter = "Audio Files|*.wav;*.mid;*.mp3;mp2;*.mod|"<br>
CommonDialog1.Flags = cdlOFNHideReadOnly<br>
CommonDialog1.CancelError = True<br>
CommonDialog1.DialogTitle = &quot;Add File&quot;<br>
CommonDialog1.FileName = ""<br>
CommonDialog1.ShowOpen<br>
<br>
List1.AddItem CommonDialog1.FileName</b></i></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors<br>
Line2: It will display only mp3,
wav, mid files etc...<br>
Line3: This will remove the &quot;Read Only&quot; checkbox at the end of the open
dialog<br>
Line4: This will handle the error you get if you click the cancel button<br>
Line5: This will set the text between the &quot;...&quot; on the titlebar of the
open dialog<br>
Line6: This will clear the previous selected Filename in the Open Dialog<br>
Line7: This will show the open dialog<br>
Line8: An Empty Line!<br>
Line9: Add the selected file to our PlayList</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to remove
the selected item from your PlayList...</font></b></p>
<p><i><b><font size="2" face="Verdana" color="#0000FF">If List1.ListIndex > -1 Then<br>
On Error Resume Next<br>
If list1.text = mediaplayer1.filename then msgbox &quot;You can't remove the
file you're playing&quot;:exit sub<br>
List1.RemoveItem List1.ListIndex<br>
End If</font></b></i></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: If the PlayList's
index is bigger than -1 go on. This means that there's actually something.<br>
Line2: Prevents the program
from giving errors<br>
Line3: If the item you're trying to remove is the current playing item show a
msgbox telling the user the item cannot be removed. Also stop the code so it
won't be removed. </font><font size="2" face="Verdana" color="#800080">This code
is needed because it will keep your PlayList working. If you remove this line
the player can't determine anymore which file was playing and it won't go next
anymore from the file you were playing. It starts the list again.</font><font size="2" face="Verdana" color="#FF0000"><br>
Line4: Remove the selected line from the PlayList<br>
Line5: Stop our check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to move the
selected item in your PlayList one line up.</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><b><i>on error resume next<br>
Dim nItem As Integer<br>
With lstItems<br>
If List1.ListIndex &lt; 0 Then Exit Sub<br>
nItem = List1.ListIndex<br>
If nItem = 0 Then Exit Sub<br>
List1.AddItem List1.Text, nItem - 1<br>
List1.RemoveItem nItem + 1<br>
List1.Selected(nItem - 1) = True<br>
End With</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors<br>
Line2: Set 'NItem&quot; as variable and tell VB that it's an Integer<br>
Line3: Tell VB that we're working with the List Items<br>
Line4: If the List Index is smaller that 0 it can't move up anymore so STOP the
code<br>
Line5: Tell VB that our variable the index of the selected item in our PlayList
is<br>
Line6: If the index is 0 STOP because the line can't move up further.<br>
Line7: Add the listindex text one item before the selected item<br>
Line8: Remove the current selected item<br>
Line9: Select the line we've just moved one line up<br>
Line10: Tell VB we're not working with the List Items anymore</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to move the
selected item in your PlayList one line down.</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><b><i>on error resume next<br>
Dim nItem As Integer<br>
With lstItems<br>
If List1.ListIndex &lt; 0 Then Exit Sub<br>
nItem = List1.ListIndex<br>
If nItem = List1.ListCount - 1 Then Exit Sub<br>
List1.AddItem List1.Text, nItem + 2<br>
List1.RemoveItem nItem<br>
List1.Selected(nItem + 1) = True<br>
End With</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors<br>
Line2: Set 'NItem&quot; as variable and tell VB that it's an Integer<br>
Line3: Tell VB that we're working with the List Items<br>
Line4: If the List Index is smaller that 0 it can't move up anymore so STOP the
code<br>
Line5: Tell VB that our variable the index of the selected item in our PlayList
is<br>
Line6: If the index is at the end of the PlayList it can't go down anymore so
STOP the code<br>
Line7: Add the listindex text one item after the current item. The listbox works
a bit strange so line1 is actually line2<br>
Line8: Remove the current selected item<br>
Line9: Select the line we've just moved one line down<br>
Line10: Tell VB we're not working with the List Items anymore</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to clear the
whole PlayList...</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><b><i>ask = MsgBox("Do you want to clear your list ?", vbQuestion + vbYesNo, "Confirm")<br>
If ask = vbYes Then<br>
List1.clear<br>
Else<br>
End If</i></b></font></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: We show a Msgbox
which asks the user if he/she really want to clear the PlayList<br>
Line2: If they answer YES...<br>
Line3: The PlayList will be cleared<br>
Line4: If they answer something else...NO in this case!<br>
Line5: Don't do anything<br>
Line6: Stop the check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to save your
PlayList...</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><b><i>On Error Resume Next<br>
CommonDialog1.Filter = "PlayList File (M3u)|*.m3u|PlayList File
(Pls)|*.pls&quot;<br>
CommonDialog1.DialogTitle = &quot;Save List"<br>
CommonDialog1.Flags = cdlOFNHideReadOnly<br>
CommonDialog1.ShowSave<br>
CommonDialog1.CancelError = True<br>
<br>
Open CommonDialog1.FileName For Output As #1<br>
For X = 0 To List1.ListCount - 1<br>
Print #1, List1.List(X)<br>
Next X<br>
Close #1</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors<br>
Line2: It can only save your PlayList as: &quot;M3u or Pls&quot;<br>
Line3: This will set the text between the &quot;...&quot; on the titlebar of the
save dialog<br>
Line4: This will remove the &quot;Read Only&quot; checkbox at the end of the
save dialog<br>
Line5: This will show the save dialog<br>
Line6: This will handle the error you get if you click the cancel button<br>
Line7: An empty line!<br>
Line8: Open the selected file and place it in a variable called #1<br>
Line9: Make sure we save the whole PlayList<br>
Line10: Write the contents to the selected file<br>
Line11: Go on till we have saved all items<br>
Line12: Close the file</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to open a
previously saved PlayList file...</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><b><i>On Error Resume Next<br>
On Error GoTo err<br>
Close #1<br>
Dim X<br>
OpenFile:<br>
CommonDialog1.Filter = "All Supported|*.m3u;*.pls|&quot;<br>
CommonDialog1.DialogTitle = &quot;Open List"<br>
CommonDialog1.Flags = cdlOFNHideReadOnly<br>
CommonDialog1.ShowOpen<br>
CommonDialog1.CancelError = True<br>
Open CommonDialog1.FileName For Input As #1<br>
List1.clear<br>
Do<br>
Input #1, X<br>
List1.AddItem (X)<br>
Loop<br>
Close #1<br>
err:<br>
Exit Sub</i></b></font></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Line1: Prevents the program
from giving errors<br>
Line2: If there's an error go the a sub called: &quot;err&quot;<br>
Line3: Close all files so we don't get any errors while trying to open a file<br>
Line4: A sub called &quot;OpenFile<br>
Line5: It will display only mp3,
wav, mid files etc...<br>
Line6: This will set the text between the &quot;...&quot; on the titlebar of the
open dialog<br>
Line7: This will remove the &quot;Read Only&quot; checkbox at the end of the
save dialog<br>
Line8: This will show the open dialog<br>
Line9: This will handle the error you get if you click the cancel button<br>
Line10: Load the file that was selected into the memory as variable #1<br>
Line11: Clear the list before adding new items in it<br>
Line12: Do function<br>
Line13: Load all contents into variable X<br>
Line14: Add X to our PlayList<br>
Line15: Gon on till we have placed all files in our PlayList<br>
Line16: Close the files which was opened<br>
Line17: Error Handler sub which is called above<br>
Line18: Stop the code. If case there's an error</b></font></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to repeat
the current playing file...<br>
</font><font size="2" face="Verdana" color="#808080">Insert a checkbox on your
form. Add the following code in the &quot;EndOfStream&quot; function of the
Media Player control.</font></b></p>
<p><b><font size="2" face="Verdana" color="#0000FF"><i>if check1.value = 1 then<br>
mediaplayer1.play<br>
else<br>
end if</i></font></b></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: If the checkbox is
checked...<br>
Line2: Play the file again!<br>
Line3: If it's not checked...<br>
Line4: Stop our check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to stop
playing after the current playing track...<br>
</font><font size="2" face="Verdana" color="#808080">Put a checkbox on your
form. Add the following code in the &quot;EndOfStream&quot; function of the
Media Player control.</font></b></p>
<p><b><i><font size="2" face="Verdana" color="#0000FF">if check2.value = 1 then<br>
mediaplayer1.stop<br>
else<br>
end if</font></i></b></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: If the checkbox is
checked...<br>
Line2: Stop Playing!<br>
Line3: If it's not checked...<br>
Line4: Stop our check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">Here's the code to play
normal. After it's done it will start playing the next track in your PlayList.<br>
</font><font size="2" face="Verdana" color="#808080">Put a checkbox on your
form. Add the following code in the &quot;EndOfStream&quot; function of the
Media Player control.</font></b></p>
<p><b><i><font size="2" face="Verdana" color="#0000FF">if check3.value = 1 then</font></i></b><font size="2" face="Verdana" color="#0000FF"><b><i><br>
list1.listindex = list1.listindex + 1<br>
mediaplayer1.filename = list1.text<br>
mediaplayer1.play<br>
text1.text = mediaplayer1.filename<br>
else<br>
end if</i></b></font></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: If the checkbox is
checked...<br>
Line2: Select the next line in the PlayList from the selected item.<br>
Line3: Tell the Media Player Control that it needs to load the selected item in
the PlayList<br>
Line4: Play the loaded file<br>
Line5: Update our Filename Window with the current playing track<br>
Line6: If it's not checked...<br>
Line7: Stop our check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">A very important thing that
you must do is the following...<br>
</font><font size="2" face="Verdana" color="#808080">You must make a function
that will select the current playing track every time it's done playing.<br>
If you don't do this the player will play the next track from the item you've
selected and not from the file you're playing.<br>
Put the following code in the &quot;EndOfStream&quot; function of the Media
Player control...</font></b></p>
<p><font size="2" face="Verdana" color="#0000FF"><b><i>filename.text =
mediaplayer1.filename</i></b></font></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: Put the filename of
the current played track in a textbox so we can look it up later.<br>
</font><font size="2" face="Verdana" color="#808080">Now add the following code
in the &quot;_Change&quot; function of the textbox called &quot;filename&quot;.</font></b></p>
<p><b><font size="2" face="Verdana" color="#0000FF"><i>If Trim(filename.Text) &lt;> "" Then<br>
For i = 0 To List1.ListCount - 1<br>
If Left(List1.List(i), Len(Trim(filename.Text))) = Trim(filename.Text) Then<br>
List1.Selected(i) = True<br>
Else<br>
List1.Selected(i) = False<br>
End If<br>
Next i<br>
End If</i></font></b></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Line1: Check if the textbox
contains some text<br>
Line2: Search the whole list<br>
Line3: Search the list for contents of the textbox<br>
Line4: If it's found select the item that was found --&gt; Now the player goes
next from this selected item. This was the playing item!<br>
Line5: If it was not found don't select anything. But the file is always found
because you were playing it :-)<br>
Line6: Stop check function<br>
Line7: Go on till end of PlayList<br>
Line8: Stop check function</font></b></p>
<hr noshade color="#000000">
<p><b><font size="2" face="Verdana" color="#FF0000">I hope you know how you can
make your own Mp3 Player after reading this tutorial.<br>
Please, Send as much feedback as you want and </font><font color="#800080"><font size="2" face="Verdana">DON'T
VORGET TO VOTE FOR MY WORK!</font></font><font size="2" face="Verdana" color="#FF0000"><br>
<br>
Some people are complaining about using the Media Player Control but I don't
care.<br>
As long as we have fun with it it's good, and another great thing is that you
don't have to use WinAmp anymore :-)</font></b></p>
<p><b><font size="2" face="Verdana" color="#FF0000">Since English is not my
Primary Language there can be some spelling faults in it.<br>
Please, Report them and I'll fix it. Dutch is my primary language...</font></b></p>
<p><font size="2" face="Verdana" color="#FF0000"><b>Enjoy your own Mp3 Player !!</b></font></p>
<p><b><font size="2" face="Verdana" color="#008000">K</font><font size="2" face="Verdana" color="#000080">E</font><font size="2" face="Verdana" color="#FF0000">V</font><font size="2" face="Verdana" color="#FF9900">I</font><font size="2" face="Verdana" color="#808080">N</font><font size="2" face="Verdana" color="#CC9900">,<br>
<a href="mailto:kevin_verp@hotpop.com">kevin_verp@hotpop.com</a></font></b></p>
</body>
</html>

