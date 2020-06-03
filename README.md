# SpotifyTaskbarFix
Band-aid fix for the never acknowledged issue with Spotify's taskbar icon appearing on the wrong monitor in a multi-monitor configuration on Windows. :pray:

### FAQ
**Q**: How do I use this?  
**A**: Download `SpotifyTaskbarFix.exe` from [here](https://github.com/aleab/SpotifyTaskbarFix/releases/latest) (or compile it yourself!) and run the program at system startup **with admin privileges**.

**Q**: How does this work?  
**A**: This thing is basically automating what you would do to make the taskbar icon move to where it should be: it moves Spotify's window from screen B where it starts in to the primary screen and back to B. This action is virtually instantaneous, so you won't even notice the window moving.

**Q**: Why does it need admin privileges?  
**A**: Because it needs to know when the Spotify process starts, and this information needs admin privileges.


### Useful Links:
[#1](https://community.spotify.com/t5/Desktop-Windows/Spotify-s-taskbar-icon-on-the-wrong-monitor-every-single-time/td-p/4669359) [#2](https://community.spotify.com/t5/Ongoing-Issues/Desktop-Spotify-s-taskbar-icon-on-the-wrong-monitor-every-single/idi-p/4888243): Community forum threads describing the issue and the steps to reproduce it.  
[#3](https://www.windowscentral.com/how-create-automated-task-using-task-scheduler-windows-10): How to create a scheduled task on Windows to run programs at startup.  
[#4](https://www.digitalcitizen.life/use-task-scheduler-launch-programs-without-uac-prompts): How to run a scheduled task as administrator.  
