Multimine Readme

Multimine was conceved early 2003. The plan was to write the game in C++ in an
effort to motivate myself into learning the language beyond the syntax. As a
result, this also meant having a crash course in the Windows SDK and WinSocks. It
certainly made things interesting when tracking down problems that spanned the
Windows Messenging system and the sockets.

That being said, there may be problems with this game, if you find one that isn't
on the list of known issues accompaning this README please let the AUTHOR know
about the problem, with a detailed report on how to replicate the issue. You may
get a mention of gratitude in the next release.

Don't forget to mention the version that you are playing with.
The version can be found in the about box :)


Thanks go out to
Quin for being a C++ and Windows programming helper
Tone for being the first person to ever play multimine (with me of course)
Conrad for thinking that we should actually make the game
Danno and Richard for wanting to actually play the game!

Who else wants some thanks?


Help! THE GAME WON'T START

Problem:
Unable to locate DLL
The dynamic link library gdiplus.dll could not be found in the specified path...

Solution:
Put gdiplus.dll into one of those directories. You can put it in the same directory
as multimine, or one of the windows directories, providing it doesnt break any licences.