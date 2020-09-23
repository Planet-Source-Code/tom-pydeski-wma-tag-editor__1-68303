<div align="center">

## WMA Tag Editor

<img src="PIC200747124464769.jpg">
</div>

### Description

WMA Tag Editor - submitted by Tom Pydeski

I saw a lot of examples for editing mp3 tags, but not much to read wma files.

there was one by Somenon at

http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61254&amp;lngWId=1

Although the original author's code worked, I wanted a way to read the file

without parsing through each character looking for a certain string.

(I kept his original routines as reference)

I utilized the class structure from InfoTag, which read WMA files, but did it in a way

that would not allow writing back to the file.

So I initially tried to dig into the file and try to read it in blocks.

It wasn't long before I realized the structure was way more complicated than I had

originally thought. I did some digging and found the attached document

"Advanced Systems Format (ASF) Specification" from Microsoft.

Using this as a guide, I built the structures neccessary for each header object.

I spent many weeks developing this and my wife hated that I was always on the 'puter,

but I wanted to finish this. It will read and write the basic tags and I put in a

treeview to display the entire file structure by object.

I'd like to add another flexgrid and read multiple files, but that's for later.

The other thing that I had hoped to accomplish was this:

When using my GotRadio submission, I found that temp files were created with the filename of

the songs played. These were located in the temporary internet directory.

The structure of these files is different than the structure of wma's that I had burned.

I was hoping to be able to learn how to modify the temp files to allow them to be played

by media player, but try as i might, when I converted the temp. file to a wma with the set

structure, media player would not play it.

I'm attaching one of those temp files before and after modifying it.

maybe someone with more knowledge can find the errors of my way
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2007-04-07 12:38:32
**By**             |[Tom Pydeski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-pydeski.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[WMA\_Tag\_Ed205892472007\.zip](https://github.com/Planet-Source-Code/tom-pydeski-wma-tag-editor__1-68303/archive/master.zip)








