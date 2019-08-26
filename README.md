# 3ds Max File Inspector

A simple WPF application shows 3ds max file (.MAX) properties without 3ds max application.

It is usuful to non-artist who want to know basic information but didn't install 3ds Max on their system. 

- Set Folder path > Gather Max files
- (Optional, use with caution) Can gather Max files from subdirectories recursively

Then it will show each files information includes,

- ACL/CRP malcious scripts affection: a warning sign will be marked if file was infected
- Filename
- Max Version: build version which file was authored in 
- Save as: file version saved as
- Show Info: same as 3ds Max > File > File Properties... > Contents
  - General
  - Mesh Totals
  - Scene Totals
  - External Dependancies
  - Objects
  - Materials
  - Used Plug-ins
  - Render Data
- Open folder: opens folder which contains picked file

## Compatibility

Autodesk 3ds Max Version 1.0 to 21.0 (Tested on 9.0 - 2019)

Support English, Chinese, and Japanese localized data
