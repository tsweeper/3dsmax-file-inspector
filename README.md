# 3ds Max File Inspector

A simple WPF application shows 3ds max file (.MAX) properties without 3ds max application.

It is useful to non-artist who want to know basic information but didn't install 3ds Max on their system. 

- Set Folder path > Gather 3ds Max files
- (Optional, use with caution) Can gather 3ds Max files from subdirectories recursively

Then it will show each files information includes,

- Detect ACL/CRP malicious scripts infection
  - A warning sign :warning: will be marked if file contains suspicious scripts
    > :information_source: This function use simple string match for just indication  
    > **:skull: CANNOT GUARANTEE detect variants and has no cleanup functionality :bangbang:**  
    > :thumbsup: For the recent (2015 and later) versions, it is recommended to use [3ds Max Scene Security Tools](https://apps.autodesk.com/3DSMAX/en/Detail/Index?id=7342616782204846316&appLang=en&os=Win64) instead
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
- Open folder: opens folder which contains selected file

## Compatibility

Autodesk 3ds Max Version 1.0 to 25.0 (Tested on 9.0 to 2023)

Support English, Chinese, and Japanese localized data
