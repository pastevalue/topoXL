# *TopoXL* Coding Standards
This document stores information about all coding standards used in the development process of the *TopoXL* project.
These standards should be applied to any code which becomes part of this project.
Any changes of the standards should be reflected in the actual code.
## Commenting
All text comments start with a space after the comment symbol
```
''' This is text comment
'' This is text comment
' This is text comment
```
Non-text comments start imeediatly after the comment symbol
```
''===
'---
```
## License
The License text:
* is included in all *.cls* and *.bas* files at their very top line
* is marked by three comment symbols (''')
* is followed by an empty line

License comment sample

```
''' <Project description>
''' <Copyright Notes>
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program.  If not, see <https://www.gnu.org/licenses/>.

```
### Class/Module comments
The class/module comment:
* is at the top of the file, after the license section
* is marked as a separate section using equal (=) symbol (total of 75 characters)
* its subsections are separated by dash (-) symbol (total of 75 characters)
* should not exceed a 75 characters long string

Class/Module comment sample
```
''========================================================================
'' Subsection 1
'' Info here
'' ...
''------------------------------------------------------------------------
'' Subsection 2
'' Info here
'' ...
''------------------------------------------------------------------------
'' ...
''========================================================================
```