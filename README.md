# CAFT: Copy As Formatted Table

**CAFT** is a Excel Add-in that copies the selected cells as a table with multiple formats.
It support below formats.

* Markdown(https://daringfireball.net/projects/markdown/)
* Trac(https://trac.edgewall.org/)
* PukiWiki(https://pukiwiki.osdn.jp/)
* XPlanner-plus(https://ja.osdn.net/projects/sfnet_xplanner-plus/)

## Installation/Uninstallation

Download master and Execute `Install.vbs`; `Install.vbs` is install script.  
Of Course, you can install manually, too.

And, if you want to uninstall this Add-In, Please execute `Uninstall.vbs`.

## Usage

1. Select Cells
2. Right Click -> `Copy as ...` -> Select Format,  
   then formatted table is copied to the clipboard.
3. Paste it wherever.

### Example: Markdown

```
| Header | Header | Header |
|:------:|:-------|-------:|
| Text   | Text   | Text   |
| Text   | Text   | Text   |
| Text   | Text   | Text   |
```


### Example: Trac

```
|| **Head** || **Head** ||
|| Text     || Text     ||
|| Text     || Text     ||
|| Text     || Text     ||
```

### Example: Pukiwiki

```
|~Header| Header| Header|
| Text  | Text  | Text  |
| Text  | Text  | Text  |
| Text  | Text  | Text  |
```

### Example: XPlanner-plus

```
| *Header* | *Header* | *Header* |
| Text     | Text     | Text     |
| Text     | Text     | Text     |
| Text     | Text     | Text     |
```

## Licence
MIT
