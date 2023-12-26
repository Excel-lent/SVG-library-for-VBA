# SVG library for VBA

<div style="text-align: right"><em>Excel is as inexhaustible as the atom, nature is infinite</em></div>
<div style="text-align: right">Lenin, Materialism and Empirio-Criticism (1908).</div>
<br>
This is a preleminary release of SVG library for VBA. Not everything is working properly. Not all properties are implemented.
The aim of the library is to export charts in SVG format, i. e. in lossless image format. <strong>Pull requests are welcome!</strong>

<br>

## Distribution
The library is distributed
- [Add-in for Excel](https://github.com/Excel-lent/SVG-library-for-VBA/releases/download/v0.0.4-alpha/SVGlib.xlam).
- [Installer for add-in](https://github.com/Excel-lent/SVG-library-for-VBA/releases/download/v0.0.4-alpha/SVGlib.Installer.xlsm). "SVGlib.xlam" should be placed to the same directory as "SVGlib installer.xlsm". Simply follow instructions and install the add-in on your computer.
- [Example of usage](https://github.com/Excel-lent/SVG-library-for-VBA/releases/download/v0.0.4-alpha/SVGlib.Example.xlsx). After installation you will get a new ribbon with a single button. Use it to export all graphs shown on the page. The graphs will be saved in your working directory.
![ExportSvg ribbon with a single button](./Images/Installed%20addin.png "ExportSvg ribbon with a single button")
- [Development table](https://github.com/Excel-lent/SVG-library-for-VBA/releases/download/v0.0.4-alpha/SVGlib.xlsm). To create add-in you have to save it as "xlam" file. Be careful, Excel tries to save it to add-in's directory! Use button "back" to return to the working directory.

## Benefits

1. Smaller size of the picture, all benefits of vector graphics (sharp, scallable images):

| <center>JPG (39 kB)</center> | <center>SVG (7 kB)</center> |
|--------------------|------------------------------------|
| <img src='./Images/Picture1.jpg' width='500'> | <img src='./Images/Picture1.svg' width='500'> |

2. The changes in the graphs can be tracked. For example, if the formula of $$5^3$$ (cell "C6") was changed to $$5^3 + 1,$$ the difference in the graph will immediately show the changes:
<img src='./Images/Git changes in the graph.png' width='500'>

3. Smaller commit size. 

## Useful references
[Creating An Add-in From An Excel Macro](https://jkp-ads.com/articles/distributemacro01.asp)
