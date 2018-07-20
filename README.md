ExcelColormap
=======

Excel C++ dll plugin able to create colormaps from an array of scalar data.

[![Examples](examples/pixel.png)](examples/gaussian.png)

What exactly?
=============
This is a simple dll library, cross-platform (it does not use any MS-only code, plain C++),
built to be called from Excel VBA editor. 
It returns a color based on the cell value, allowing you to set the cell background and directly create colormaps from Excel.

* MATLAB Jet **colorscheme**
* Creation of **any size you want** colormaps
* Two styles as of now **Excel colormap** and **Gaussian colormap** (thanks to OpenCV)
* It relies on OpenCV (cross-platform visual library) to generate the Gaussian colormap
* Quick drop-in-your-project : only two files (.cpp and .h)
* Save the Gaussian colormap to a path you set
* Excel 64bit and 32bit supported


Installing
----------

It relies on OpenCV dll (opencv_worldxxx.dll, latest version), so you need the dll in the SYSTEM Path or you can drop it inside the plugin dll execution folder, which appears to be User/CurrentUser/Documents (at least for me).

Add the source files to your project and **be sure to build the dll with the same bitness as your Excel installation**.
Simply put : Excel x64 -> plugin x64, same for x86.


You are ready to go and use the library! :)


How to use
----------

We need to add some code in the Excel VBA panel in order to call the plugin.

Read OpenCV documentation about Gaussian blur : https://docs.opencv.org/3.4/d4/d86/group__imgproc__filter.html#gaabe8c836e97159a9193fb0b11ac52cf1

```
Private Type CellInfo
 row As Long
 col As Long
 value As Single
End Type

Private Type Settings
 minValue As Single
 maxValue As Single
 width As Single
 height As Single
 rowsCount As Long
 columnsCount As Long
 kernelX As Long
 kernelY As Long
 sigmaX As Single
 sigmaY As Single
 path As LongPtr
End Type

#If VBA7 Then
Private Declare PtrSafe Function InterpolateSingleValue Lib "path_to_plugin_dll" (ByVal value As Single, ByVal min As Single, ByVal max As Single) As Long
Private Declare PtrSafe Function Interpolate Lib "path_to_plugin_dll" (ByRef x As CellInfo, ByRef v As Settings) As Long
#Else
Private Declare Function InterpolateSingleValue Lib "path_to_plugin_dll" (ByVal value As Single, ByVal min As Single, ByVal max As Single) As Long
Private Declare Function Interpolate Lib "path_to_plugin_dll" (ByRef x As CellInfo, ByRef v As Settings) As Long
#End If

Sub CreateImage()

    'we are selecting a range of values from the spreadsheet
    Dim rng As Range, cell As Range
    Set rng = Range("A1:H20")
    
    'we set up the needed settings for the plugin dll
    Dim v As Settings
    v.minValue = 30 'lowest value
    v.maxValue = 70 'highest value
    v.width = 500 'Gaussian image width
    v.height = 500 'Gaussian image height
    v.rowsCount = rng.rows.Count
    v.columnsCount = rng.Columns.Count
    v.kernelX = 71 'Gaussian smoothing parameter - must be odd!
    v.kernelY = 71 'Gaussian smoothing parameter - must be odd!
    v.sigmaX = 35 'Gaussian smoothing parameter
    v.sigmaY = 0 'Gaussian smoothing parameter
    v.path = StrPtr("D://test.png") 'Optional parameter to save the image to disk
    
    
    For Each cell In rng
        Dim cell_ As CellInfo
        cell_.col = cell.Column
        cell_.row = cell.row
        cell_.value = cell.value
        
        'we are just setting cells background near the data source
        Cells(cell.row, cell.Column + 9).Interior.color = Interpolate(cell_, v)
        'Cells(cell.row, cell.Column + 9).Interior.color = InterpolateSingleValue(cell.value, 30, 70)
    Next cell

End Sub
```

Now by executing CreateImage you will get a result similar to the first images, depending on the parameters you chose and dataset.
