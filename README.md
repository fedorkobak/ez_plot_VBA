# ez_plot_VBA
Macro for momental plot building.

## build_special_plot SUB
 
It waints string with address of range for building plot. It builds simple line plot with markers and titles for every point. 
Tiles have only 2 digits after point. Last five markers have some extra size. It will looks line that.

<img src="https://github.com/Dranikf/ez_plot_VBA/blob/main/example.JPG">

## ez_plot_full SUB

Finds where the row of data ends to the left of the highlighted cell and use ```build_special_plot``` for  plot it.

## ez_plot_ten SUB

Gets 10 more cells after selected and use ```build_special_plot``` for building some plot.

## file ez_plot_old.bas 
uses old funcions for creating the same thigns. It contains funcitons with the same names but with postscript "old"