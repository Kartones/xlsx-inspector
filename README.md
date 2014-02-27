XLSX Inspector
==============

Current version: 0.1

# Purpose #

Being able to quickly detect in Microsoft Excel XLSX files are too big for operating with them
(for example with gems like `Roo`).

Many solutions scan the whole file before letting you do anything with it. This small class extracts the first worksheet
only, scans part of "header" and obtaining the `<dimension>` tag and its info, estimates the numer of total cells inside.

Opens and scans a file in less than 5 seconds in the same machine that needs several minutes to open the same file with
Roo and just get the number of columns and rows.

*This is not meant for operating with the XLSX*, its sole purpose is to estimate the number of cells.

# Usage #

Check `sample1.rb` in the examples folder, but for the lazy ones, it is damm simple:

``` ruby
require_relative '../lib/xlsx_inspector.rb'

myTest = XLSXInspector::Inspector.new()
puts myTest.inspect('200krows.xlsx')
```

It either returns `nil` or an integer of the total estimated cells. Estimation is done by multiplying # columns x # rows.

# Roadmap #

* documentation (here)
* tests
* error handling
* allow to choose which sheet to scan, or "scan all"