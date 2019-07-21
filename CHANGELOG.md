#Changelog
All notable changes (Added features,depracated features,removed features,changed functionalities and bug fixes) will be documented in this file.

##[Unreleased]
## [0.6.6] 2019-07-21
##Added
-Feature to select a predefined excel template.

##Fixed
-Minor bugs in extract function

## [0.5.5] 2019-07-21
##Fixed
-Minor bug fixes in sqlscript in cases when table not found and single digit values tables.

## [0.5.4] 2019-07-21
##Added
-Checkbox feature.
-Listbox to view selected signals.
-Time selection.
-Suppports multiple signals.
-Support for multiple signal excel layout 

##Changed
-Now use faster method to locate the table for MV signals (using the last two digits of UID).
-Tree population changed from lazy to eager to support checkbox feature properly.
-For each catgeory sperate tree object so each node no longer stores a variable for measurement.
-Added default value to update on focusout for frequency entry if left empty.
-Extraction algorithm now cleaner,removed redundant if-else clauses and logic to support to support multiple signals.
-Switched from xlsxwriter -> openpyxl to support future feature of adding to templates.

##Fixed
-Logic on categorization of signals.
-Typ5 criteria for metering signals.
-Consumption report available only for metering signals.
-Minor bug fix for avg_aux list out of index error in ParseData

## [0.1.1] 2019-07-10
## Changed
-Replaced the long method of using regular expressions after creating a dumnp file to locate the table associated with a certain object uid with just using the last two digits of the uid.
## Fixed
-Logic on seggegation of different signal types.

## [0.1.0] 2019-07-10
### Added
-This CHANGELOG file
-Application in base stage, supports base features such as 
----Segreggation between Electrical and System signals
----Seggregation between events and mv signals
----Extracting the latest available data according to user defined frequency interval
----Providing an option to dipslay average and consumption during each interval

### Changed
-Renamed main2.1 to main

##Depracated
-Removing use of module xlsxwriter and swiitching to openpyxl to support editing predefined excel templates.

