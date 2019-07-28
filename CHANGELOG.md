#Changelog
All notable changes (Added features,depracated features,removed features,changed functionalities and bug fixes) will be documented in this file.

##[Unreleased]
## [0.11.10] 2019-07-28
##Added
-Chart support for 4 types of charts (Area,Vertical-Var,Horizontal-Bar,Line).

##Fixed
-Small bugs related to building charts (Fixed correct step size when iterating through data columns).
-Edge case when first found was false after exit of unparsed data loop.

## [0.11.8] 2019-07-25
##Added
-Preference option to always view multiple signals on multiple worksheets.
-Chart support for line chart only.Remaining charts on next version.

##Changed
-Removed redundant method of copying template file to current directory and then editing.
-Preference layout strcuture for better layout managment.
-Removed completely option for consumption data for MV values.
-Code design involcing creating table. Now seperate functions for single sheet and multiple sheet layout.

##Fixed
-Template not being used after the first worksheet in a workbook of multiplke signals.

## [0.10.7] 2019-07-24
##Added
-View template path chosen,clear path option.
-Preferences menu option using pickle module to remember preferences.

##Changed
-Updated setup.py for msi installer build.

## [0.8.7] 2019-07-23
##Added
-Timezone selection feature. Will automatically convert and check database and convert back while printing report.
-Feature for multiple signals raw data presented in seperate worksheets

##Fixed
-Boundary condition when o further data is available but next_dt is less than tdate

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

