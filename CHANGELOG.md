#Changelog
All notable changes (Added features,depracated features,removed features,changed functionalities and bug fixes) will be documented in this file.

##[Unreleased]

## [0.1.1] 2019-07-10
## Fixed
-Logic on seggegation of different signal types.
## Changed
-Replaced the long method of using regular expressions after creating a dumnp file to locate the table associated with a certain object uid with just using the last two digits of the uid.

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

