# emis-exporter

Tuvalu and similar Databases are supported on master (UIS 2016 Extract)
Kiribati-2017 branch supports KEMIS db and UIS 2017 Extract

Changes were required to the BaseSQL format due to limits in disaggregating student data in KEMIS.  
Updates to the Tuvalu Base SQL is required before merging the changes back to master



# Install Instructions

> git clone https://github.com/JeremyKells/emis-exporter.git

> cd emis-exporter

> start EmisExporter.sln

Open the `EmisExporter.exe.config` file, and enter your DB Credentials.

Build, Run.
