# Skjema - Bestilling laboratorieanalyser og dataforvaltning

*Skjema - Bestilling laboratorieanalyser og dataforvaltning* is an Excel template for joint ordering of chemical analyzes from NIVAlab and Data management from Miljøinformatikk.

The user documentation for the order form is found within the sheet itself and at [Kilden](http://doc-o1.niva.no/symfoni/infoportal/publikasjon.nsf/redirectIntra?ReadForm&Url=http://doc-o1.niva.no/symfoni/infoportal/publikasjon.nsf/.vieShowWeb/51AE723016C35DA4C1257EB90045AFA4?OpenDocument "Infrastruktur > Miljøinformatikk > Dataforvaltning > Veiledning bestilling")


### Order form home in NIVApro (TQM)

The approved versions of the document must be uploaded to NIVApro, where the NIVA end users can download it. Due to a bug in TQM, the version number and approved/revised dates must be set manually in top text of excel sheet, before uploading to TQM. In TQM, field "Read version" must be set to "Original". Se email "Fwd Ny melding fra 4humanTQM.no - Send oss en melding.pdf" for information from 4human TQM.

How to upload to TQM:

1.	Set new document version, revision data and approved date in top text of the updated Excel orderform worksbook
2.	Search for document id 17060 in nivapro/TQM
3.	Open menu “Revisions” under settings icon for the document
4.	Press “Add revision”
5.	Fill in comment for changes, and press “Add revision”
6.	Press “Save”
7.	In document settings, make sure “Read version” is set to “Original”
8.	Open setting icon again, and choose “Replace document”, and replace it
9.	Open setting icon again, and choose «Klar til godkjenning». Select “Klar til godkjenning», and optionally write a comment . After saving, an e-mail will be sent to the approver.


### Source code and version control

The source code is imported / exported to the two folders "Class Modules" and "Modules" with the macros UpdateFromGithub / PrepareForGithub. By doing this it will be easier to track the changes.

If you get an error "Programmatic Access To Visual Basic Project Is Not Trusted", you should try this:

https://stackoverflow.com/questions/25638344/programmatic-access-to-visual-basic-project-is-not-trusted


### protection of workbook and worksheets

The workbook is protected, to avoid users from changing and adding fields, and writing comments outside of existing fields. The password for unprotecting is encrypted. It is also possible to protect and unprotect by running VBA procedures protect_wb_and_ws and unprotect_wb_and_ws respectively. The former should always be run after changes in the workbook.