# Skjema - Bestilling laboratorieanalyser og dataforvaltning

*Skjema - Bestilling laboratorieanalyser og dataforvaltning* is an Excel template for joint ordering of chemical analyzes from NIVAlab and Data management from Miljøinformatikk.

The user documentation for the order form is found within the sheet itself and at [Kilden](http://doc-o1.niva.no/symfoni/infoportal/publikasjon.nsf/redirectIntra?ReadForm&Url=http://doc-o1.niva.no/symfoni/infoportal/publikasjon.nsf/.vieShowWeb/51AE723016C35DA4C1257EB90045AFA4?OpenDocument "Infrastruktur > Miljøinformatikk > Dataforvaltning > Veiledning bestilling")



### Order form home in NIVApro (TQM)

The approved versions of the document must be uploaded to [NIVApro](https://tqm2.tqmenterprise.no/NIVA/Publishing/Document/LoadLocalContent/17060?forOL1=niva) where the NIVA end users can download it. Due to a bug in TQM,the version number and approved/revised dates must be set manually in top text of excel sheet before uploading to TQM. In TQM, field "*Leseversjon*" must be set to "*Original*". Se email "Fwd Ny melding fra 4humanTQM.no - Send oss en melding.pdf" for information from *4human TQM*.



### Source code and version control

The source code is imported / exported to the two folders "Class Modules" and "Modules" with the macros UpdateFromGithub / PrepareForGithub. By doing this it will be easier to track the changes.

If you get an error "Programmatic Access To Visual Basic Project Is Not Trusted", you should try this:

https://stackoverflow.com/questions/25638344/programmatic-access-to-visual-basic-project-is-not-trusted
