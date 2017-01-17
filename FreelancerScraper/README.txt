
For every new type of file you need to scrape, you should follow this pattern:


1) Create a new data object in ScraperModel.Models (See ScraperModel.Models.GovUkSupplyUseCrudeGasFeedstocks for an example)
	a) This needs to inherit from ScraperModel.ScrapeTable
	b) You need to apply attributes following the "code first" entity framework convention. A unique key must be applied across fields in your object
	c) You need to add a DbSet property to ScraperModel.ScraperDbContext

* NOTE YOU MAY NEED MORE THAN ONE TYPE OF OBJECT FOR THE SAME FILE. FOR EXAMPLE, AN EXCEL FILE CAN HAVE MULTIPLE SHEETS WITH DIFFERENT DATA
* THE NAME OF THE OBJECT SHOULD BE DESCRIPTIVE OF THE SOURCE AND CONTENT
* IF YOU DO REUSE AN OBJECT YOU MUST ADD A "SOURCE" FIELD TO SAY WHICH FILE THE DATA CAME FROM


2) Create a new Scrape for each file (See ScraperFreelancerLib.GovUkSupplyUseCrudeGasFeedstocksScraper for an example)
	a) The URL the file can be used against should be a constant in the file (if more than one URL can be scraped with the code have a seperate const for each one)
	b) Inherit from HtmlScrapeData, ExcelXlsxScrapeData (Office 2007+), ExcelXlsScrapeData (Office 2003 and before) or PdfScrapeData depending on the type of file being scraped. This loads the relevant object for you.
	c) Scrape from the data and save anything using the dbContext.SaveData(list, lastModified) call (again you may need to call this multiple times for different data in different parts of the target document)


3) Run the scrape (see FreelancerScraper.Program for an example)
	a) You can run this either by getting the remote file on the fly (using WebGrabData) or from a locally saved file (using LocalFileGrabData). The later is recommended while developing.
	b) Scrapes should fail gracefully and if they find a data error, alert on which file, sheet, row (as appropriate) the issue is found on by outputting to the Console


The end results should be a library of scrapes in ScraperFreelancerLib that can handle each Excel file and PDF found in the below sources, and a set of objects in ScraperModel (with DbSet entries in ScraperDbContext).

NOTE 1 - In many instances there are identical files for different time periods, so you can reuse the same scrape.
NOTE 2 - Sometimes there are zip files that contain multiple spreadsheets\pdf files which also need to be scraped.
NOTE 3 - Sometimes there are PDF and Excel files representing the same data. In these cases you only need to scrape one of them (probably the spreadsheet as it is easy to do)

https://www.gov.uk/government/statistical-data-sets/crude-oil-and-petroleum-products-imports-by-product-1920-to-2011

https://www.gov.uk/government/statistical-data-sets/crude-oil-and-petroleum-production-imports-and-exports-1890-to-2011

https://www.gov.uk/government/statistics/petroleum-chapter-3-digest-of-united-kingdom-energy-statistics-dukes

http://www.cores.es/en/estadisticas

http://dgsaie.mise.gov.it/dgerm/bollettino_nuovo/indice.asp?anno=2016

http://dgsaie.mise.gov.it/dgerm/bollettino_nuovo/indice.asp?anno=2015

http://dgsaie.mise.gov.it/dgerm/bollettino_nuovo/indice.asp?anno=2014

https://ssb.no/energi-og-industri/statistikker/petroleumsalg/maaned/2016-11-15?fane=tabell#content

https://ssb.no/energi-og-industri/statistikker/petroleumsalg/aar/2016-04-05?fane=tabell#content

https://www.eia.gov/outlooks/steo/

https://www.eia.gov/analysis/projection-data.cfm

http://www.anp.gov.br/wwwanp/dados-estatisticos

http://www.capp.ca/publications-and-statistics/crude-oil-forecast

https://www.aer.ca/data-and-publications/statistical-reports/st3
