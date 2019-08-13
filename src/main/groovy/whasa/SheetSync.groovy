package whasa

import com.xlson.groovycsv.CsvParser
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellType

public class SheetSync {

	static List<ProductOnHand> PRODUCTS_ON_HAND = []
	static List<ProductOnHand> PRODUCTS_CSV = []
	static List<ProductOnHand> PRODUCTS_XLS = []

	static String INPUT_CSV = 'BC_output.csv'
	static String INPUT_XLS = 'QB_ProductServiceList.xls'


	static void main(String[] args) {
		File csvFile, xlsFile
		if(args.size() == 0) {
			println "Loading default input files"
			csvFile = new File(INPUT_CSV)
			xlsFile = new File(INPUT_XLS)
		}
		else if(args.size() == 2) {
			println "Loading custom input files ${args}"
			csvFile = new File(args[0])
			xlsFile = new File(args[1])
		}
		else {
			usage()
			System.exit(1)
		}
		[csvFile, xlsFile].each {
			if(!it.isFile()) {
				System.err.println("Input file [${it.toString()}] is missing or invalid.")
				System.exit(2)
			}
		}

		println "Loading products from file [${csvFile.toString()}]"
		PRODUCTS_CSV = loadCsvProducts(csvFile)

		println "Loading products from file [${xlsFile.toString()}]"
		PRODUCTS_XLS = loadXlsProducts(xlsFile)

		println "Syncing sheets"
		syncProducts()

		println "Generating output spreadsheet"
		try
		{
			File outputFile = new File('./SyncOutput.xls')
			FileOutputStream outputStream = new FileOutputStream(outputFile);
			SyncOutputGenerator.generateWorkbook(PRODUCTS_ON_HAND).write(outputStream);
			println "\tOutput spreadsheet [${outputFile.getCanonicalPath()}] generated"
		}
		catch(Exception e)
		{
			System.err.println("\tFailed to write out results to XLS file");
			e.printStackTrace();
			System.exit(4);
		}

		println "Checking for missing SKUs"
		for(ProductOnHand poh : PRODUCTS_XLS) {
			if(!PRODUCTS_CSV.contains(poh)) {
				System.err.println("\tProduct [${poh.SKU}] is missing from input [${csvFile.toString()}]")
			}
		}
		for(ProductOnHand poh : PRODUCTS_CSV) {
			if(!PRODUCTS_XLS.contains(poh)) {
				System.err.println("\tProduct [${poh.SKU}] is missing from input [${xlsFile.toString()}]")
			}
		}
	}

	private static def loadCsvProducts(File f) {
		def ret = []
		f.withReader { reader ->
			Iterator iterator = CsvParser.parseCsv(reader)
			while(iterator.hasNext()) {
				def line = iterator.next()
				String sku = line."SKU"
				String product = line."Product/Service Name"
				int quantity = (line."Quantity On Hand").toInteger()
				ProductOnHand poh = new ProductOnHand(SKU: sku, PRODUCT: product, QUANTITY: quantity)
				int index = ret.indexOf(poh)
				if(index > -1) {
					println "\tFile [${f.toString()}] contains both [${ret[index]}] and [${poh}]. Using the lesser value."
					if(ret[index].QUANTITY > poh.QUANTITY) {
						ret[index] = poh
					}
				}
				else {
					println "\tLoaded product [${poh.toString()}] from file [${f.toString()}]"
					ret.add(poh)
				}
			}
		}
		return ret
	}

	private static def loadXlsProducts(File f) {
		def ret =[]
		FileInputStream inputStream = new FileInputStream(f);
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		HSSFSheet worksheet = workbook.getSheetAt(0);

		int skuIndex = -1
		int productIndex = -1
		int quantityIndex = -1

		HSSFRow firstRow = worksheet.getRow(0)
		for(int i = firstRow.getFirstCellNum(); i < firstRow.getLastCellNum(); i++) {
			switch(firstRow.getCell(i).getStringCellValue().toLowerCase()) {
				case 'sku':
					skuIndex = i
					break
				case 'product/service name':
					productIndex = i
					break
				case 'quantity on hand':
					quantityIndex = i
					break
			}
		}
		if(skuIndex == -1 || productIndex == -1 || quantityIndex == -1) {
			throw new Exception("Input file [${f.toString()}] did not contain the required columns")
		}

		for(int i = 1; true; i++) {
			HSSFRow row = worksheet.getRow(i)
			if(row == null) break
			row.getCell(skuIndex)?.setCellType(CellType.STRING)
			row.getCell(productIndex)?.setCellType(CellType.STRING)
			row.getCell(quantityIndex)?.setCellType(CellType.STRING)
			String sku = row.getCell(skuIndex)?:''.toString()
			String product = row.getCell(productIndex)?:''.toString()
			int quantity = (row.getCell(quantityIndex)?:'0').toString().toInteger()

			ProductOnHand poh = new ProductOnHand(SKU: sku, PRODUCT: product, QUANTITY: quantity)
			int index = ret.indexOf(poh)
			if(index > -1) {
				println "\tFile [${f.toString()}] contains both [${ret[index]}] and [${poh}]. Using the lesser value."
				if(ret[index].QUANTITY > poh.QUANTITY) {
					ret[index] = poh
				}
			}
			else {
				println "\tLoaded product [${poh.toString()}] from file [${f.toString()}]"
				ret.add(poh)
			}
		}
		return ret
	}

	private static void syncProducts() {
		PRODUCTS_XLS.each { poh ->
			int index = PRODUCTS_ON_HAND.indexOf(poh)
			if(index > -1) {
				if(PRODUCTS_ON_HAND[index].QUANTITY > poh.QUANTITY) {
					println "\tReplacing [${PRODUCTS_ON_HAND[index].toString()}] with [${poh.toString()}]"
					PRODUCTS_ON_HAND[index] = poh
				}
			}
			else {
				println "\tAdding [${poh.toString()}]"
				PRODUCTS_ON_HAND.add(poh)
			}
		}

		PRODUCTS_CSV.each { poh ->
			int index = PRODUCTS_ON_HAND.indexOf(poh)
			if(index > -1) {
				if(PRODUCTS_ON_HAND[index].QUANTITY > poh.QUANTITY) {
					println "\tReplacing [${PRODUCTS_ON_HAND[index].toString()}] with [${poh.toString()}]"
					PRODUCTS_ON_HAND[index].QUANTITY = poh.QUANTITY
				}
			}
			else {
				println "\tAdding [${poh.toString()}]"
				PRODUCTS_ON_HAND.add(poh)
			}
		}
	}

	private static void usage() {
		System.err.println("""\
			Usage 1:
			java -jar SheetSync.jar
				Assumes BC_Output.csv and QBProdcutServiceList.xls in current folder
			Usage 2:
			java -jar SheetSync.jar <CSV File> <XLS File>
				CSV file must be first argument
			Output:
			./sheetSync.xls - The synced and merged input files in XLS format""".stripIndent())
	}
}