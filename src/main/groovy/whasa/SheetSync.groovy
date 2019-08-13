package whasa

import com.xlson.groovycsv.CsvParser
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.filesystem.NotOLE2FileException
import org.apache.poi.ss.usermodel.CellType

public class SheetSync {

	static List<ProductOnHand> PRODUCTS_ON_HAND = []

	public static void main(String[] args) {
		if(args.size() != 2) {
			usage()
			System.exit(1)
		}
		args.each { arg ->
			File f = new File(arg)
			if(!f.isFile()) {
				System.err.println("Input file [${arg}] is missing or invalid.")
				System.exit(2)
			}

			println "Loading products from file [${arg}]"
			try {
				println "Attempting to load as a spreadsheet"
				addXlsProducts(f)
				println "Finished loading products from file [${arg}]"
			} catch(NotOLE2FileException e1) {
				try {
					println "File was not a spreadsheet. Attempting to load as a csv"
					addCsvProducts(f)
					println "Finished loading products from file [${arg}]"
				} catch(Exception e2) {
					System.err.println('Unknown file format')
					e2.printStackTrace()
					System.exit(5)
				}
			}
			catch(Exception e3) {
				System.err.println('Unknown error')
				e3.printStackTrace()
				System.exit(6)
			}
		}

		println "Generating output spreadsheet"
		try
		{
			File outputFile = new File('./SyncOutput.xls')
			FileOutputStream outputStream = new FileOutputStream(outputFile);
			SyncOutputGenerator.generateWorkbook(PRODUCTS_ON_HAND).write(outputStream);
			println "Output spreadsheet [${outputFile.getCanonicalPath()}] generated"
		}
		catch(Exception e)
		{
			System.err.println("Failed to write out results to XLS file");
			e.printStackTrace();
			System.exit(4);
		}
	}

	private static void addCsvProducts(File f) {
		f.withReader {reader ->
			Iterator iterator = CsvParser.parseCsv(reader)
			while(iterator.hasNext()) {
				def line = iterator.next()
				String sku = line."SKU"
				String product = line."Product/Service Name"
				int quantity = (line."Quantity On Hand").toInteger()
				ProductOnHand poh = new ProductOnHand(SKU: sku, PRODUCT: product, QUANTITY: quantity)
				int index = PRODUCTS_ON_HAND.indexOf(poh)
				if(index > -1) {
					if(PRODUCTS_ON_HAND[index].QUANTITY > poh.QUANTITY) {
						println "Replacing [${PRODUCTS_ON_HAND[index].toString()}] with [${poh.toString()}]"
						PRODUCTS_ON_HAND[index] = poh
					}
				}
				else {
					println "Adding [${poh.toString()}]"
					PRODUCTS_ON_HAND.add(poh)
				}
			}
		}
	}

	private static void addXlsProducts(File f) {
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
			System.err.println("Input file [${f.toString()}] did not contain the required columns")
			System.exit(3)
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
			int index = PRODUCTS_ON_HAND.indexOf(poh)
			if(index > -1) {
				if(PRODUCTS_ON_HAND[index].QUANTITY > poh.QUANTITY) {
					println "Replacing [${PRODUCTS_ON_HAND[index].toString()}] with [${poh.toString()}]"
					PRODUCTS_ON_HAND[index] = poh
				}
			}
			else {
				println "Ading [${poh.toString()}]"
				PRODUCTS_ON_HAND.add(poh)
			}
		}
	}

	public static void usage() {
		System.err.println("""\
			Usage:
			java -jar whasa.SheetSync.jar <spreadsheet1> <spreadsheet2>
			Output:
			./sheetSync.xls - The synced and merged input files in XLS format""".stripIndent())
	}
}