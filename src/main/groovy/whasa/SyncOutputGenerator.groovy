package whasa

import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Font

class SyncOutputGenerator {
    private static HSSFWorkbook WORKBOOK = new HSSFWorkbook()
    private static HSSFSheet WORKSHEET = WORKBOOK.createSheet()
    private static HSSFFont BOLD_UNDERLINE_FONT = WORKBOOK.createFont();
    private static HSSFCellStyle HEADER_STYLE = WORKBOOK.createCellStyle();

    public static HSSFWorkbook generateWorkbook(List<ProductOnHand> products) {
        BOLD_UNDERLINE_FONT.setBold(true)
        BOLD_UNDERLINE_FONT.setUnderline(Font.U_SINGLE)
        HEADER_STYLE.setFont(BOLD_UNDERLINE_FONT)
        HSSFRow firstRow = WORKSHEET.createRow(0)
        for(int i = 0; i < 3; i++) {
            firstRow.createCell(i)
            firstRow.getCell(i).setCellStyle(HEADER_STYLE)
        }

        WORKSHEET.getRow(0).getCell(0).setCellValue('Product/Service Name')
        WORKSHEET.getRow(0).getCell(1).setCellValue('SKU')
        WORKSHEET.getRow(0).getCell(2).setCellValue('Quantity On Hand')

        for(int i = 0; i < products.size(); i++) {
            ProductOnHand poh = products.get(i)
            int r = i+1
            WORKSHEET.createRow(r)
            for(int j = 0; j < 3; j++) {
                WORKSHEET.getRow(r).createCell(j)
            }
            WORKSHEET.getRow(r).getCell(0).setCellValue(poh.PRODUCT)
            WORKSHEET.getRow(r).getCell(1).setCellValue(poh.SKU)
            WORKSHEET.getRow(r).getCell(2).setCellValue(poh.QUANTITY)
        }

        for (int i = 0; i < 3; i++)
        {
            WORKSHEET.autoSizeColumn(i);
        }

        return WORKBOOK
    }
}
